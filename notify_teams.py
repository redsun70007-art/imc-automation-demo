"""
IMC 일정 알림 — Teams DM (Adaptive Card)

매주 월요일 09:00 KST 실행.
- data.json에서 (BRAND, CHANNEL)별 마지막 END_DATE 추출
- 이미 종료 / 7일 이내 종료 예정 항목 분류
- Teams Workflow Webhook으로 Adaptive Card 발송

dry-run: DRY_RUN=1 환경변수 주면 발송 없이 출력만
"""

import os
import json
import sys
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

KST = timezone(timedelta(hours=9))
ALERT_WINDOW_DAYS = 7
ROOT = Path(__file__).parent


def load_data():
    with open(ROOT / "data.json", encoding="utf-8") as f:
        return json.load(f)


def last_end_dates(items):
    last = {}
    for item in items:
        brand = item.get("BRAND")
        channel = item.get("CHANNEL")
        end = item.get("END_DATE")
        if not (brand and channel and end):
            continue
        key = (brand, channel)
        if key not in last or end > last[key]:
            last[key] = end
    return last


def parse_date(s):
    """다양한 형식 허용: 2026-05-10, 2026.5.10, 2026/5/10, 2026. 5. 10. 등"""
    if not isinstance(s, str):
        return None
    cleaned = s.replace(".", "-").replace("/", "-").replace(" ", "").rstrip("-")
    parts = cleaned.split("-")
    if len(parts) != 3:
        return None
    try:
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        return datetime(y, m, d).date()
    except ValueError:
        return None


def classify(end_str, today):
    end = parse_date(end_str)
    if end is None:
        return None, None
    delta = (end - today).days
    if delta < 0:
        return "expired", abs(delta)
    if delta <= ALERT_WINDOW_DAYS:
        return "expiring", delta
    return None, None


def build_adaptive_card(expired, expiring, today):
    body = [
        {
            "type": "TextBlock",
            "size": "Large",
            "weight": "Bolder",
            "text": "🔔 IMC 일정 등록 알림"
        },
        {
            "type": "TextBlock",
            "wrap": True,
            "isSubtle": True,
            "text": f"기준일: {today.strftime('%Y-%m-%d')} (KST)"
        },
        {
            "type": "TextBlock",
            "wrap": True,
            "text": "담당 채널 중 **일정 등록이 필요한 항목**이 있어 알려드립니다."
        }
    ]

    if expired:
        body.append({
            "type": "TextBlock",
            "weight": "Bolder",
            "color": "Attention",
            "text": "🚨 이미 종료된 채널 (재등록 필요)"
        })
        lines = []
        for brand, channel, days, end in expired:
            lines.append(f"• **[{brand}] {channel}** — {end} 종료  `({days}일 경과)`")
        body.append({
            "type": "TextBlock",
            "wrap": True,
            "text": "\n\n".join(lines)
        })

    if expiring:
        body.append({
            "type": "TextBlock",
            "weight": "Bolder",
            "color": "Warning",
            "text": f"⚠️ {ALERT_WINDOW_DAYS}일 이내 종료 예정"
        })
        lines = []
        for brand, channel, days, end in expiring:
            d_label = "D-Day" if days == 0 else f"D-{days}"
            lines.append(f"• **[{brand}] {channel}** — {end} 종료  `({d_label})`")
        body.append({
            "type": "TextBlock",
            "wrap": True,
            "text": "\n\n".join(lines)
        })

    body.append({
        "type": "TextBlock",
        "wrap": True,
        "isSubtle": True,
        "size": "Small",
        "text": "📅 매주 월요일 09:00 KST 자동 발송"
    })

    return {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": body
    }


def send_to_teams(card, webhook_url):
    payload = {
        "type": "message",
        "attachments": [{
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": card
        }]
    }
    req = urllib.request.Request(
        webhook_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"}
    )
    with urllib.request.urlopen(req) as res:
        return res.status


def main():
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    dry_run = os.environ.get("DRY_RUN") == "1"

    if not webhook_url and not dry_run:
        sys.exit("❌ TEAMS_WEBHOOK_URL 환경변수가 필요합니다. (또는 DRY_RUN=1)")

    items = load_data()
    today = datetime.now(KST).date()
    print(f"[info] today (KST) = {today}, dry_run = {dry_run}")

    last = last_end_dates(items)

    expired, expiring = [], []
    for (brand, channel), end_str in last.items():
        status, days = classify(end_str, today)
        if status == "expired":
            expired.append((brand, channel, days, end_str))
        elif status == "expiring":
            expiring.append((brand, channel, days, end_str))

    expired.sort(key=lambda x: -x[2])
    expiring.sort(key=lambda x: x[2])

    print(f"[info] expired={len(expired)}, expiring={len(expiring)}")

    if not expired and not expiring:
        print("✅ 알림 대상 없음. 발송 스킵.")
        return

    card = build_adaptive_card(expired, expiring, today)

    if dry_run:
        print("[dry-run] 발송 생략. Adaptive Card 미리보기:")
        print(json.dumps(card, ensure_ascii=False, indent=2))
        return

    status = send_to_teams(card, webhook_url)
    print(f"✅ Teams 발송 완료 (HTTP {status})")


if __name__ == "__main__":
    main()
