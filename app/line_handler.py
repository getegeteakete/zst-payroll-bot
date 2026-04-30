"""
LINE Messaging API ハンドラー
- 署名検証
- メッセージコンテンツ取得（CSV添付ファイルDL）
- 返信・プッシュメッセージ
"""
import os
import hmac
import hashlib
import base64
import json
import requests
from typing import Optional


LINE_API = "https://api.line.me/v2/bot"
LINE_DATA_API = "https://api-data.line.me/v2/bot"


def get_channel_secret() -> str:
    val = os.environ.get("LINE_CHANNEL_SECRET")
    if not val:
        raise RuntimeError("LINE_CHANNEL_SECRET 未設定")
    return val


def get_access_token() -> str:
    val = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
    if not val:
        raise RuntimeError("LINE_CHANNEL_ACCESS_TOKEN 未設定")
    return val


def verify_signature(body: bytes, signature: str) -> bool:
    """LINE Webhook署名検証"""
    secret = get_channel_secret()
    h = hmac.new(secret.encode("utf-8"), body, hashlib.sha256).digest()
    expected = base64.b64encode(h).decode("utf-8")
    return hmac.compare_digest(expected, signature or "")


def download_message_content(message_id: str) -> bytes:
    """LINE Message Content APIで添付ファイルをDL"""
    url = f"{LINE_DATA_API}/message/{message_id}/content"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.content


def reply_message(reply_token: str, messages: list):
    """応答メッセージ送信（30秒以内）"""
    url = f"{LINE_API}/message/reply"
    headers = {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json",
    }
    payload = {"replyToken": reply_token, "messages": messages}
    r = requests.post(url, headers=headers, json=payload, timeout=10)
    if r.status_code >= 400:
        print(f"[LINE reply error] {r.status_code} {r.text}")
    return r


def push_message(to_user_id: str, messages: list):
    """プッシュメッセージ送信（処理完了後の通知）"""
    url = f"{LINE_API}/message/push"
    headers = {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json",
    }
    payload = {"to": to_user_id, "messages": messages}
    r = requests.post(url, headers=headers, json=payload, timeout=10)
    if r.status_code >= 400:
        print(f"[LINE push error] {r.status_code} {r.text}")
    return r


# ===== メッセージビルダー =====

def text_message(text: str) -> dict:
    return {"type": "text", "text": text[:5000]}


def result_flex_message(driver_name: str, year_month: tuple, summary: dict,
                        xlsx_url: str, pdf_url: str,
                        low_conf_count: int = 0) -> dict:
    """処理結果のFlex Message"""
    y, m = year_month
    title = f"{driver_name} {y}年{m}月"

    summary_rows = []
    for label, val in summary.items():
        if isinstance(val, (int, float)):
            val_str = f"{val:,}"
        else:
            val_str = str(val)
        summary_rows.append({
            "type": "box", "layout": "horizontal",
            "contents": [
                {"type": "text", "text": label, "size": "sm", "color": "#666666", "flex": 5},
                {"type": "text", "text": val_str, "size": "sm", "weight": "bold", "align": "end", "flex": 4},
            ],
            "margin": "sm"
        })

    warning_section = []
    if low_conf_count > 0:
        warning_section = [{
            "type": "box", "layout": "vertical",
            "backgroundColor": "#FFF8E1",
            "paddingAll": "12px",
            "cornerRadius": "8px",
            "contents": [
                {"type": "text", "text": f"⚠ {low_conf_count}件は推定不可。Excelで手動確認ください",
                 "size": "xs", "color": "#B8860B", "wrap": True}
            ],
            "margin": "md"
        }]

    return {
        "type": "flex",
        "altText": f"給与明細表が完成しました（{title}）",
        "contents": {
            "type": "bubble",
            "header": {
                "type": "box", "layout": "vertical",
                "backgroundColor": "#0A1F44", "paddingAll": "16px",
                "contents": [
                    {"type": "text", "text": "✓ 計算完了",
                     "color": "#FFD200", "size": "xs", "weight": "bold"},
                    {"type": "text", "text": title,
                     "color": "#FFFFFF", "size": "lg", "weight": "bold", "wrap": True},
                ]
            },
            "body": {
                "type": "box", "layout": "vertical",
                "contents": summary_rows + warning_section,
                "spacing": "sm"
            },
            "footer": {
                "type": "box", "layout": "vertical", "spacing": "sm",
                "contents": [
                    {
                        "type": "button", "style": "primary", "color": "#06C755",
                        "action": {"type": "uri", "label": "📊 Excel をダウンロード", "uri": xlsx_url}
                    },
                    {
                        "type": "button", "style": "secondary",
                        "action": {"type": "uri", "label": "📄 PDF をダウンロード", "uri": pdf_url}
                    },
                ]
            }
        }
    }


def error_flex_message(error: str) -> dict:
    """エラー通知"""
    return {
        "type": "flex",
        "altText": "エラーが発生しました",
        "contents": {
            "type": "bubble",
            "body": {
                "type": "box", "layout": "vertical",
                "contents": [
                    {"type": "text", "text": "✕ 処理失敗", "color": "#C04A1B", "weight": "bold"},
                    {"type": "text", "text": error, "size": "sm", "wrap": True, "margin": "md"},
                    {"type": "text",
                     "text": "CSVのカラムは「日付・便・運行行程・荷姿」を含めてください",
                     "size": "xs", "color": "#888", "wrap": True, "margin": "md"}
                ]
            }
        }
    }
