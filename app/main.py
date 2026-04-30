"""
LINE Bot メインアプリ（FastAPI）
- POST /webhook : LINE webhook
- GET /files/{name} : ローカルフォールバック用ファイル配信
- GET /health : ヘルスチェック
"""
import os
import json
import traceback
from pathlib import Path
from fastapi import FastAPI, Request, BackgroundTasks, HTTPException, Header
from fastapi.responses import FileResponse, JSONResponse, PlainTextResponse

from app.line_handler import (
    verify_signature, download_message_content,
    reply_message, push_message,
    text_message, result_flex_message, error_flex_message,
)
from app.payroll_orchestrator import process_payroll
from app.storage import upload_file_to_supabase


app = FastAPI(title="ZST Payroll LINE Bot")

PUBLIC_FILES_DIR = Path("/tmp/public_files")
PUBLIC_FILES_DIR.mkdir(exist_ok=True)


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/files/{filename}")
async def serve_file(filename: str):
    """ローカルフォールバック用配信。本番ではSupabase Storage推奨"""
    if "/" in filename or ".." in filename:
        raise HTTPException(400, "invalid filename")
    p = PUBLIC_FILES_DIR / filename
    if not p.exists():
        raise HTTPException(404)
    return FileResponse(p, filename=filename)


@app.post("/webhook")
async def webhook(
    request: Request,
    background: BackgroundTasks,
    x_line_signature: str = Header(None, alias="x-line-signature"),
):
    body = await request.body()

    # 署名検証
    if not verify_signature(body, x_line_signature or ""):
        raise HTTPException(401, "invalid signature")

    payload = json.loads(body.decode("utf-8"))
    events = payload.get("events", [])

    for event in events:
        try:
            handle_event(event, background)
        except Exception as e:
            traceback.print_exc()
            # エラー時の応答試行
            reply_token = event.get("replyToken")
            if reply_token:
                try:
                    reply_message(reply_token, [error_flex_message(f"内部エラー: {e}")])
                except Exception:
                    pass

    return PlainTextResponse("OK")


def handle_event(event: dict, background: BackgroundTasks):
    event_type = event.get("type")
    if event_type != "message":
        return

    msg = event["message"]
    msg_type = msg.get("type")
    reply_token = event.get("replyToken")
    source = event.get("source", {})
    user_id = source.get("userId")

    if msg_type == "text":
        text = msg.get("text", "").strip()
        if text in ("ヘルプ", "help", "使い方"):
            reply_message(reply_token, [text_message(_help_text())])
            return
        reply_message(reply_token, [text_message(
            "配車CSVファイル（.csv または .txt）を送ってください。\n"
            "「ヘルプ」と送ると使い方が表示されます。"
        )])
        return

    if msg_type == "file":
        file_name = msg.get("fileName", "")
        file_size = msg.get("fileSize", 0)
        message_id = msg.get("id")

        if file_size > 5 * 1024 * 1024:
            reply_message(reply_token, [text_message("⚠ ファイルサイズは5MB以下にしてください")])
            return

        if not file_name.lower().endswith((".csv", ".txt")):
            reply_message(reply_token, [text_message(
                f"⚠ CSVファイル(.csv)を送ってください。受信: {file_name}"
            )])
            return

        # 即座にACK
        reply_message(reply_token, [text_message(
            f"📥 「{file_name}」を受信しました。\n計算中です…通常10〜30秒で完了します。"
        )])

        # バックグラウンド処理
        background.add_task(process_in_background, message_id, user_id, file_name)
        return

    # その他
    reply_message(reply_token, [text_message(
        "配車CSVファイルを送ってください。「ヘルプ」で使い方表示。"
    )])


def process_in_background(message_id: str, user_id: str, file_name: str):
    """重い処理をバックグラウンドで"""
    try:
        # CSVダウンロード
        content = download_message_content(message_id)

        # エンコーディング判定（BOM/UTF-8/Shift_JIS 自動）
        text = _decode_csv(content)

        # 給与処理
        result = process_payroll(text)

        # ストレージにアップロード
        xlsx_url = upload_file_to_supabase(result["xlsx_path"])
        pdf_url = upload_file_to_supabase(result["pdf_path"])

        # LINE Push（結果通知）
        push_message(user_id, [
            result_flex_message(
                driver_name=result["driver_name"],
                year_month=result["target_year_month"],
                summary=result["summary"],
                xlsx_url=xlsx_url,
                pdf_url=pdf_url,
                low_conf_count=result["low_conf_count"],
            )
        ])
    except ValueError as e:
        push_message(user_id, [error_flex_message(str(e))])
    except Exception as e:
        traceback.print_exc()
        push_message(user_id, [error_flex_message(f"処理中にエラー: {e}")])


def _decode_csv(b: bytes) -> str:
    """CSVをデコード"""
    for enc in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSVのエンコーディングを認識できません（UTF-8またはShift_JISで保存してください）")


def _help_text() -> str:
    return (
        "📋 ZST給与計算Bot 使い方\n\n"
        "1. 配車CSVを LINE に送信\n"
        "2. 自動でポイント明細表（Excel）と PDF を生成\n"
        "3. ダウンロードボタンから保存\n\n"
        "【CSVのカラム】\n"
        "・日付（必須）例: 2025-06-01\n"
        "・便（任意）①/②/③\n"
        "・運行行程（必須）例: 東大阪市～小牧市\n"
        "・荷姿（任意）P/L、バラ、かご台車\n"
        "・乗務員（任意）\n\n"
        "「ヘルプ」と送るとこのメッセージが表示されます。"
    )
