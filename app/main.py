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

        if file_size > 10 * 1024 * 1024:
            reply_message(reply_token, [text_message("⚠ ファイルサイズは10MB以下にしてください")])
            return

        fname_lower = file_name.lower()
        is_csv = fname_lower.endswith((".csv", ".txt"))
        is_xlsx = fname_lower.endswith(".xlsx")

        if not (is_csv or is_xlsx):
            reply_message(reply_token, [text_message(
                f"⚠ CSV(.csv) または Excel(.xlsx) を送ってください。\n受信: {file_name}"
            )])
            return

        # 即座にACK
        if is_xlsx:
            ack_msg = (
                f"📥 「{file_name}」を受信しました。\n"
                f"📘 Excel直接処理モード\n"
                f"  ・空欄の追加ポイントを自動補完\n"
                f"  ・既存の入力値はそのまま尊重\n"
                f"処理中…1〜3分かかります。"
            )
        else:
            ack_msg = (
                f"📥 「{file_name}」を受信しました。\n"
                f"📄 CSV処理モード\n"
                f"計算中…通常10〜30秒で完了します。"
            )
        reply_message(reply_token, [text_message(ack_msg)])

        # バックグラウンド処理（拡張子で分岐）
        if is_xlsx:
            background.add_task(process_xlsx_in_background, message_id, user_id, file_name)
        else:
            background.add_task(process_in_background, message_id, user_id, file_name)
        return

    # その他
    reply_message(reply_token, [text_message(
        "配車CSVファイルを送ってください。「ヘルプ」で使い方表示。"
    )])


def process_in_background(message_id: str, user_id: str, file_name: str):
    """CSV処理をバックグラウンドで"""
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


def process_xlsx_in_background(message_id: str, user_id: str, file_name: str):
    """Excel直接処理：受信したxlsxを補完して返す"""
    import tempfile
    from pathlib import Path
    from app.xlsx_processor import process_payroll_xlsx
    
    try:
        # ファイルダウンロード
        content = download_message_content(message_id)
        
        # 一時保存
        tmpdir = Path(tempfile.mkdtemp())
        input_path = tmpdir / file_name
        input_path.write_bytes(content)
        
        # 処理
        result = process_payroll_xlsx(input_path)
        
        # アップロード
        xlsx_url = upload_file_to_supabase(result["completed_xlsx_path"])
        pdf_url = upload_file_to_supabase(result["completed_pdf_path"])
        
        # サマリ作成
        year, month = result["year_month"]
        total_pt = sum(r["summary"].get("ポイント合計") or 0 for r in result["results"])
        total_filled = sum(r["routes_filled"] for r in result["results"])
        total_unknown = sum(r["routes_unknown"] for r in result["results"])
        
        msg = (
            f"✅ Excel処理完了 {year}年{month}月\n"
            f"乗務員: {len(result['results'])}名\n"
            f"合計ポイント: {total_pt:,}pt\n"
            f"自動補完: {total_filled}件\n"
        )
        if total_unknown > 0:
            msg += f"⚠ 不明ルート: {total_unknown}件（要手動確認）\n"
        msg += f"\n📊 完成版Excel:\n{xlsx_url}\n\n"
        msg += f"📄 PDF:\n{pdf_url}"
        
        # 5名以下なら内訳も
        if len(result['results']) <= 5:
            msg += "\n\n【内訳】"
            for r in result["results"]:
                pt = r['summary'].get('ポイント合計', 0)
                msg += f"\n・{r['driver_name']}: {pt:,}pt ({r['work_days']}日稼働)"
        
        push_message(user_id, [text_message(msg)])
    except ValueError as e:
        push_message(user_id, [error_flex_message(str(e))])
    except Exception as e:
        traceback.print_exc()
        push_message(user_id, [error_flex_message(f"Excel処理中にエラー: {e}")])


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
        "▼ 方法1: CSV送信（自動生成）\n"
        "配車CSVを送ると、明細表（Excel/PDF）を自動生成\n\n"
        "【CSVのカラム】\n"
        "・日付（必須）例: 2024-08-01\n"
        "・乗務員（複数人時必須）\n"
        "・運行行程（必須）例: 東大阪市～小牧市\n"
        "・便（任意）①/②/③\n"
        "・荷姿（任意）P/L、バラ、かご台車\n\n"
        "▼ 方法2: Excel送信（補完モード）\n"
        "ポイント明細表のExcelに運行行程・荷姿を入力した状態で送ると\n"
        "・空欄の追加ポイントを自動補完\n"
        "・既存の入力値はそのまま尊重\n"
        "・基本ポイント・集計を自動計算\n"
        "→ 完成版Excel/PDFを返却\n\n"
        "「ヘルプ」と送るとこのメッセージが表示されます。"
    )
