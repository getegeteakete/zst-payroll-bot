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
async def serve_file(filename: str, name: str = None):
    """ローカルフォールバック用配信。本番ではSupabase Storage推奨
    ?name= クエリで日本語表示名を指定可能（ダウンロード時のファイル名）
    """
    if "/" in filename or ".." in filename:
        raise HTTPException(400, "invalid filename")
    p = PUBLIC_FILES_DIR / filename
    if not p.exists():
        raise HTTPException(404)
    # 表示名（ダウンロード時のファイル名）。指定なければ実ファイル名
    display_name = name or filename
    return FileResponse(p, filename=display_name)


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
        
        # ヘルプ
        if text in ("ヘルプ", "help", "使い方", "メニュー"):
            reply_message(reply_token, [text_message(_help_text())])
            return
        
        # ウィザード進行中の入力かチェック
        from app.wizard import (
            is_in_session, handle_wizard_input, handle_evaluation,
            start_wizard, cancel_wizard
        )
        from app.history_search import handle_history_search
        
        # 1. ウィザード入力中
        if is_in_session(user_id):
            wizard_response = handle_wizard_input(user_id, text)
            if wizard_response:
                reply_message(reply_token, [text_message(wizard_response)])
                return
            # 計算後の追加コマンド（金額/比較）
            eval_response = handle_evaluation(user_id, text)
            if eval_response:
                reply_message(reply_token, [text_message(eval_response)])
                return
        
        # 2. ウィザード開始トリガー
        if text in ("見積", "見積もり", "見積り", "原価", "原価計算", "計算"):
            reply_message(reply_token, [text_message(start_wizard(user_id))])
            return
        
        # 3. キャンセル
        if text in ("キャンセル", "終了", "中止"):
            reply_message(reply_token, [text_message(cancel_wizard(user_id))])
            return
        
        # 4. 履歴検索
        if text.startswith("履歴"):
            reply_message(reply_token, [text_message(handle_history_search(text))])
            return
        
        # その他
        reply_message(reply_token, [text_message(
            "📋 使い方:\n"
            "・配車CSV / Excel送信 → 給与明細表生成\n"
            "・「見積」 → 原価計算ウィザード\n"
            "・「履歴 [地名]」 → 過去実績検索\n"
            "・「ヘルプ」 → 詳細"
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
        "📋 ZST業務支援Bot 使い方\n"
        "━━━━━━━━━━━━━━━\n\n"
        "【🚛 給与計算】\n"
        "▶ CSV送信\n"
        "  配車CSVを送ると明細表自動生成\n\n"
        "▶ Excel送信\n"
        "  ポイント明細表に運行行程入力した\n"
        "  Excelを送ると追加ポイント自動補完\n\n"
        "【💰 原価計算（営業向け）】\n"
        "▶「見積」と送信\n"
        "  ウィザードが7つの質問\n"
        "  → 最低受注額・推奨単価を自動算出\n\n"
        "▶ 計算後の追加コマンド\n"
        "・「金額 70000」採算判定\n"
        "・「比較 60000 70000 80000」複数比較\n\n"
        "【📚 過去実績検索】\n"
        "▶「履歴 [地名]」と送信\n"
        "  例: 「履歴 春日井市」\n"
        "      「履歴 大阪 名古屋」\n\n"
        "【🚪 終了】\n"
        "「キャンセル」「終了」"
    )
