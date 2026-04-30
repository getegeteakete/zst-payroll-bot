"""
Supabase Storageでファイルホスト
LINEから直接ダウンロードできる署名付きURLを発行
"""
import os
import requests
from pathlib import Path
from datetime import datetime
import secrets
import urllib.parse


def upload_file_to_supabase(local_path: Path, bucket: str = "payroll", expires_in: int = 86400) -> str:
    """
    Supabase Storageにアップロード → 署名付きURL返却
    expires_in: URL有効期限（秒）デフォルト24時間
    
    SUPABASE_URL, SUPABASE_SERVICE_ROLE環境変数が必要
    """
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_ROLE")

    if not url or not key:
        # フォールバック: ローカル一時保存（開発用）
        return _local_fallback_url(local_path)

    # オブジェクトパス（衝突回避でランダムプレフィックス）
    obj_name = f"{datetime.now().strftime('%Y%m%d')}/{secrets.token_hex(8)}/{local_path.name}"

    # アップロード
    upload_url = f"{url}/storage/v1/object/{bucket}/{obj_name}"
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": _content_type(local_path),
    }
    with open(local_path, "rb") as f:
        r = requests.post(upload_url, headers=headers, data=f.read(), timeout=60)
    r.raise_for_status()

    # 署名付きURL発行
    sign_url = f"{url}/storage/v1/object/sign/{bucket}/{obj_name}"
    sign_headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    r = requests.post(sign_url, headers=sign_headers, json={"expiresIn": expires_in}, timeout=15)
    r.raise_for_status()
    signed_path = r.json()["signedURL"]
    return f"{url}/storage/v1{signed_path}"


def _content_type(path: Path) -> str:
    suffix = path.suffix.lower()
    return {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".pdf": "application/pdf",
        ".csv": "text/csv",
    }.get(suffix, "application/octet-stream")


def _local_fallback_url(local_path: Path) -> str:
    """
    開発用フォールバック
    ファイル名はASCII safe（token+拡張子）にして、
    ダウンロード時の日本語名は ?name= クエリで指定する
    """
    base = os.environ.get("PUBLIC_BASE_URL", "http://localhost:8000")
    public_dir = Path("/tmp/public_files")
    public_dir.mkdir(exist_ok=True)
    
    # ASCII safe な保存ファイル名（token + 拡張子のみ）
    token = secrets.token_urlsafe(12)
    suffix = local_path.suffix.lower()
    safe_name = f"{token}{suffix}"
    
    target = public_dir / safe_name
    target.write_bytes(local_path.read_bytes())
    
    # 表示用ファイル名はURLエンコードしてクエリパラメータに
    # スペースは _ に置換しておく（LINEのリンク認識対策）
    display_name = local_path.name.replace(" ", "_").replace("　", "_")
    encoded_display = urllib.parse.quote(display_name)
    
    return f"{base}/files/{safe_name}?name={encoded_display}"
