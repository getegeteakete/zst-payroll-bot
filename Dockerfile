FROM python:3.12-slim

# LibreOffice + 日本語フォントをインストール（PDF変換用）
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-core \
    fonts-noto-cjk \
    fonts-ipafont \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /code

# 依存関係
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリ本体
COPY app/ ./app/
COPY templates/ ./templates/

# PORT環境変数（Render/Railway/Cloud Run対応）
ENV PORT=8000
EXPOSE 8000

CMD uvicorn app.main:app --host 0.0.0.0 --port ${PORT}
