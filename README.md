# ZST 給与計算 LINE Bot

配車CSVを公式LINEに送信すると、AIが自動でポイント明細表（Excel）とPDFを生成して返すLINE Bot。

## 流れ

```
ドライバー / 事務員
  ↓ 配車CSVをLINEに送信
LINE Bot (このアプリ)
  ↓ 計算 → Excel生成 → PDF変換
  ↓ Supabase Storage にアップ
  ↓ 署名付きURLをLINEに返信
ユーザーがダウンロード
```

## 構成

```
line_payroll_bot/
├── app/
│   ├── main.py                    # FastAPI（LINE webhook）
│   ├── line_handler.py            # LINE API呼び出し
│   ├── payroll_engine.py          # ポイント計算ロジック
│   ├── payroll_orchestrator.py    # CSV→Excel→PDFオーケストレーション
│   ├── storage.py                 # Supabase Storage連携
│   └── learned_route_dict.json    # 過去実績から学習したルート辞書
├── templates/
│   └── blank_template.xlsx        # ポイント明細表のひな形
├── Dockerfile
├── render.yaml                    # Render用設定
├── requirements.txt
└── .env.example
```

## ローカル起動

```bash
# 1. 環境変数
cp .env.example .env
# LINE_CHANNEL_SECRET, LINE_CHANNEL_ACCESS_TOKEN を入力

# 2. Python仮想環境
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 3. LibreOfficeをインストール（PDF変換用）
# Mac:    brew install --cask libreoffice
# Ubuntu: sudo apt install libreoffice

# 4. 起動
uvicorn app.main:app --reload --port 8000

# 5. ngrokで外部公開（LINE webhookに設定するURL）
ngrok http 8000
```

LINE Developers コンソールで Webhook URL を `https://xxx.ngrok.app/webhook` に設定。

## 本番デプロイ

### Render（推奨）

1. GitHubリポジトリにpush
2. https://dashboard.render.com → New → Web Service
3. リポジトリ選択 → Docker検出 → Deploy
4. Environment Variables に必要なキーを設定：
   - `LINE_CHANNEL_SECRET`
   - `LINE_CHANNEL_ACCESS_TOKEN`
   - `SUPABASE_URL` / `SUPABASE_SERVICE_ROLE`（推奨）
5. デプロイ完了後 `https://xxx.onrender.com/webhook` をLINE側に登録

### Railway

```bash
npm i -g @railway/cli
railway login
railway init
railway up
```

### Google Cloud Run

```bash
gcloud run deploy zst-payroll-bot \
  --source . \
  --region asia-northeast1 \
  --allow-unauthenticated \
  --memory 1Gi
```

## LINE Developers 設定

1. https://developers.line.biz/console/ でMessaging APIチャネル作成
2. **Channel secret** をコピー → `LINE_CHANNEL_SECRET`
3. **Channel access token** を発行 → `LINE_CHANNEL_ACCESS_TOKEN`
4. **Webhook URL** にデプロイ先のURLを設定 + Webhook利用ON
5. **応答メッセージ** OFF / **あいさつメッセージ** はお好みで

## Supabase Storage 設定

1. Supabaseで `payroll` バケットを作成（Private推奨）
2. Service Role Keyを取得 → `SUPABASE_SERVICE_ROLE`
3. ファイルは署名付きURLで配信（24時間有効）

## CSVフォーマット

| 日付 | 便 | 運行行程 | 荷姿 | 乗務員 |
|---|---|---|---|---|
| 2025-06-01 | ① | 東大阪市〜小牧市 | P/L | 岡田悠嗣 |
| 2025-06-01 | ② | 小牧市〜尼崎市 | P/L | 岡田悠嗣 |
| 2025-06-02 | ① | 休み | | 岡田悠嗣 |

- カラム名は柔軟に解釈（「日付」「Date」「年月日」など）
- エンコーディングはUTF-8 / Shift_JIS どちらも自動判定
- 日付形式: `2025-06-01` / `2025/6/1` / `2025年6月1日` / Excelシリアル値

## 計算ロジック

- **基本ポイント**: 9,000pt/日（運行行程がある日）
- **追加ポイント**: 過去実績辞書 → 都道府県カテゴリ判定
  - 関西〜関西: 2,000〜4,500（市区町村ペアで決定）
  - 中距離: 6,500〜10,000
  - 長距離: 12,500〜15,500
- **荷姿手当**: バラ全数5,000 / バラ半分・P/L半分 2,500 / その他 0
- **出勤加算**: 月所定労働日数を超えたら +5,000pt × (超過日数)
- **配車貢献**: 月所定を満たせば 10,000、満たさなければ 0
- **愛車・無事故**: 経営者が月次で手入力（デフォルト 0）

## 月別所定労働日数（2025年6月〜2026年5月）

| 6月 | 7月 | 8月 | 9月 | 10月 | 11月 | 12月 | 1月 | 2月 | 3月 | 4月 | 5月 |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 21 | 22 | 21 | 22 | 22 | 20 | 22 | 22 | 20 | 21 | 22 | 22 |

## トラブルシューティング

### LibreOfficeでPDF変換に失敗する
- 日本語フォントが入っていない可能性。Dockerfileでは `fonts-noto-cjk` を導入済み

### LINE で「処理中...」のあと応答がない
- Render Free Plan はスリープ復帰に時間がかかる。Starter ($7/月) 以上推奨
- LibreOffice起動が初回30秒前後。タイムアウトに注意

### #VALUE! エラー
- テンプレ複製時の元シートに空白由来のエラーがある場合あり。`Q53/S53` を 0 で初期化済み

## ライセンス

(C) 2026 エーライフ / ZST社向け
