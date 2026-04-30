"""
営業向けウィザード式 原価計算
LINE上で対話的に必要情報を集めて原価計算

状態管理: プロセス内辞書（簡易版）
- TTL 1時間で自動削除
- Renderの再起動で消えるが、ウィザードは10分以内に終わる前提なので問題なし
"""
import time
from typing import Optional
from app.costing_engine import calc_costing, evaluate_offer, compare_offers, ROUTE_PRESETS


# セッション保管庫
_sessions: dict[str, dict] = {}
SESSION_TTL_SECONDS = 3600  # 1時間


def _cleanup_expired():
    """期限切れセッションを削除"""
    now = time.time()
    expired = [k for k, v in _sessions.items() if now - v["updated_at"] > SESSION_TTL_SECONDS]
    for k in expired:
        del _sessions[k]


def get_session(user_id: str) -> Optional[dict]:
    _cleanup_expired()
    return _sessions.get(user_id)


def set_session(user_id: str, session: dict):
    session["updated_at"] = time.time()
    _sessions[user_id] = session


def clear_session(user_id: str):
    _sessions.pop(user_id, None)


# ========================================================
# ウィザードのステップ定義
# ========================================================
STEPS = [
    {
        "key": "vehicle_class",
        "question": (
            "📋 原価計算ウィザード Step 1/7\n\n"
            "車格を選んでください\n"
            "① 2t\n"
            "② 4t\n"
            "③ 大型\n\n"
            "※「キャンセル」で中止"
        ),
        "options": {"1": "2t", "2": "4t", "3": "大型",
                    "2t": "2t", "4t": "4t", "大型": "大型"},
    },
    {
        "key": "route_type",
        "question": (
            "Step 2/7\n\n"
            "運行種別は？\n"
            "① 長距離（1日700km級）\n"
            "② 中距離（1日400km級）\n"
            "③ 地場（1日250km級）"
        ),
        "options": {"1": "長距離", "2": "中距離", "3": "地場",
                    "長距離": "長距離", "中距離": "中距離", "地場": "地場"},
    },
    {
        "key": "is_leased",
        "question": (
            "Step 3/7\n\n"
            "リース車両ですか？\n"
            "① はい\n"
            "② いいえ"
        ),
        "options": {"1": True, "2": False, "はい": True, "いいえ": False,
                    "yes": True, "no": False},
    },
    {
        "key": "monthly_days",
        "question_dynamic": True,  # 直前の入力に基づきデフォルト値を提示
        "validator": "positive_int",
    },
    {
        "key": "distance_per_day",
        "question_dynamic": True,
        "validator": "positive_int",
    },
    {
        "key": "toll_per_day",
        "question_dynamic": True,
        "validator": "non_negative_int",
    },
    {
        "key": "fuel_price",
        "question_dynamic": True,
        "validator": "positive_int",
    },
]


def _build_dynamic_question(step_index: int, data: dict) -> str:
    """直前までの入力に基づき、デフォルト値付きの質問を組み立てる"""
    route_type = data.get("route_type", "中距離")
    is_leased = data.get("is_leased", False)
    preset = ROUTE_PRESETS.get((route_type, is_leased), ROUTE_PRESETS[("地場", False)])
    
    key = STEPS[step_index]["key"]
    
    if key == "monthly_days":
        return (
            f"Step 4/7\n\n"
            f"月稼働日数は？\n"
            f"（数字のみ入力 / 目安: {preset['default_days']}日）\n"
            f"例: 22"
        )
    elif key == "distance_per_day":
        return (
            f"Step 5/7\n\n"
            f"1日の平均走行距離は？(km)\n"
            f"（数字のみ / 目安: {preset['default_distance_per_day']}km）\n"
            f"例: 300"
        )
    elif key == "toll_per_day":
        return (
            f"Step 6/7\n\n"
            f"1日の通行料（高速代）は？(円)\n"
            f"（数字のみ / 目安: {preset['default_toll_per_day']:,}円）\n"
            f"※高速使わなければ 0"
        )
    elif key == "fuel_price":
        return (
            f"Step 7/7\n\n"
            f"燃料単価は？(円/L)\n"
            f"（数字のみ / 直近相場: {preset['default_fuel_price']}円/L）"
        )
    return "（入力をお願いします）"


def _validate_input(value: str, validator: str) -> Optional[float]:
    """数値バリデーション"""
    try:
        # カンマ・全角を取り除く
        cleaned = value.replace(",", "").replace(",", "").translate(
            str.maketrans("0123456789", "0123456789")
        )
        n = float(cleaned)
        if validator == "positive_int" and n <= 0:
            return None
        if validator == "non_negative_int" and n < 0:
            return None
        return n
    except ValueError:
        return None


# ========================================================
# メイン処理
# ========================================================
def start_wizard(user_id: str) -> str:
    """ウィザード開始"""
    set_session(user_id, {
        "mode": "wizard",
        "step": 0,
        "data": {},
    })
    return STEPS[0]["question"]


def cancel_wizard(user_id: str) -> str:
    clear_session(user_id)
    return "📋 ウィザードを終了しました。\n再開するには「見積」と送ってください。"


def handle_wizard_input(user_id: str, text: str) -> Optional[str]:
    """
    ウィザード中の入力を処理
    Returns: 返信メッセージ（Noneならウィザード対象外）
    """
    session = get_session(user_id)
    if not session or session.get("mode") != "wizard":
        return None
    
    text = text.strip()
    
    # キャンセル
    if text in ("キャンセル", "中止", "終了", "やめる", "cancel"):
        return cancel_wizard(user_id)
    
    step_index = session["step"]
    if step_index >= len(STEPS):
        return None
    
    step = STEPS[step_index]
    key = step["key"]
    
    # 入力値の解釈
    if "options" in step:
        # 選択肢
        value = step["options"].get(text)
        if value is None:
            return f"⚠ 認識できません: {text}\n\n" + step["question"]
    else:
        # 数値
        value = _validate_input(text, step.get("validator", "positive_int"))
        if value is None:
            return f"⚠ 数字で入力してください: {text}"
    
    # 値を保存
    session["data"][key] = value
    session["step"] = step_index + 1
    set_session(user_id, session)
    
    # 次のステップ
    next_index = step_index + 1
    if next_index >= len(STEPS):
        # 全ステップ完了 → 計算実行
        return _finalize_wizard(user_id)
    
    next_step = STEPS[next_index]
    if next_step.get("question_dynamic"):
        return _build_dynamic_question(next_index, session["data"])
    else:
        return next_step["question"]


def _finalize_wizard(user_id: str) -> str:
    """全入力完了後の原価計算"""
    session = get_session(user_id)
    data = session["data"]
    
    result = calc_costing(
        vehicle_class=data["vehicle_class"],
        route_type=data["route_type"],
        is_leased=data["is_leased"],
        monthly_days=int(data["monthly_days"]),
        distance_per_day=float(data["distance_per_day"]),
        toll_per_day=float(data["toll_per_day"]),
        fuel_price=float(data["fuel_price"]),
    )
    
    # セッション更新（後続の「金額」「比較」コマンドのために結果を保持）
    session["mode"] = "results"
    session["costing"] = result
    set_session(user_id, session)
    
    # メッセージ生成
    lease_str = "リースあり" if data["is_leased"] else "リース無し"
    msg = [
        "✅ 計算完了\n",
        f"📋 条件",
        f"・車格: {data['vehicle_class']} ({data['route_type']}・{lease_str})",
        f"・月稼働: {int(data['monthly_days'])}日",
        f"・走行距離: {int(data['distance_per_day'])}km/日",
        f"・通行料: {int(data['toll_per_day']):,}円/日",
        f"・燃料単価: {int(data['fuel_price'])}円/L",
        "",
        "💰 1日あたりコスト",
    ]
    for k, v in result["daily"].items():
        msg.append(f"・{k}: {v:>10,.0f}円")
    msg.append(f"━━━━━━━━━━━")
    msg.append(f"📊 最低原価: {result['min_daily_cost']:,.0f}円/日")
    msg.append(f"📊 最低受注単価: {result['min_daily_revenue']:,.0f}円/日 ⚠")
    msg.append("")
    msg.append("✅ 推奨単価")
    msg.append(f"・利益10%確保: {result['recommended_10']:,.0f}円/日")
    msg.append(f"・利益20%確保: {result['recommended_20']:,.0f}円/日")
    msg.append("")
    msg.append("💡 月額イメージ")
    msg.append(f"・最低: {result['min_daily_revenue']*int(data['monthly_days']):,.0f}円/月")
    msg.append(f"・利益10%: {result['recommended_10']*int(data['monthly_days']):,.0f}円/月")
    msg.append("")
    msg.append("━━━━━━━━━━━")
    msg.append("次にできること:")
    msg.append("「金額 70000」→この単価での利益判定")
    msg.append("「比較 60000 70000 80000」→複数比較")
    msg.append("「履歴 春日井市」→実績検索")
    msg.append("「終了」→セッション終了")
    
    return "\n".join(msg)


def handle_evaluation(user_id: str, text: str) -> Optional[str]:
    """「金額 70000」「比較 60000 70000」コマンド処理"""
    session = get_session(user_id)
    if not session or "costing" not in session:
        return None
    
    text = text.strip()
    
    if text.startswith("金額") or text.startswith("単価"):
        # 単一単価の判定
        parts = text.replace("金額", "").replace("単価", "").strip().split()
        if not parts:
            return "⚠ 例: 「金額 70000」"
        try:
            price = float(parts[0].replace(",", ""))
        except ValueError:
            return "⚠ 数字で入力してください: 「金額 70000」"
        
        eval_result = evaluate_offer(session["costing"], price)
        days = session["costing"]["input"]["monthly_days"]
        
        return (
            f"📈 採算判定: {price:,.0f}円/日\n\n"
            f"・1日コスト: {eval_result['daily_cost']:,.0f}円\n"
            f"・1日収入: {price:,.0f}円\n"
            f"・1日利益: {eval_result['daily_profit']:+,.0f}円\n"
            f"・利益率: {eval_result['profit_rate']*100:.1f}%\n"
            f"・月{days}日換算: {eval_result['monthly_profit']:+,.0f}円\n\n"
            f"判定: {eval_result['judgment']}"
        )
    
    if text.startswith("比較"):
        # 複数単価の比較
        parts = text.replace("比較", "").strip().split()
        try:
            prices = [float(p.replace(",", "")) for p in parts]
        except ValueError:
            return "⚠ 例: 「比較 60000 70000 80000」"
        
        if not prices:
            return "⚠ 例: 「比較 60000 70000 80000」"
        
        results = compare_offers(session["costing"], prices)
        days = session["costing"]["input"]["monthly_days"]
        
        msg = [f"📊 単価比較 (月{days}日換算)\n"]
        msg.append(f"{'単価':>9} {'月利益':>11} {'利益率':>7}")
        msg.append("-" * 31)
        for r in results:
            msg.append(
                f"{r['daily_revenue']:>9,.0f} "
                f"{r['monthly_profit']:>+11,.0f} "
                f"{r['profit_rate']*100:>6.1f}%"
            )
        msg.append("")
        msg.append("判定:")
        for r in results:
            msg.append(f"・{r['daily_revenue']:>7,.0f}円: {r['judgment']}")
        return "\n".join(msg)
    
    return None


def is_in_session(user_id: str) -> bool:
    """ユーザーがセッション中か"""
    return get_session(user_id) is not None


def session_help_text() -> str:
    return (
        "📋 原価計算機能\n\n"
        "「見積」と送ると計算ウィザード開始。\n"
        "7つの質問に答えるだけで:\n"
        "・最低受注金額\n"
        "・推奨単価\n"
        "・月額利益見込み\n"
        "を自動算出します。\n\n"
        "ウィザード後の追加コマンド:\n"
        "・「金額 70000」→ 採算判定\n"
        "・「比較 60000 70000 80000」→ 複数単価比較\n"
        "・「履歴 地名」→ 過去実績検索"
    )
