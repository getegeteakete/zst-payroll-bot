"""
過去実績ルート検索
キーワード（市区町村名）から該当ルートを検索
"""
import json
from pathlib import Path

_HISTORY_DATA = None


def _load():
    global _HISTORY_DATA
    if _HISTORY_DATA is None:
        path = Path(__file__).parent / "route_history.json"
        if path.exists():
            with open(path, encoding="utf-8") as f:
                _HISTORY_DATA = json.load(f)
        else:
            _HISTORY_DATA = {}
    return _HISTORY_DATA


def search_routes(keywords: list[str], limit: int = 10) -> list[dict]:
    """
    キーワードで過去ルート検索
    全キーワードがルート名に含まれるものを返す
    """
    data = _load()
    matches = []
    for route, stats in data.items():
        if all(kw in route for kw in keywords):
            matches.append({
                "route": route,
                **stats,
            })
    # 出現回数で降順
    matches.sort(key=lambda x: -x["n"])
    return matches[:limit]


def format_search_results(keywords: list[str], results: list[dict]) -> str:
    """検索結果を整形"""
    if not results:
        return (
            f"📚 「{' '.join(keywords)}」の検索結果\n\n"
            f"該当する過去実績がありません。\n"
            f"別のキーワードで検索してみてください。\n\n"
            f"例: 「履歴 大阪」「履歴 春日井」"
        )
    
    msg = [f"📚 「{' '.join(keywords)}」検索結果（{len(results)}件）\n"]
    
    for i, r in enumerate(results[:10], 1):
        emoji = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"][i-1]
        msg.append(f"\n{emoji} {r['route']}")
        msg.append(f"  {r['n']}回運行 / {r['mode_pt']:,}pt（最頻）")
        if r['min_pt'] != r['max_pt']:
            msg.append(f"  範囲: {r['min_pt']:,}〜{r['max_pt']:,}pt")
        msg.append(f"  主な担当: {', '.join(r['top_drivers'][:3])}")
        msg.append(f"  期間: {r['first_seen']} 〜 {r['last_seen']}")
    
    if len(results) > 10:
        msg.append(f"\n（他{len(results)-10}件あり）")
    
    return "\n".join(msg)


def handle_history_search(text: str) -> str:
    """「履歴 春日井市 此花区」のようなコマンドを処理"""
    parts = text.replace("履歴", "").strip().split()
    if not parts:
        return (
            "📚 履歴検索\n\n"
            "「履歴 [地名]」で過去実績を検索できます\n\n"
            "例:\n"
            "・「履歴 大阪」\n"
            "・「履歴 春日井市」\n"
            "・「履歴 加須市 西淀川区」"
        )
    
    results = search_routes(parts, limit=10)
    return format_search_results(parts, results)
