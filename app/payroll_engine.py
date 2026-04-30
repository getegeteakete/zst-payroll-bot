"""
ZST社 給与ポイント自動計算エンジン
- 運行行程から地域カテゴリを判定
- 手当一覧のルールに従ってポイントを算出
- 既存テンプレートに値を流し込む
"""
from openpyxl import load_workbook
from datetime import datetime, timedelta
import re
import csv
import json
import os

# 学習辞書ロード（あれば）
LEARNED_DICT = {}
_dict_path = os.path.join(os.path.dirname(__file__), "learned_route_dict.json")
if os.path.exists(_dict_path):
    with open(_dict_path, encoding="utf-8") as f:
        LEARNED_DICT = json.load(f)

# ===== 都道府県カテゴリ（手当一覧 Sheet "手当項目" / "運行加算" より）=====
# 関西 = 関西〜関西の地場対象
PREF_KANSAI = {"大阪", "京都", "兵庫", "奈良", "和歌山", "滋賀"}

# 中距離6500
PREF_MID_6500 = {"三重", "岐阜", "愛知", "岡山"}
# 中距離8500
PREF_MID_8500 = {"鳥取", "石川", "香川", "徳島"}
# 中距離10000
PREF_MID_10000 = {"静岡", "長野", "広島", "島根", "富山", "愛媛", "高知"}
# 長距離12500
PREF_LONG_12500 = {"神奈川", "山梨", "山口"}
# 長距離14000
PREF_LONG_14000 = {"東京", "埼玉", "千葉"}
# 長距離15500
PREF_LONG_15500 = {"群馬", "栃木", "茨城"}

# 荷姿手当
PACKING_POINTS = {
    "バラ": 5000,
    "バラ全数": 5000,
    "P/L→バラ": 2500,  # P/L→バラ or バラ→P/L (半数)
    "バラ→P/L": 2500,
    "P/L": 0,
    "かご台車": 0,
    "リフト": 0,
}

# 市区町村 → 都道府県（主要都市のみ。本番では辞書を拡充）
CITY_TO_PREF = {
    # 大阪府
    "住之江区": "大阪", "西淀川区": "大阪", "東大阪市": "大阪",
    "岸和田市": "大阪", "泉大津市": "大阪", "和泉市": "大阪",
    "大東市": "大阪", "門真市": "大阪", "茨木市": "大阪",
    "高槻市": "大阪", "枚方市": "大阪", "豊中市": "大阪",
    "吹田市": "大阪", "守口市": "大阪", "堺市": "大阪",
    "貝塚市": "大阪", "八尾市": "大阪", "羽曳野市": "大阪",
    "藤井寺市": "大阪", "松原市": "大阪", "柏原市": "大阪",
    "港区": "大阪", "西区": "大阪", "北区": "大阪",
    "中央区": "大阪", "都島区": "大阪", "淀川区": "大阪",
    "西成区": "大阪", "天王寺区": "大阪", "阿倍野区": "大阪",
    "東淀川区": "大阪", "鶴見区": "大阪", "城東区": "大阪",
    "旭区": "大阪", "東成区": "大阪", "生野区": "大阪",
    # 兵庫県
    "尼崎市": "兵庫", "西宮市": "兵庫", "神戸市": "兵庫",
    "須磨区": "兵庫", "東灘区": "兵庫", "灘区": "兵庫",
    "兵庫区": "兵庫", "長田区": "兵庫", "北区": "兵庫",
    "三木市": "兵庫", "小野市": "兵庫", "姫路市": "兵庫",
    "丹波市": "兵庫", "西区": "兵庫",
    # 京都府
    "京都市": "京都", "八幡市": "京都", "久御山町": "京都",
    "宇治市": "京都",
    # 奈良県
    "天理市": "奈良", "奈良市": "奈良", "生駒市": "奈良",
    "桜井市": "奈良", "橿原市": "奈良",
    # 和歌山県
    "和歌山市": "和歌山",
    # 滋賀県
    "湖南市": "滋賀", "草津市": "滋賀", "大津市": "滋賀",
    # 三重県
    "四日市市": "三重", "津市": "三重",
    # 愛知県
    "名古屋市": "愛知", "守山区": "愛知", "西尾市": "愛知",
    "豊田市": "愛知", "安城市": "愛知", "小牧市": "愛知",
    "春日井市": "愛知", "丹羽郡": "愛知", "大口町": "愛知",
    "加茂郡": "愛知", "犬山市": "愛知", "瀬戸市": "愛知",
    # 岐阜県
    "岐阜市": "岐阜", "大垣市": "岐阜",
    # 静岡県
    "静岡市": "静岡", "浜松市": "静岡",
    # 関西空港
    "関西空港": "大阪",
}


def detect_pref(text):
    """文字列から都道府県を判定"""
    if not text:
        return None
    # まず直接都道府県名を探す
    for pref in (PREF_KANSAI | PREF_MID_6500 | PREF_MID_8500 | PREF_MID_10000
                 | PREF_LONG_12500 | PREF_LONG_14000 | PREF_LONG_15500):
        if pref in text:
            return pref
    # 市区町村名から推定
    for city, pref in CITY_TO_PREF.items():
        if city in text:
            return pref
    return None


def category_of_pref(pref):
    """都道府県 → カテゴリ・ポイント"""
    if pref in PREF_KANSAI: return ("関西", 0)  # 地場は別途距離判定
    if pref in PREF_MID_6500: return ("中距離6500", 6500)
    if pref in PREF_MID_8500: return ("中距離8500", 8500)
    if pref in PREF_MID_10000: return ("中距離10000", 10000)
    if pref in PREF_LONG_12500: return ("長距離12500", 12500)
    if pref in PREF_LONG_14000: return ("長距離14000", 14000)
    if pref in PREF_LONG_15500: return ("長距離15500", 15500)
    return (None, 0)


def calc_route_points(route_text, packing_text=""):
    """
    運行行程テキストから追加ポイントを算出
    優先順位: ①学習辞書(過去実績) ②都道府県カテゴリ判定 ③地場推定
    
    返り値: (推定ポイント, カテゴリ, 信頼度)
        信頼度: "high" = 過去実績または中・長距離確定
                "medium" = 関西地場(目安)
                "low" = 不明
    """
    if not route_text or "休み" in str(route_text) or "特別休暇" in str(route_text):
        return (0, "休み", "high")

    route_str = str(route_text).strip()

    # ★優先①: 学習辞書に完全一致する過去実績があれば使う
    if route_str in LEARNED_DICT:
        return (LEARNED_DICT[route_str], "過去実績", "high")

    # 「～」で出発地・到着地を分離
    parts = re.split(r"[～~〜]", route_str)
    if len(parts) < 2:
        # 単独地名のみ（例: "リフト作業"）
        return (None, "要手動確認", "low")

    pref_from = detect_pref(parts[0])
    pref_to = detect_pref(parts[-1])

    if pref_from is None and pref_to is None:
        return (None, "要手動確認", "low")

    # どちらか不明 → 判明している方を使う
    pref_either = pref_from or pref_to
    pref_other = pref_to if pref_from else pref_from

    # 両方関西 → 地場
    if pref_from in PREF_KANSAI and pref_to in PREF_KANSAI:
        # 距離判定は今回はデフォルト3000として実装
        return (3000, "関西地場(目安)", "medium")

    # 中・長距離判定（関西以外のいずれかから取る）
    target_pref = pref_other if pref_other and pref_other not in PREF_KANSAI else pref_either
    if target_pref in PREF_KANSAI:
        target_pref = pref_other  # 両方関西だったケースは上で処理済み

    cat, pts = category_of_pref(target_pref)
    if pts > 0:
        return (pts, cat, "high")

    return (None, "要手動確認", "low")


def calc_packing_bonus(packing_text):
    """荷姿手当を返す"""
    if not packing_text:
        return 0
    txt = str(packing_text)
    if "バラ" in txt and "P/L" not in txt and "→" not in txt:
        return 5000
    if "→" in txt or "・" in txt:
        return 2500
    return 0


# ===== セルマップ =====
# 1日3行構成。1〜16日が左ブロック、17〜末日が右ブロック
def cell_map_for_day(day_index_in_block, side):
    """
    day_index_in_block: 0始まり (左:0=1日, 1=2日... / 右:0=17日, 1=18日...)
    side: "left" or "right"
    """
    base_row = 4 + day_index_in_block * 3
    if side == "left":
        return {
            "date": ("A", base_row),
            "route": [("C", base_row + i) for i in range(3)],
            "packing": [("D", base_row + i) for i in range(3)],
            "base_pt": ("F", base_row),
            "add_pt": [("H", base_row + i) for i in range(3)],
        }
    else:
        return {
            "date": ("L", base_row),
            "route": [("N", base_row + i) for i in range(3)],
            "packing": [("O", base_row + i) for i in range(3)],
            "base_pt": ("Q", base_row),
            "add_pt": [("S", base_row + i) for i in range(3)],
        }


def serial_to_date(serial):
    return datetime(1899, 12, 30) + timedelta(days=int(serial))


def date_to_serial(d):
    return (d - datetime(1899, 12, 30)).days
