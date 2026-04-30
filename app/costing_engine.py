"""
車両原価計算エンジン
ZST社の「再々修正_車両仮原価表.xlsx」のロジックをPythonに移植

入力項目:
- 車格: "2t" / "4t" / "大型"
- 運行種別: "長距離" / "中距離" / "地場"
- リース有無: bool
- 月稼働日数, 走行距離/日, 通行料/日, 燃料単価
- (オプション) 残業時間/月, 基本給, 仮収入

出力:
- 1日あたり各種コスト
- 最低原価費用
- 推奨単価
- 黒字判定
"""
from typing import Optional


# 車格別のデフォルト値（Excelから抽出）
VEHICLE_PRESETS = {
    "2t": {
        "tire_cost_per_km": 5.0,         # タイヤ費 円/km（Excel タイヤシートより）
        "fuel_efficiency": 5.0,           # 燃費 km/L
        "depreciation_years": 4,          # 償却年数
        "default_purchase_price": 5_000_000,
    },
    "4t": {
        "tire_cost_per_km": 5.0,
        "fuel_efficiency": 4.0,
        "depreciation_years": 5,
        "default_purchase_price": 8_000_000,
    },
    "大型": {
        "tire_cost_per_km": 6.39,
        "fuel_efficiency": 3.0,
        "depreciation_years": 5,
        "default_purchase_price": 12_000_000,
    },
}

# 運行種別 × リース有無 別のデフォルト値
ROUTE_PRESETS = {
    ("長距離", True): {
        "default_days": 22,
        "default_distance_per_day": 700,
        "default_toll_per_day": 21000,
        "default_fuel_price": 133,
        "default_overtime_hours": 30,
        "default_base_salary": 500_000,
        "lease_cost_monthly": 80_000,  # リース料金 月額（推定）
    },
    ("長距離", False): {
        "default_days": 22,
        "default_distance_per_day": 700,
        "default_toll_per_day": 21000,
        "default_fuel_price": 133,
        "default_overtime_hours": 30,
        "default_base_salary": 500_000,
        "lease_cost_monthly": 0,
    },
    ("中距離", True): {
        "default_days": 24,
        "default_distance_per_day": 387,
        "default_toll_per_day": 9776,
        "default_fuel_price": 128,
        "default_overtime_hours": 30,
        "default_base_salary": 490_808,
        "lease_cost_monthly": 80_000,
    },
    ("中距離", False): {
        "default_days": 24,
        "default_distance_per_day": 366,
        "default_toll_per_day": 13556,
        "default_fuel_price": 127,
        "default_overtime_hours": 150,
        "default_base_salary": 580_000,
        "lease_cost_monthly": 0,
    },
    ("地場", True): {
        "default_days": 22,
        "default_distance_per_day": 250,
        "default_toll_per_day": 7000,
        "default_fuel_price": 125,
        "default_overtime_hours": 30,
        "default_base_salary": 397_600,
        "lease_cost_monthly": 80_000,
    },
    ("地場", False): {
        "default_days": 22,
        "default_distance_per_day": 250,
        "default_toll_per_day": 7000,
        "default_fuel_price": 125,
        "default_overtime_hours": 30,
        "default_base_salary": 397_600,
        "lease_cost_monthly": 0,
    },
}


def calc_costing(
    vehicle_class: str,        # "2t" / "4t" / "大型"
    route_type: str,           # "長距離" / "中距離" / "地場"
    is_leased: bool,
    monthly_days: int,
    distance_per_day: float,
    toll_per_day: float,
    fuel_price: float,
    overtime_hours: Optional[float] = None,
    base_salary: Optional[float] = None,
    purchase_price: Optional[float] = None,
) -> dict:
    """
    原価計算メイン
    
    Returns:
        {
          "input": {...},        # 入力値
          "monthly": {...},      # 月額コスト内訳
          "daily": {...},        # 日額コスト内訳
          "min_daily_revenue": # 最低受注単価
          "recommended_10": # 利益10%確保の推奨単価
          "recommended_20": # 利益20%確保の推奨単価
        }
    """
    # 車格・運行種別のプリセット取得
    v_preset = VEHICLE_PRESETS.get(vehicle_class, VEHICLE_PRESETS["大型"])
    r_preset = ROUTE_PRESETS.get((route_type, is_leased), ROUTE_PRESETS[("地場", False)])
    
    # デフォルト値補完
    overtime_hours = overtime_hours if overtime_hours is not None else r_preset["default_overtime_hours"]
    base_salary = base_salary if base_salary is not None else r_preset["default_base_salary"]
    purchase_price = purchase_price if purchase_price is not None else v_preset["default_purchase_price"]
    
    # ====== 月額コスト計算 ======
    
    # 1. 人件費
    total_work_hours = 8 * monthly_days + overtime_hours  # 所定 + 残業
    bonus_provision = 200_000 / 12  # 賞与引当
    legal_welfare = base_salary * 0.15876  # 法定福利費（給与×15.876%）
    bonus_welfare = bonus_provision * 0.15876
    transit_allowance = 3000  # 通勤手当
    family_allowance = 0
    
    personnel_total = (
        base_salary
        + family_allowance
        + transit_allowance
        + bonus_provision
        + legal_welfare
        + bonus_welfare
    )
    
    # 2. 運行三費
    monthly_distance = distance_per_day * monthly_days
    fuel_cost = monthly_distance / v_preset["fuel_efficiency"] * fuel_price
    oil_cost = fuel_cost * 0.09  # 油脂費 = 燃料費の9%
    
    repair_cost = 38_554  # 修繕費（年額×12÷12、Excel値）
    
    # タイヤ費 = 月間距離 × 円/km
    tire_cost = monthly_distance * v_preset["tire_cost_per_km"]
    
    monthly_toll = toll_per_day * monthly_days
    
    operation_total = fuel_cost + oil_cost + repair_cost + tire_cost + monthly_toll
    
    # 3. 諸税・保険
    insurance_compulsory = 39_100 / 12  # 自賠責
    weight_tax = 31_200 / 12  # 重量税
    auto_tax = 63_000 / 12  # 自動車税
    voluntary_insurance = 190_000 / 12  # 任意保険
    
    tax_insurance_total = insurance_compulsory + weight_tax + auto_tax + voluntary_insurance
    
    # 4. 車両償却費 or リース費
    if is_leased:
        vehicle_cost = r_preset["lease_cost_monthly"]
    else:
        vehicle_cost = purchase_price / v_preset["depreciation_years"] / 12
    
    # 取得税（年額の月割、購入時のみ）
    acquisition_tax = (purchase_price * 0.01) / 5 / 12 if not is_leased else 0
    
    # 運行原価合計
    operating_total = (
        personnel_total
        + operation_total
        + tax_insurance_total
        + vehicle_cost
        + acquisition_tax
    )
    
    # ====== 日額換算 ======
    daily = {
        "人件費": personnel_total / monthly_days,
        "燃料費": fuel_cost / monthly_days,
        "油脂費": oil_cost / monthly_days,
        "修繕費": repair_cost / monthly_days,
        "タイヤ費": tire_cost / monthly_days,
        "通行料": monthly_toll / monthly_days,
        "諸税・保険": tax_insurance_total / monthly_days,
        "車両費": (vehicle_cost + acquisition_tax) / monthly_days,
    }
    
    monthly = {
        "人件費": personnel_total,
        "燃料費": fuel_cost,
        "油脂費": oil_cost,
        "修繕費": repair_cost,
        "タイヤ費": tire_cost,
        "通行料": monthly_toll,
        "諸税・保険": tax_insurance_total,
        "車両費": vehicle_cost + acquisition_tax,
        "運行原価合計": operating_total,
    }
    
    min_daily_cost = operating_total / monthly_days
    
    # 一般管理費（収入の約3%）を含めた最低単価
    min_daily_revenue = min_daily_cost / 0.97  # 一般管理費3%控除前
    recommended_10 = min_daily_revenue * 1.10
    recommended_20 = min_daily_revenue * 1.20
    
    return {
        "input": {
            "vehicle_class": vehicle_class,
            "route_type": route_type,
            "is_leased": is_leased,
            "monthly_days": monthly_days,
            "distance_per_day": distance_per_day,
            "toll_per_day": toll_per_day,
            "fuel_price": fuel_price,
            "overtime_hours": overtime_hours,
            "base_salary": base_salary,
            "purchase_price": purchase_price,
        },
        "monthly": monthly,
        "daily": daily,
        "min_daily_cost": min_daily_cost,
        "min_daily_revenue": min_daily_revenue,
        "recommended_10": recommended_10,
        "recommended_20": recommended_20,
        "monthly_distance": monthly_distance,
    }


def evaluate_offer(costing: dict, daily_revenue: float) -> dict:
    """
    受注金額に対する黒字判定
    
    Args:
        costing: calc_costing() の結果
        daily_revenue: 1日あたりの提示金額
    
    Returns:
        {
          "daily_profit": 日利益,
          "monthly_profit": 月利益,
          "profit_rate": 利益率,
          "is_profitable": 黒字か,
          "judgment": "高収益" / "推奨" / "ぎりぎり" / "赤字",
        }
    """
    monthly_days = costing["input"]["monthly_days"]
    min_daily_revenue = costing["min_daily_revenue"]
    
    # 一般管理費3%控除
    daily_net_revenue = daily_revenue * 0.97
    daily_cost = costing["min_daily_cost"]
    daily_profit = daily_net_revenue - daily_cost
    monthly_profit = daily_profit * monthly_days
    profit_rate = (daily_profit / daily_revenue) if daily_revenue > 0 else 0
    
    if profit_rate >= 0.20:
        judgment = "🌟 高収益（利益20%以上）"
    elif profit_rate >= 0.10:
        judgment = "✅ 推奨（利益10%以上）"
    elif profit_rate >= 0.03:
        judgment = "⚠ ぎりぎり（利益3-10%）"
    elif profit_rate >= 0:
        judgment = "🟡 損益分岐ライン（利益3%未満）"
    else:
        judgment = "🔴 赤字（受注すべきでない）"
    
    return {
        "daily_revenue": daily_revenue,
        "daily_cost": daily_cost,
        "daily_profit": daily_profit,
        "monthly_profit": monthly_profit,
        "profit_rate": profit_rate,
        "is_profitable": daily_profit > 0,
        "judgment": judgment,
        "daily_net_revenue": daily_net_revenue,
    }


def compare_offers(costing: dict, prices: list[float]) -> list[dict]:
    """複数単価を比較"""
    return [evaluate_offer(costing, p) for p in prices]
