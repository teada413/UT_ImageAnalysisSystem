"""座標変換・キロ程計算ロジック"""

import re

# 画像の元サイズ
IMAGE_W = 1274
IMAGE_H = 992
# トリミング表示高さ（Y=900px以下をカット）
IMAGE_H_TRIMMED = 900

# キャンバス表示サイズ
CANVAS_W = 850
CANVAS_H = 800

# 初期ズーム倍率（850 / 1274 ≒ 0.667）
INITIAL_ZOOM = 0.65

# X軸の有効ピクセル範囲と実距離
X_PX_MIN = 72
X_PX_MAX = 916
X_RANGE_M = 20.0

# 3つの作業エリア定義
WORK_AREAS = {
    "左軌間外": {"x_min": 72, "x_max": 916, "y_min": 270, "y_max": 449, "stops": [270, 302, 351, 400, 449]},
    "軌間内":   {"x_min": 72, "x_max": 916, "y_min": 490, "y_max": 669, "stops": [490, 522, 571, 620, 669]},
    "右軌間外": {"x_min": 72, "x_max": 916, "y_min": 710, "y_max": 889, "stops": [710, 742, 791, 840, 889]},
}

# エリア名略称
AREA_SHORT = {"左軌間外": "左", "軌間内": "内", "右軌間外": "右"}

# 種別サフィックス
CATEGORY_SUFFIX = {"ゆるみ": "（ゆ）", "空洞": "（空）"}

# 丸数字
CIRCLED_NUMBERS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def circled_number(n):
    """整数を丸数字に変換（1→①、21以降は(21)形式）"""
    if 1 <= n <= 20:
        return CIRCLED_NUMBERS[n - 1]
    return f"({n})"


# ------------------------------------------------------------------
# ピクセル ↔ メートル変換
# ------------------------------------------------------------------

def px_to_m_x(px):
    """X軸ピクセル → メートル変換（線形）"""
    return (px - X_PX_MIN) / (X_PX_MAX - X_PX_MIN) * X_RANGE_M


def px_to_m_y(py, area_name):
    """Y軸ピクセル → メートル変換（非線形、各エリアのstopsに基づく）"""
    s = WORK_AREAS[area_name]["stops"]
    if py <= s[1]:
        return (py - s[0]) / 32.0 * 0.5
    elif py <= s[2]:
        return 0.5 + (py - s[1]) / 49.0 * 0.5
    elif py <= s[3]:
        return 1.0 + (py - s[2]) / 49.0 * 0.5
    else:
        return 1.5 + (py - s[3]) / 49.0 * 0.5


def m_to_px_x(m):
    """X軸メートル → ピクセル変換（線形）"""
    return X_PX_MIN + (m / X_RANGE_M) * (X_PX_MAX - X_PX_MIN)


def m_to_px_y(m, area_name):
    """Y軸メートル → ピクセル変換（非線形、px_to_m_yの逆変換）"""
    s = WORK_AREAS[area_name]["stops"]
    if m <= 0.5:
        return s[0] + m / 0.5 * 32.0
    elif m <= 1.0:
        return s[1] + (m - 0.5) / 0.5 * 49.0
    elif m <= 1.5:
        return s[2] + (m - 1.0) / 0.5 * 49.0
    else:
        return s[3] + (m - 1.5) / 0.5 * 49.0


# ------------------------------------------------------------------
# キロ程文字列
# ------------------------------------------------------------------

def _kilo_to_str(k_m):
    """メートル値をキロ程文字列に変換（例: 32987.5 → '32k987.5m', 12008.5 → '12k008.5m'）"""
    km = int(k_m) // 1000
    m = k_m % 1000
    return f"{km}k{m:05.1f}m"


def _calc_kilo_range(base_kilo_m, direction, min_x, max_x):
    """範囲のキロ程部分を計算"""
    if direction == "起点→終点":
        k_a = base_kilo_m + min_x
        k_b = base_kilo_m + max_x
    else:
        k_a = base_kilo_m - 20.0 + min_x
        k_b = base_kilo_m - 20.0 + max_x
    return min(k_a, k_b), max(k_a, k_b)


def calc_location_string(base_kilo_m, direction, min_x, max_x, min_y, max_y):
    """キロ程文字列を生成する（範囲＋深さ結合）"""
    start_k, end_k = _calc_kilo_range(base_kilo_m, direction, min_x, max_x)
    return f"{_kilo_to_str(start_k)} ～ {_kilo_to_str(end_k)}　ー　{min_y:.1f}m ～ {max_y:.1f}m"


def calc_range_string(base_kilo_m, direction, min_x, max_x):
    """範囲のみのキロ程文字列（GUI表示用）"""
    start_k, end_k = _calc_kilo_range(base_kilo_m, direction, min_x, max_x)
    return f"{_kilo_to_str(start_k)} ～ {_kilo_to_str(end_k)}"


def calc_depth_string(min_y, max_y):
    """深さ範囲の文字列（GUI表示用）"""
    return f"{min_y:.1f}m ～ {max_y:.1f}m"


def format_table_entry(mgmt_number, area, base_kilo_m, direction, min_x, max_x, min_y, max_y, category):
    """Excel表の行テキストを生成する。
    出力例: ① 右：32k987.5m ～ 32k991.0m　ー　0.6m ～ 1.3m（ゆ）
    """
    num_str = circled_number(mgmt_number) if mgmt_number else "-"
    area_str = AREA_SHORT.get(area, area)
    suffix = CATEGORY_SUFFIX.get(category, "")

    if direction == "起点→終点":
        k_a = base_kilo_m + min_x
        k_b = base_kilo_m + max_x
    else:
        k_a = base_kilo_m - 20.0 + min_x
        k_b = base_kilo_m - 20.0 + max_x

    start_k = min(k_a, k_b)
    end_k = max(k_a, k_b)

    return (
        f"{num_str} {area_str}："
        f"{_kilo_to_str(start_k)} ～ {_kilo_to_str(end_k)}"
        f"　ー　{min_y:.1f}m ～ {max_y:.1f}m{suffix}"
    )


def parse_kilo(k):
    """キロ程文字列（例: '012k120m'）を数値に変換"""
    match = re.match(r'(\d+)k(\d+)m', k)
    if match:
        return int(match.group(1)) * 1000 + int(match.group(2))
    return 0
