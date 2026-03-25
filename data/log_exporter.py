"""集計表ログ出力（詳細.log + 集計.log）"""

import os

from core.calc_utils import (
    parse_kilo, px_to_m_x, px_to_m_y, AREA_SHORT,
)


def _kilo_number(kilo_str):
    """キロ程文字列 → 整数メートル値 (例: '012k540m' → 12540)"""
    return parse_kilo(kilo_str)


def _abs_position(base_kilo_m, direction, offset_m):
    """オフセットメートルを絶対キロ程メートルに変換"""
    if direction == "起点→終点":
        return base_kilo_m + offset_m
    else:
        return base_kilo_m - 20.0 + offset_m


def _fmt_pos(val):
    """位置を出力用にフォーマット（末尾の不要な0を除去）"""
    s = f"{val:.1f}"
    # 12540.0 → 12540.0, 12546.5 → 12546.5
    # 整数の場合は小数なし: 12547.0 → 12547
    if s.endswith('.0'):
        return s[:-2]
    return s


def _image_basename_no_ext(image_path):
    """画像パスからベース名（拡張子なし）を取得"""
    return os.path.splitext(os.path.basename(image_path))[0]


def _image_basename(image_path):
    """画像パスからベース名（拡張子付き）を取得"""
    return os.path.basename(image_path)


def generate_detail_log(kilo_str, image_path, drawings, base_kilo_m, direction):
    """詳細.log の内容を生成して文字列で返す"""
    image_name = _image_basename(image_path) if image_path else ""
    lines = [image_name]

    # 変状と除外区間を分離
    defects = [d for d in drawings if d.get('category') != '除外区間']
    exclusions = [d for d in drawings if d.get('category') == '除外区間']

    # 変状: 管理番号順
    defects.sort(key=lambda d: d.get('mgmt_number') or 9999)
    for d in defects:
        mgmt = d.get('mgmt_number', 0)
        sort_num = 100 + mgmt  # ①→101, ②→102
        category = d.get('category', 'ゆるみ')
        area = d.get('area', '')
        area_short = AREA_SHORT.get(area, area)

        lx0, lx1 = d['lx0'], d['lx1']
        ly0, ly1 = d['ly0'], d['ly1']
        min_lx, max_lx = min(lx0, lx1), max(lx0, lx1)
        min_ly, max_ly = min(ly0, ly1), max(ly0, ly1)

        offset_start = px_to_m_x(min_lx)
        offset_end = px_to_m_x(max_lx)
        depth_start = px_to_m_y(min_ly, area)
        depth_end = px_to_m_y(max_ly, area)

        abs_start = _abs_position(base_kilo_m, direction, offset_start)
        abs_end = _abs_position(base_kilo_m, direction, offset_end)
        start_pos = min(abs_start, abs_end)
        end_pos = max(abs_start, abs_end)

        line = (
            f"{image_name}_異状マーク{sort_num}_{category}"
            f"_{_fmt_pos(start_pos)}_{_fmt_pos(end_pos)}"
            f"_{depth_start:.1f}_{depth_end:.1f}_{area_short}"
        )
        lines.append(line)

    # 除外区間: 管理番号順
    exclusions.sort(key=lambda d: d.get('mgmt_number') or 9999)
    for d in exclusions:
        mgmt = d.get('mgmt_number', 0)
        area = d.get('area', '')
        area_short = AREA_SHORT.get(area, area)
        reason = d.get('exclusion_reason', '設備等による反射')

        lx0, lx1 = d['lx0'], d['lx1']
        min_lx, max_lx = min(lx0, lx1), max(lx0, lx1)

        offset_start = px_to_m_x(min_lx)
        offset_end = px_to_m_x(max_lx)

        abs_start = _abs_position(base_kilo_m, direction, offset_start)
        abs_end = _abs_position(base_kilo_m, direction, offset_end)
        start_pos = min(abs_start, abs_end)
        end_pos = max(abs_start, abs_end)

        line = (
            f"{image_name}_除外{mgmt}_{area_short}_{reason}"
            f"__{_fmt_pos(start_pos)}_{_fmt_pos(end_pos)}"
        )
        lines.append(line)

    return "\n".join(lines) + "\n"


def generate_summary_log(kilo_str, drawings):
    """集計.log の内容を生成して文字列で返す"""
    kilo_m = _kilo_number(kilo_str)

    defects = [d for d in drawings if d.get('category') != '除外区間']
    exclusions = [d for d in drawings if d.get('category') == '除外区間']

    has_exclusion = "除外あり" if exclusions else "0"
    cavity_count = sum(1 for d in defects if d.get('category') == '空洞')
    loosening_count = sum(1 for d in defects if d.get('category') == 'ゆるみ')

    line = f"除外,{has_exclusion},空洞,{cavity_count},ゆるみ,{loosening_count},沈下,0,{kilo_m}"
    return line + "\n"


def build_log_filename(kilo_str, image_path, suffix):
    """ログファイル名を生成
    例: 1000012540_20250927_002_両毛線_s_s_012k540m_c{suffix}
    suffix: '詳細.log' or '集計.log'
    """
    kilo_m = _kilo_number(kilo_str)
    prefix = 1000000 + kilo_m
    basename = _image_basename_no_ext(image_path) if image_path else kilo_str
    return f"{prefix}_{basename}{suffix}"


def export_logs(output_dir, image_groups, sorted_kilos, db):
    """全キロ程の詳細.log と 集計.log を出力する

    Returns:
        出力ファイル数
    """
    count = 0
    for kilo in sorted_kilos:
        group = image_groups[kilo]
        image_path = group.get('marked')
        direction = group.get('direction', '起点→終点')
        base_kilo_m = parse_kilo(kilo)

        drawings = db.load_drawings(kilo)

        # 詳細.log
        detail_content = generate_detail_log(kilo, image_path, drawings, base_kilo_m, direction)
        detail_name = build_log_filename(kilo, image_path, "詳細.log")
        detail_path = os.path.join(output_dir, detail_name)
        with open(detail_path, 'w', encoding='cp932') as f:
            f.write(detail_content)
        count += 1

        # 集計.log
        summary_content = generate_summary_log(kilo, drawings)
        summary_name = build_log_filename(kilo, image_path, "集計.log")
        summary_path = os.path.join(output_dir, summary_name)
        with open(summary_path, 'w', encoding='cp932') as f:
            f.write(summary_content)
        count += 1

    return count
