"""除外区間CSVの一括取り込み"""

import os
import re
import csv

from core.calc_utils import parse_kilo, m_to_px_x, WORK_AREAS, calc_location_string

# CSVのエリア略称 → 内部エリア名
AREA_FROM_CSV = {
    "左": "左軌間外",
    "内": "軌間内",
    "右": "右軌間外",
}


def parse_exclusion_csv(csv_path):
    """除外区間CSVを読み込み、除外区間のリストを返す。

    Returns:
        list of dict: [{'area': '左軌間外', 'start_m': 5.0, 'end_m': 6.0}, ...]
        除外設定が無い場合は空リスト。
    """
    exclusions = []
    in_exclusion_section = False

    with open(csv_path, 'r', encoding='cp932') as f:
        for line in f:
            line = line.strip()

            if not line:
                continue

            # 除外設定セクション開始
            if line.startswith("除外設定"):
                in_exclusion_section = True
                continue

            # 除外設定セクション終了（NO.で始まる行＝データセクション開始）
            if in_exclusion_section and line.startswith("NO."):
                break

            if in_exclusion_section:
                parts = line.split(',')
                if len(parts) >= 3:
                    area_short = parts[0].strip()
                    area_name = AREA_FROM_CSV.get(area_short)
                    if area_name is None:
                        continue
                    try:
                        start_m = float(parts[1].strip())
                        end_m = float(parts[2].strip())
                    except ValueError:
                        continue
                    exclusions.append({
                        'area': area_name,
                        'start_m': start_m,
                        'end_m': end_m,
                    })

    return exclusions


def find_csv_for_kilo(csv_folder, kilo_str, image_groups):
    """キロ程に対応するCSVファイルを探す。

    画像ファイル名のベース(拡張子なし) + ".csv" を探す。
    例: 20250927_002_両毛線_s_s_012k120m.csv

    Returns:
        csv_path or None
    """
    group = image_groups.get(kilo_str)
    if not group:
        return None

    # マーキングあり画像のファイル名をベースにする
    marked_path = group.get('marked')
    if not marked_path:
        return None

    basename = os.path.splitext(os.path.basename(marked_path))[0]

    # _c 等のサフィックスを除去してCSV名を探す
    # 例: 20250927_002_両毛線_s_s_012k120m_c.jpg → 20250927_002_両毛線_s_s_012k120m.csv
    # パターン: キロ程以降のサフィックスを除去
    match = re.match(r'(.+_\d+k\d+m)', basename)
    if match:
        csv_base = match.group(1)
    else:
        csv_base = basename

    csv_name = csv_base + '.csv'
    csv_path = os.path.join(csv_folder, csv_name)
    if os.path.exists(csv_path):
        return csv_path

    # サフィックス付きでも探す
    csv_name_full = basename + '.csv'
    csv_path_full = os.path.join(csv_folder, csv_name_full)
    if os.path.exists(csv_path_full):
        return csv_path_full

    return None


def import_exclusions_from_csv(csv_folder, image_groups, sorted_kilos, db, overwrite=False):
    """CSVフォルダから全キロ程の除外区間を一括取り込み。

    Args:
        csv_folder: CSVファイルが格納されたフォルダ
        image_groups: 画像グループ辞書
        sorted_kilos: ソート済みキロ程リスト
        db: DatabaseManager
        overwrite: True=既存除外区間を削除して上書き

    Returns:
        (imported_count, skipped_count)
    """
    imported = 0
    skipped = 0

    for kilo in sorted_kilos:
        csv_path = find_csv_for_kilo(csv_folder, kilo, image_groups)
        if csv_path is None:
            skipped += 1
            continue

        exclusions = parse_exclusion_csv(csv_path)
        if not exclusions:
            skipped += 1
            continue

        group = image_groups[kilo]
        direction = group.get('direction', '起点→終点')
        base_kilo_m = parse_kilo(kilo)

        # 既存除外区間の処理
        if overwrite:
            # 既存の除外区間を全て削除
            existing = db.load_drawings(kilo)
            for d in existing:
                if d.get('category') == '除外区間':
                    db.delete_drawing(d['db_id'])

        # 除外区間番号の開始値
        next_num = db.get_next_mgmt_number(kilo, category_filter='除外区間')

        for exc in exclusions:
            area = exc['area']
            start_m = exc['start_m']
            end_m = exc['end_m']

            # ピクセル座標に変換
            area_info = WORK_AREAS[area]
            lx0 = m_to_px_x(start_m)
            lx1 = m_to_px_x(end_m)
            ly0 = area_info["y_min"]
            ly1 = area_info["y_max"]

            drawing_dict = {
                'type': 'rectangle',
                'category': '除外区間',
                'area': area,
                'lx0': lx0, 'ly0': ly0, 'lx1': lx1, 'ly1': ly1,
                'tx': 0, 'ty': 0, 'text': '',
            }

            loc_str = calc_location_string(base_kilo_m, direction, start_m, end_m, 0.0, 2.0)

            db.insert_drawing(
                kilo, area, drawing_dict, loc_str,
                category='除外区間', mgmt_number=next_num,
                exclusion_reason='設備等による反射',
            )
            next_num += 1
            imported += 1

    return imported, skipped
