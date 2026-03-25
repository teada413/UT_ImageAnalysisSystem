"""フォルダ読み込み・画像ペアリング"""

import os
import re

from core.calc_utils import parse_kilo


def load_image_groups(parent_folder):
    """親フォルダからマーキングあり/なしのペアを読み込む。

    Returns:
        (image_groups, error_message)
        成功時: (dict, None)
        失敗時: (None, str)
    """
    dir_marked, dir_unmarked = None, None
    for d in os.listdir(parent_folder):
        path = os.path.join(parent_folder, d)
        if os.path.isdir(path):
            if "マーキングあり" in d:
                dir_marked = path
            elif "マーキングなし" in d:
                dir_unmarked = path

    if not dir_marked or not dir_unmarked:
        return None, "エラー: サブフォルダが見つかりません"

    image_groups = {}
    for filename in os.listdir(dir_marked):
        if not filename.lower().endswith(('.jpg', '.jpeg')):
            continue

        path_marked = os.path.join(dir_marked, filename)
        path_unmarked = os.path.join(dir_unmarked, filename)
        if not os.path.exists(path_unmarked):
            path_unmarked = None

        match = re.search(r'(\d+k\d+m)', filename)
        if match:
            kilo_str = match.group(1)

            direction = "起点→終点"
            if "終点→起点" in filename:
                direction = "終点→起点"

            image_groups[kilo_str] = {
                'marked': path_marked,
                'unmarked': path_unmarked,
                'direction': direction,
            }

    return image_groups, None


def sort_kilos(image_groups):
    """キロ程キーを数値順にソートして返す"""
    return sorted(image_groups.keys(), key=parse_kilo)
