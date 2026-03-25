"""SQLite操作クラス"""

import sqlite3


class DatabaseManager:
    def __init__(self):
        self.conn = None

    def setup(self, db_path):
        """DB接続・テーブル作成"""
        if self.conn:
            self.conn.close()

        self.conn = sqlite3.connect(db_path)
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS drawings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kilo TEXT,
                area TEXT,
                shape_type TEXT,
                lx0 REAL, ly0 REAL, lx1 REAL, ly1 REAL,
                tx REAL, ty REAL, text_lbl TEXT,
                location_str TEXT
            )
        ''')
        # 既存DBへのマイグレーション
        for col, typedef in [
            ("category", "TEXT DEFAULT 'ゆるみ'"),
            ("mgmt_number", "INTEGER DEFAULT NULL"),
            ("exclusion_reason", "TEXT DEFAULT ''"),
        ]:
            try:
                cursor.execute(f"ALTER TABLE drawings ADD COLUMN {col} {typedef}")
            except sqlite3.OperationalError:
                pass  # カラムが既に存在
        self.conn.commit()

        # 既存レコードの location_str を現在のフォーマットで再生成
        self._refresh_location_strings()

    def _refresh_location_strings(self):
        """全レコードの location_str を現在の calc_location_string で再生成"""
        from core.calc_utils import parse_kilo, px_to_m_x, px_to_m_y, calc_location_string
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, kilo, area, lx0, ly0, lx1, ly1, category FROM drawings")
        rows = cursor.fetchall()
        for db_id, kilo, area, lx0, ly0, lx1, ly1, category in rows:
            base_kilo_m = parse_kilo(kilo)
            # direction は DB に保存していないため起点→終点で統一
            # （元のデータも同じ direction で生成されている前提）
            min_lx, max_lx = min(lx0, lx1), max(lx0, lx1)
            min_ly, max_ly = min(ly0, ly1), max(ly0, ly1)
            min_m_x = px_to_m_x(min_lx)
            max_m_x = px_to_m_x(max_lx)
            if category == '除外区間':
                min_m_y, max_m_y = 0.0, 2.0
            else:
                min_m_y = px_to_m_y(min_ly, area)
                max_m_y = px_to_m_y(max_ly, area)
            loc_str = calc_location_string(base_kilo_m, "起点→終点", min_m_x, max_m_x, min_m_y, max_m_y)
            cursor.execute("UPDATE drawings SET location_str=? WHERE id=?", (loc_str, db_id))
        self.conn.commit()

    @property
    def is_connected(self):
        return self.conn is not None

    def insert_drawing(self, kilo, area, drawing_dict, location_str,
                       category="ゆるみ", mgmt_number=None, exclusion_reason=""):
        """図形データを挿入し、挿入されたIDを返す"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO drawings
                (kilo, area, shape_type, lx0, ly0, lx1, ly1, tx, ty, text_lbl,
                 location_str, category, mgmt_number, exclusion_reason)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            kilo, area, drawing_dict['type'],
            drawing_dict['lx0'], drawing_dict['ly0'], drawing_dict['lx1'], drawing_dict['ly1'],
            drawing_dict.get('tx', 0), drawing_dict.get('ty', 0),
            drawing_dict.get('text', ''), location_str,
            category, mgmt_number, exclusion_reason,
        ))
        self.conn.commit()
        return cursor.lastrowid

    def load_drawings(self, kilo):
        """キロ程指定でレコードを取得"""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT id, area, shape_type, lx0, ly0, lx1, ly1, tx, ty, text_lbl, "
            "location_str, category, mgmt_number, exclusion_reason "
            "FROM drawings WHERE kilo=?",
            (kilo,)
        )
        rows = cursor.fetchall()

        drawings = []
        for row in rows:
            (db_id, area, shape_type, lx0, ly0, lx1, ly1, tx, ty, text_lbl,
             location_str, category, mgmt_number, exclusion_reason) = row
            d = {
                'db_id': db_id, 'area': area, 'type': shape_type,
                'lx0': lx0, 'ly0': ly0, 'lx1': lx1, 'ly1': ly1,
                'tx': tx, 'ty': ty, 'text': text_lbl,
                'location_str': location_str,
                'category': category or 'ゆるみ',
                'mgmt_number': mgmt_number,
                'exclusion_reason': exclusion_reason or '',
            }
            drawings.append(d)
        return drawings

    def update_drawing_coords(self, db_id, lx0, ly0, lx1, ly1, location_str):
        """図形の座標と位置文字列を更新"""
        cursor = self.conn.cursor()
        cursor.execute(
            "UPDATE drawings SET lx0=?, ly0=?, lx1=?, ly1=?, location_str=? WHERE id=?",
            (lx0, ly0, lx1, ly1, location_str, db_id),
        )
        self.conn.commit()

    def delete_drawing(self, db_id):
        """レコード削除"""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM drawings WHERE id=?", (db_id,))
        self.conn.commit()

    def is_mgmt_number_taken(self, kilo, number, category_filter=None):
        """同一キロ程内で管理番号が既に使われているか
        category_filter: '除外区間' or None(変状)で検索対象を分離"""
        cursor = self.conn.cursor()
        if category_filter == '除外区間':
            cursor.execute(
                "SELECT COUNT(*) FROM drawings WHERE kilo=? AND mgmt_number=? AND category='除外区間'",
                (kilo, number),
            )
        else:
            cursor.execute(
                "SELECT COUNT(*) FROM drawings WHERE kilo=? AND mgmt_number=? AND category!='除外区間'",
                (kilo, number),
            )
        return cursor.fetchone()[0] > 0

    def get_next_mgmt_number(self, kilo, category_filter=None):
        """次の管理番号を提案（変状と除外区間で別カウント）"""
        cursor = self.conn.cursor()
        if category_filter == '除外区間':
            cursor.execute(
                "SELECT MAX(mgmt_number) FROM drawings WHERE kilo=? AND mgmt_number IS NOT NULL AND category='除外区間'",
                (kilo,),
            )
        else:
            cursor.execute(
                "SELECT MAX(mgmt_number) FROM drawings WHERE kilo=? AND mgmt_number IS NOT NULL AND category!='除外区間'",
                (kilo,),
            )
        result = cursor.fetchone()[0]
        return (result or 0) + 1

    def get_exclusion_zones(self, kilo, area):
        """指定キロ程・エリアの除外区間の範囲一覧を返す [(lx0, lx1), ...]"""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT lx0, lx1 FROM drawings WHERE kilo=? AND area=? AND category='除外区間'",
            (kilo, area),
        )
        return [(min(r[0], r[1]), max(r[0], r[1])) for r in cursor.fetchall()]

    def close(self):
        """接続クローズ"""
        if self.conn:
            self.conn.close()
            self.conn = None
