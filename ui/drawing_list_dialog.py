"""変状一覧編集ダイアログ（現在ページ、変状/除外タブ分離）"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QAbstractItemView, QMessageBox, QTabWidget, QWidget,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor

from core.calc_utils import extract_line_type, strip_line_prefix, line_type_label

EXCLUSION_BG = QColor(255, 200, 200)
EXCLUSION_FG = QColor(80, 0, 0)


class DrawingListDialog(QDialog):
    """現在のキロ程の変状一覧を表示・編集できるダイアログ"""

    def __init__(self, parent=None, db=None, kilo=""):
        super().__init__(parent)
        lt = extract_line_type(kilo)
        bare = strip_line_prefix(kilo)
        self.setWindowTitle(f"変状一覧 - {line_type_label(lt)} {bare}")
        self.resize(800, 450)
        self._db = db
        self._kilo = kilo
        self._modified = False
        self._deleted_ids = []

        font = QFont("Meiryo", 10)
        layout = QVBoxLayout(self)

        # タブ
        self._tabs = QTabWidget()
        self._tabs.setFont(font)

        # 変状タブ
        self._defect_table = self._create_table(
            ["管理番号", "種別", "エリア", "範囲", "深さ"], [80, 70, 80, 250], editable_cols={0}
        )
        defect_w = QWidget()
        dl = QVBoxLayout(defect_w)
        dl.setContentsMargins(0, 0, 0, 0)
        dl.addWidget(self._defect_table)
        self._tabs.addTab(defect_w, "変状")

        # 除外タブ
        self._exclusion_table = self._create_table(
            ["管理番号", "エリア", "範囲", "除外理由"], [80, 80, 250], editable_cols={0, 3}
        )
        exclusion_w = QWidget()
        el = QVBoxLayout(exclusion_w)
        el.setContentsMargins(0, 0, 0, 0)
        el.addWidget(self._exclusion_table)
        self._tabs.addTab(exclusion_w, "除外区間")

        layout.addWidget(self._tabs)

        # ボタン
        btn_layout = QHBoxLayout()

        del_btn = QPushButton("選択した項目を削除")
        del_btn.setFont(QFont("Meiryo", 12, QFont.Bold))
        del_btn.setStyleSheet(
            "QPushButton { background-color: #d32f2f; color: white; padding: 8px 24px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #b71c1c; }"
        )
        del_btn.clicked.connect(self._delete_selected)
        btn_layout.addWidget(del_btn)

        btn_layout.addStretch()

        save_btn = QPushButton("保存")
        save_btn.setFont(QFont("Meiryo", 12, QFont.Bold))
        save_btn.setStyleSheet(
            "QPushButton { background-color: #2ea043; color: white; padding: 8px 24px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #238636; }"
        )
        save_btn.clicked.connect(self._save)
        btn_layout.addWidget(save_btn)

        close_btn = QPushButton("閉じる")
        close_btn.setFont(QFont("Meiryo", 12, QFont.Bold))
        close_btn.setStyleSheet("QPushButton { padding: 8px 24px; border-radius: 4px; }")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)
        self._load_data()

    def _create_table(self, headers, section_widths, editable_cols=None):
        self._editable_cols = editable_cols or set()
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setFont(QFont("Meiryo", 10))
        h = table.horizontalHeader()
        h.setFont(QFont("Meiryo", 10, QFont.Bold))
        for i, w in enumerate(section_widths):
            h.resizeSection(i, w)
        h.setStretchLastSection(True)
        # ダブルクリックで編集対象列のみ編集
        ec = editable_cols or set()
        table.cellDoubleClicked.connect(lambda r, c, t=table, e=ec: self._edit_cell(t, r, c, e))
        table.cellChanged.connect(lambda r, c: self._on_changed(c))
        return table

    @staticmethod
    def _edit_cell(table, row, col, editable_cols):
        if col in editable_cols:
            table.editItem(table.item(row, col))

    def _on_changed(self, col):
        self._modified = True

    def _load_data(self):
        from core.calc_utils import parse_kilo, px_to_m_x, px_to_m_y, calc_range_string, calc_depth_string
        self._base_kilo_m = parse_kilo(self._kilo)

        drawings = self._db.load_drawings(self._kilo)
        defects = [d for d in drawings if d.get('category') != '除外区間']
        exclusions = [d for d in drawings if d.get('category') == '除外区間']
        self._populate_defect(defects)
        self._populate_exclusion(exclusions)

    def _populate_defect(self, defects):
        from core.calc_utils import px_to_m_x, px_to_m_y, calc_range_string, calc_depth_string
        t = self._defect_table
        t.blockSignals(True)
        t.setRowCount(0)
        for d in defects:
            row = t.rowCount()
            t.insertRow(row)

            mgmt = d.get('mgmt_number')
            mi = QTableWidgetItem(str(mgmt) if mgmt is not None else "")
            mi.setTextAlignment(Qt.AlignCenter)
            mi.setData(Qt.UserRole, d['db_id'])
            mi.setData(Qt.UserRole + 1, d.get('category', ''))
            mi.setData(Qt.UserRole + 2, mgmt)
            t.setItem(row, 0, mi)

            ci = QTableWidgetItem(d.get('category', ''))
            ci.setFlags(ci.flags() & ~Qt.ItemIsEditable)
            ci.setTextAlignment(Qt.AlignCenter)
            t.setItem(row, 1, ci)

            ai = QTableWidgetItem(d.get('area', ''))
            ai.setFlags(ai.flags() & ~Qt.ItemIsEditable)
            ai.setTextAlignment(Qt.AlignCenter)
            t.setItem(row, 2, ai)

            area = d.get('area', '')
            lx0, lx1 = d['lx0'], d['lx1']
            ly0, ly1 = d['ly0'], d['ly1']
            range_str = calc_range_string(self._base_kilo_m, "起点→終点",
                                          px_to_m_x(min(lx0, lx1)), px_to_m_x(max(lx0, lx1)))
            depth_str = calc_depth_string(px_to_m_y(min(ly0, ly1), area),
                                          px_to_m_y(max(ly0, ly1), area))

            ri = QTableWidgetItem(range_str)
            ri.setFlags(ri.flags() & ~Qt.ItemIsEditable)
            t.setItem(row, 3, ri)

            di = QTableWidgetItem(depth_str)
            di.setFlags(di.flags() & ~Qt.ItemIsEditable)
            t.setItem(row, 4, di)
        t.blockSignals(False)

    def _populate_exclusion(self, exclusions):
        from core.calc_utils import px_to_m_x, calc_range_string
        t = self._exclusion_table
        t.blockSignals(True)
        t.setRowCount(0)
        for d in exclusions:
            row = t.rowCount()
            t.insertRow(row)

            mgmt = d.get('mgmt_number')
            mi = QTableWidgetItem(str(mgmt) if mgmt is not None else "")
            mi.setTextAlignment(Qt.AlignCenter)
            mi.setData(Qt.UserRole, d['db_id'])
            mi.setData(Qt.UserRole + 1, '除外区間')
            mi.setData(Qt.UserRole + 2, mgmt)
            t.setItem(row, 0, mi)

            ai = QTableWidgetItem(d.get('area', ''))
            ai.setFlags(ai.flags() & ~Qt.ItemIsEditable)
            ai.setTextAlignment(Qt.AlignCenter)
            t.setItem(row, 1, ai)

            lx0, lx1 = d['lx0'], d['lx1']
            range_str = calc_range_string(self._base_kilo_m, "起点→終点",
                                          px_to_m_x(min(lx0, lx1)), px_to_m_x(max(lx0, lx1)))
            ri_range = QTableWidgetItem(range_str)
            ri_range.setFlags(ri_range.flags() & ~Qt.ItemIsEditable)
            t.setItem(row, 2, ri_range)

            ri = QTableWidgetItem(d.get('exclusion_reason', ''))
            t.setItem(row, 3, ri)

            for col in range(4):
                item = t.item(row, col)
                if item:
                    item.setBackground(EXCLUSION_BG)
                    item.setForeground(EXCLUSION_FG)
        t.blockSignals(False)

    def _delete_selected(self):
        """現在アクティブなタブの選択行を削除"""
        idx = self._tabs.currentIndex()
        table = self._defect_table if idx == 0 else self._exclusion_table

        rows = sorted({item.row() for item in table.selectedItems()}, reverse=True)
        if not rows:
            return

        reply = QMessageBox.question(
            self, "削除の確認",
            f"{len(rows)} 件を削除しますか？\n（保存時にDBに反映されます）",
        )
        if reply != QMessageBox.Yes:
            return

        for row in rows:
            mi = table.item(row, 0)
            if mi:
                db_id = mi.data(Qt.UserRole)
                if db_id is not None:
                    self._deleted_ids.append(db_id)
            table.removeRow(row)
        self._modified = True

    def _save(self):
        errors = []
        updates = []

        # 管理番号の変更を収集
        for table, is_excl in [(self._defect_table, False), (self._exclusion_table, True)]:
            for row in range(table.rowCount()):
                mi = table.item(row, 0)
                if mi is None:
                    continue

                db_id = mi.data(Qt.UserRole)
                category = mi.data(Qt.UserRole + 1)
                original = mi.data(Qt.UserRole + 2)

                text = mi.text().strip()
                new_val = None
                if text:
                    try:
                        new_val = int(text)
                    except ValueError:
                        errors.append(f"'{text}' は整数ではありません")
                        continue

                if new_val != original:
                    updates.append(('mgmt_number', db_id, category, new_val, original))

                # 除外理由の変更
                if is_excl:
                    ri = table.item(row, 3)
                    if ri:
                        updates.append(('exclusion_reason', db_id, None, ri.text(), None))

        # 管理番号の重複チェック（入れ替えを考慮: 変更後の値同士で重複がないか確認）
        mgmt_updates = [u for u in updates if u[0] == 'mgmt_number' and u[3] is not None]
        # 変更後の番号をカテゴリ別に集計
        new_numbers_defect = {}  # number -> db_id
        new_numbers_excl = {}
        for _, db_id, category, new_val, original in mgmt_updates:
            bucket = new_numbers_excl if category == '除外区間' else new_numbers_defect
            if new_val in bucket:
                errors.append(f"管理番号 {new_val} が重複しています")
            bucket[new_val] = db_id

        # 変更のないレコードの既存番号との重複チェック
        changed_defect_ids = {u[1] for u in mgmt_updates if u[2] != '除外区間'}
        changed_excl_ids = {u[1] for u in mgmt_updates if u[2] == '除外区間'}

        all_drawings = self._db.load_drawings(self._kilo)
        for d in all_drawings:
            if d['db_id'] in self._deleted_ids:
                continue
            mgmt = d.get('mgmt_number')
            if mgmt is None:
                continue
            cat = d.get('category', '')
            if cat == '除外区間':
                if d['db_id'] not in changed_excl_ids and mgmt in new_numbers_excl:
                    errors.append(f"除外管理番号 {mgmt} は既に使用されています")
            else:
                if d['db_id'] not in changed_defect_ids and mgmt in new_numbers_defect:
                    errors.append(f"管理番号 {mgmt} は既に使用されています")

        if errors:
            QMessageBox.warning(self, "保存エラー", "\n".join(set(errors)))
            return

        # DB更新
        cursor = self._db.conn.cursor()

        # 削除
        for db_id in self._deleted_ids:
            cursor.execute("DELETE FROM drawings WHERE id=?", (db_id,))

        # 管理番号更新（入れ替え対応: 一旦NULLにしてから設定）
        mgmt_ids = [u[1] for u in mgmt_updates]
        if mgmt_ids:
            cursor.executemany(
                "UPDATE drawings SET mgmt_number=NULL WHERE id=?",
                [(did,) for did in mgmt_ids],
            )
        for field, db_id, category, new_val, original in updates:
            if field == 'mgmt_number':
                cursor.execute("UPDATE drawings SET mgmt_number=? WHERE id=?", (new_val, db_id))
            elif field == 'exclusion_reason':
                cursor.execute("UPDATE drawings SET exclusion_reason=? WHERE id=?", (new_val, db_id))

        self._db.conn.commit()

        # original 値を更新
        for table in [self._defect_table, self._exclusion_table]:
            for row in range(table.rowCount()):
                mi = table.item(row, 0)
                if mi:
                    text = mi.text().strip()
                    mi.setData(Qt.UserRole + 2, int(text) if text else None)

        self._deleted_ids.clear()
        self._modified = False
        count = len(updates) + len(self._deleted_ids)
        QMessageBox.information(self, "保存完了", "保存しました。")

    def closeEvent(self, event):
        if self._modified:
            reply = QMessageBox.question(
                self, "未保存の変更",
                "未保存の変更があります。保存せずに閉じますか？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                event.ignore()
                return
        event.accept()
