"""UI部品（QTableWidget + キロ程一覧）"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget,
    QTableWidgetItem, QHeaderView, QPushButton, QAbstractItemView,
    QListWidget, QListWidgetItem,
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QColor

from core.calc_utils import circled_number

EXCLUSION_BG = QColor(255, 200, 200)
EXCLUSION_FG = QColor(80, 0, 0)


class DrawingTable(QWidget):
    """描画情報を表示するテーブル（編集対応）"""

    row_selected = Signal(int)
    data_edited = Signal(int, str, object)

    def __init__(self, parent=None, on_delete_callback=None, ui_font_family="Meiryo"):
        super().__init__(parent)
        self.on_delete_callback = on_delete_callback

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # テーブル（6列: 管理番号, 種別, エリア, 範囲, 深さ, 除外理由）
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["管理番号", "種別", "エリア", "範囲", "深さ", "除外理由"]
        )
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setMaximumHeight(160)

        self.table.setStyleSheet(
            "QTableWidget::item:selected { background-color: #4a90d9; color: white; }"
        )

        header_font = QFont(ui_font_family, 11)
        header_font.setBold(True)
        self.table.horizontalHeader().setFont(header_font)

        data_font = QFont(ui_font_family, 10)
        self.table.setFont(data_font)

        header = self.table.horizontalHeader()
        header.resizeSection(0, 80)
        header.resizeSection(1, 80)
        header.resizeSection(2, 80)
        header.resizeSection(3, 350)
        header.resizeSection(4, 150)
        header.setStretchLastSection(True)

        self.table.itemSelectionChanged.connect(self._on_selection_changed)
        self.table.cellDoubleClicked.connect(self._on_cell_double_click)
        self.table.cellChanged.connect(self._on_cell_changed)

        layout.addWidget(self.table)

        # ボタン行
        self.btn_layout = QHBoxLayout()
        self.btn_layout.setContentsMargins(0, 0, 0, 0)

        self.delete_btn = QPushButton("選択した図形を削除")
        self.delete_btn.setFont(QFont(ui_font_family, 12, QFont.Bold))
        self.delete_btn.setStyleSheet(
            "QPushButton { background-color: #d32f2f; color: white; padding: 6px 16px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #b71c1c; }"
        )
        self.delete_btn.clicked.connect(self._on_delete)
        self.btn_layout.addStretch()
        self.btn_layout.addWidget(self.delete_btn)

        layout.addLayout(self.btn_layout)

    def _on_selection_changed(self):
        rows = set()
        for item in self.table.selectedItems():
            rows.add(item.row())
        if len(rows) == 1:
            row = next(iter(rows))
            id_item = self.table.item(row, 0)
            if id_item and id_item.data(Qt.UserRole) is not None:
                self.row_selected.emit(int(id_item.data(Qt.UserRole)))
                return
        self.row_selected.emit(-1)

    def _on_cell_double_click(self, row, col):
        """管理番号(col=0)と除外理由(col=5)のみ編集可能"""
        if col == 0:
            item = self.table.item(row, col)
            if item:
                self.table.editItem(item)
        elif col == 5:
            cat_item = self.table.item(row, 1)
            if cat_item and cat_item.text() == "除外区間":
                item = self.table.item(row, col)
                if item:
                    self.table.editItem(item)

    def _on_cell_changed(self, row, col):
        if col == 0:
            item = self.table.item(row, 0)
            if item:
                db_id = item.data(Qt.UserRole)
                text = item.text().strip()
                try:
                    val = int(text) if text else None
                except ValueError:
                    return
                self.data_edited.emit(db_id, 'mgmt_number', val)
        elif col == 5:
            item = self.table.item(row, 5)
            id_item = self.table.item(row, 0)
            if item and id_item:
                db_id = id_item.data(Qt.UserRole)
                self.data_edited.emit(db_id, 'exclusion_reason', item.text())

    def _on_delete(self):
        selected_rows = set()
        for item in self.table.selectedItems():
            selected_rows.add(item.row())

        selected_ids = []
        for row in selected_rows:
            id_item = self.table.item(row, 0)
            if id_item and id_item.data(Qt.UserRole) is not None:
                selected_ids.append(str(id_item.data(Qt.UserRole)))

        if selected_ids and self.on_delete_callback:
            self.on_delete_callback(selected_ids)

    def clear(self):
        self.table.setRowCount(0)

    def insert_row(self, db_id, area, range_str="", depth_str="",
                   category="ゆるみ", mgmt_number=None, exclusion_reason=""):
        self.table.blockSignals(True)
        row = self.table.rowCount()
        self.table.insertRow(row)

        is_exclusion = (category == "除外区間")

        # col0: 管理番号
        if mgmt_number is None:
            num_display = "-"
        else:
            num_display = circled_number(mgmt_number)
        id_item = QTableWidgetItem(num_display)
        id_item.setTextAlignment(Qt.AlignCenter)
        id_item.setData(Qt.UserRole, db_id)
        self.table.setItem(row, 0, id_item)

        # col1: 種別
        cat_item = QTableWidgetItem(category)
        cat_item.setTextAlignment(Qt.AlignCenter)
        cat_item.setFlags(cat_item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row, 1, cat_item)

        # col2: エリア
        area_item = QTableWidgetItem(area)
        area_item.setTextAlignment(Qt.AlignCenter)
        area_item.setFlags(area_item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row, 2, area_item)

        # col3: 範囲
        range_item = QTableWidgetItem(range_str)
        range_item.setFlags(range_item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row, 3, range_item)

        # col4: 深さ
        depth_item = QTableWidgetItem(depth_str)
        depth_item.setFlags(depth_item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row, 4, depth_item)

        # col5: 除外理由
        reason_text = exclusion_reason if is_exclusion else ""
        reason_item = QTableWidgetItem(reason_text)
        if not is_exclusion:
            reason_item.setFlags(reason_item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row, 5, reason_item)

        if is_exclusion:
            for col in range(6):
                item = self.table.item(row, col)
                if item:
                    item.setBackground(EXCLUSION_BG)
                    item.setForeground(EXCLUSION_FG)

        self.table.scrollToBottom()
        self.table.blockSignals(False)

    def delete_row(self, db_id):
        db_id_int = int(db_id)
        for row in range(self.table.rowCount()):
            id_item = self.table.item(row, 0)
            if id_item and id_item.data(Qt.UserRole) == db_id_int:
                self.table.removeRow(row)
                break

    def select_by_db_id(self, db_id):
        self.table.blockSignals(True)
        if db_id < 0:
            self.table.clearSelection()
        else:
            for row in range(self.table.rowCount()):
                id_item = self.table.item(row, 0)
                if id_item and id_item.data(Qt.UserRole) == db_id:
                    self.table.selectRow(row)
                    break
        self.table.blockSignals(False)


class KiloListWidget(QWidget):
    """キロ程一覧（右サイドバー）"""

    kilo_selected = Signal(int)

    def __init__(self, parent=None, ui_font_family="Meiryo"):
        super().__init__(parent)
        self.setFixedWidth(140)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        label = QPushButton("キロ程一覧")
        label.setFont(QFont(ui_font_family, 10, QFont.Bold))
        label.setEnabled(False)
        label.setStyleSheet("QPushButton:disabled { color: white; background: #444; }")
        layout.addWidget(label)

        self.list_widget = QListWidget()
        self.list_widget.setFont(QFont(ui_font_family, 10))
        self.list_widget.currentRowChanged.connect(self._on_row_changed)
        layout.addWidget(self.list_widget)

    def set_kilos(self, kilo_list):
        self.list_widget.blockSignals(True)
        self.list_widget.clear()
        for kilo in kilo_list:
            self.list_widget.addItem(kilo)
        self.list_widget.blockSignals(False)

    def set_current_index(self, index):
        self.list_widget.blockSignals(True)
        self.list_widget.setCurrentRow(index)
        self.list_widget.blockSignals(False)

    def _on_row_changed(self, row):
        if row >= 0:
            self.kilo_selected.emit(row)
