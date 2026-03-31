"""Excel出力範囲選択ダイアログ（QTreeWidget + ドラッグ連続ON/OFF）"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QDialogButtonBox, QTreeWidget, QTreeWidgetItem, QCheckBox,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from core.calc_utils import parse_kilo, extract_line_type, strip_line_prefix, line_type_short


class _DragCheckTree(QTreeWidget):
    """マウスドラッグでチェックボックスを一括切替できるQTreeWidget。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._dragging = False
        self._check_state = None

    def mousePressEvent(self, event):
        item = self.itemAt(event.pos())
        if item and event.button() == Qt.MouseButton.LeftButton:
            rect = self.visualItemRect(item)
            indent = self.indentation() + 20
            if event.pos().x() < rect.left() + indent:
                current = item.checkState(0)
                if current == Qt.CheckState.Checked:
                    self._check_state = Qt.CheckState.Unchecked
                else:
                    self._check_state = Qt.CheckState.Checked
                item.setCheckState(0, self._check_state)
                self._dragging = True
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging and self._check_state is not None:
            item = self.itemAt(event.pos())
            if item:
                item.setCheckState(0, self._check_state)
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._dragging = False
        self._check_state = None
        super().mouseReleaseEvent(event)


class ExportRangeDialog(QDialog):
    """出力するキロ程をチェックボックスで選択するダイアログ"""

    def __init__(self, parent=None, sorted_kilos=None):
        super().__init__(parent)
        self.setWindowTitle("Excel出力範囲の選択")
        self.setMinimumWidth(300)
        self.setMinimumHeight(500)
        self._sorted_kilos = sorted_kilos or []

        font = QFont("Meiryo", 11)
        layout = QVBoxLayout(self)

        # 全選択/全解除
        btn_layout = QHBoxLayout()
        all_on_btn = QPushButton("全選択")
        all_on_btn.setFont(font)
        all_on_btn.clicked.connect(lambda: self._set_all(Qt.CheckState.Checked))
        btn_layout.addWidget(all_on_btn)

        all_off_btn = QPushButton("全解除")
        all_off_btn.setFont(font)
        all_off_btn.clicked.connect(lambda: self._set_all(Qt.CheckState.Unchecked))
        btn_layout.addWidget(all_off_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        hint = QLabel("※ ドラッグで連続ON/OFF可能")
        hint.setFont(QFont("Meiryo", 9))
        hint.setStyleSheet("color: gray;")
        layout.addWidget(hint)

        # ツリー（1列、チェックボックス付き）
        self._tree = _DragCheckTree()
        self._tree.setHeaderHidden(True)
        self._tree.setFont(font)
        self._tree.setIndentation(0)

        self._items = []
        for kilo in self._sorted_kilos:
            lt = extract_line_type(kilo)
            bare = strip_line_prefix(kilo)
            short = line_type_short(lt)
            display = f"{short}_{bare}" if short else bare
            item = QTreeWidgetItem([display])
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(0, Qt.CheckState.Checked)
            self._tree.addTopLevelItem(item)
            self._items.append(item)

        layout.addWidget(self._tree, stretch=1)

        # 個別ファイル出力オプション
        self._individual_check = QCheckBox("個別ファイルで出力する（キロ程ごとに1ファイル）")
        self._individual_check.setFont(font)
        layout.addWidget(self._individual_check)

        # ボタン
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _set_all(self, state):
        for item in self._items:
            item.setCheckState(0, state)

    def selected_kilos(self):
        return [
            kilo for kilo, item in zip(self._sorted_kilos, self._items)
            if item.checkState(0) == Qt.CheckState.Checked
        ]

    def filename_range(self):
        selected = self.selected_kilos()
        if not selected:
            return "", ""
        values = [(parse_kilo(k), k) for k in selected]
        values.sort(key=lambda x: x[0])
        return values[0][1], values[-1][1]

    def is_individual(self):
        return self._individual_check.isChecked()
