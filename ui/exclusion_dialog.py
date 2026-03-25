"""除外区間入力ダイアログ"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QDoubleSpinBox, QSpinBox, QDialogButtonBox, QMessageBox,
    QLineEdit, QCheckBox, QGroupBox,
)
from PySide6.QtGui import QFont

ALL_AREAS = ["左軌間外", "軌間内", "右軌間外"]


class ExclusionDialog(QDialog):
    """除外区間の開始・終了位置、管理番号、除外理由を入力するダイアログ"""

    def __init__(self, parent=None, area_name="", existing_zones=None,
                 suggested_number=1, existing_numbers=None):
        super().__init__(parent)
        self.setWindowTitle(f"除外区間の設定 - {area_name}")
        self.setMinimumWidth(420)
        self._area_name = area_name
        self._existing_zones = existing_zones or []
        self._existing_numbers = existing_numbers or set()

        font = QFont("Meiryo", 12)
        layout = QVBoxLayout(self)

        # 管理番号
        num_layout = QHBoxLayout()
        num_label = QLabel("管理番号:")
        num_label.setFont(font)
        num_layout.addWidget(num_label)

        self._number_spin = QSpinBox()
        self._number_spin.setFont(font)
        self._number_spin.setRange(1, 9999)
        self._number_spin.setValue(suggested_number)
        num_layout.addWidget(self._number_spin)
        layout.addLayout(num_layout)

        # 開始位置
        start_layout = QHBoxLayout()
        start_label = QLabel("開始位置 (m):")
        start_label.setFont(font)
        start_layout.addWidget(start_label)

        self._start_spin = QDoubleSpinBox()
        self._start_spin.setFont(font)
        self._start_spin.setRange(0.0, 20.0)
        self._start_spin.setSingleStep(0.1)
        self._start_spin.setDecimals(1)
        self._start_spin.setValue(0.0)
        start_layout.addWidget(self._start_spin)
        layout.addLayout(start_layout)

        # 終了位置
        end_layout = QHBoxLayout()
        end_label = QLabel("終了位置 (m):")
        end_label.setFont(font)
        end_layout.addWidget(end_label)

        self._end_spin = QDoubleSpinBox()
        self._end_spin.setFont(font)
        self._end_spin.setRange(0.0, 20.0)
        self._end_spin.setSingleStep(0.1)
        self._end_spin.setDecimals(1)
        self._end_spin.setValue(20.0)
        end_layout.addWidget(self._end_spin)
        layout.addLayout(end_layout)

        self._start_spin.valueChanged.connect(self._on_start_changed)

        # 除外理由
        reason_layout = QHBoxLayout()
        reason_label = QLabel("除外理由:")
        reason_label.setFont(font)
        reason_layout.addWidget(reason_label)

        self._reason_edit = QLineEdit()
        self._reason_edit.setFont(font)
        self._reason_edit.setText("設備等による反射")
        reason_layout.addWidget(self._reason_edit)
        layout.addLayout(reason_layout)

        # 他エリア同時入力
        other_areas = [a for a in ALL_AREAS if a != area_name]
        if other_areas:
            group = QGroupBox("他エリアにも同時入力")
            group.setFont(font)
            group_layout = QVBoxLayout(group)
            self._area_checks = {}
            for a in other_areas:
                cb = QCheckBox(f"{a} にも入力する")
                cb.setFont(font)
                group_layout.addWidget(cb)
                self._area_checks[a] = cb
            layout.addWidget(group)
        else:
            self._area_checks = {}

        # ボタン
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_start_changed(self, value):
        if self._end_spin.value() < value:
            self._end_spin.setValue(value)
        self._end_spin.setMinimum(value)

    def _validate_and_accept(self):
        start = self._start_spin.value()
        end = self._end_spin.value()
        number = self._number_spin.value()

        if end <= start:
            QMessageBox.warning(self, "入力エラー", "終了位置は開始位置より後にしてください。")
            return

        for ex_start, ex_end in self._existing_zones:
            if not (end <= ex_start or start >= ex_end):
                QMessageBox.warning(
                    self, "入力エラー",
                    f"既存の除外区間 ({ex_start:.1f}m ～ {ex_end:.1f}m) と重複しています。",
                )
                return

        if number in self._existing_numbers:
            QMessageBox.warning(
                self, "番号重複",
                f"管理番号 {number} は既に使用されています。別の番号を入力してください。",
            )
            return

        self.accept()

    def mgmt_number(self):
        return self._number_spin.value()

    def start_pos(self):
        return self._start_spin.value()

    def end_pos(self):
        return self._end_spin.value()

    def reason(self):
        return self._reason_edit.text().strip() or "設備等による反射"

    def additional_areas(self):
        """同時入力する追加エリアのリストを返す"""
        return [a for a, cb in self._area_checks.items() if cb.isChecked()]
