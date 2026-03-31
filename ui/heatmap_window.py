"""変状ヒートマップウィンドウ"""

import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QCheckBox,
    QPushButton, QFileDialog, QMessageBox, QDialog, QGroupBox,
    QRadioButton, QButtonGroup,
)
from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import (
    QPainter, QColor, QPen, QFont, QBrush, QImage,
    QWheelEvent, QMouseEvent, QPaintEvent,
)
from PIL import Image as PILImage

from core.calc_utils import parse_kilo, px_to_m_x, WORK_AREAS

AREA_NAMES = ["左軌間外", "軌間内", "右軌間外"]

CATEGORY_COLORS = {
    "ゆるみ": QColor(60, 120, 255, 180),
    "空洞": QColor(255, 60, 60, 180),
}

IMAGE_SPAN_M = 20.0

# ストリップのネイティブサイズ
STRIP_W_NATIVE = 844  # X_PX_MAX - X_PX_MIN
STRIP_H_NATIVE = 179  # 各エリアの高さ

# レイアウト定数
ROW_H = 180
EXCLUSION_H = 22
GAP_H = 12
LABEL_W = 120
SCALE_H = 45
TOP_MARGIN = 40


def _kilo_format(m):
    """メートル値をキロ程表記に変換"""
    km = int(m) // 1000
    rem = m % 1000
    return f"{km}k{rem:05.1f}m"


class HeatmapCanvas(QWidget):
    """ヒートマップ描画キャンバス"""

    # 表示フィルタ
    show_yurumi = True
    show_kudo = True
    show_exclusion = True
    show_gap = True
    show_waveform = True

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(900, 850)
        self.setMouseTracking(True)

        self._segments = []
        self._drawings = {}
        self._total_start = 0
        self._total_end = 0

        self._view_start = 0.0
        self._view_end = 1000.0

        self._last_pan_x = None
        self._mouse_m = None
        self._mouse_y = 0
        self._crosshair_label = None

        self._image_groups = {}
        self._sorted_kilos = []
        self._area_strips = {}  # {(kilo, area_name): QImage}
        self._native_ratio = False
        self._image_key = 'marked'  # 'marked' or 'unmarked'

    def _effective_row_h(self):
        """等倍モード時はネイティブ縦横比に基づく行高さを返す"""
        if not self._native_ratio or not self._segments:
            return ROW_H
        view_span = self._view_end - self._view_start
        if view_span <= 0:
            return ROW_H
        canvas_w = self.width() - LABEL_W
        px_per_20m = IMAGE_SPAN_M / view_span * canvas_w
        h = px_per_20m * (STRIP_H_NATIVE / STRIP_W_NATIVE)
        return max(30, min(h, 600))

    def set_data(self, image_groups, sorted_kilos, db):
        self._segments = []
        self._drawings = {}
        self._image_groups = image_groups
        self._sorted_kilos = sorted_kilos

        for kilo in sorted_kilos:
            base_m = parse_kilo(kilo)
            self._segments.append((base_m, base_m + IMAGE_SPAN_M, kilo))

            drawings = db.load_drawings(kilo)
            group = image_groups.get(kilo, {})
            direction = group.get('direction', '起点→終点')

            processed = []
            for d in drawings:
                area = d.get('area', '')
                category = d.get('category', 'ゆるみ')
                lx0, lx1 = d['lx0'], d['lx1']
                min_lx, max_lx = min(lx0, lx1), max(lx0, lx1)
                offset_start = px_to_m_x(min_lx)
                offset_end = px_to_m_x(max_lx)

                if direction == "起点→終点":
                    abs_start = base_m + offset_start
                    abs_end = base_m + offset_end
                else:
                    abs_start = base_m - IMAGE_SPAN_M + offset_start
                    abs_end = base_m - IMAGE_SPAN_M + offset_end

                processed.append({
                    'area': area,
                    'category': category,
                    'abs_start': min(abs_start, abs_end),
                    'abs_end': max(abs_start, abs_end),
                    'mgmt_number': d.get('mgmt_number'),
                })
            self._drawings[kilo] = processed

        if self._segments:
            self._segments.sort(key=lambda s: s[0])
            self._total_start = self._segments[0][0]
            self._total_end = self._segments[-1][1]
            self._view_start = self._total_start
            self._view_end = self._total_end

        self._reload_strips()
        self.update()

    def set_image_key(self, key):
        """表示画像を切り替え ('marked' or 'unmarked')"""
        if key != self._image_key:
            self._image_key = key
            self._reload_strips()
            self.update()

    def _reload_strips(self):
        """現在の _image_key に基づいてエリアストリップを再読み込み"""
        self._area_strips = {}
        for kilo in self._sorted_kilos:
            group = self._image_groups.get(kilo, {})
            img_path = group.get(self._image_key)
            if not img_path or not os.path.exists(img_path):
                continue
            direction = group.get('direction', '起点→終点')
            try:
                pil_img = PILImage.open(img_path)
                for area_name in AREA_NAMES:
                    area_def = WORK_AREAS[area_name]
                    strip = pil_img.crop((
                        area_def['x_min'], area_def['y_min'],
                        area_def['x_max'], area_def['y_max'],
                    ))
                    if direction == '終点→起点':
                        strip = strip.transpose(PILImage.Transpose.FLIP_LEFT_RIGHT)
                    strip_rgb = strip.convert('RGB')
                    raw = strip_rgb.tobytes('raw', 'RGB')
                    qimg = QImage(raw, strip_rgb.width, strip_rgb.height,
                                  strip_rgb.width * 3, QImage.Format_RGB888)
                    self._area_strips[(kilo, area_name)] = qimg.copy()
                pil_img.close()
            except Exception:
                pass

    def _m_to_x(self, m):
        w = self.width() - LABEL_W
        if self._view_end <= self._view_start:
            return LABEL_W
        return LABEL_W + (m - self._view_start) / (self._view_end - self._view_start) * w

    def _x_to_m(self, x):
        w = self.width() - LABEL_W
        if w <= 0:
            return self._view_start
        return self._view_start + (x - LABEL_W) / w * (self._view_end - self._view_start)

    def _content_height(self):
        rh = self._effective_row_h()
        return TOP_MARGIN + 4 * (rh + GAP_H) + SCALE_H

    def _y_offset(self):
        total = self._content_height()
        return max(0, (self.height() - total) // 2)

    def _row_y(self, row_idx):
        rh = self._effective_row_h()
        return self._y_offset() + TOP_MARGIN + row_idx * (rh + GAP_H)

    # ------------------------------------------------------------------
    # 描画
    # ------------------------------------------------------------------

    def paintEvent(self, event: QPaintEvent):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        rh = self._effective_row_h()

        painter.fillRect(0, 0, w, h, QColor(30, 30, 30))

        if not self._segments:
            painter.setPen(QColor("white"))
            painter.setFont(QFont("Meiryo", 14))
            painter.drawText(QRectF(0, 0, w, h), Qt.AlignCenter, "データがありません")
            painter.end()
            return

        row_configs = [
            ("全体", None),
            ("左軌間外", "左軌間外"),
            ("軌間内", "軌間内"),
            ("右軌間外", "右軌間外"),
        ]

        painter.setClipRect(QRectF(LABEL_W, 0, w - LABEL_W, h))

        self._draw_top_scale(painter, w)

        for row_idx, (_, area_filter) in enumerate(row_configs):
            y = self._row_y(row_idx)
            self._draw_row(painter, y, w, area_filter, rh)

        self._draw_scale(painter, w, rh)

        # クロスヘア
        if self._mouse_m is not None:
            cx = self._m_to_x(self._mouse_m)
            if LABEL_W <= cx <= w:
                painter.setPen(QPen(QColor(255, 255, 255, 150), 1, Qt.DashLine))
                painter.drawLine(QPointF(cx, 0), QPointF(cx, h))

                tip_text = _kilo_format(self._mouse_m)
                tip_font = QFont("Consolas", 11)
                tip_font.setBold(True)
                painter.setFont(tip_font)
                fm = painter.fontMetrics()
                tw = fm.horizontalAdvance(tip_text) + 12
                th = fm.height() + 6
                tx = cx + 12
                ty = self._mouse_y - th // 2
                if tx + tw > w:
                    tx = cx - tw - 8
                painter.fillRect(QRectF(tx, ty, tw, th), QColor(20, 20, 20, 220))
                painter.setPen(QPen(QColor(0, 255, 136), 1))
                painter.drawRect(QRectF(tx, ty, tw, th))
                painter.setPen(QColor(0, 255, 136))
                painter.drawText(QRectF(tx, ty, tw, th), Qt.AlignCenter, tip_text)

        painter.setClipping(False)

        # 左ラベル
        painter.fillRect(QRectF(0, 0, LABEL_W, h), QColor(30, 30, 30))
        label_font = QFont("Meiryo", 13)
        label_font.setBold(True)
        painter.setFont(label_font)
        painter.setPen(QColor("white"))

        for row_idx, (label, _) in enumerate(row_configs):
            y = self._row_y(row_idx)
            painter.drawText(
                QRectF(4, y, LABEL_W - 8, rh),
                Qt.AlignRight | Qt.AlignVCenter, label,
            )

        painter.end()

    def _draw_row(self, painter, y, w, area_filter, rh):
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(45, 45, 50))
        painter.drawRect(QRectF(LABEL_W, y, w - LABEL_W, rh))

        excl_h = min(EXCLUSION_H, int(rh * 0.15))

        for seg_start, seg_end, kilo in self._segments:
            x0 = self._m_to_x(seg_start)
            x1 = self._m_to_x(seg_end)
            if x1 < LABEL_W or x0 > w:
                continue
            painter.fillRect(QRectF(x0, y, x1 - x0, rh), QColor(55, 55, 65))

            if area_filter and self.show_waveform:
                strip = self._area_strips.get((kilo, area_filter))
                if strip is not None:
                    target = QRectF(x0, y, x1 - x0, rh)
                    source = QRectF(0, 0, strip.width(), strip.height())
                    painter.drawImage(target, strip, source)

            painter.setPen(QPen(QColor(70, 70, 80), 1))
            painter.setBrush(Qt.NoBrush)
            painter.drawRect(QRectF(x0, y, x1 - x0, rh))

        if self.show_gap:
            for i in range(len(self._segments) - 1):
                _, prev_end, _ = self._segments[i]
                next_start, _, _ = self._segments[i + 1]
                if next_start > prev_end + 0.1:
                    gx0 = self._m_to_x(prev_end)
                    gx1 = self._m_to_x(next_start)
                    painter.fillRect(QRectF(gx0, y, gx1 - gx0, rh), QColor(80, 60, 20, 60))
                    painter.setPen(QPen(QColor("orange"), 2, Qt.DashLine))
                    painter.drawLine(QPointF(gx0, y), QPointF(gx0, y + rh))
                    painter.drawLine(QPointF(gx1, y), QPointF(gx1, y + rh))

        if self.show_exclusion:
            for kilo_str in self._drawings:
                for d in self._drawings[kilo_str]:
                    if d['category'] != '除外区間':
                        continue
                    if area_filter and d['area'] != area_filter:
                        continue
                    x0 = self._m_to_x(d['abs_start'])
                    x1 = self._m_to_x(d['abs_end'])
                    if x1 < LABEL_W or x0 > w:
                        continue
                    painter.fillRect(
                        QRectF(x0, y, max(x1 - x0, 2), excl_h),
                        QColor(255, 100, 100, 100),
                    )
                    painter.setPen(QPen(QColor(255, 80, 80, 200), 1))
                    painter.setBrush(Qt.NoBrush)
                    painter.drawRect(QRectF(x0, y, max(x1 - x0, 2), excl_h))

        bar_y = y + excl_h + 3
        bar_h = rh - excl_h - 6

        for kilo_str in self._drawings:
            for d in self._drawings[kilo_str]:
                cat = d['category']
                if cat == '除外区間':
                    continue
                if cat == 'ゆるみ' and not self.show_yurumi:
                    continue
                if cat == '空洞' and not self.show_kudo:
                    continue
                if area_filter and d['area'] != area_filter:
                    continue

                x0 = self._m_to_x(d['abs_start'])
                x1 = self._m_to_x(d['abs_end'])
                if x1 < LABEL_W or x0 > w:
                    continue

                color = CATEGORY_COLORS.get(cat, QColor(200, 200, 200, 120))
                painter.fillRect(QRectF(x0, bar_y, max(x1 - x0, 3), bar_h), color)
                painter.setPen(QPen(color.darker(130), 1))
                painter.setBrush(Qt.NoBrush)
                painter.drawRect(QRectF(x0, bar_y, max(x1 - x0, 3), bar_h))

    def _calc_tick_step(self):
        view_span = self._view_end - self._view_start
        if view_span <= 0:
            return 100
        for step in [10, 20, 50, 100, 200, 500, 1000, 2000, 5000]:
            if view_span / step <= 20:
                return step
        return 5000

    def _draw_top_scale(self, painter, w):
        view_span = self._view_end - self._view_start
        if view_span <= 0:
            return
        step = self._calc_tick_step()
        font = QFont("Meiryo", 9)
        painter.setFont(font)

        top_y = self._y_offset() + TOP_MARGIN
        start_tick = int(self._view_start / step) * step
        m = start_tick
        while m <= self._view_end:
            x = self._m_to_x(m)
            if x >= LABEL_W:
                painter.setPen(QColor(150, 150, 150))
                painter.drawLine(QPointF(x, top_y - 2), QPointF(x, top_y - 10))
                km = int(m) // 1000
                rem = m % 1000
                label = f"{km}k{rem:05.1f}"
                painter.drawText(QRectF(x - 45, top_y - 28, 90, 16), Qt.AlignCenter, label)
            m += step

    def _draw_scale(self, painter, w, rh):
        scale_y = self._y_offset() + TOP_MARGIN + 4 * (rh + GAP_H)
        view_span = self._view_end - self._view_start
        if view_span <= 0:
            return
        step = self._calc_tick_step()

        font = QFont("Meiryo", 10)
        painter.setFont(font)

        start_tick = int(self._view_start / step) * step
        m = start_tick
        while m <= self._view_end:
            x = self._m_to_x(m)
            if x >= LABEL_W:
                painter.setPen(QPen(QColor(80, 80, 80), 1))
                painter.drawLine(QPointF(x, self._row_y(0)), QPointF(x, scale_y))
                painter.setPen(QColor(180, 180, 180))
                painter.drawLine(QPointF(x, scale_y), QPointF(x, scale_y + 10))
                km = int(m) // 1000
                rem = m % 1000
                label = f"{km}k{rem:05.1f}"
                painter.drawText(QRectF(x - 45, scale_y + 12, 90, 18), Qt.AlignCenter, label)
            m += step

    # ------------------------------------------------------------------
    # ズーム・パン・ホバー
    # ------------------------------------------------------------------

    def wheelEvent(self, event: QWheelEvent):
        if not self._segments:
            return
        pos_m = self._x_to_m(event.position().x())
        factor = 0.8 if event.angleDelta().y() > 0 else 1.25
        span = self._view_end - self._view_start
        new_span = max(20.0, min(span * factor, (self._total_end - self._total_start) * 1.1))
        ratio = (pos_m - self._view_start) / span if span > 0 else 0.5
        self._view_start = pos_m - new_span * ratio
        self._view_end = pos_m + new_span * (1 - ratio)
        self.update()

    def mousePressEvent(self, event: QMouseEvent):
        if event.button() in (Qt.LeftButton, Qt.MiddleButton, Qt.RightButton):
            self._last_pan_x = event.position().x()

    def mouseMoveEvent(self, event: QMouseEvent):
        x = event.position().x()
        self._mouse_y = event.position().y()
        if x >= LABEL_W:
            self._mouse_m = self._x_to_m(x)
            if self._crosshair_label:
                self._crosshair_label.setText(f"  {_kilo_format(self._mouse_m)}  ")
        else:
            self._mouse_m = None
            if self._crosshair_label:
                self._crosshair_label.setText("")

        if self._last_pan_x is not None:
            dx = x - self._last_pan_x
            self._last_pan_x = x
            w = self.width() - LABEL_W
            if w > 0:
                dm = -dx / w * (self._view_end - self._view_start)
                self._view_start += dm
                self._view_end += dm

        self.update()

    def mouseReleaseEvent(self, event: QMouseEvent):
        self._last_pan_x = None

    def leaveEvent(self, event):
        self._mouse_m = None
        if self._crosshair_label:
            self._crosshair_label.setText("")
        self.update()


# ------------------------------------------------------------------
# 出力レイヤー選択ダイアログ
# ------------------------------------------------------------------

class WaveformExportDialog(QDialog):
    """波形Excel出力のレイヤー選択ダイアログ"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("波形Excel出力設定")
        self.setMinimumWidth(300)
        layout = QVBoxLayout(self)
        ctrl_font = QFont("Meiryo", 12)

        # 画像種別選択
        img_group = QGroupBox("波形画像")
        img_group.setFont(ctrl_font)
        img_layout = QVBoxLayout(img_group)
        self._img_btn_group = QButtonGroup(self)
        self._rb_unmarked = QRadioButton("マーキングなし")
        self._rb_unmarked.setFont(ctrl_font)
        self._rb_unmarked.setChecked(True)
        self._rb_marked = QRadioButton("マーキングあり")
        self._rb_marked.setFont(ctrl_font)
        self._img_btn_group.addButton(self._rb_unmarked)
        self._img_btn_group.addButton(self._rb_marked)
        img_layout.addWidget(self._rb_unmarked)
        img_layout.addWidget(self._rb_marked)
        layout.addWidget(img_group)

        # エリア選択
        area_group = QGroupBox("出力エリア")
        area_group.setFont(ctrl_font)
        area_layout = QVBoxLayout(area_group)
        self._cb_left = QCheckBox("左軌間外")
        self._cb_left.setFont(ctrl_font)
        self._cb_left.setChecked(True)
        area_layout.addWidget(self._cb_left)
        self._cb_inner = QCheckBox("軌間内")
        self._cb_inner.setFont(ctrl_font)
        self._cb_inner.setChecked(True)
        area_layout.addWidget(self._cb_inner)
        self._cb_right = QCheckBox("右軌間外")
        self._cb_right.setFont(ctrl_font)
        self._cb_right.setChecked(True)
        area_layout.addWidget(self._cb_right)
        layout.addWidget(area_group)

        # オーバーレイ選択
        overlay_group = QGroupBox("オーバーレイ")
        overlay_group.setFont(ctrl_font)
        overlay_layout = QVBoxLayout(overlay_group)
        self._cb_yurumi = QCheckBox("ゆるみ")
        self._cb_yurumi.setFont(ctrl_font)
        self._cb_yurumi.setChecked(True)
        self._cb_yurumi.setStyleSheet("color: #3c78ff;")
        overlay_layout.addWidget(self._cb_yurumi)
        self._cb_kudo = QCheckBox("空洞")
        self._cb_kudo.setFont(ctrl_font)
        self._cb_kudo.setChecked(True)
        self._cb_kudo.setStyleSheet("color: #ff3c3c;")
        overlay_layout.addWidget(self._cb_kudo)
        self._cb_exclusion = QCheckBox("除外区間")
        self._cb_exclusion.setFont(ctrl_font)
        self._cb_exclusion.setChecked(True)
        self._cb_exclusion.setStyleSheet("color: #ff6464;")
        overlay_layout.addWidget(self._cb_exclusion)
        layout.addWidget(overlay_group)

        # ボタン
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("出力")
        ok_btn.setFont(QFont("Meiryo", 12, QFont.Bold))
        ok_btn.setFixedHeight(36)
        ok_btn.setStyleSheet(
            "QPushButton { background-color: #2ea043; color: white; border-radius: 4px; }"
            "QPushButton:hover { background-color: #238636; }"
        )
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.setFont(QFont("Meiryo", 12))
        cancel_btn.setFixedHeight(36)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

    def image_key(self):
        return 'unmarked' if self._rb_unmarked.isChecked() else 'marked'

    def selected_areas(self):
        areas = []
        if self._cb_left.isChecked():
            areas.append("左軌間外")
        if self._cb_inner.isChecked():
            areas.append("軌間内")
        if self._cb_right.isChecked():
            areas.append("右軌間外")
        return areas

    def selected_overlays(self):
        overlays = []
        if self._cb_yurumi.isChecked():
            overlays.append("ゆるみ")
        if self._cb_kudo.isChecked():
            overlays.append("空洞")
        if self._cb_exclusion.isChecked():
            overlays.append("除外区間")
        return overlays


# ------------------------------------------------------------------
# メインウィンドウ
# ------------------------------------------------------------------

class HeatmapWindow(QMainWindow):
    """変状ヒートマップウィンドウ"""

    def __init__(self, parent=None, image_groups=None, sorted_kilos=None, db=None,
                 parent_folder=""):
        super().__init__(parent)
        self.setWindowTitle("変状ヒートマップ")
        self.resize(1500, 950)

        self._parent_folder = parent_folder

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(5, 5, 5, 5)

        # --- 上段: 凡例 + フィルタ + キロ程表示 ---
        top_layout = QHBoxLayout()
        ctrl_font = QFont("Meiryo", 11)

        # フィルタチェックボックス
        self._cb_yurumi = QCheckBox("ゆるみ")
        self._cb_yurumi.setFont(ctrl_font)
        self._cb_yurumi.setChecked(True)
        self._cb_yurumi.setStyleSheet("color: #3c78ff;")
        top_layout.addWidget(self._cb_yurumi)

        self._cb_kudo = QCheckBox("空洞")
        self._cb_kudo.setFont(ctrl_font)
        self._cb_kudo.setChecked(True)
        self._cb_kudo.setStyleSheet("color: #ff3c3c;")
        top_layout.addWidget(self._cb_kudo)

        self._cb_exclusion = QCheckBox("除外区間")
        self._cb_exclusion.setFont(ctrl_font)
        self._cb_exclusion.setChecked(True)
        self._cb_exclusion.setStyleSheet("color: #ff6464;")
        top_layout.addWidget(self._cb_exclusion)

        self._cb_gap = QCheckBox("途切れ")
        self._cb_gap.setFont(ctrl_font)
        self._cb_gap.setChecked(True)
        self._cb_gap.setStyleSheet("color: orange;")
        top_layout.addWidget(self._cb_gap)

        self._cb_waveform = QCheckBox("波形画像")
        self._cb_waveform.setFont(ctrl_font)
        self._cb_waveform.setChecked(True)
        self._cb_waveform.setStyleSheet("color: #00ccaa;")
        top_layout.addWidget(self._cb_waveform)

        # 画像種別切替
        sep = QLabel("｜")
        sep.setFont(ctrl_font)
        top_layout.addWidget(sep)

        img_style = (
            "QPushButton { border: 2px solid #888; border-radius: 4px; "
            "padding: 3px 8px; }"
            "QPushButton:checked { background-color: #555; color: white; "
            "border-color: white; }"
        )
        self._img_btn_group = QButtonGroup(self)
        self._img_btn_group.setExclusive(True)

        self._marked_btn = QPushButton("マーキングあり")
        self._marked_btn.setFont(ctrl_font)
        self._marked_btn.setCheckable(True)
        self._marked_btn.setChecked(True)
        self._marked_btn.setStyleSheet(img_style)
        self._img_btn_group.addButton(self._marked_btn, 0)
        top_layout.addWidget(self._marked_btn)

        self._unmarked_btn = QPushButton("マーキングなし")
        self._unmarked_btn.setFont(ctrl_font)
        self._unmarked_btn.setCheckable(True)
        self._unmarked_btn.setStyleSheet(img_style)
        self._img_btn_group.addButton(self._unmarked_btn, 1)
        top_layout.addWidget(self._unmarked_btn)

        self._img_btn_group.idClicked.connect(self._change_image_source)

        # 等倍ボタン
        native_style = (
            "QPushButton { border: 2px solid #888; border-radius: 4px; "
            "padding: 3px 10px; font-weight: bold; }"
            "QPushButton:checked { background-color: #00ccaa; color: #111; "
            "border-color: #00ccaa; }"
        )
        self._native_btn = QPushButton("等倍")
        self._native_btn.setFont(ctrl_font)
        self._native_btn.setCheckable(True)
        self._native_btn.setChecked(False)
        self._native_btn.setFixedSize(60, 30)
        self._native_btn.setStyleSheet(native_style)
        self._native_btn.toggled.connect(self._toggle_native_ratio)
        top_layout.addWidget(self._native_btn)

        top_layout.addStretch()

        # Excel出力ボタン
        self._export_btn = QPushButton("連続波形画像Excel出力")
        self._export_btn.setFont(QFont("Meiryo", 11, QFont.Bold))
        self._export_btn.setFixedSize(190, 34)
        self._export_btn.setStyleSheet(
            "QPushButton { background-color: #2ea043; color: white; border-radius: 4px; }"
            "QPushButton:hover { background-color: #238636; }"
        )
        self._export_btn.clicked.connect(self._export_waveform_excel)
        top_layout.addWidget(self._export_btn)

        # キロ程リアルタイム表示
        self._pos_label = QLabel("")
        pos_font = QFont("Consolas", 14)
        pos_font.setBold(True)
        self._pos_label.setFont(pos_font)
        self._pos_label.setStyleSheet(
            "color: #00ff88; background: #222; padding: 2px 8px; border-radius: 3px;"
        )
        self._pos_label.setMinimumWidth(180)
        self._pos_label.setAlignment(Qt.AlignCenter)
        top_layout.addWidget(self._pos_label)

        hint = QLabel("ホイール: ズーム　ドラッグ: スクロール")
        hint.setFont(QFont("Meiryo", 10))
        hint.setStyleSheet("color: gray;")
        top_layout.addWidget(hint)

        layout.addLayout(top_layout)

        # --- キャンバス ---
        self._canvas = HeatmapCanvas()
        self._canvas._crosshair_label = self._pos_label
        layout.addWidget(self._canvas, stretch=1)

        # フィルタ接続
        self._cb_yurumi.toggled.connect(self._update_filter)
        self._cb_kudo.toggled.connect(self._update_filter)
        self._cb_exclusion.toggled.connect(self._update_filter)
        self._cb_gap.toggled.connect(self._update_filter)
        self._cb_waveform.toggled.connect(self._update_filter)

        self._image_groups = image_groups or {}
        self._sorted_kilos = sorted_kilos or []
        self._db = db

        if image_groups and sorted_kilos and db:
            self._canvas.set_data(image_groups, sorted_kilos, db)

    def _update_filter(self):
        self._canvas.show_yurumi = self._cb_yurumi.isChecked()
        self._canvas.show_kudo = self._cb_kudo.isChecked()
        self._canvas.show_exclusion = self._cb_exclusion.isChecked()
        self._canvas.show_gap = self._cb_gap.isChecked()
        self._canvas.show_waveform = self._cb_waveform.isChecked()
        self._canvas.update()

    def _change_image_source(self, btn_id):
        key = 'marked' if btn_id == 0 else 'unmarked'
        self._canvas.set_image_key(key)

    def _toggle_native_ratio(self, checked):
        self._canvas._native_ratio = checked
        self._canvas.update()

    def _export_waveform_excel(self):
        """連続波形画像をExcelに出力（レイヤー選択付き）"""
        if not self._sorted_kilos:
            QMessageBox.warning(self, "データなし", "出力するデータがありません。")
            return

        dialog = WaveformExportDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        selected_areas = dialog.selected_areas()
        selected_overlays = dialog.selected_overlays()
        image_key = dialog.image_key()
        if not selected_areas:
            QMessageBox.warning(self, "選択なし", "出力するエリアが選択されていません。")
            return

        # 作業フォルダを初期表示
        default_name = "連続波形画像.xlsx"
        default_path = (
            os.path.join(self._parent_folder, default_name)
            if self._parent_folder else default_name
        )
        output_path, _ = QFileDialog.getSaveFileName(
            self, "波形Excel出力先を選択", default_path, "Excel ファイル (*.xlsx)",
        )
        if not output_path:
            return
        try:
            from data.waveform_exporter import WaveformExcelExporter
            exporter = WaveformExcelExporter(
                self._image_groups, self._sorted_kilos, self._db,
            )
            exporter.export(
                output_path,
                areas=selected_areas,
                overlays=selected_overlays,
                image_key=image_key,
            )
            QMessageBox.information(self, "完了", f"波形Excelを保存しました:\n{output_path}")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"波形Excel出力に失敗しました:\n{e}")
