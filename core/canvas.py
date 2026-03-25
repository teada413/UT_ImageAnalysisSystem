"""描画キャンバスクラス (PySide6)"""

import os
from PySide6.QtWidgets import QWidget, QSizePolicy
from PySide6.QtCore import Qt, QRectF, QPointF, QTimer
from PySide6.QtGui import (
    QPainter, QPixmap, QColor, QPen, QFont, QFontMetrics,
    QWheelEvent, QMouseEvent, QPaintEvent, QResizeEvent, QCursor,
)

from core.calc_utils import (
    WORK_AREAS, IMAGE_W, IMAGE_H, IMAGE_H_TRIMMED, CANVAS_W, CANVAS_H, INITIAL_ZOOM,
    px_to_m_x, px_to_m_y, calc_location_string, circled_number,
)

_RESIZE_DELAY = 50
_HANDLE_SIZE = 4  # バウンディングボックスハンドルの半径（ディスプレイpx）

# 種別ごとの色定義
CATEGORY_COLORS = {
    "ゆるみ": QColor("blue"),
    "空洞": QColor("red"),
    "除外区間": QColor("red"),
}
EXCLUSION_FILL = QColor(255, 0, 0, 77)  # 赤30%透過

# ハンドル位置定数
_H_TL, _H_T, _H_TR = 0, 1, 2
_H_L, _H_R = 3, 4
_H_BL, _H_B, _H_BR = 5, 6, 7


class DrawingCanvas(QWidget):
    def __init__(self, parent=None, on_draw_callback=None):
        super().__init__(parent)
        self.on_draw_callback = on_draw_callback
        self.on_exclusion_click_callback = None
        self.on_drawing_modified_callback = None
        self.on_selection_changed_callback = None  # 選択変更通知
        self.setMinimumSize(400, 300)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.StrongFocus)

        self._pixmap = None
        self.twin = None

        self.base_kilo_m = 0
        self.direction = "起点→終点"

        self.zoom = INITIAL_ZOOM
        self.pan_x = 0.0
        self.pan_y = 0.0
        self._last_pan_pos = None

        self.drawings = []
        self.current_drawing = None
        self.draw_mode = "rectangle"
        self.draw_category = "ゆるみ"

        # 編集モード: "draw" or "move"
        self.edit_mode = "draw"
        self._selected_idx = -1       # 選択中の図形index
        self._drag_handle = -1        # ドラッグ中のハンドル (-1=移動, 0-7=リサイズ)
        self._drag_start_lx = 0.0
        self._drag_start_ly = 0.0
        self._drag_orig = None        # ドラッグ開始時の座標コピー

        self._placeholder_text = "画像がありません"
        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._on_resize)

    # ------------------------------------------------------------------
    # サイズ取得・リサイズ
    # ------------------------------------------------------------------

    def _canvas_size(self):
        w, h = self.width(), self.height()
        if w <= 1 or h <= 1:
            return CANVAS_W, CANVAS_H
        return w, h

    def resizeEvent(self, event: QResizeEvent):
        self._resize_timer.start(_RESIZE_DELAY)
        super().resizeEvent(event)

    def _on_resize(self):
        if self._pixmap:
            fit = self._calc_fit_zoom()
            if self.zoom < fit:
                self.zoom = fit
            self.clamp_pan()
        self.update()

    # ------------------------------------------------------------------
    # 表示
    # ------------------------------------------------------------------

    def display_text(self, text, subtitle=""):
        self._placeholder_text = f"{subtitle}\n{text}" if subtitle else text
        self._pixmap = None
        self.update()

    def _calc_fit_zoom(self):
        cw, ch = self._canvas_size()
        zoom_x = cw / IMAGE_W
        zoom_y = ch / IMAGE_H_TRIMMED
        return min(zoom_x, zoom_y)

    def set_image(self, image_path, base_kilo_m=0, direction="起点→終点"):
        self.drawings = []
        self.current_drawing = None
        self._selected_idx = -1
        self.pan_x = 0.0
        self.pan_y = 0.0
        self.base_kilo_m = base_kilo_m
        self.direction = direction

        if image_path and os.path.exists(image_path):
            pm = QPixmap(image_path)
            if pm.isNull():
                self._pixmap = None
                self.display_text("読込エラー")
            else:
                self._pixmap = pm
                self.zoom = self._calc_fit_zoom()
        else:
            self._pixmap = None
            self.display_text("画像なし")
        self.update()

    def set_drawings(self, drawings_list):
        self.drawings = [d.copy() for d in drawings_list]
        self._selected_idx = -1
        self.update()

    def remove_drawing(self, db_id):
        self.drawings = [d for d in self.drawings if d.get('db_id') != db_id]
        self._selected_idx = -1
        self.update()

    # ------------------------------------------------------------------
    # 座標変換
    # ------------------------------------------------------------------

    def logical_to_display(self, lx, ly):
        return (lx - self.pan_x) * self.zoom, (ly - self.pan_y) * self.zoom

    def display_to_logical(self, dx, dy):
        return (dx / self.zoom) + self.pan_x, (dy / self.zoom) + self.pan_y

    def clamp_logical(self, lx, ly, area_name):
        b = WORK_AREAS[area_name]
        return max(b["x_min"], min(b["x_max"], lx)), max(b["y_min"], min(b["y_max"], ly))

    def clamp_pan(self):
        cw, ch = self._canvas_size()
        view_w, view_h = cw / self.zoom, ch / self.zoom
        max_pan_x = max(0, IMAGE_W - view_w)
        max_pan_y = max(0, IMAGE_H_TRIMMED - view_h)
        self.pan_x = max(0, min(self.pan_x, max_pan_x))
        self.pan_y = max(0, min(self.pan_y, max_pan_y))

    # ------------------------------------------------------------------
    # エリア判定
    # ------------------------------------------------------------------

    def get_area_at(self, lx, ly):
        for name, b in WORK_AREAS.items():
            if b["x_min"] <= lx <= b["x_max"] and b["y_min"] <= ly <= b["y_max"]:
                return name
        return None

    # ------------------------------------------------------------------
    # ヒットテスト (移動モード用)
    # ------------------------------------------------------------------

    def _get_drawing_rect(self, d):
        """drawing dict から正規化された (min_lx, min_ly, max_lx, max_ly) を返す"""
        return min(d['lx0'], d['lx1']), min(d['ly0'], d['ly1']), max(d['lx0'], d['lx1']), max(d['ly0'], d['ly1'])

    def _hit_test_drawing(self, lx, ly, d, tolerance=5):
        """論理座標が図形上にあるか"""
        x0, y0, x1, y1 = self._get_drawing_rect(d)
        # 矩形エリア内 + マージン
        return (x0 - tolerance <= lx <= x1 + tolerance and
                y0 - tolerance <= ly <= y1 + tolerance)

    def _find_drawing_at(self, lx, ly):
        """論理座標にある図形のインデックスを返す（上から順に検索）"""
        for i in range(len(self.drawings) - 1, -1, -1):
            if self._hit_test_drawing(lx, ly, self.drawings[i]):
                return i
        return -1

    def _get_handle_positions(self, d):
        """選択図形の8つのハンドル位置を論理座標で返す"""
        x0, y0, x1, y1 = self._get_drawing_rect(d)
        mx, my = (x0 + x1) / 2, (y0 + y1) / 2
        return [
            (x0, y0), (mx, y0), (x1, y0),  # TL, T, TR
            (x0, my), (x1, my),             # L, R
            (x0, y1), (mx, y1), (x1, y1),  # BL, B, BR
        ]

    def _hit_test_handle(self, dx, dy, d):
        """ディスプレイ座標がどのハンドルに当たっているか (-1=なし)"""
        handles = self._get_handle_positions(d)
        for i, (hx, hy) in enumerate(handles):
            hdx, hdy = self.logical_to_display(hx, hy)
            if abs(dx - hdx) <= _HANDLE_SIZE and abs(dy - hdy) <= _HANDLE_SIZE:
                return i
        return -1

    # ------------------------------------------------------------------
    # 描画 (paintEvent)
    # ------------------------------------------------------------------

    def paintEvent(self, event: QPaintEvent):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        cw, ch = self._canvas_size()

        painter.fillRect(0, 0, cw, ch, QColor(25, 25, 25))

        if not self._pixmap:
            painter.setPen(QColor("white"))
            painter.setFont(QFont("Meiryo", 16))
            painter.drawText(QRectF(0, 0, cw, ch), Qt.AlignCenter, self._placeholder_text)
            painter.end()
            return

        # 画像描画
        view_w = cw / self.zoom
        view_h = ch / self.zoom
        actual_w = min(view_w, IMAGE_W - self.pan_x)
        actual_h = min(view_h, IMAGE_H_TRIMMED - self.pan_y)
        src_rect = QRectF(self.pan_x, self.pan_y, actual_w, actual_h)
        dst_rect = QRectF(0, 0, actual_w * self.zoom, actual_h * self.zoom)
        painter.drawPixmap(dst_rect, self._pixmap, src_rect)

        # ガイド線
        pen_guide = QPen(QColor("cyan"), 1, Qt.DashLine)
        painter.setPen(pen_guide)
        painter.setBrush(Qt.NoBrush)
        for b in WORK_AREAS.values():
            x0, y0 = self.logical_to_display(b["x_min"], b["y_min"])
            x1, y1 = self.logical_to_display(b["x_max"], b["y_max"])
            painter.drawRect(QRectF(QPointF(x0, y0), QPointF(x1, y1)))

        # 図形
        all_drawings = self.drawings + ([self.current_drawing] if self.current_drawing else [])
        for d in all_drawings:
            self._paint_drawing(painter, d)

        # 選択中のバウンディングボックス
        if self.edit_mode == "move" and 0 <= self._selected_idx < len(self.drawings):
            self._paint_bounding_box(painter, self.drawings[self._selected_idx])

        painter.end()

    def _paint_drawing(self, painter, d):
        dx0, dy0 = self.logical_to_display(d['lx0'], d['ly0'])
        dx1, dy1 = self.logical_to_display(d['lx1'], d['ly1'])
        category = d.get('category', 'ゆるみ')
        color = CATEGORY_COLORS.get(category, QColor("yellow"))

        rect = QRectF(QPointF(dx0, dy0), QPointF(dx1, dy1))

        if category == "除外区間":
            painter.setPen(QPen(color, 2))
            painter.setBrush(EXCLUSION_FILL)
            painter.drawRect(rect)
        else:
            painter.setPen(QPen(color, 2))
            painter.setBrush(Qt.NoBrush)
            if d['type'] == "rectangle":
                painter.drawRect(rect)
            else:
                painter.drawEllipse(rect)

        # 管理番号ラベル
        mgmt = d.get('mgmt_number')
        if mgmt and category != "除外区間":
            num_text = circled_number(mgmt)
            font = QFont("Meiryo", 11)
            font.setBold(True)
            painter.setFont(font)
            fm = QFontMetrics(font)
            num_rect = fm.boundingRect(num_text)
            nx = max(dx0, dx1) + 3
            ny = min(dy0, dy1) - 2
            num_rect.moveTopLeft(QPointF(nx, ny).toPoint())
            bg = num_rect.adjusted(-2, -1, 2, 1)
            painter.fillRect(bg, QColor(0, 0, 0, 180))
            painter.setPen(QColor("white"))
            painter.drawText(num_rect, Qt.AlignCenter, num_text)

    def _paint_bounding_box(self, painter, d):
        """選択図形のバウンディングボックスとハンドルを描画"""
        handles = self._get_handle_positions(d)
        # 破線枠
        pen = QPen(QColor("white"), 1, Qt.DashLine)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        x0, y0, x1, y1 = self._get_drawing_rect(d)
        dx0, dy0 = self.logical_to_display(x0, y0)
        dx1, dy1 = self.logical_to_display(x1, y1)
        painter.drawRect(QRectF(QPointF(dx0, dy0), QPointF(dx1, dy1)))

        # ハンドル
        painter.setPen(QPen(QColor("white"), 1))
        painter.setBrush(QColor(100, 100, 255))
        hs = _HANDLE_SIZE
        for hx, hy in handles:
            hdx, hdy = self.logical_to_display(hx, hy)
            painter.drawRect(QRectF(hdx - hs, hdy - hs, hs * 2, hs * 2))

    # ------------------------------------------------------------------
    # マウスイベント
    # ------------------------------------------------------------------

    def mousePressEvent(self, event: QMouseEvent):
        if not self._pixmap:
            return

        if event.button() in (Qt.MiddleButton, Qt.RightButton):
            self._last_pan_pos = event.position()
            return

        if event.button() != Qt.LeftButton:
            return

        if self.edit_mode == "move":
            self._move_mode_press(event)
        else:
            self._draw_mode_press(event)

    def mouseMoveEvent(self, event: QMouseEvent):
        if not self._pixmap:
            return

        # パン操作
        if self._last_pan_pos is not None:
            dx = event.position().x() - self._last_pan_pos.x()
            dy = event.position().y() - self._last_pan_pos.y()
            self._last_pan_pos = event.position()
            self.pan_x -= dx / self.zoom
            self.pan_y -= dy / self.zoom
            self.clamp_pan()
            self.update()
            if self.twin:
                self.twin.sync_pan(self.pan_x, self.pan_y)
            return

        if self.edit_mode == "move":
            self._move_mode_move(event)
        else:
            self._draw_mode_move(event)

    def mouseReleaseEvent(self, event: QMouseEvent):
        if event.button() in (Qt.MiddleButton, Qt.RightButton):
            self._last_pan_pos = None
            return

        if event.button() != Qt.LeftButton:
            return

        if self.edit_mode == "move":
            self._move_mode_release(event)
        else:
            self._draw_mode_release(event)

    # ------------------------------------------------------------------
    # 描画モード マウスイベント
    # ------------------------------------------------------------------

    def _draw_mode_press(self, event):
        lx, ly = self.display_to_logical(event.position().x(), event.position().y())
        area = self.get_area_at(lx, ly)
        if not area:
            return

        if self.draw_category == "除外区間":
            if self.on_exclusion_click_callback:
                self.on_exclusion_click_callback(area)
            return

        lx, ly = self.clamp_logical(lx, ly, area)
        self.current_drawing = {
            'type': self.draw_mode, 'category': self.draw_category,
            'area': area,
            'lx0': lx, 'ly0': ly, 'lx1': lx, 'ly1': ly,
        }
        self.update()
        if self.twin:
            self.twin.sync_press(self.current_drawing)

    def _draw_mode_move(self, event):
        if self.current_drawing:
            lx, ly = self.display_to_logical(event.position().x(), event.position().y())
            area = self.current_drawing['area']
            lx, ly = self.clamp_logical(lx, ly, area)
            self.current_drawing['lx1'] = lx
            self.current_drawing['ly1'] = ly
            self.update()
            if self.twin:
                self.twin.sync_drag(lx, ly)

    def _draw_mode_release(self, event):
        if not self.current_drawing:
            return

        area = self.current_drawing['area']
        lx0, ly0 = self.current_drawing['lx0'], self.current_drawing['ly0']
        lx1, ly1 = self.current_drawing['lx1'], self.current_drawing['ly1']

        if abs(lx1 - lx0) < 2 and abs(ly1 - ly0) < 2:
            self.current_drawing = None
            self.update()
            if self.twin:
                self.twin.sync_release(self.drawings)
            return

        min_lx, max_lx = min(lx0, lx1), max(lx0, lx1)
        min_ly, max_ly = min(ly0, ly1), max(ly0, ly1)
        min_m_x, max_m_x = px_to_m_x(min_lx), px_to_m_x(max_lx)
        min_m_y, max_m_y = px_to_m_y(min_ly, area), px_to_m_y(max_ly, area)

        self.current_drawing['text'] = f"{min_m_x:.1f}~{max_m_x:.1f}m\n{min_m_y:.1f}~{max_m_y:.1f}m"
        self.current_drawing['tx'] = max_lx
        b = WORK_AREAS[area]
        self.current_drawing['ty'] = max_ly + 15 if max_ly + 30 <= b["y_max"] else min_ly - 15

        if self.on_draw_callback:
            loc_str = calc_location_string(
                self.base_kilo_m, self.direction, min_m_x, max_m_x, min_m_y, max_m_y,
            )
            result = self.on_draw_callback(area, loc_str, self.current_drawing)
            if result:
                db_id, mgmt_number = result
                self.current_drawing['db_id'] = db_id
                self.current_drawing['mgmt_number'] = mgmt_number
            else:
                # キャンセル → シェイプを破棄
                self.current_drawing = None
                self.update()
                if self.twin:
                    self.twin.sync_release(self.drawings)
                return

        self.drawings.append(self.current_drawing)
        self.current_drawing = None
        self.update()
        if self.twin:
            self.twin.sync_release(self.drawings)

    # ------------------------------------------------------------------
    # 移動モード マウスイベント
    # ------------------------------------------------------------------

    def _move_mode_press(self, event):
        dx, dy = event.position().x(), event.position().y()
        lx, ly = self.display_to_logical(dx, dy)

        # 選択中の図形がある場合、まずハンドルをチェック
        if 0 <= self._selected_idx < len(self.drawings):
            handle = self._hit_test_handle(dx, dy, self.drawings[self._selected_idx])
            if handle >= 0:
                self._drag_handle = handle
                self._drag_start_lx = lx
                self._drag_start_ly = ly
                d = self.drawings[self._selected_idx]
                # 正規化した矩形を保存（ハンドル位置と一致させる）
                x0, y0, x1, y1 = self._get_drawing_rect(d)
                self._drag_orig = (x0, y0, x1, y1)
                return

            # 選択中の図形自体をクリック → 移動開始（選択維持）
            if self._hit_test_drawing(lx, ly, self.drawings[self._selected_idx]):
                self._drag_handle = -1
                self._drag_start_lx = lx
                self._drag_start_ly = ly
                d = self.drawings[self._selected_idx]
                x0, y0, x1, y1 = self._get_drawing_rect(d)
                self._drag_orig = (x0, y0, x1, y1)
                return

        # 別の図形をクリック → 選択切替
        idx = self._find_drawing_at(lx, ly)
        if idx >= 0:
            self._selected_idx = idx
            self._drag_handle = -1
            self._drag_start_lx = lx
            self._drag_start_ly = ly
            d = self.drawings[idx]
            x0, y0, x1, y1 = self._get_drawing_rect(d)
            self._drag_orig = (x0, y0, x1, y1)
            self.update()
            if self.twin:
                self.twin.sync_select(idx)
            self._notify_selection(idx)
        else:
            # シェイプ外クリック → 選択解除
            self._selected_idx = -1
            self._drag_orig = None
            self.update()
            if self.twin:
                self.twin.sync_select(-1)
            self._notify_selection(-1)

    def _notify_selection(self, idx):
        """選択変更をコールバックで通知"""
        if self.on_selection_changed_callback:
            if 0 <= idx < len(self.drawings):
                db_id = self.drawings[idx].get('db_id', -1)
                self.on_selection_changed_callback(db_id)
            else:
                self.on_selection_changed_callback(-1)

    def _move_mode_move(self, event):
        if self._drag_orig is None or self._selected_idx < 0:
            # カーソル形状の変更
            if 0 <= self._selected_idx < len(self.drawings):
                dx, dy = event.position().x(), event.position().y()
                handle = self._hit_test_handle(dx, dy, self.drawings[self._selected_idx])
                if handle >= 0:
                    cursors = {
                        _H_TL: Qt.SizeFDiagCursor, _H_BR: Qt.SizeFDiagCursor,
                        _H_TR: Qt.SizeBDiagCursor, _H_BL: Qt.SizeBDiagCursor,
                        _H_T: Qt.SizeVerCursor, _H_B: Qt.SizeVerCursor,
                        _H_L: Qt.SizeHorCursor, _H_R: Qt.SizeHorCursor,
                    }
                    self.setCursor(cursors.get(handle, Qt.ArrowCursor))
                else:
                    self.setCursor(Qt.SizeAllCursor if self._find_drawing_at(*self.display_to_logical(dx, dy)) >= 0 else Qt.ArrowCursor)
            return

        lx, ly = self.display_to_logical(event.position().x(), event.position().y())
        dlx = lx - self._drag_start_lx
        dly = ly - self._drag_start_ly
        o = self._drag_orig
        d = self.drawings[self._selected_idx]
        area = d.get('area')
        b = WORK_AREAS.get(area)

        # o = (min_x, min_y, max_x, max_y) — 正規化済み
        if self._drag_handle == -1:
            # 移動（エリア制限付き）
            new_x0 = o[0] + dlx
            new_y0 = o[1] + dly
            new_x1 = o[2] + dlx
            new_y1 = o[3] + dly
            if b:
                # エリア内にクランプ
                if new_x0 < b["x_min"]:
                    shift = b["x_min"] - new_x0
                    new_x0 += shift
                    new_x1 += shift
                if new_x1 > b["x_max"]:
                    shift = new_x1 - b["x_max"]
                    new_x0 -= shift
                    new_x1 -= shift
                if new_y0 < b["y_min"]:
                    shift = b["y_min"] - new_y0
                    new_y0 += shift
                    new_y1 += shift
                if new_y1 > b["y_max"]:
                    shift = new_y1 - b["y_max"]
                    new_y0 -= shift
                    new_y1 -= shift
            d['lx0'] = new_x0
            d['ly0'] = new_y0
            d['lx1'] = new_x1
            d['ly1'] = new_y1
        else:
            # リサイズ（エリア制限付き）
            # o[0]=min_x, o[1]=min_y, o[2]=max_x, o[3]=max_y
            new_x0, new_y0, new_x1, new_y1 = o[0], o[1], o[2], o[3]
            h_idx = self._drag_handle
            # TL/L/BL → min_x辺を動かす
            if h_idx in (_H_TL, _H_L, _H_BL):
                new_x0 = o[0] + dlx
            # TR/R/BR → max_x辺を動かす
            if h_idx in (_H_TR, _H_R, _H_BR):
                new_x1 = o[2] + dlx
            # TL/T/TR → min_y辺を動かす
            if h_idx in (_H_TL, _H_T, _H_TR):
                new_y0 = o[1] + dly
            # BL/B/BR → max_y辺を動かす
            if h_idx in (_H_BL, _H_B, _H_BR):
                new_y1 = o[3] + dly
            if b:
                new_x0 = max(b["x_min"], min(b["x_max"], new_x0))
                new_x1 = max(b["x_min"], min(b["x_max"], new_x1))
                new_y0 = max(b["y_min"], min(b["y_max"], new_y0))
                new_y1 = max(b["y_min"], min(b["y_max"], new_y1))
            d['lx0'] = min(new_x0, new_x1)
            d['ly0'] = min(new_y0, new_y1)
            d['lx1'] = max(new_x0, new_x1)
            d['ly1'] = max(new_y0, new_y1)

        self.update()
        if self.twin:
            self.twin.sync_move(self._selected_idx, d['lx0'], d['ly0'], d['lx1'], d['ly1'])

    def _move_mode_release(self, event):
        if self._drag_orig is None or self._selected_idx < 0:
            return

        d = self.drawings[self._selected_idx]
        o = self._drag_orig

        # 変更があった場合のみコールバック
        if (d['lx0'], d['ly0'], d['lx1'], d['ly1']) != o:
            if self.on_drawing_modified_callback and d.get('db_id'):
                self.on_drawing_modified_callback(d)
            if self.twin:
                self.twin.sync_release(self.drawings)

        # ドラッグ状態をリセットするが、選択は維持（連続変形対応）
        self._drag_orig = None
        self._drag_handle = -1
        self.update()

    # ------------------------------------------------------------------
    # ズーム
    # ------------------------------------------------------------------

    def wheelEvent(self, event: QWheelEvent):
        if not self._pixmap:
            return
        zoom_factor = 1.1 if event.angleDelta().y() > 0 else (1 / 1.1)
        min_zoom = min(self._calc_fit_zoom(), 0.4)
        new_zoom = max(min_zoom, min(self.zoom * zoom_factor, 5.0))
        if new_zoom == self.zoom:
            return

        pos = event.position()
        lx, ly = self.display_to_logical(pos.x(), pos.y())
        self.zoom = new_zoom
        self.pan_x = lx - pos.x() / self.zoom
        self.pan_y = ly - pos.y() / self.zoom

        self.clamp_pan()
        self.update()
        if self.twin:
            self.twin.sync_zoom(self.zoom, self.pan_x, self.pan_y)

    # ------------------------------------------------------------------
    # Twin同期
    # ------------------------------------------------------------------

    def sync_zoom(self, zoom, pan_x, pan_y):
        self.zoom = zoom
        self.pan_x = pan_x
        self.pan_y = pan_y
        self.update()

    def sync_pan(self, pan_x, pan_y):
        self.pan_x = pan_x
        self.pan_y = pan_y
        self.update()

    def sync_press(self, drawing_dict):
        self.current_drawing = drawing_dict.copy() if drawing_dict else None
        self.update()

    def sync_drag(self, lx1, ly1):
        if self.current_drawing:
            self.current_drawing['lx1'] = lx1
            self.current_drawing['ly1'] = ly1
            self.update()

    def sync_release(self, drawings_list):
        self.drawings = [d.copy() for d in drawings_list]
        self.current_drawing = None
        self.update()

    def sync_select(self, idx):
        self._selected_idx = idx
        self.update()

    def sync_move(self, idx, lx0, ly0, lx1, ly1):
        if 0 <= idx < len(self.drawings):
            self.drawings[idx]['lx0'] = lx0
            self.drawings[idx]['ly0'] = ly0
            self.drawings[idx]['lx1'] = lx1
            self.drawings[idx]['ly1'] = ly1
            self.update()
