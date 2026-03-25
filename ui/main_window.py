"""メインウィンドウ（ImageViewerApp） - PySide6版"""

import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox,
    QButtonGroup, QInputDialog,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from core.canvas import DrawingCanvas
from core.calc_utils import (
    parse_kilo, WORK_AREAS, m_to_px_x, px_to_m_x, px_to_m_y,
    calc_location_string, calc_range_string, calc_depth_string,
    circled_number,
)
from data.db_manager import DatabaseManager
from data.excel_exporter import ExcelExporter
from data.file_loader import load_image_groups, sort_kilos
from ui.components import DrawingTable, KiloListWidget
from ui.exclusion_dialog import ExclusionDialog


class ImageViewerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("アトラス解析システム")
        self.resize(1920, 1000)

        self.ui_font = QFont("Meiryo", 14)
        self.ui_font.setBold(True)
        self.btn_font = QFont("Meiryo", 14)
        self.btn_font.setBold(True)

        self.image_groups = {}
        self.sorted_kilos = []
        self.current_index = 0
        self.parent_folder = ""

        self.db = DatabaseManager()

        self._create_widgets()

    def _create_widgets(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(10, 10, 10, 15)

        # === 左側メインエリア ===
        main_layout = QVBoxLayout()

        # --- コントロールパネル ---
        control_layout = QHBoxLayout()

        self.select_btn = QPushButton("フォルダを選択")
        self.select_btn.setFont(self.btn_font)
        self.select_btn.setFixedSize(160, 40)
        self.select_btn.clicked.connect(self.load_folder)
        control_layout.addWidget(self.select_btn)

        self.info_label = QLabel("フォルダが選択されていません")
        self.info_label.setFont(self.ui_font)
        control_layout.addWidget(self.info_label)
        control_layout.addStretch()

        self.excel_btn = QPushButton("Excel出力")
        self.excel_btn.setFont(self.btn_font)
        self.excel_btn.setFixedSize(130, 40)
        self.excel_btn.setEnabled(False)
        self.excel_btn.setStyleSheet(
            "QPushButton { background-color: #2ea043; color: white; border-radius: 4px; }"
            "QPushButton:hover { background-color: #238636; }"
            "QPushButton:disabled { background-color: #555; color: #999; }"
        )
        self.excel_btn.clicked.connect(self.export_excel)
        control_layout.addWidget(self.excel_btn)

        self.pdf_btn = QPushButton("PDF出力")
        self.pdf_btn.setFont(self.btn_font)
        self.pdf_btn.setFixedSize(120, 40)
        self.pdf_btn.setEnabled(False)
        self.pdf_btn.setStyleSheet(
            "QPushButton { background-color: #c62828; color: white; border-radius: 4px; }"
            "QPushButton:hover { background-color: #b71c1c; }"
            "QPushButton:disabled { background-color: #555; color: #999; }"
        )
        self.pdf_btn.clicked.connect(self.export_pdf)
        control_layout.addWidget(self.pdf_btn)

        self.log_btn = QPushButton("集計表出力")
        self.log_btn.setFont(self.btn_font)
        self.log_btn.setFixedSize(130, 40)
        self.log_btn.setEnabled(False)
        self.log_btn.setStyleSheet(
            "QPushButton { background-color: #1976d2; color: white; border-radius: 4px; }"
            "QPushButton:hover { background-color: #1565c0; }"
            "QPushButton:disabled { background-color: #555; color: #999; }"
        )
        self.log_btn.clicked.connect(self.export_logs)
        control_layout.addWidget(self.log_btn)

        self.csv_import_btn = QPushButton("除外csv一括取込")
        self.csv_import_btn.setFont(self.btn_font)
        self.csv_import_btn.setFixedSize(180, 40)
        self.csv_import_btn.setEnabled(False)
        self.csv_import_btn.setStyleSheet(
            "QPushButton { background-color: #e65100; color: white; border-radius: 4px; }"
            "QPushButton:hover { background-color: #bf360c; }"
            "QPushButton:disabled { background-color: #555; color: #999; }"
        )
        self.csv_import_btn.clicked.connect(self.import_exclusion_csv)
        control_layout.addWidget(self.csv_import_btn)

        # ページ送り（10単位）
        self.prev10_btn = QPushButton("≪10")
        self.prev10_btn.setFont(self.btn_font)
        self.prev10_btn.setFixedSize(70, 40)
        self.prev10_btn.setEnabled(False)
        self.prev10_btn.clicked.connect(lambda: self._jump(-10))
        control_layout.addWidget(self.prev10_btn)

        self.prev_btn = QPushButton("＜")
        self.prev_btn.setFont(self.btn_font)
        self.prev_btn.setFixedSize(50, 40)
        self.prev_btn.setEnabled(False)
        self.prev_btn.clicked.connect(lambda: self._jump(-1))
        control_layout.addWidget(self.prev_btn)

        self.next_btn = QPushButton("＞")
        self.next_btn.setFont(self.btn_font)
        self.next_btn.setFixedSize(50, 40)
        self.next_btn.setEnabled(False)
        self.next_btn.clicked.connect(lambda: self._jump(1))
        control_layout.addWidget(self.next_btn)

        self.next10_btn = QPushButton("10≫")
        self.next10_btn.setFont(self.btn_font)
        self.next10_btn.setFixedSize(70, 40)
        self.next10_btn.setEnabled(False)
        self.next10_btn.clicked.connect(lambda: self._jump(10))
        control_layout.addWidget(self.next10_btn)

        main_layout.addLayout(control_layout)

        # --- ツールパネル ---
        tool_layout = QHBoxLayout()

        # モード切替
        self._mode_btn_group = QButtonGroup(self)
        self._mode_btn_group.setExclusive(True)
        mode_style = (
            "QPushButton { border: 2px solid #888; border-radius: 4px; padding: 4px 12px; }"
            "QPushButton:checked { background-color: #444; color: white; border-color: white; }"
        )

        self.draw_mode_btn = QPushButton("入力モード")
        self.draw_mode_btn.setFont(self.ui_font)
        self.draw_mode_btn.setCheckable(True)
        self.draw_mode_btn.setChecked(True)
        self.draw_mode_btn.setFixedHeight(36)
        self.draw_mode_btn.setStyleSheet(mode_style)
        self._mode_btn_group.addButton(self.draw_mode_btn, 0)
        tool_layout.addWidget(self.draw_mode_btn)

        self.move_mode_btn = QPushButton("移動モード")
        self.move_mode_btn.setFont(self.ui_font)
        self.move_mode_btn.setCheckable(True)
        self.move_mode_btn.setFixedHeight(36)
        self.move_mode_btn.setStyleSheet(mode_style)
        self._mode_btn_group.addButton(self.move_mode_btn, 1)
        tool_layout.addWidget(self.move_mode_btn)
        self._mode_btn_group.idClicked.connect(self._change_edit_mode)

        sep0 = QLabel("｜")
        sep0.setFont(self.ui_font)
        tool_layout.addWidget(sep0)

        # 形状選択
        self._draw_btn_group = QButtonGroup(self)
        self._draw_btn_group.setExclusive(True)

        self.rect_btn = QPushButton("四角を描画")
        self.rect_btn.setFont(self.ui_font)
        self.rect_btn.setCheckable(True)
        self.rect_btn.setChecked(True)
        self.rect_btn.setFixedHeight(36)
        self._draw_btn_group.addButton(self.rect_btn, 0)
        tool_layout.addWidget(self.rect_btn)

        self.oval_btn = QPushButton("円を描画")
        self.oval_btn.setFont(self.ui_font)
        self.oval_btn.setCheckable(True)
        self.oval_btn.setFixedHeight(36)
        self._draw_btn_group.addButton(self.oval_btn, 1)
        tool_layout.addWidget(self.oval_btn)
        self._draw_btn_group.idClicked.connect(self._change_draw_mode)

        sep1 = QLabel("｜")
        sep1.setFont(self.ui_font)
        tool_layout.addWidget(sep1)

        # 種別選択
        self._category_btn_group = QButtonGroup(self)
        self._category_btn_group.setExclusive(True)
        cat_btn_style = (
            "QPushButton {{ color: {color}; border: 2px solid {color}; border-radius: 4px; padding: 4px 12px; }}"
            "QPushButton:checked {{ background-color: {color}; color: white; }}"
        )

        self.yurumi_btn = QPushButton("ゆるみ")
        self.yurumi_btn.setFont(self.ui_font)
        self.yurumi_btn.setCheckable(True)
        self.yurumi_btn.setChecked(True)
        self.yurumi_btn.setFixedHeight(36)
        self.yurumi_btn.setStyleSheet(cat_btn_style.format(color="blue"))
        self._category_btn_group.addButton(self.yurumi_btn, 0)
        tool_layout.addWidget(self.yurumi_btn)

        self.kudo_btn = QPushButton("空洞")
        self.kudo_btn.setFont(self.ui_font)
        self.kudo_btn.setCheckable(True)
        self.kudo_btn.setFixedHeight(36)
        self.kudo_btn.setStyleSheet(cat_btn_style.format(color="red"))
        self._category_btn_group.addButton(self.kudo_btn, 1)
        tool_layout.addWidget(self.kudo_btn)

        self.exclusion_btn = QPushButton("除外区間")
        self.exclusion_btn.setFont(self.ui_font)
        self.exclusion_btn.setCheckable(True)
        self.exclusion_btn.setFixedHeight(36)
        self.exclusion_btn.setStyleSheet(cat_btn_style.format(color="#cc0000"))
        self._category_btn_group.addButton(self.exclusion_btn, 2)
        tool_layout.addWidget(self.exclusion_btn)
        self._category_btn_group.idClicked.connect(self._change_category)

        tool_layout.addStretch()
        main_layout.addLayout(tool_layout)

        # --- 画像表示エリア ---
        image_layout = QHBoxLayout()
        image_layout.setSpacing(20)

        left_layout = QVBoxLayout()
        left_label = QLabel("【マーキングあり】")
        left_label.setFont(self.ui_font)
        left_label.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(left_label)

        self.canvas_l = DrawingCanvas(on_draw_callback=self.add_table_row)
        self.canvas_l.on_exclusion_click_callback = self._on_exclusion_click
        self.canvas_l.on_drawing_modified_callback = self._on_drawing_modified
        self.canvas_l.on_selection_changed_callback = self._on_canvas_selection_changed
        left_layout.addWidget(self.canvas_l, stretch=1)
        image_layout.addLayout(left_layout, stretch=1)

        right_layout = QVBoxLayout()
        right_label = QLabel("【マーキングなし】")
        right_label.setFont(self.ui_font)
        right_label.setAlignment(Qt.AlignCenter)
        right_layout.addWidget(right_label)

        self.canvas_r = DrawingCanvas(on_draw_callback=self.add_table_row)
        self.canvas_r.on_exclusion_click_callback = self._on_exclusion_click
        self.canvas_r.on_drawing_modified_callback = self._on_drawing_modified
        self.canvas_r.on_selection_changed_callback = self._on_canvas_selection_changed
        right_layout.addWidget(self.canvas_r, stretch=1)
        image_layout.addLayout(right_layout, stretch=1)

        main_layout.addLayout(image_layout, stretch=1)

        # 左右同期
        self.canvas_l.twin = self.canvas_r
        self.canvas_r.twin = self.canvas_l
        self.all_canvases = [self.canvas_l, self.canvas_r]

        # --- 下段: テーブル ---
        self.drawing_table = DrawingTable(
            on_delete_callback=self.delete_selected, ui_font_family="Meiryo",
        )
        self.drawing_table.row_selected.connect(self._on_table_row_selected)
        self.drawing_table.data_edited.connect(self._on_table_data_edited)

        # 変状一覧ボタンをテーブルのボタン行(左側)に追加
        self.list_dialog_btn = QPushButton("変状一覧")
        self.list_dialog_btn.setFont(QFont("Meiryo", 12, QFont.Bold))
        self.list_dialog_btn.setStyleSheet(
            "QPushButton { background-color: #1976d2; color: white; padding: 6px 16px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #1565c0; }"
        )
        self.list_dialog_btn.clicked.connect(self._open_drawing_list)
        self.drawing_table.btn_layout.insertWidget(0, self.list_dialog_btn)

        self.heatmap_btn = QPushButton("ヒートマップ")
        self.heatmap_btn.setFont(QFont("Meiryo", 12, QFont.Bold))
        self.heatmap_btn.setStyleSheet(
            "QPushButton { background-color: #6a1b9a; color: white; padding: 6px 16px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #4a148c; }"
        )
        self.heatmap_btn.clicked.connect(self._open_heatmap)
        self.drawing_table.btn_layout.insertWidget(1, self.heatmap_btn)

        main_layout.addWidget(self.drawing_table)

        root_layout.addLayout(main_layout, stretch=1)

        # === 右側: キロ程一覧 ===
        self.kilo_list = KiloListWidget(ui_font_family="Meiryo")
        self.kilo_list.kilo_selected.connect(self._on_kilo_selected)
        root_layout.addWidget(self.kilo_list)

    # ------------------------------------------------------------------
    # 描画モード・種別
    # ------------------------------------------------------------------

    def _change_edit_mode(self, btn_id):
        mode = "draw" if btn_id == 0 else "move"
        for canvas in self.all_canvases:
            canvas.edit_mode = mode
            canvas._selected_idx = -1
            canvas.update()
        is_move = (mode == "move")
        self.rect_btn.setEnabled(not is_move)
        self.oval_btn.setEnabled(not is_move)
        self.yurumi_btn.setEnabled(not is_move)
        self.kudo_btn.setEnabled(not is_move)
        self.exclusion_btn.setEnabled(not is_move)

    def _change_draw_mode(self, btn_id):
        mode = "rectangle" if btn_id == 0 else "oval"
        for canvas in self.all_canvases:
            canvas.draw_mode = mode

    def _change_category(self, btn_id):
        category_map = {0: "ゆるみ", 1: "空洞", 2: "除外区間"}
        category = category_map[btn_id]
        for canvas in self.all_canvases:
            canvas.draw_category = category
        is_exclusion = (category == "除外区間")
        self.rect_btn.setEnabled(not is_exclusion)
        self.oval_btn.setEnabled(not is_exclusion)

    # ------------------------------------------------------------------
    # フォルダ読み込み
    # ------------------------------------------------------------------

    def load_folder(self):
        parent_folder = QFileDialog.getExistingDirectory(self, "作業フォルダを選択")
        if not parent_folder:
            return

        db_path = os.path.join(parent_folder, "drawings.db")
        if not os.path.exists(db_path):
            reply = QMessageBox.question(
                self, "データベースの作成",
                "選択したフォルダにデータベースが見つかりません。\n新規作成してよろしいですか？",
            )
            if reply != QMessageBox.Yes:
                return

        self.db.setup(db_path)
        self.parent_folder = parent_folder

        groups, error = load_image_groups(parent_folder)
        if error:
            self.info_label.setText(error)
            return

        self.image_groups = groups
        self.sorted_kilos = sort_kilos(groups)
        self.current_index = 0

        if self.sorted_kilos:
            for btn in [self.prev_btn, self.next_btn, self.prev10_btn, self.next10_btn,
                        self.excel_btn, self.pdf_btn, self.log_btn, self.csv_import_btn]:
                btn.setEnabled(True)
            self.kilo_list.set_kilos(self.sorted_kilos)
            self.update_display()
        else:
            self.info_label.setText("対象の画像が見つかりませんでした")

    # ------------------------------------------------------------------
    # 表示更新
    # ------------------------------------------------------------------

    def update_display(self):
        if not self.sorted_kilos:
            return

        current_kilo = self.sorted_kilos[self.current_index]
        total = len(self.sorted_kilos)
        self.info_label.setText(f"キロ程: {current_kilo} ({self.current_index + 1} / {total})")

        group = self.image_groups[current_kilo]
        direction = group.get('direction', '起点→終点')
        base_kilo_m = parse_kilo(current_kilo)

        self.canvas_l.set_image(group.get('marked'), base_kilo_m, direction)
        self.canvas_r.set_image(group.get('unmarked'), base_kilo_m, direction)

        self.kilo_list.set_current_index(self.current_index)

        if self.db.is_connected:
            self.load_drawings_from_db(current_kilo)

    def load_drawings_from_db(self, kilo, preserve_selection_db_id=-1):
        self.drawing_table.clear()

        group = self.image_groups.get(kilo, {})
        direction = group.get('direction', '起点→終点')
        base_kilo_m = parse_kilo(kilo)

        drawings = self.db.load_drawings(kilo)
        for d in drawings:
            area = d.get('area', '')
            category = d.get('category', 'ゆるみ')
            lx0, lx1 = d['lx0'], d['lx1']
            ly0, ly1 = d['ly0'], d['ly1']
            min_lx, max_lx = min(lx0, lx1), max(lx0, lx1)
            min_ly, max_ly = min(ly0, ly1), max(ly0, ly1)

            range_str = calc_range_string(base_kilo_m, direction,
                                          px_to_m_x(min_lx), px_to_m_x(max_lx))
            if category == '除外区間':
                depth_str = "0.0m ～ 2.0m"
            else:
                depth_str = calc_depth_string(
                    px_to_m_y(min_ly, area), px_to_m_y(max_ly, area))

            self.drawing_table.insert_row(
                d['db_id'], area, range_str, depth_str,
                category=category,
                mgmt_number=d.get('mgmt_number'),
                exclusion_reason=d.get('exclusion_reason', ''),
            )

        drawing_shapes = [
            {k: v for k, v in d.items() if k != 'location_str'} for d in drawings
        ]
        self.canvas_l.set_drawings(drawing_shapes)
        self.canvas_r.set_drawings(drawing_shapes)

        # 選択状態を復元
        if preserve_selection_db_id >= 0:
            for canvas in self.all_canvases:
                for i, d in enumerate(canvas.drawings):
                    if d.get('db_id') == preserve_selection_db_id:
                        canvas._selected_idx = i
                        break
                canvas.update()
            self.drawing_table.select_by_db_id(preserve_selection_db_id)

    # ------------------------------------------------------------------
    # テーブル行選択 → キャンバスのバウンディングボックス表示
    # ------------------------------------------------------------------

    def _on_table_row_selected(self, db_id):
        """テーブル行選択 → キャンバスのシェイプを選択"""
        for canvas in self.all_canvases:
            if db_id < 0:
                canvas._selected_idx = -1
            else:
                idx = -1
                for i, d in enumerate(canvas.drawings):
                    if d.get('db_id') == db_id:
                        idx = i
                        break
                canvas._selected_idx = idx
                canvas.edit_mode = "move"
            canvas.update()
        if db_id >= 0:
            self.move_mode_btn.setChecked(True)

    def _on_canvas_selection_changed(self, db_id):
        """キャンバスのシェイプ選択 → テーブル行を選択"""
        self.drawing_table.select_by_db_id(db_id)

    def _on_table_data_edited(self, db_id, field, new_value):
        """テーブル上で管理番号や除外理由が編集されたときDB反映"""
        from PySide6.QtCore import QTimer

        if not self.db.is_connected or db_id is None:
            return
        kilo = self.sorted_kilos[self.current_index]

        if field == 'mgmt_number':
            if new_value is not None:
                cursor = self.db.conn.cursor()
                cursor.execute("SELECT category FROM drawings WHERE id=?", (db_id,))
                row = cursor.fetchone()
                if row:
                    cat_filter = '除外区間' if row[0] == '除外区間' else None
                    cursor.execute(
                        "SELECT COUNT(*) FROM drawings WHERE kilo=? AND mgmt_number=? AND id!=? AND category{}".format(
                            "='除外区間'" if cat_filter == '除外区間' else "!='除外区間'"
                        ),
                        (kilo, new_value, db_id),
                    )
                    if cursor.fetchone()[0] > 0:
                        QMessageBox.warning(self, "番号重複", f"管理番号 {new_value} は既に使用されています。")
                        QTimer.singleShot(0, lambda: self.load_drawings_from_db(kilo))
                        return
            self.db.conn.execute("UPDATE drawings SET mgmt_number=? WHERE id=?", (new_value, db_id))
            self.db.conn.commit()
            # 編集エディタのコミット完了後にテーブル再構築
            QTimer.singleShot(0, lambda: self.load_drawings_from_db(kilo))

        elif field == 'exclusion_reason':
            self.db.conn.execute("UPDATE drawings SET exclusion_reason=? WHERE id=?", (new_value, db_id))
            self.db.conn.commit()

    # ------------------------------------------------------------------
    # DB操作コールバック（通常描画）
    # ------------------------------------------------------------------

    def add_table_row(self, area, location_str, drawing_dict):
        if not self.db.is_connected:
            return None

        kilo = self.sorted_kilos[self.current_index]
        category = drawing_dict.get('category', 'ゆるみ')

        mgmt_number = None
        if category != "除外区間":
            mgmt_number = self._ask_mgmt_number(kilo)
            if mgmt_number is None:
                return None

        db_id = self.db.insert_drawing(
            kilo, area, drawing_dict, location_str,
            category=category, mgmt_number=mgmt_number,
        )

        # テーブルには範囲と深さを分離して渡す
        lx0, lx1 = drawing_dict['lx0'], drawing_dict['lx1']
        ly0, ly1 = drawing_dict['ly0'], drawing_dict['ly1']
        base_kilo_m = parse_kilo(kilo)
        group = self.image_groups.get(kilo, {})
        direction = group.get('direction', '起点→終点')
        range_str = calc_range_string(base_kilo_m, direction,
                                      px_to_m_x(min(lx0, lx1)), px_to_m_x(max(lx0, lx1)))
        depth_str = calc_depth_string(
            px_to_m_y(min(ly0, ly1), area), px_to_m_y(max(ly0, ly1), area))

        self.drawing_table.insert_row(
            db_id, area, range_str, depth_str,
            category=category, mgmt_number=mgmt_number,
        )
        return (db_id, mgmt_number)

    def _ask_mgmt_number(self, kilo):
        suggested = self.db.get_next_mgmt_number(kilo, category_filter=None)
        while True:
            number, ok = QInputDialog.getInt(
                self, "管理番号の入力",
                "管理番号を入力してください:",
                suggested, 1, 9999,
            )
            if not ok:
                return None
            if self.db.is_mgmt_number_taken(kilo, number, category_filter=None):
                QMessageBox.warning(
                    self, "番号重複",
                    f"管理番号 {number} は既に使用されています。別の番号を入力してください。",
                )
                continue
            return number

    # ------------------------------------------------------------------
    # 除外区間入力
    # ------------------------------------------------------------------

    def _on_exclusion_click(self, area):
        if not self.db.is_connected or not self.sorted_kilos:
            return

        kilo = self.sorted_kilos[self.current_index]
        group = self.image_groups[kilo]
        direction = group.get('direction', '起点→終点')
        base_kilo_m = parse_kilo(kilo)

        existing_px = self.db.get_exclusion_zones(kilo, area)
        existing_m = [(px_to_m_x(s), px_to_m_x(e)) for s, e in existing_px]

        suggested = self.db.get_next_mgmt_number(kilo, category_filter='除外区間')

        all_drawings = self.db.load_drawings(kilo)
        existing_ex_numbers = {
            d['mgmt_number'] for d in all_drawings
            if d.get('category') == '除外区間' and d.get('mgmt_number') is not None
        }

        dialog = ExclusionDialog(
            self, area_name=area, existing_zones=existing_m,
            suggested_number=suggested, existing_numbers=existing_ex_numbers,
        )
        if dialog.exec() != ExclusionDialog.Accepted:
            return

        start_m = dialog.start_pos()
        end_m = dialog.end_pos()
        reason = dialog.reason()
        mgmt_number = dialog.mgmt_number()
        additional_areas = dialog.additional_areas()

        # 入力対象エリアリスト（メインエリア + 追加エリア）
        target_areas = [area] + additional_areas
        next_num = mgmt_number

        for target_area in target_areas:
            area_info = WORK_AREAS[target_area]
            lx0 = m_to_px_x(start_m)
            lx1 = m_to_px_x(end_m)
            ly0 = area_info["y_min"]
            ly1 = area_info["y_max"]

            drawing_dict = {
                'type': 'rectangle', 'category': '除外区間', 'area': target_area,
                'lx0': lx0, 'ly0': ly0, 'lx1': lx1, 'ly1': ly1,
                'tx': 0, 'ty': 0, 'text': '',
            }

            loc_str = calc_location_string(base_kilo_m, direction, start_m, end_m, 0.0, 2.0)

            self.db.insert_drawing(
                kilo, target_area, drawing_dict, loc_str,
                category='除外区間', mgmt_number=next_num,
                exclusion_reason=reason,
            )
            next_num += 1

        self.load_drawings_from_db(kilo)

    # ------------------------------------------------------------------
    # 図形移動・リサイズ後のDB更新
    # ------------------------------------------------------------------

    def _on_drawing_modified(self, drawing_dict):
        db_id = drawing_dict.get('db_id')
        if not db_id or not self.db.is_connected:
            return
        area = drawing_dict.get('area', '')
        lx0, ly0 = drawing_dict['lx0'], drawing_dict['ly0']
        lx1, ly1 = drawing_dict['lx1'], drawing_dict['ly1']
        min_lx, max_lx = min(lx0, lx1), max(lx0, lx1)
        min_ly, max_ly = min(ly0, ly1), max(ly0, ly1)
        min_m_x, max_m_x = px_to_m_x(min_lx), px_to_m_x(max_lx)
        min_m_y, max_m_y = px_to_m_y(min_ly, area), px_to_m_y(max_ly, area)
        loc_str = calc_location_string(
            self.canvas_l.base_kilo_m, self.canvas_l.direction,
            min_m_x, max_m_x, min_m_y, max_m_y,
        )
        self.db.update_drawing_coords(db_id, lx0, ly0, lx1, ly1, loc_str)
        kilo = self.sorted_kilos[self.current_index]
        # 選択状態を復元してバウンディングボックスを維持
        self.load_drawings_from_db(kilo, preserve_selection_db_id=db_id)

    # ------------------------------------------------------------------
    # 削除
    # ------------------------------------------------------------------

    def delete_selected(self, selected_items):
        reply = QMessageBox.question(
            self, "削除の確認",
            "選択した図形を削除しますか？\n（この操作は元に戻せません）",
        )
        if reply != QMessageBox.Yes:
            return

        for item_id in selected_items:
            db_id = int(item_id)
            self.db.delete_drawing(db_id)
            self.drawing_table.delete_row(item_id)
            for canvas in self.all_canvases:
                canvas.remove_drawing(db_id)

    # ------------------------------------------------------------------
    # ナビゲーション
    # ------------------------------------------------------------------

    def _jump(self, delta):
        new_idx = max(0, min(self.current_index + delta, len(self.sorted_kilos) - 1))
        if new_idx != self.current_index:
            self.current_index = new_idx
            self.update_display()

    def _on_kilo_selected(self, index):
        if 0 <= index < len(self.sorted_kilos) and index != self.current_index:
            self.current_index = index
            self.update_display()

    # ------------------------------------------------------------------
    # 変状一覧ダイアログ
    # ------------------------------------------------------------------

    def _open_heatmap(self):
        if not self.sorted_kilos or not self.db.is_connected:
            return
        from ui.heatmap_window import HeatmapWindow
        win = HeatmapWindow(
            self,
            image_groups=self.image_groups,
            sorted_kilos=self.sorted_kilos,
            db=self.db,
        )
        win.show()

    def _open_drawing_list(self):
        if not self.sorted_kilos or not self.db.is_connected:
            return
        from ui.drawing_list_dialog import DrawingListDialog
        kilo = self.sorted_kilos[self.current_index]
        dialog = DrawingListDialog(self, db=self.db, kilo=kilo)
        dialog.exec()
        # ダイアログを閉じた後、現在の表示を更新
        self.load_drawings_from_db(kilo)

    # ------------------------------------------------------------------
    # Excel出力
    # ------------------------------------------------------------------

    def _get_template_path(self):
        """テンプレートパスを取得。見つからなければユーザーに選択を促す。"""
        from data.excel_exporter import DEFAULT_TEMPLATE_PATH
        template_path = DEFAULT_TEMPLATE_PATH
        if not os.path.exists(template_path):
            QMessageBox.warning(
                self, "テンプレート未検出",
                f"デフォルトのテンプレートが見つかりません:\n{template_path}\n\nテンプレートファイルを選択してください。",
            )
            template_path, _ = QFileDialog.getOpenFileName(
                self, "テンプレートファイルを選択", "", "Excel ファイル (*.xlsx)",
            )
        return template_path or None

    def _image_basename(self, kilo):
        """キロ程に対応する画像のベース名（拡張子なし）を返す"""
        group = self.image_groups.get(kilo, {})
        marked = group.get('marked')
        if marked:
            return os.path.splitext(os.path.basename(marked))[0]
        return kilo

    def export_excel(self):
        from ui.export_range_dialog import ExportRangeDialog

        if not self.sorted_kilos:
            return

        range_dialog = ExportRangeDialog(self, sorted_kilos=self.sorted_kilos)
        if range_dialog.exec() != ExportRangeDialog.Accepted:
            return
        selected_kilos = range_dialog.selected_kilos()
        if not selected_kilos:
            QMessageBox.warning(self, "選択なし", "出力するキロ程が選択されていません。")
            return
        individual = range_dialog.is_individual()

        template_path = self._get_template_path()
        if not template_path:
            return

        if individual:
            # 個別ファイル出力 → フォルダ選択
            out_dir = QFileDialog.getExistingDirectory(self, "出力先フォルダを選択")
            if not out_dir:
                return
            try:
                for kilo in selected_kilos:
                    name = self._image_basename(kilo) + ".xlsx"
                    out_path = os.path.join(out_dir, name)
                    exporter = ExcelExporter(self.image_groups, [kilo], self.db)
                    exporter.export(out_path, template_path=template_path)
                QMessageBox.information(
                    self, "完了", f"{len(selected_kilos)} 件のExcelファイルを出力しました:\n{out_dir}")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"Excel出力に失敗しました:\n{e}")
        else:
            # 一括ファイル出力
            start_kilo, end_kilo = range_dialog.filename_range()
            default_name = f"解析報告書（{start_kilo}～{end_kilo}）.xlsx"
            output_path, _ = QFileDialog.getSaveFileName(
                self, "名前を付けて保存", default_name, "Excel ファイル (*.xlsx)",
            )
            if not output_path:
                return
            try:
                exporter = ExcelExporter(self.image_groups, selected_kilos, self.db)
                exporter.export(output_path, template_path=template_path)
                QMessageBox.information(self, "完了", f"Excelファイルを保存しました:\n{output_path}")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"Excel出力に失敗しました:\n{e}")

    # ------------------------------------------------------------------
    # PDF出力（Excel経由）
    # ------------------------------------------------------------------

    def export_pdf(self):
        from ui.export_range_dialog import ExportRangeDialog

        if not self.sorted_kilos:
            return

        range_dialog = ExportRangeDialog(self, sorted_kilos=self.sorted_kilos)
        if range_dialog.exec() != ExportRangeDialog.Accepted:
            return
        selected_kilos = range_dialog.selected_kilos()
        if not selected_kilos:
            QMessageBox.warning(self, "選択なし", "出力するキロ程が選択されていません。")
            return
        individual = range_dialog.is_individual()

        template_path = self._get_template_path()
        if not template_path:
            return

        if individual:
            # 個別ファイル出力 → フォルダ選択
            out_dir = QFileDialog.getExistingDirectory(self, "PDF出力先フォルダを選択")
            if not out_dir:
                return
            self._run_pdf_export_individual(selected_kilos, template_path, out_dir)
        else:
            # 一括ファイル出力
            start_kilo, end_kilo = range_dialog.filename_range()
            default_name = f"解析報告書（{start_kilo}～{end_kilo}）.pdf"
            output_path, _ = QFileDialog.getSaveFileName(
                self, "名前を付けて保存", default_name, "PDF ファイル (*.pdf)",
            )
            if not output_path:
                return
            self._run_pdf_export(selected_kilos, template_path, output_path)

    def _run_pdf_export(self, selected_kilos, template_path, pdf_path):
        """Excel生成(メインスレッド) → PDF変換(ワーカースレッド)"""
        import tempfile
        from PySide6.QtCore import QThread, Signal
        from PySide6.QtWidgets import QProgressDialog

        # Step 1: Excel生成（メインスレッド、DB操作あり）
        try:
            tmp = tempfile.NamedTemporaryFile(
                suffix='.xlsx', delete=False,
            )
            tmp_xlsx = tmp.name
            tmp.close()

            exporter = ExcelExporter(self.image_groups, selected_kilos, self.db)
            exporter.export(tmp_xlsx, template_path=template_path)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"Excel生成に失敗しました:\n{e}")
            return

        # Step 2: PDF変換（ワーカースレッド、COM操作のみ）
        progress = QProgressDialog("PDF変換中...", None, 0, 0, self)
        progress.setWindowTitle("処理中")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setCancelButton(None)
        progress.setMinimumWidth(300)
        progress.show()

        class PdfWorker(QThread):
            finished = Signal(str)

            def __init__(self, xlsx_path, pdf_out):
                super().__init__()
                self.xlsx_path = xlsx_path
                self.pdf_out = pdf_out

            def run(self):
                try:
                    from data.pdf_exporter import excel_to_pdf
                    excel_to_pdf(self.xlsx_path, self.pdf_out)
                    self.finished.emit("")
                except Exception as e:
                    self.finished.emit(str(e))

        worker = PdfWorker(tmp_xlsx, pdf_path)

        def on_finished(error_msg):
            progress.close()
            # 一時ファイル削除
            try:
                os.remove(tmp_xlsx)
            except Exception:
                pass
            worker.deleteLater()
            if error_msg:
                QMessageBox.critical(self, "エラー", f"PDF変換に失敗しました:\n{error_msg}")
            else:
                QMessageBox.information(self, "完了", f"PDFファイルを保存しました:\n{pdf_path}")

        worker.finished.connect(on_finished)
        worker.start()

    def _run_pdf_export_individual(self, selected_kilos, template_path, out_dir):
        """個別PDF出力: キロ程ごとにExcel生成→PDF変換"""
        import tempfile
        from PySide6.QtCore import QThread, Signal
        from PySide6.QtWidgets import QProgressDialog

        # Step 1: 全キロ程のExcel一時ファイルを生成（メインスレッド）
        tmp_pairs = []  # [(tmp_xlsx, pdf_path), ...]
        try:
            for kilo in selected_kilos:
                tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
                tmp_xlsx = tmp.name
                tmp.close()
                exporter = ExcelExporter(self.image_groups, [kilo], self.db)
                exporter.export(tmp_xlsx, template_path=template_path)
                pdf_name = self._image_basename(kilo) + ".pdf"
                pdf_path = os.path.join(out_dir, pdf_name)
                tmp_pairs.append((tmp_xlsx, pdf_path))
        except Exception as e:
            for xlsx, _ in tmp_pairs:
                try:
                    os.remove(xlsx)
                except Exception:
                    pass
            QMessageBox.critical(self, "エラー", f"Excel生成に失敗しました:\n{e}")
            return

        # Step 2: PDF変換（ワーカースレッド）
        progress = QProgressDialog(f"PDF変換中... (0/{len(tmp_pairs)})", None, 0, 0, self)
        progress.setWindowTitle("処理中")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setCancelButton(None)
        progress.setMinimumWidth(300)
        progress.show()

        class BatchPdfWorker(QThread):
            finished = Signal(str)
            progress_update = Signal(int, int)  # current, total

            def __init__(self, pairs):
                super().__init__()
                self.pairs = pairs

            def run(self):
                try:
                    from data.pdf_exporter import excel_to_pdf
                    for i, (xlsx, pdf) in enumerate(self.pairs):
                        excel_to_pdf(xlsx, pdf)
                        self.progress_update.emit(i + 1, len(self.pairs))
                    self.finished.emit("")
                except Exception as e:
                    self.finished.emit(str(e))

        worker = BatchPdfWorker(tmp_pairs)

        def on_progress(current, total):
            progress.setLabelText(f"PDF変換中... ({current}/{total})")

        def on_finished(error_msg):
            progress.close()
            for xlsx, _ in tmp_pairs:
                try:
                    os.remove(xlsx)
                except Exception:
                    pass
            worker.deleteLater()
            if error_msg:
                QMessageBox.critical(self, "エラー", f"PDF変換に失敗しました:\n{error_msg}")
            else:
                QMessageBox.information(
                    self, "完了",
                    f"{len(tmp_pairs)} 件のPDFファイルを出力しました:\n{out_dir}",
                )

        worker.progress_update.connect(on_progress)
        worker.finished.connect(on_finished)
        worker.start()

    # ------------------------------------------------------------------
    # 集計表出力
    # ------------------------------------------------------------------

    def export_logs(self):
        from data.log_exporter import export_logs

        if not self.parent_folder or not self.sorted_kilos:
            return

        output_dir = os.path.join(self.parent_folder, "集計表")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        else:
            existing_files = [f for f in os.listdir(output_dir) if f.endswith('.log')]
            if existing_files:
                reply = QMessageBox.question(
                    self, "上書き確認",
                    f"集計表フォルダに既に {len(existing_files)} 件のログファイルが存在します。\n上書きしますか？",
                )
                if reply != QMessageBox.Yes:
                    return

        try:
            count = export_logs(output_dir, self.image_groups, self.sorted_kilos, self.db)
            QMessageBox.information(
                self, "完了",
                f"集計表を出力しました:\n{output_dir}\n({count} ファイル)",
            )
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"集計表出力に失敗しました:\n{e}")

    # ------------------------------------------------------------------
    # 除外CSV一括取込
    # ------------------------------------------------------------------

    def import_exclusion_csv(self):
        from data.csv_importer import import_exclusions_from_csv

        if not self.sorted_kilos or not self.db.is_connected:
            return

        csv_folder = QFileDialog.getExistingDirectory(self, "除外CSVフォルダを選択")
        if not csv_folder:
            return

        # 既存除外区間の確認
        has_existing = False
        for kilo in self.sorted_kilos:
            drawings = self.db.load_drawings(kilo)
            if any(d.get('category') == '除外区間' for d in drawings):
                has_existing = True
                break

        overwrite = False
        if has_existing:
            msg = QMessageBox(self)
            msg.setWindowTitle("既存の除外区間")
            msg.setText("既に除外区間が設定されているキロ程があります。")
            msg.setInformativeText("既存の除外区間を上書きしますか？\nまたは追加で取り込みますか？")
            btn_overwrite = msg.addButton("上書き", QMessageBox.YesRole)
            btn_append = msg.addButton("追加", QMessageBox.NoRole)
            btn_cancel = msg.addButton("キャンセル", QMessageBox.RejectRole)
            msg.setDefaultButton(btn_append)
            msg.exec()
            if msg.clickedButton() == btn_cancel:
                return
            overwrite = (msg.clickedButton() == btn_overwrite)

        try:
            imported, skipped = import_exclusions_from_csv(
                csv_folder, self.image_groups, self.sorted_kilos, self.db,
                overwrite=overwrite,
            )
            QMessageBox.information(
                self, "取込完了",
                f"除外区間を取り込みました。\n取込: {imported} 件\nスキップ: {skipped} キロ程",
            )
            # 現在ページを更新
            kilo = self.sorted_kilos[self.current_index]
            self.load_drawings_from_db(kilo)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSV取込に失敗しました:\n{e}")
