import re
import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLineEdit, QPushButton, QDateEdit, QTextEdit,
    QFileDialog, QMessageBox, QLabel, QGroupBox, QSizePolicy,
    QComboBox, QCheckBox, QFrame, QInputDialog, QDialog, QTableWidget,
    QTableWidgetItem, QRadioButton
)
from PySide6.QtCore import QThread, Signal, QDate, Qt
from PySide6.QtGui import QFont, QTextCursor

from config import DB_PATH, RETENTION_DAYS, XLSX_SUFFIX, CSV_SUFFIX
from file_parser import parse_filename
from db import LogDatabase, LogDatabaseError
from merger import perform_merge


class ScanWorker(QThread):
    """扫描线程：查找待合并文件并查重"""
    log_signal = Signal(str)
    scan_finished = Signal(int, list)   # count, list of (Path, ParsedBoxInfo)

    def __init__(self, folder: Path, shipment_date: str):
        super().__init__()
        self.folder = folder
        self.shipment_date = shipment_date

    def run(self):
        try:
            db = LogDatabase(DB_PATH)
            db.initialize()

            # 检查数据库结构是否为旧版，若需要迁移则直接停止扫描
            if db.check_schema_status():
                self.log_signal.emit("数据库结构为旧版本，请先点击右上角“更新数据库”按钮进行更新，然后再执行合并。")
                self.scan_finished.emit(0, [])
                db.close()
                return

            db.cleanup_old_logs(RETENTION_DAYS)
            existing_box_codes = db.load_existing_box_codes()
            db.close()
        except LogDatabaseError as e:
            self.log_signal.emit(f"日志数据库错误：{e}")
            self.scan_finished.emit(0, [])
            return
        except Exception as e:
            self.log_signal.emit(f"未知错误：{e}")
            self.scan_finished.emit(0, [])
            return

        file_infos = []
        seen_box_codes = set()
        try:
            for f in self.folder.iterdir():
                if f.suffix.lower() not in (XLSX_SUFFIX, CSV_SUFFIX):
                    continue

                if not re.search(r'\(v\d+\)', f.stem):
                    self.log_signal.emit(f"文件名中未提取到版本号，跳过文件：{f.name}")
                    continue

                parsed = parse_filename(f)
                if not parsed:
                    self.log_signal.emit(f"文件名不符合规则，跳过：{f.name}")
                    continue

                if parsed.date_from_mtime:
                    self.log_signal.emit(
                        f"注意：文件 {f.name} 文件名中未找到有效日期，已使用文件修改日期 {parsed.package_date}"
                    )
                if parsed.package_date != self.shipment_date:
                    continue
                if parsed.box_code in existing_box_codes:
                    self.log_signal.emit(f"箱码已合并过，跳过：{f.name}")
                    continue
                if parsed.box_code in seen_box_codes:
                    self.log_signal.emit(f"同一批次中箱码重复，跳过：{f.name}")
                    continue

                seen_box_codes.add(parsed.box_code)
                file_infos.append((f, parsed))
                self.log_signal.emit(f"发现待合并文件：{f.name}")
        except Exception as e:
            self.log_signal.emit(f"扫描文件夹失败：{e}")
            self.scan_finished.emit(0, [])
            return

        self.scan_finished.emit(len(file_infos), file_infos)


class MergeWorker(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(str)
    merge_finished = Signal(bool, list, str)

    def __init__(self, file_infos, output_dir: Path, output_name_prefix: str,
                 shipment_date: str, auto_export_boxes: bool = True):
        super().__init__()
        self.file_infos = file_infos
        self.output_dir = output_dir
        self.output_name_prefix = output_name_prefix
        self.shipment_date = shipment_date
        self.auto_export_boxes = auto_export_boxes

    def run(self):
        try:
            success, files, message = perform_merge(
                self.file_infos,
                self.output_dir,
                self.output_name_prefix,
                self.shipment_date,
                log_func=self.log_signal.emit,
                progress_callback=self.progress_signal.emit,
                auto_export_boxes=self.auto_export_boxes
            )
            self.merge_finished.emit(success, files, message)
        except Exception as e:
            self.log_signal.emit(f"合并过程发生异常：{e}")
            self.merge_finished.emit(False, [], str(e))


class ExportWorker(QThread):
    log_signal = Signal(str)
    export_finished = Signal(bool, str)

    def __init__(self, task_type: str, output_path: str, data_file_path: str = None, date_filter: str = None):
        super().__init__()
        self.task_type = task_type  # 'boxes', 'sns_simple', 'sns_full'
        self.output_path = output_path
        self.data_file_path = data_file_path
        self.date_filter = date_filter

    def run(self):
        try:
            db = LogDatabase(DB_PATH)
            db.initialize()
            if self.task_type == 'boxes':
                count = db.export_boxes_to_excel(self.output_path, self.date_filter)
                message = f"箱码导出成功，共导出 {count} 条记录"
            elif self.task_type == 'sns_simple':
                count = db.export_sns_to_excel(self.output_path, self.date_filter)
                message = f"SN 简单导出成功，共导出 {count} 条记录"
            elif self.task_type == 'sns_full':
                if not self.data_file_path:
                    raise ValueError("完整导出需要提供原始数据文件")
                count = db.export_sns_with_data(self.output_path, self.data_file_path, self.date_filter)
                message = f"SN 完整导出成功，共导出 {count} 条记录"
            else:
                raise ValueError("未知的导出任务类型")
            db.close()
            self.export_finished.emit(True, message)
        except Exception as e:
            self.export_finished.emit(False, f"导出失败：{e}")


class UndoWorker(QThread):
    log_signal = Signal(str)
    undo_finished = Signal(bool, str)

    def run(self):
        try:
            db = LogDatabase(DB_PATH)
            db.initialize()
            deleted_count = db.undo_last_merge()
            db.close()
            if deleted_count == 0:
                self.undo_finished.emit(False, "没有可撤销的合并记录")
            else:
                self.undo_finished.emit(True, f"撤销成功，已删除 {deleted_count} 条箱码记录及其关联的 SN 记录")
        except Exception as e:
            self.undo_finished.emit(False, f"撤销失败：{e}")


class UpdateDbWorker(QThread):
    log_signal = Signal(str)
    update_finished = Signal(bool, str)

    def run(self):
        try:
            db = LogDatabase(DB_PATH)
            db.initialize()
            success, message = db.migrate_schema()
            db.close()
            self.update_finished.emit(success, message)
        except Exception as e:
            self.update_finished.emit(False, f"更新数据库失败：{e}")


class DatabaseManageDialog(QDialog):
    """数据库管理对话框，提供批次查询和按批次号撤销功能"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("数据库模块 - 批次管理")
        self.resize(700, 500)
        self.db = LogDatabase(DB_PATH)
        self._init_ui()
        self._on_query()  # 初始化时自动查询全部批次

    def _init_ui(self):
        layout = QVBoxLayout(self)

        # 查询区域
        query_group = QGroupBox("查询批次记录")
        query_layout = QHBoxLayout()
        self.query_all_radio = QRadioButton("全部")
        self.query_all_radio.setChecked(True)
        self.query_date_radio = QRadioButton("按日期")
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setEnabled(False)
        self.query_btn = QPushButton("查询")
        self.query_btn.clicked.connect(self._on_query)

        self.query_all_radio.toggled.connect(self._on_radio_toggled)

        query_layout.addWidget(self.query_all_radio)
        query_layout.addWidget(self.query_date_radio)
        query_layout.addWidget(self.date_edit)
        query_layout.addWidget(self.query_btn)
        query_group.setLayout(query_layout)
        layout.addWidget(query_group)

        # 结果表格
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["批次号", "合并时间", "箱码数", "箱码信息"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.table)

        # 撤销操作区域
        undo_group = QGroupBox("按批次号撤销")
        undo_layout = QHBoxLayout()
        undo_layout.addWidget(QLabel("批次号 (batch_id)："))
        self.batch_id_edit = QLineEdit()
        self.batch_id_edit.setPlaceholderText("输入要撤销的批次号")
        self.undo_btn = QPushButton("撤销该批次")
        self.undo_btn.clicked.connect(self._on_undo)
        undo_layout.addWidget(self.batch_id_edit)
        undo_layout.addWidget(self.undo_btn)
        undo_group.setLayout(undo_layout)
        layout.addWidget(undo_group)

        # 日志输出
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        self._append_log("数据库模块已启动，显示全部批次记录。")

    def _on_radio_toggled(self):
        self.date_edit.setEnabled(self.query_date_radio.isChecked())

    def _on_query(self):
        date_filter = None
        if self.query_date_radio.isChecked():
            date_filter = self.date_edit.date().toPython().isoformat()
        try:
            self.db.initialize()
            batches = self.db.get_batches(date_filter)
            self._populate_table(batches)
            # 输出所有批次号到日志框
            if batches:
                self._append_log(f"查询到 {len(batches)} 个批次：")
                for b in batches:
                    self._append_log(f"  {b['batch_id']}")
            else:
                self._append_log("没有查询到批次记录。")
        except Exception as e:
            QMessageBox.warning(self, "查询失败", str(e))

    def _populate_table(self, batches):
        self.table.setRowCount(len(batches))
        for row_idx, batch in enumerate(batches):
            self.table.setItem(row_idx, 0, QTableWidgetItem(batch['batch_id']))
            self.table.setItem(row_idx, 1, QTableWidgetItem(batch['merge_timestamp']))
            self.table.setItem(row_idx, 2, QTableWidgetItem(str(batch['box_count'])))
            self.table.setItem(row_idx, 3, QTableWidgetItem(batch['box_codes_preview']))

    def _on_undo(self):
        batch_id = self.batch_id_edit.text().strip()
        if not batch_id:
            QMessageBox.warning(self, "输入错误", "请输入批次号")
            return
        password, ok = QInputDialog.getText(self, "撤销批次", "请输入密码：", QLineEdit.Password, "")
        if not ok:
            return
        if password != "zdc123":
            QMessageBox.warning(self, "密码错误", "密码错误，无法执行撤销")
            return
        ret = QMessageBox.question(
            self,
            "确认撤销",
            f"即将撤销批次 {batch_id} 的所有记录，此操作不可恢复。\n确定继续吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if ret != QMessageBox.Yes:
            return
        try:
            self.db.initialize()
            deleted = self.db.undo_batch_by_id(batch_id)
            self._append_log(f"撤销批次 {batch_id} 成功，删除 {deleted} 条箱码记录。")
            self._on_query()  # 刷新表格
        except Exception as e:
            QMessageBox.warning(self, "撤销失败", str(e))

    def _append_log(self, message):
        self.log_text.append(message)
        self.log_text.moveCursor(QTextCursor.End)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("出货数据合并工具")
        self.resize(1100, 800)
        self.setMinimumSize(1000, 700)

        self.scan_thread = None
        self.merge_thread = None
        self.export_thread = None
        self.undo_thread = None
        self.update_thread = None

        self._init_ui()
        self._apply_style()
        self._center_window()

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 顶部标题栏
        top_layout = QHBoxLayout()
        title_layout = QVBoxLayout()
        title_label = QLabel("出货数据合并工具")
        title_label.setObjectName("titleLabel")
        subtitle_label = QLabel("批量合并 Csv/Xlsx 出货箱，自动去重并记录日志")
        subtitle_label.setObjectName("subtitleLabel")
        title_layout.addWidget(title_label)
        title_layout.addWidget(subtitle_label)
        top_layout.addLayout(title_layout)

        top_layout.addStretch(1)

        # 按钮容器
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        # 设置按钮
        self.settings_btn = QPushButton("⚙️ 设置")
        self.settings_btn.setObjectName("settingsButton")
        self.settings_btn.setFixedWidth(100)
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.clicked.connect(self._toggle_settings_panel)
        btn_layout.addWidget(self.settings_btn)

        # 撤销合并按钮
        self.undo_btn = QPushButton("↩️ 撤销合并")
        self.undo_btn.setObjectName("settingsButton")
        self.undo_btn.setFixedWidth(120)
        self.undo_btn.setCursor(Qt.PointingHandCursor)
        self.undo_btn.clicked.connect(self._on_undo_clicked)
        btn_layout.addWidget(self.undo_btn)

        # 更新数据库按钮
        self.update_db_btn = QPushButton("🔄 更新数据库")
        self.update_db_btn.setObjectName("settingsButton")
        self.update_db_btn.setFixedWidth(120)
        self.update_db_btn.setCursor(Qt.PointingHandCursor)
        self.update_db_btn.clicked.connect(self._on_update_db_clicked)
        btn_layout.addWidget(self.update_db_btn)

        # 数据库模块按钮
        self.database_btn = QPushButton("🗄️ 数据库模块")
        self.database_btn.setObjectName("settingsButton")
        self.database_btn.setFixedWidth(120)
        self.database_btn.setCursor(Qt.PointingHandCursor)
        self.database_btn.clicked.connect(self._open_database_dialog)
        btn_layout.addWidget(self.database_btn)

        top_layout.addLayout(btn_layout)
        main_layout.addLayout(top_layout)

        # 设置面板
        self.settings_panel = QFrame()
        self.settings_panel.setObjectName("settingsPanel")
        self.settings_panel.setVisible(False)
        settings_layout = QVBoxLayout(self.settings_panel)
        settings_layout.setContentsMargins(10, 10, 10, 10)
        self.auto_export_check = QCheckBox("文件合并时自动生成对应的箱码记录")
        self.auto_export_check.setChecked(True)
        self.auto_export_check.setObjectName("autoExportCheckBox")
        settings_layout.addWidget(self.auto_export_check)
        main_layout.addWidget(self.settings_panel)

        # 左右分栏
        content_layout = QHBoxLayout()
        content_layout.setSpacing(20)
        left_layout = QVBoxLayout()
        left_layout.setSpacing(15)

        # 文件选择
        file_group = QGroupBox("文件选择")
        file_group.setObjectName("groupBox")
        file_layout = QGridLayout()
        file_layout.setSpacing(12)
        file_layout.setContentsMargins(15, 15, 15, 15)
        file_layout.addWidget(QLabel("待出库文件夹："), 0, 0, Qt.AlignRight)
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("请选择包含 xlsx 文件的文件夹")
        file_layout.addWidget(self.folder_edit, 0, 1)
        folder_btn = QPushButton("浏览")
        folder_btn.setObjectName("browseButton")
        folder_btn.clicked.connect(self._select_folder)
        file_layout.addWidget(folder_btn, 0, 2)
        file_layout.addWidget(QLabel("出库日期："), 1, 0, Qt.AlignRight)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setFixedWidth(150)
        file_layout.addWidget(self.date_edit, 1, 1, alignment=Qt.AlignLeft)
        file_group.setLayout(file_layout)
        left_layout.addWidget(file_group)

        # 输出设置
        output_group = QGroupBox("输出设置")
        output_group.setObjectName("groupBox")
        output_layout = QGridLayout()
        output_layout.setSpacing(12)
        output_layout.setContentsMargins(15, 15, 15, 15)
        output_layout.addWidget(QLabel("汇总存放路径："), 0, 0, Qt.AlignRight)
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("请选择汇总文件保存的文件夹")
        output_layout.addWidget(self.output_dir_edit, 0, 1)
        output_dir_btn = QPushButton("浏览")
        output_dir_btn.setObjectName("browseButton")
        output_dir_btn.clicked.connect(self._select_output_dir)
        output_layout.addWidget(output_dir_btn, 0, 2)
        output_layout.addWidget(QLabel("汇总文件名："), 1, 0, Qt.AlignRight)
        self.output_name_edit = QLineEdit()
        self.output_name_edit.setPlaceholderText("请输入文件名前缀，例如：001")
        output_layout.addWidget(self.output_name_edit, 1, 1, 1, 2)
        hint1 = QLabel("        汇总文件会自动命名为：输入的前缀-六位出库日期-版本号。并自动导出对应的箱码记录。若有文件重名，会自动添加（1）等后缀")
        hint1.setObjectName("hintLabel")
        hint1.setWordWrap(True)
        hint2 = QLabel("       例：输入001，出库日期为8月25日，则会生成 001-260825-v16.xlsx 、001-260825-v16.xlsx-箱码记录 等多个汇总表")
        hint2.setObjectName("hintLabel")
        hint2.setWordWrap(True)
        output_layout.addWidget(hint1, 2, 0, 1, 3)
        output_layout.addWidget(hint2, 3, 0, 1, 3)
        output_group.setLayout(output_layout)
        left_layout.addWidget(output_group)

        self.merge_btn = QPushButton("开始合并")
        self.merge_btn.setObjectName("mergeButton")
        self.merge_btn.setMinimumHeight(45)
        self.merge_btn.setCursor(Qt.PointingHandCursor)
        self.merge_btn.clicked.connect(self._on_merge_clicked)
        left_layout.addWidget(self.merge_btn)

        content_layout.addLayout(left_layout, stretch=7)

        # 右侧面板
        right_layout = QVBoxLayout()
        right_layout.setSpacing(15)
        export_group = QGroupBox("数据导出")
        export_group.setObjectName("groupBox")
        export_layout = QVBoxLayout()
        export_layout.setSpacing(10)
        export_layout.setContentsMargins(15, 15, 15, 15)
        scope_layout = QHBoxLayout()
        scope_layout.addWidget(QLabel("导出范围（按出库日期）:"))
        self.export_scope_combo = QComboBox()
        self.export_scope_combo.addItem("全部数据")
        self.export_scope_combo.addItem("按日期导出")
        self.export_scope_combo.currentIndexChanged.connect(self._on_export_scope_changed)
        scope_layout.addWidget(self.export_scope_combo)
        export_layout.addLayout(scope_layout)
        self.export_date_edit = QDateEdit()
        self.export_date_edit.setCalendarPopup(True)
        self.export_date_edit.setDate(QDate.currentDate())
        self.export_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.export_date_edit.setVisible(False)
        export_layout.addWidget(self.export_date_edit)
        self.export_boxes_btn = QPushButton("导出箱码")
        self.export_boxes_btn.setObjectName("exportBoxesButton")
        self.export_boxes_btn.setMinimumHeight(35)
        self.export_boxes_btn.setCursor(Qt.PointingHandCursor)
        self.export_boxes_btn.clicked.connect(self._on_export_boxes_clicked)
        export_layout.addWidget(self.export_boxes_btn)
        self.export_sns_simple_btn = QPushButton("简单导出 SN")
        self.export_sns_simple_btn.setObjectName("exportSnsSimpleButton")
        self.export_sns_simple_btn.setMinimumHeight(35)
        self.export_sns_simple_btn.setCursor(Qt.PointingHandCursor)
        self.export_sns_simple_btn.clicked.connect(self._on_export_sns_simple_clicked)
        export_layout.addWidget(self.export_sns_simple_btn)
        self.export_sns_full_btn = QPushButton("完整导出 SN")
        self.export_sns_full_btn.setObjectName("exportSnsFullButton")
        self.export_sns_full_btn.setMinimumHeight(35)
        self.export_sns_full_btn.setCursor(Qt.PointingHandCursor)
        self.export_sns_full_btn.clicked.connect(self._on_export_sns_full_clicked)
        export_layout.addWidget(self.export_sns_full_btn)
        export_group.setLayout(export_layout)
        right_layout.addWidget(export_group)

        view_group = QGroupBox("数据查看")
        view_group.setObjectName("groupBox")
        view_layout = QVBoxLayout()
        view_layout.setSpacing(10)
        view_layout.setContentsMargins(15, 15, 15, 15)
        view_desc = QLabel("选择查看内容：")
        view_layout.addWidget(view_desc)
        self.view_combo = QComboBox()
        self.view_combo.addItem("查看数据库信息")
        self.view_combo.addItem("查看最新的10条日志记录")
        self.view_combo.setMinimumHeight(60)
        view_layout.addWidget(self.view_combo)
        self.view_btn = QPushButton("查看")
        self.view_btn.setObjectName("viewButton")
        self.view_btn.setMinimumHeight(35)
        self.view_btn.setCursor(Qt.PointingHandCursor)
        self.view_btn.clicked.connect(self._on_view_clicked)
        view_layout.addWidget(self.view_btn)
        view_group.setLayout(view_layout)
        right_layout.addWidget(view_group)

        content_layout.addLayout(right_layout, stretch=3)
        main_layout.addLayout(content_layout)

        log_group = QGroupBox("输出信息")
        log_group.setObjectName("groupBox")
        log_layout = QVBoxLayout()
        log_layout.setContentsMargins(10, 10, 10, 10)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("logText")
        self.log_text.setFont(QFont("Menlo", 10))
        self.log_text.setMinimumHeight(150)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group, 1)

        self.status_label = QLabel("就绪")
        self.status_label.setObjectName("statusLabel")
        main_layout.addWidget(self.status_label)

    def _apply_style(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f0f4f8; }
            QGroupBox {
                background-color: #ffffff; border: 1px solid #d0d7de; border-radius: 10px;
                margin-top: 12px; font-weight: bold; font-size: 13px; color: #2c3e50;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 15px; padding: 0 5px; }
            QLabel { color: #333; font-size: 13px; }
            QLabel#titleLabel { font-size: 24px; font-weight: bold; color: #1a5276; }
            QLabel#subtitleLabel { font-size: 14px; color: #5d6d7e; margin-bottom: 5px; }
            QLabel#statusLabel { color: #7f8c8d; font-size: 12px; padding-left: 5px; }
            QLabel#hintLabel { font-size: 11px; color: #888888; margin-top: 2px; }
            QLineEdit {
                background-color: #f5f7fa; border: 1px solid #a0a8b0; border-radius: 5px;
                padding: 7px 10px; font-size: 13px; color: #2c3e50;
                placeholder-text-color: #999999; selection-background-color: #3498db;
            }
            QLineEdit:focus { border: 2px solid #3498db; background-color: #ffffff; }
            QDateEdit {
                background-color: #e3f2fd; border: 1px solid #1976d2; border-radius: 5px;
                padding: 5px 8px; font-size: 13px; color: #0d47a1;
                selection-background-color: #3498db;
            }
            QDateEdit:focus { border: 2px solid #ff9800; background-color: #ffffff; }
            QDateEdit::drop-down {
                subcontrol-origin: padding; subcontrol-position: top right; width: 20px;
                border-left: 1px solid #1976d2; background-color: #bbdefb;
            }
            QDateEdit::down-arrow {
                image: none; border-left: 4px solid transparent; border-right: 4px solid transparent;
                border-top: 5px solid #0d47a1; margin-right: 5px;
            }
            QComboBox {
                background-color: #f5f7fa; border: 1px solid #a0a8b0; border-radius: 5px;
                padding: 5px 10px; font-size: 13px; min-height: 25px; color: #2c3e50;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding; subcontrol-position: top right; width: 20px;
                border-left: 1px solid #a0a8b0; background-color: #eaecee;
            }
            QPushButton {
                background-color: #eaecee; border: 1px solid #cbd2d9; border-radius: 5px;
                padding: 7px 15px; font-size: 13px; color: #2c3e50;
            }
            QPushButton:hover { background-color: #dde1e3; }
            QPushButton:pressed { background-color: #cfd4d8; }
            QPushButton#browseButton { min-width: 80px; }
            QPushButton#mergeButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #3498db, stop:1 #2980b9);
                color: white; font-size: 16px; font-weight: bold; border: none; border-radius: 8px; padding: 10px;
            }
            QPushButton#mergeButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #3fa3e0, stop:1 #2c89c9); }
            QPushButton#mergeButton:pressed { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2c89c9, stop:1 #1f6da0); }
            QPushButton#mergeButton:disabled { background: #a9cce3; color: #eaf2f8; }
            QPushButton#exportBoxesButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #9b59b6, stop:1 #8e44ad);
                color: white; font-size: 14px; font-weight: bold; border: none; border-radius: 8px; padding: 8px;
            }
            QPushButton#exportBoxesButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #af7ac5, stop:1 #9b59b6); }
            QPushButton#exportBoxesButton:pressed { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #8e44ad, stop:1 #7d3c98); }
            QPushButton#exportBoxesButton:disabled { background: #c39bd3; color: #f5eef8; }
            QPushButton#exportSnsSimpleButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #3498db, stop:1 #2980b9);
                color: white; font-size: 14px; font-weight: bold; border: none; border-radius: 8px; padding: 8px;
            }
            QPushButton#exportSnsSimpleButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #3fa3e0, stop:1 #2c89c9); }
            QPushButton#exportSnsSimpleButton:pressed { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2c89c9, stop:1 #1f6da0); }
            QPushButton#exportSnsSimpleButton:disabled { background: #a9cce3; color: #eaf2f8; }
            QPushButton#exportSnsFullButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f39c12, stop:1 #e67e22);
                color: white; font-size: 14px; font-weight: bold; border: none; border-radius: 8px; padding: 8px;
            }
            QPushButton#exportSnsFullButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f5b041, stop:1 #f39c12); }
            QPushButton#exportSnsFullButton:pressed { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #e67e22, stop:1 #d35400); }
            QPushButton#exportSnsFullButton:disabled { background: #f5cba7; color: #fef9e7; }
            QPushButton#viewButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #16a085, stop:1 #1abc9c);
                color: white; font-size: 14px; font-weight: bold; border: none; border-radius: 8px; padding: 8px;
            }
            QPushButton#viewButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1abc9c, stop:1 #17a589); }
            QPushButton#viewButton:pressed { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #17a589, stop:1 #148f77); }
            QPushButton#viewButton:disabled { background: #a3e4d7; color: #eafaf5; }
            QPushButton#settingsButton {
                background-color: #eaecee; border: 1px solid #cbd2d9; border-radius: 5px;
                padding: 5px 10px; font-size: 13px; color: #2c3e50;
            }
            QPushButton#settingsButton:hover { background-color: #dde1e3; }
            QPushButton#settingsButton:pressed { background-color: #cfd4d8; }
            QFrame#settingsPanel {
                background-color: #f5f7fa; border: 1px solid #d0d7de; border-radius: 5px;
            }
            QCheckBox { font-size: 13px; color: #2c3e50; }
            QTextEdit#logText {
                background-color: #f8f9fa; border: 1px solid #d0d7de; border-radius: 5px;
                font-family: "Menlo", "Monaco", "Courier New", monospace; font-size: 12px; padding: 5px;
            }
        """)

    def _center_window(self):
        screen = QApplication.primaryScreen()
        if screen:
            screen_geometry = screen.availableGeometry()
            window_geometry = self.frameGeometry()
            center_point = screen_geometry.center()
            window_geometry.moveCenter(center_point)
            self.move(window_geometry.topLeft())

    # ---------- 数据库模块 ----------
    def _open_database_dialog(self):
        dialog = DatabaseManageDialog(self)
        dialog.exec()

    # ---------- 更新数据库 ----------
    def _on_update_db_clicked(self):
        password, ok = QInputDialog.getText(self, "更新数据库", "请输入密码：", QLineEdit.Password, "")
        if not ok:
            return
        if password != "zdc123":
            QMessageBox.warning(self, "密码错误", "密码错误，无法更新数据库")
            return
        ret = QMessageBox.question(self, "确认更新", "即将检测并更新数据库结构，请确保已备份数据。\n确定继续吗？",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret != QMessageBox.Yes:
            return
        self.update_db_btn.setEnabled(False)
        self._append_log("开始更新数据库结构...")
        self.update_thread = UpdateDbWorker()
        self.update_thread.log_signal.connect(self._append_log)
        self.update_thread.update_finished.connect(self._on_update_finished)
        self.update_thread.start()

    def _on_update_finished(self, success, message):
        self.update_db_btn.setEnabled(True)
        self._append_log(message)
        if success:
            QMessageBox.information(self, "更新完成", message)
        else:
            QMessageBox.warning(self, "更新失败", message)

    # ---------- 撤销合并 ----------
    def _on_undo_clicked(self):
        password, ok = QInputDialog.getText(self, "撤销合并", "请输入密码：", QLineEdit.Password, "")
        if not ok:
            return
        if password != "zdc123":
            QMessageBox.warning(self, "密码错误", "密码错误，无法执行撤销操作")
            return
        ret = QMessageBox.question(self, "确认撤销", "即将撤销最近一次合并插入到数据库的记录，此操作不可恢复。\n确定继续吗？",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret != QMessageBox.Yes:
            return
        self.undo_btn.setEnabled(False)
        self._append_log("正在撤销最近一次合并...")
        self.undo_thread = UndoWorker()
        self.undo_thread.log_signal.connect(self._append_log)
        self.undo_thread.undo_finished.connect(self._on_undo_finished)
        self.undo_thread.start()

    def _on_undo_finished(self, success, message):
        self.undo_btn.setEnabled(True)
        self._append_log(message)
        if success:
            QMessageBox.information(self, "撤销完成", message)
        else:
            QMessageBox.warning(self, "撤销失败", message)

    # ---------- 设置面板切换 ----------
    def _toggle_settings_panel(self):
        self.settings_panel.setVisible(not self.settings_panel.isVisible())

    # ---------- 文件/路径选择 ----------
    def _select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择待出库文件夹")
        if folder:
            self.folder_edit.setText(folder)
            self.status_label.setText(f"已选择待出库文件夹：{folder}")

    def _select_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "选择汇总文件存放路径")
        if folder:
            self.output_dir_edit.setText(folder)
            self.status_label.setText(f"已选择汇总存放路径：{folder}")

    # ---------- 日志辅助 ----------
    def _append_log(self, message: str):
        self.log_text.append(message)
        self.log_text.moveCursor(QTextCursor.End)
        self.status_label.setText(message)

    def _update_status(self, message: str):
        self.status_label.setText(message)

    def _on_export_scope_changed(self):
        if self.export_scope_combo.currentIndex() == 1:
            self.export_date_edit.setVisible(True)
        else:
            self.export_date_edit.setVisible(False)

    # ---------- 合并操作 ----------
    def _on_merge_clicked(self):
        folder_str = self.folder_edit.text().strip()
        output_dir_str = self.output_dir_edit.text().strip()
        output_name = self.output_name_edit.text().strip()

        if not folder_str:
            QMessageBox.warning(self, "输入错误", "请选择待出库文件夹")
            return
        if not output_dir_str:
            QMessageBox.warning(self, "输入错误", "请选择汇总文件存放路径")
            return
        if not output_name:
            QMessageBox.warning(self, "输入错误", "请填写汇总文件名")
            return

        if output_name.lower().endswith('.xlsx'):
            output_name = output_name[:-5]
            self.output_name_edit.setText(output_name)

        folder = Path(folder_str)
        output_dir = Path(output_dir_str)

        if not folder.exists() or not folder.is_dir():
            QMessageBox.warning(self, "路径错误", "待出库文件夹不存在")
            return
        if not output_dir.exists():
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, "路径错误", f"无法创建汇总存放路径：{e}")
                return

        shipment_date = self.date_edit.date().toPython().isoformat()

        try:
            test_file = output_dir / '.write_test'
            test_file.touch()
            test_file.unlink()
        except Exception as e:
            QMessageBox.warning(self, "路径错误", f"汇总存放路径不可写：{e}")
            return

        self.merge_btn.setEnabled(False)
        self._append_log("开始扫描待合并文件...")

        self.scan_thread = ScanWorker(folder, shipment_date)
        self.scan_thread.log_signal.connect(self._append_log)
        self.scan_thread.scan_finished.connect(self._on_scan_finished)
        self.scan_thread.start()

    def _on_scan_finished(self, count, file_infos):
        self.merge_btn.setEnabled(True)
        if count == 0:
            self._append_log("没有需要合并的新文件")
            QMessageBox.information(self, "提示", "没有需要合并的新文件")
            return

        ret = QMessageBox.question(
            self,
            "确认合并",
            f"共合并 {count} 个文件，请确认！",
            QMessageBox.Ok | QMessageBox.Cancel,
            QMessageBox.Cancel
        )
        if ret != QMessageBox.Ok:
            self._append_log("用户取消合并")
            return

        output_name = self.output_name_edit.text().strip()
        if not output_name:
            QMessageBox.warning(self, "输入错误", "请填写汇总文件名")
            return
        if output_name.lower().endswith('.xlsx'):
            output_name = output_name[:-5]

        output_dir = Path(self.output_dir_edit.text().strip())
        shipment_date_obj = self.date_edit.date().toPython()
        merge_date_str = shipment_date_obj.strftime('%y%m%d')

        versions = set()
        for _, info in file_infos:
            versions.add(info.version)

        warning_files = []
        for ver in versions:
            base_filename = f"{output_name}-{merge_date_str}-v{ver}"
            final_path = _get_unique_filename(output_dir, base_filename, '.xlsx')
            if final_path.name != (base_filename + '.xlsx'):
                warning_files.append(base_filename + '.xlsx')
            if self.auto_export_check.isChecked():
                box_base = f"{output_name}-{merge_date_str}-v{ver}-箱码记录"
                box_path = _get_unique_filename(output_dir, box_base, '.xlsx')
                if box_path.name != (box_base + '.xlsx'):
                    warning_files.append(box_base + '.xlsx')

        if warning_files:
            file_list = "\n".join(warning_files)
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle("文件重名提示")
            msg_box.setText("以下文件名已存在，保存时将自动添加后缀，请注意！")
            msg_box.setInformativeText(f"{file_list}\n\n点击“确定”继续合并，或“取消”中止操作。")
            msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            msg_box.setDefaultButton(QMessageBox.Ok)
            ret_confirm = msg_box.exec()
            if ret_confirm != QMessageBox.Ok:
                self._append_log("用户取消合并（因重名提示）")
                return

        self.merge_btn.setEnabled(False)
        self._append_log("开始合并数据...")

        shipment_date = self.date_edit.date().toPython().isoformat()
        auto_export = self.auto_export_check.isChecked()

        self.merge_thread = MergeWorker(
            file_infos, output_dir, output_name, shipment_date,
            auto_export_boxes=auto_export
        )
        self.merge_thread.log_signal.connect(self._append_log)
        self.merge_thread.progress_signal.connect(self._update_status)
        self.merge_thread.merge_finished.connect(self._on_merge_finished)
        self.merge_thread.start()

    def _on_merge_finished(self, success, files, message):
        self.merge_btn.setEnabled(True)
        self.status_label.setText("合并完成" if success else "合并失败")
        if success:
            if files:
                file_list = "\n".join(files)
                QMessageBox.information(self, "完成", f"{message}\n\n生成的文件：\n{file_list}")
            else:
                QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.warning(self, "合并失败", message)

    # ---------- 导出操作 ----------
    def _on_export_boxes_clicked(self):
        date_filter = None
        if self.export_scope_combo.currentIndex() == 1:
            date_filter = self.export_date_edit.date().toPython().isoformat()
            date_str = self.export_date_edit.date().toString('yyMMdd')
            default_name = f"{date_str}-箱码记录.xlsx"
        else:
            default_name = "箱码记录（全部数据）.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出箱码记录", default_name, "Excel 文件 (*.xlsx)"
        )
        if not file_path:
            return
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'
        self._start_export('boxes', file_path, date_filter=date_filter)

    def _on_export_sns_simple_clicked(self):
        date_filter = None
        if self.export_scope_combo.currentIndex() == 1:
            date_filter = self.export_date_edit.date().toPython().isoformat()
            date_str = self.export_date_edit.date().toString('yyMMdd')
            default_name = f"{date_str}-SN简单导出.xlsx"
        else:
            default_name = "SN简单导出（全部数据）.xlsx"

        save_path, _ = QFileDialog.getSaveFileName(
            self, "简单导出 SN", default_name, "Excel 文件 (*.xlsx)"
        )
        if not save_path:
            return
        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'
        self._start_export('sns_simple', save_path, date_filter=date_filter)

    def _on_export_sns_full_clicked(self):
        date_filter = None
        if self.export_scope_combo.currentIndex() == 1:
            date_filter = self.export_date_edit.date().toPython().isoformat()
            date_str = self.export_date_edit.date().toString('yyMMdd')
            default_name = f"{date_str}-SN完整导出.xlsx"
        else:
            default_name = "SN完整导出（全部数据）.xlsx"

        data_file, _ = QFileDialog.getOpenFileName(
            self, "请选择包含原始数据列的 Excel 文件", "", "Excel 文件 (*.xlsx)"
        )
        if not data_file:
            self._append_log("已取消完整导出")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "完整导出 SN", default_name, "Excel 文件 (*.xlsx)"
        )
        if not save_path:
            self._append_log("已取消保存，完整导出中止")
            return
        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'

        self._start_export('sns_full', save_path, data_file_path=data_file, date_filter=date_filter)

    def _start_export(self, task_type: str, output_path: str, data_file_path: str = None, date_filter: str = None):
        self._append_log(f"开始导出{'箱码' if task_type == 'boxes' else 'SN'}数据...")
        if task_type == 'boxes':
            self.export_boxes_btn.setEnabled(False)
        elif task_type == 'sns_simple':
            self.export_sns_simple_btn.setEnabled(False)
        elif task_type == 'sns_full':
            self.export_sns_full_btn.setEnabled(False)

        self.export_thread = ExportWorker(task_type, output_path, data_file_path, date_filter)
        self.export_thread.log_signal.connect(self._append_log)
        self.export_thread.export_finished.connect(self._on_export_finished)
        self.export_thread.start()

    def _on_export_finished(self, success, message):
        self.export_boxes_btn.setEnabled(True)
        self.export_sns_simple_btn.setEnabled(True)
        self.export_sns_full_btn.setEnabled(True)
        self._append_log(message)
        if success:
            QMessageBox.information(self, "导出完成", message)
        else:
            QMessageBox.warning(self, "导出失败", message)

    # ---------- 查看操作 ----------
    def _on_view_clicked(self):
        selected = self.view_combo.currentIndex()
        try:
            db = LogDatabase(DB_PATH)
            db.initialize()
            if selected == 0:
                boxes_count, sns_count, mtime = db.get_database_info()
                self._append_log("--- 数据库信息 ---")
                self._append_log(f"箱码记录数：{boxes_count}")
                self._append_log(f"SN 记录数：{sns_count}")
                self._append_log(f"数据库最后修改时间：{mtime}")
                self._append_log("------------------")
            elif selected == 1:
                records = db.get_latest_boxes(10)
                self._append_log("--- 最新10条箱码日志 ---")
                if not records:
                    self._append_log("暂无日志记录")
                for rec in records:
                    self._append_log(
                        f"箱码: {rec['box_code']} | 文件: {rec['original_filename']} | "
                        f"合并日期: {rec['merge_date']} {rec['merge_time']} | "
                        f"有效行数: {rec['valid_rows']}，总行数: {rec['total_rows']}"
                    )
                self._append_log("------------------------")
            db.close()
        except LogDatabaseError as e:
            self._append_log(f"数据库错误：{e}")
        except Exception as e:
            self._append_log(f"查看时发生异常：{e}")


# 辅助函数：生成唯一文件名
def _get_unique_filename(directory: Path, base_name: str, extension: str = '.xlsx') -> Path:
    final_path = directory / (base_name + extension)
    counter = 1
    while final_path.exists():
        final_path = directory / (f"{base_name} ({counter})" + extension)
        counter += 1
    return final_path