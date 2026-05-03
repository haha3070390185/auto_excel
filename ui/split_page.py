import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QGroupBox, QRadioButton,
    QLineEdit, QFileDialog, QMessageBox, QProgressBar,
    QComboBox, QSpinBox, QCheckBox, QButtonGroup
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

from core.excel_splitter import ExcelSplitter
from database import DatabaseManager


class SplitWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(
        self,
        file_path,
        output_dir,
        split_type,
        split_column=None,
        rows_per_file=None,
        sheet_name=None,
        file_prefix="",
        include_headers=True
    ):
        super().__init__()
        self.file_path = file_path
        self.output_dir = output_dir
        self.split_type = split_type
        self.split_column = split_column
        self.rows_per_file = rows_per_file
        self.sheet_name = sheet_name
        self.file_prefix = file_prefix
        self.include_headers = include_headers
        self.splitter = ExcelSplitter()
        self.splitter.set_progress_callback(self._on_progress)
        self.splitter.set_error_callback(self._on_error)
        self.error_message = ""
        
    def _on_progress(self, progress: int, message: str):
        self.progress_signal.emit(progress, message)
        
    def _on_error(self, message: str):
        self.error_message = message
        
    def run(self):
        try:
            success = False
            
            if self.split_type == "by_column":
                success = self.splitter.split_by_column(
                    file_path=self.file_path,
                    output_dir=self.output_dir,
                    split_column=self.split_column,
                    sheet_name=self.sheet_name if self.sheet_name else None,
                    file_prefix=self.file_prefix,
                    include_headers=self.include_headers
                )
            elif self.split_type == "by_rows":
                success = self.splitter.split_by_row_count(
                    file_path=self.file_path,
                    output_dir=self.output_dir,
                    rows_per_file=self.rows_per_file,
                    sheet_name=self.sheet_name if self.sheet_name else None,
                    file_prefix=self.file_prefix if self.file_prefix else "part",
                    include_headers=self.include_headers
                )
            elif self.split_type == "by_sheets":
                success = self.splitter.split_sheets_to_files(
                    file_path=self.file_path,
                    output_dir=self.output_dir,
                    file_prefix=self.file_prefix
                )
            
            if success:
                self.finished_signal.emit(True, f"拆分完成！文件已保存至: {self.output_dir}")
            else:
                self.finished_signal.emit(False, self.error_message or "拆分失败")
                
        except Exception as e:
            self.finished_signal.emit(False, f"拆分过程中发生错误: {str(e)}")


class SplitPage(QWidget):
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.selected_file = ""
        self.available_columns = []
        self.available_sheets = []
        self.worker = None
        
        self._init_ui()
        
    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        file_group = QGroupBox("选择Excel文件")
        file_group.setFont(QFont("Microsoft YaHei", 11))
        file_layout = QHBoxLayout(file_group)
        
        file_label = QLabel("源文件:")
        file_label.setFont(QFont("Microsoft YaHei", 10))
        file_layout.addWidget(file_label)
        
        self.file_edit = QLineEdit()
        self.file_edit.setFont(QFont("Microsoft YaHei", 10))
        self.file_edit.setPlaceholderText("选择要拆分的Excel文件...")
        self.file_edit.setReadOnly(True)
        file_layout.addWidget(self.file_edit)
        
        self.browse_file_btn = QPushButton("浏览...")
        self.browse_file_btn.setFont(QFont("Microsoft YaHei", 10))
        self.browse_file_btn.clicked.connect(self._browse_file)
        file_layout.addWidget(self.browse_file_btn)
        
        main_layout.addWidget(file_group)
        
        split_type_group = QGroupBox("拆分方式")
        split_type_group.setFont(QFont("Microsoft YaHei", 11))
        split_type_layout = QVBoxLayout(split_type_group)
        
        self.split_type_group = QButtonGroup(self)
        
        self.by_column_rb = QRadioButton("按列值拆分（如按部门、地区等）")
        self.by_column_rb.setFont(QFont("Microsoft YaHei", 10))
        self.by_column_rb.setChecked(True)
        self.split_type_group.addButton(self.by_column_rb)
        split_type_layout.addWidget(self.by_column_rb)
        
        column_layout = QHBoxLayout()
        column_label = QLabel("选择拆分列:")
        column_label.setFont(QFont("Microsoft YaHei", 10))
        column_layout.addWidget(column_label)
        
        self.column_combo = QComboBox()
        self.column_combo.setFont(QFont("Microsoft YaHei", 10))
        self.column_combo.setMinimumWidth(200)
        column_layout.addWidget(self.column_combo)
        
        self.refresh_columns_btn = QPushButton("刷新列")
        self.refresh_columns_btn.setFont(QFont("Microsoft YaHei", 10))
        self.refresh_columns_btn.clicked.connect(self._refresh_columns)
        column_layout.addWidget(self.refresh_columns_btn)
        
        column_layout.addStretch()
        split_type_layout.addLayout(column_layout)
        
        self.by_rows_rb = QRadioButton("按行数拆分（每个文件固定行数）")
        self.by_rows_rb.setFont(QFont("Microsoft YaHei", 10))
        self.split_type_group.addButton(self.by_rows_rb)
        split_type_layout.addWidget(self.by_rows_rb)
        
        rows_layout = QHBoxLayout()
        rows_label = QLabel("每个文件行数:")
        rows_label.setFont(QFont("Microsoft YaHei", 10))
        rows_layout.addWidget(rows_label)
        
        self.rows_spin = QSpinBox()
        self.rows_spin.setFont(QFont("Microsoft YaHei", 10))
        self.rows_spin.setRange(1, 1000000)
        self.rows_spin.setValue(1000)
        self.rows_spin.setEnabled(False)
        rows_layout.addWidget(self.rows_spin)
        
        rows_layout.addStretch()
        split_type_layout.addLayout(rows_layout)
        
        self.by_sheets_rb = QRadioButton("按工作表拆分（每个工作表一个文件）")
        self.by_sheets_rb.setFont(QFont("Microsoft YaHei", 10))
        self.split_type_group.addButton(self.by_sheets_rb)
        split_type_layout.addWidget(self.by_sheets_rb)
        
        self.split_type_group.buttonClicked.connect(self._on_split_type_changed)
        
        main_layout.addWidget(split_type_group)
        
        options_group = QGroupBox("高级选项")
        options_group.setFont(QFont("Microsoft YaHei", 11))
        options_layout = QHBoxLayout(options_group)
        
        sheet_label = QLabel("指定工作表（可选）:")
        sheet_label.setFont(QFont("Microsoft YaHei", 10))
        options_layout.addWidget(sheet_label)
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.setFont(QFont("Microsoft YaHei", 10))
        self.sheet_combo.setMinimumWidth(150)
        self.sheet_combo.addItem("（使用第一个工作表）")
        options_layout.addWidget(self.sheet_combo)
        
        prefix_label = QLabel("文件名前缀:")
        prefix_label.setFont(QFont("Microsoft YaHei", 10))
        options_layout.addWidget(prefix_label)
        
        self.prefix_edit = QLineEdit()
        self.prefix_edit.setFont(QFont("Microsoft YaHei", 10))
        self.prefix_edit.setPlaceholderText("可选")
        self.prefix_edit.setMaximumWidth(150)
        options_layout.addWidget(self.prefix_edit)
        
        self.include_headers_cb = QCheckBox("包含表头")
        self.include_headers_cb.setFont(QFont("Microsoft YaHei", 10))
        self.include_headers_cb.setChecked(True)
        options_layout.addWidget(self.include_headers_cb)
        
        options_layout.addStretch()
        main_layout.addWidget(options_group)
        
        output_group = QGroupBox("输出设置")
        output_group.setFont(QFont("Microsoft YaHei", 11))
        output_layout = QHBoxLayout(output_group)
        
        output_label = QLabel("输出目录:")
        output_label.setFont(QFont("Microsoft YaHei", 10))
        output_layout.addWidget(output_label)
        
        self.output_edit = QLineEdit()
        self.output_edit.setFont(QFont("Microsoft YaHei", 10))
        self.output_edit.setPlaceholderText("选择输出目录...")
        output_layout.addWidget(self.output_edit)
        
        self.browse_output_btn = QPushButton("浏览...")
        self.browse_output_btn.setFont(QFont("Microsoft YaHei", 10))
        self.browse_output_btn.clicked.connect(self._browse_output)
        output_layout.addWidget(self.browse_output_btn)
        
        main_layout.addWidget(output_group)
        
        progress_group = QGroupBox("处理进度")
        progress_group.setFont(QFont("Microsoft YaHei", 11))
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setFont(QFont("Microsoft YaHei", 10))
        self.progress_bar.setMinimumHeight(25)
        progress_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("就绪")
        self.status_label.setFont(QFont("Microsoft YaHei", 10))
        progress_layout.addWidget(self.status_label)
        
        main_layout.addWidget(progress_group)
        
        action_layout = QHBoxLayout()
        action_layout.addStretch()
        
        self.split_btn = QPushButton("开始拆分")
        self.split_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.split_btn.setMinimumHeight(45)
        self.split_btn.setMinimumWidth(150)
        self.split_btn.clicked.connect(self._start_split)
        action_layout.addWidget(self.split_btn)
        
        self.open_output_btn = QPushButton("打开输出目录")
        self.open_output_btn.setFont(QFont("Microsoft YaHei", 10))
        self.open_output_btn.setMinimumHeight(45)
        self.open_output_btn.clicked.connect(self._open_output_dir)
        action_layout.addWidget(self.open_output_btn)
        
        action_layout.addStretch()
        main_layout.addLayout(action_layout)
        
    def _browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls *.xlsm);;所有文件 (*.*)"
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_edit.setText(file_path)
            self._load_file_info()
            
    def _load_file_info(self):
        if not self.selected_file:
            return
        
        splitter = ExcelSplitter()
        
        self.available_sheets = splitter.get_sheet_names(self.selected_file)
        self.sheet_combo.clear()
        self.sheet_combo.addItem("（使用第一个工作表）")
        self.sheet_combo.addItems(self.available_sheets)
        
        self._refresh_columns()
        
    def _refresh_columns(self):
        if not self.selected_file:
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return
        
        splitter = ExcelSplitter()
        
        sheet_name = None
        if self.sheet_combo.currentIndex() > 0:
            sheet_name = self.sheet_combo.currentText()
        
        self.available_columns = splitter.get_columns(self.selected_file, sheet_name)
        
        self.column_combo.clear()
        self.column_combo.addItems(self.available_columns)
        
    def _on_split_type_changed(self, button):
        if button == self.by_column_rb:
            self.column_combo.setEnabled(True)
            self.refresh_columns_btn.setEnabled(True)
            self.rows_spin.setEnabled(False)
        elif button == self.by_rows_rb:
            self.column_combo.setEnabled(False)
            self.refresh_columns_btn.setEnabled(False)
            self.rows_spin.setEnabled(True)
        elif button == self.by_sheets_rb:
            self.column_combo.setEnabled(False)
            self.refresh_columns_btn.setEnabled(False)
            self.rows_spin.setEnabled(False)
            
    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择输出目录"
        )
        
        if folder:
            self.output_edit.setText(folder)
            
    def _get_split_type(self):
        if self.by_column_rb.isChecked():
            return "by_column"
        elif self.by_rows_rb.isChecked():
            return "by_rows"
        else:
            return "by_sheets"
            
    def _start_split(self):
        if not self.selected_file:
            QMessageBox.warning(self, "警告", "请先选择要拆分的Excel文件")
            return
        
        if not os.path.exists(self.selected_file):
            QMessageBox.warning(self, "警告", "所选文件不存在")
            return
        
        output_dir = self.output_edit.text().strip()
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录")
            return
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"无法创建输出目录: {str(e)}")
                return
        
        split_type = self._get_split_type()
        
        if split_type == "by_column":
            if not self.available_columns:
                QMessageBox.warning(self, "警告", "未能读取到列信息，请先刷新列")
                return
        
        sheet_name = None
        if self.sheet_combo.currentIndex() > 0:
            sheet_name = self.sheet_combo.currentText()
        
        self.split_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在准备拆分...")
        
        self.worker = SplitWorker(
            file_path=self.selected_file,
            output_dir=output_dir,
            split_type=split_type,
            split_column=self.column_combo.currentText() if split_type == "by_column" else None,
            rows_per_file=self.rows_spin.value() if split_type == "by_rows" else None,
            sheet_name=sheet_name,
            file_prefix=self.prefix_edit.text().strip(),
            include_headers=self.include_headers_cb.isChecked()
        )
        
        self.worker.progress_signal.connect(self._on_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.start()
        
    def _on_progress(self, progress: int, message: str):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        
    def _on_finished(self, success: bool, message: str):
        self.split_btn.setEnabled(True)
        
        options = {
            'split_type': self._get_split_type(),
            'split_column': self.column_combo.currentText() if self.by_column_rb.isChecked() else "",
            'rows_per_file': self.rows_spin.value() if self.by_rows_rb.isChecked() else 0,
            'file_prefix': self.prefix_edit.text().strip(),
            'include_headers': self.include_headers_cb.isChecked()
        }
        
        self.db_manager.add_history_record(
            operation_type="表格拆分",
            source_files=[self.selected_file],
            target_file=self.output_edit.text().strip(),
            options=options,
            status="completed" if success else "failed",
            message=message
        )
        
        if success:
            QMessageBox.information(self, "成功", message)
            self.progress_bar.setValue(100)
        else:
            QMessageBox.critical(self, "失败", message)
            self.progress_bar.setValue(0)
            
        self.status_label.setText(message if success else "拆分失败")
        
    def _open_output_dir(self):
        output_dir = self.output_edit.text().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            QMessageBox.information(self, "提示", "请先选择输出目录")
