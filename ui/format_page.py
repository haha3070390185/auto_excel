import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem,
    QLabel, QGroupBox, QCheckBox, QLineEdit,
    QFileDialog, QMessageBox, QProgressBar,
    QComboBox, QSpinBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

from core.excel_formatter import ExcelFormatter
from database import DatabaseManager


class FormatWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(
        self,
        file_paths,
        output_dir,
        options
    ):
        super().__init__()
        self.file_paths = file_paths
        self.output_dir = output_dir
        self.options = options
        self.formatter = ExcelFormatter()
        self.formatter.set_progress_callback(self._on_progress)
        self.formatter.set_error_callback(self._on_error)
        self.error_message = ""
        
    def _on_progress(self, progress: int, message: str):
        self.progress_signal.emit(progress, message)
        
    def _on_error(self, message: str):
        self.error_message = message
        
    def run(self):
        try:
            success = self.formatter.format_files(
                file_paths=self.file_paths,
                output_dir=self.output_dir,
                options=self.options
            )
            
            if success:
                self.finished_signal.emit(True, f"格式规整完成！文件已保存至: {self.output_dir}")
            else:
                self.finished_signal.emit(False, self.error_message or "格式规整失败")
                
        except Exception as e:
            self.finished_signal.emit(False, f"格式规整过程中发生错误: {str(e)}")


class FormatPage(QWidget):
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.selected_files = []
        self.worker = None
        
        self._init_ui()
        
    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        files_group = QGroupBox("选择Excel文件")
        files_group.setFont(QFont("Microsoft YaHei", 11))
        files_layout = QVBoxLayout(files_group)
        
        files_btn_layout = QHBoxLayout()
        
        self.add_files_btn = QPushButton("添加文件")
        self.add_files_btn.setFont(QFont("Microsoft YaHei", 10))
        self.add_files_btn.setMinimumHeight(35)
        self.add_files_btn.clicked.connect(self._add_files)
        files_btn_layout.addWidget(self.add_files_btn)
        
        self.add_folder_btn = QPushButton("添加文件夹")
        self.add_folder_btn.setFont(QFont("Microsoft YaHei", 10))
        self.add_folder_btn.setMinimumHeight(35)
        self.add_folder_btn.clicked.connect(self._add_folder)
        files_btn_layout.addWidget(self.add_folder_btn)
        
        self.clear_btn = QPushButton("清空列表")
        self.clear_btn.setFont(QFont("Microsoft YaHei", 10))
        self.clear_btn.setMinimumHeight(35)
        self.clear_btn.clicked.connect(self._clear_files)
        files_btn_layout.addWidget(self.clear_btn)
        
        files_btn_layout.addStretch()
        files_layout.addLayout(files_btn_layout)
        
        self.files_list = QListWidget()
        self.files_list.setFont(QFont("Microsoft YaHei", 9))
        self.files_list.setMinimumHeight(120)
        files_layout.addWidget(self.files_list)
        
        self.files_count_label = QLabel("已选择: 0 个文件")
        self.files_count_label.setFont(QFont("Microsoft YaHei", 9))
        files_layout.addWidget(self.files_count_label)
        
        main_layout.addWidget(files_group)
        
        options_group = QGroupBox("格式规整选项")
        options_group.setFont(QFont("Microsoft YaHei", 11))
        options_layout = QVBoxLayout(options_group)
        
        basic_options_layout = QHBoxLayout()
        
        self.trim_spaces_cb = QCheckBox("清除多余空格")
        self.trim_spaces_cb.setFont(QFont("Microsoft YaHei", 10))
        self.trim_spaces_cb.setChecked(True)
        basic_options_layout.addWidget(self.trim_spaces_cb)
        
        self.center_align_cb = QCheckBox("自动居中对齐")
        self.center_align_cb.setFont(QFont("Microsoft YaHei", 10))
        self.center_align_cb.setChecked(True)
        basic_options_layout.addWidget(self.center_align_cb)
        
        self.add_borders_cb = QCheckBox("添加边框")
        self.add_borders_cb.setFont(QFont("Microsoft YaHei", 10))
        self.add_borders_cb.setChecked(True)
        basic_options_layout.addWidget(self.add_borders_cb)
        
        self.auto_fit_cb = QCheckBox("自动调整列宽")
        self.auto_fit_cb.setFont(QFont("Microsoft YaHei", 10))
        self.auto_fit_cb.setChecked(True)
        basic_options_layout.addWidget(self.auto_fit_cb)
        
        self.style_header_cb = QCheckBox("美化表头")
        self.style_header_cb.setFont(QFont("Microsoft YaHei", 10))
        self.style_header_cb.setChecked(True)
        basic_options_layout.addWidget(self.style_header_cb)
        
        basic_options_layout.addStretch()
        options_layout.addLayout(basic_options_layout)
        
        advanced_options_layout = QHBoxLayout()
        
        self.remove_duplicates_cb = QCheckBox("删除重复行")
        self.remove_duplicates_cb.setFont(QFont("Microsoft YaHei", 10))
        advanced_options_layout.addWidget(self.remove_duplicates_cb)
        
        self.drop_empty_rows_cb = QCheckBox("删除空行")
        self.drop_empty_rows_cb.setFont(QFont("Microsoft YaHei", 10))
        advanced_options_layout.addWidget(self.drop_empty_rows_cb)
        
        self.drop_empty_cols_cb = QCheckBox("删除空列")
        self.drop_empty_cols_cb.setFont(QFont("Microsoft YaHei", 10))
        advanced_options_layout.addWidget(self.drop_empty_cols_cb)
        
        self.fill_missing_cb = QCheckBox("填充缺失值为:")
        self.fill_missing_cb.setFont(QFont("Microsoft YaHei", 10))
        advanced_options_layout.addWidget(self.fill_missing_cb)
        
        self.fill_value_edit = QLineEdit()
        self.fill_value_edit.setFont(QFont("Microsoft YaHei", 10))
        self.fill_value_edit.setPlaceholderText("0 或 空")
        self.fill_value_edit.setMaximumWidth(80)
        self.fill_value_edit.setEnabled(False)
        advanced_options_layout.addWidget(self.fill_value_edit)
        
        self.fill_missing_cb.toggled.connect(lambda checked: self.fill_value_edit.setEnabled(checked))
        
        advanced_options_layout.addStretch()
        options_layout.addLayout(advanced_options_layout)
        
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
        
        self.format_btn = QPushButton("开始格式规整")
        self.format_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.format_btn.setMinimumHeight(45)
        self.format_btn.setMinimumWidth(150)
        self.format_btn.clicked.connect(self._start_format)
        action_layout.addWidget(self.format_btn)
        
        self.open_output_btn = QPushButton("打开输出目录")
        self.open_output_btn.setFont(QFont("Microsoft YaHei", 10))
        self.open_output_btn.setMinimumHeight(45)
        self.open_output_btn.clicked.connect(self._open_output_dir)
        action_layout.addWidget(self.open_output_btn)
        
        action_layout.addStretch()
        main_layout.addLayout(action_layout)
        
    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls *.xlsm);;所有文件 (*.*)"
        )
        
        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    item = QListWidgetItem(file)
                    item.setToolTip(file)
                    self.files_list.addItem(item)
            
            self._update_files_count()
            
    def _add_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择包含Excel文件的文件夹"
        )
        
        if folder:
            excel_files = []
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.endswith(('.xlsx', '.xls', '.xlsm')):
                        full_path = os.path.join(root, file)
                        if full_path not in self.selected_files:
                            excel_files.append(full_path)
            
            if excel_files:
                for file in excel_files:
                    self.selected_files.append(file)
                    item = QListWidgetItem(file)
                    item.setToolTip(file)
                    self.files_list.addItem(item)
                
                self._update_files_count()
                QMessageBox.information(
                    self,
                    "添加完成",
                    f"已从文件夹添加 {len(excel_files)} 个Excel文件"
                )
            else:
                QMessageBox.information(
                    self,
                    "未找到文件",
                    "所选文件夹中没有找到Excel文件"
                )
                
    def _clear_files(self):
        self.selected_files.clear()
        self.files_list.clear()
        self._update_files_count()
        
    def _update_files_count(self):
        self.files_count_label.setText(f"已选择: {len(self.selected_files)} 个文件")
        
    def _browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择输出目录"
        )
        
        if folder:
            self.output_edit.setText(folder)
            
    def _get_options(self) -> dict:
        options = {
            'trim_spaces': self.trim_spaces_cb.isChecked(),
            'center_align': self.center_align_cb.isChecked(),
            'add_borders': self.add_borders_cb.isChecked(),
            'auto_fit_columns': self.auto_fit_cb.isChecked(),
            'style_header': self.style_header_cb.isChecked(),
            'apply_styles': True,
            'use_pandas': False,
            'remove_duplicates': self.remove_duplicates_cb.isChecked(),
            'drop_empty_rows': self.drop_empty_rows_cb.isChecked(),
            'drop_empty_columns': self.drop_empty_cols_cb.isChecked(),
            'fill_missing': self.fill_missing_cb.isChecked(),
            'fill_value': self.fill_value_edit.text().strip() if self.fill_missing_cb.isChecked() else ''
        }
        return options
            
    def _start_format(self):
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择要处理的Excel文件")
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
        
        self.format_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在准备处理...")
        
        options = self._get_options()
        
        self.worker = FormatWorker(
            file_paths=self.selected_files,
            output_dir=output_dir,
            options=options
        )
        
        self.worker.progress_signal.connect(self._on_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.start()
        
    def _on_progress(self, progress: int, message: str):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        
    def _on_finished(self, success: bool, message: str):
        self.format_btn.setEnabled(True)
        
        options = self._get_options()
        
        self.db_manager.add_history_record(
            operation_type="格式规整",
            source_files=self.selected_files,
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
            
        self.status_label.setText(message if success else "格式规整失败")
        
    def _open_output_dir(self):
        output_dir = self.output_edit.text().strip()
        if output_dir and os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            QMessageBox.information(self, "提示", "请先选择输出目录")
