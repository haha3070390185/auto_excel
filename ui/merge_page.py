import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem,
    QLabel, QGroupBox, QRadioButton, QCheckBox,
    QLineEdit, QFileDialog, QMessageBox, QSpinBox,
    QProgressBar, QFrame
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

from core.excel_merger import ExcelMerger
from database import DatabaseManager


class MergeWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(
        self,
        file_paths,
        output_path,
        merge_type,
        sheet_name,
        add_source_column,
        skip_rows,
        header_row
    ):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path
        self.merge_type = merge_type
        self.sheet_name = sheet_name
        self.add_source_column = add_source_column
        self.skip_rows = skip_rows
        self.header_row = header_row
        self.merger = ExcelMerger()
        self.merger.set_progress_callback(self._on_progress)
        self.merger.set_error_callback(self._on_error)
        self.error_message = ""
        
    def _on_progress(self, progress: int, message: str):
        self.progress_signal.emit(progress, message)
        
    def _on_error(self, message: str):
        self.error_message = message
        
    def run(self):
        try:
            success = self.merger.merge_files(
                file_paths=self.file_paths,
                output_path=self.output_path,
                merge_type=self.merge_type,
                sheet_name=self.sheet_name,
                add_source_column=self.add_source_column,
                skip_rows=self.skip_rows,
                header_row=self.header_row
            )
            
            if success:
                self.finished_signal.emit(True, f"合并成功！文件已保存至: {self.output_path}")
            else:
                self.finished_signal.emit(False, self.error_message or "合并失败")
                
        except Exception as e:
            self.finished_signal.emit(False, f"合并过程中发生错误: {str(e)}")


class MergePage(QWidget):
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
        self.files_list.setMinimumHeight(150)
        self.files_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        files_layout.addWidget(self.files_list)
        
        self.files_count_label = QLabel("已选择: 0 个文件")
        self.files_count_label.setFont(QFont("Microsoft YaHei", 9))
        files_layout.addWidget(self.files_count_label)
        
        main_layout.addWidget(files_group)
        
        options_group = QGroupBox("合并选项")
        options_group.setFont(QFont("Microsoft YaHei", 11))
        options_layout = QVBoxLayout(options_group)
        
        merge_type_layout = QHBoxLayout()
        merge_type_label = QLabel("合并方式:")
        merge_type_label.setFont(QFont("Microsoft YaHei", 10))
        merge_type_layout.addWidget(merge_type_label)
        
        self.merge_first_sheet_rb = QRadioButton("仅合并第一个工作表")
        self.merge_first_sheet_rb.setFont(QFont("Microsoft YaHei", 10))
        self.merge_first_sheet_rb.setChecked(True)
        merge_type_layout.addWidget(self.merge_first_sheet_rb)
        
        self.merge_all_sheets_rb = QRadioButton("合并所有工作表")
        self.merge_all_sheets_rb.setFont(QFont("Microsoft YaHei", 10))
        merge_type_layout.addWidget(self.merge_all_sheets_rb)
        
        self.merge_specific_rb = QRadioButton("合并指定工作表:")
        self.merge_specific_rb.setFont(QFont("Microsoft YaHei", 10))
        merge_type_layout.addWidget(self.merge_specific_rb)
        
        self.sheet_name_edit = QLineEdit()
        self.sheet_name_edit.setFont(QFont("Microsoft YaHei", 10))
        self.sheet_name_edit.setPlaceholderText("输入工作表名称")
        self.sheet_name_edit.setEnabled(False)
        merge_type_layout.addWidget(self.sheet_name_edit)
        
        self.merge_specific_rb.toggled.connect(lambda checked: self.sheet_name_edit.setEnabled(checked))
        
        merge_type_layout.addStretch()
        options_layout.addLayout(merge_type_layout)
        
        advanced_layout = QHBoxLayout()
        
        self.add_source_col_cb = QCheckBox("添加来源文件列")
        self.add_source_col_cb.setFont(QFont("Microsoft YaHei", 10))
        advanced_layout.addWidget(self.add_source_col_cb)
        
        skip_rows_label = QLabel("跳过前N行:")
        skip_rows_label.setFont(QFont("Microsoft YaHei", 10))
        advanced_layout.addWidget(skip_rows_label)
        
        self.skip_rows_spin = QSpinBox()
        self.skip_rows_spin.setFont(QFont("Microsoft YaHei", 10))
        self.skip_rows_spin.setRange(0, 100)
        self.skip_rows_spin.setValue(0)
        advanced_layout.addWidget(self.skip_rows_spin)
        
        header_label = QLabel("表头行:")
        header_label.setFont(QFont("Microsoft YaHei", 10))
        advanced_layout.addWidget(header_label)
        
        self.header_row_spin = QSpinBox()
        self.header_row_spin.setFont(QFont("Microsoft YaHei", 10))
        self.header_row_spin.setRange(0, 10)
        self.header_row_spin.setValue(0)
        advanced_layout.addWidget(self.header_row_spin)
        
        advanced_layout.addStretch()
        options_layout.addLayout(advanced_layout)
        
        main_layout.addWidget(options_group)
        
        output_group = QGroupBox("输出设置")
        output_group.setFont(QFont("Microsoft YaHei", 11))
        output_layout = QHBoxLayout(output_group)
        
        output_label = QLabel("输出文件:")
        output_label.setFont(QFont("Microsoft YaHei", 10))
        output_layout.addWidget(output_label)
        
        self.output_edit = QLineEdit()
        self.output_edit.setFont(QFont("Microsoft YaHei", 10))
        self.output_edit.setPlaceholderText("选择输出文件路径...")
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
        
        self.merge_btn = QPushButton("开始合并")
        self.merge_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.merge_btn.setMinimumHeight(45)
        self.merge_btn.setMinimumWidth(150)
        self.merge_btn.clicked.connect(self._start_merge)
        action_layout.addWidget(self.merge_btn)
        
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
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存合并后的文件",
            "",
            "Excel文件 (*.xlsx);;所有文件 (*.*)"
        )
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_edit.setText(file_path)
            
    def _get_merge_type(self):
        if self.merge_first_sheet_rb.isChecked():
            return "first_sheet"
        elif self.merge_all_sheets_rb.isChecked():
            return "all_sheets"
        else:
            return "specific_sheet"
            
    def _start_merge(self):
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择要合并的Excel文件")
            return
        
        output_path = self.output_edit.text().strip()
        if not output_path:
            QMessageBox.warning(self, "警告", "请选择输出文件路径")
            return
        
        merge_type = self._get_merge_type()
        sheet_name = self.sheet_name_edit.text().strip() if self.merge_specific_rb.isChecked() else None
        
        if merge_type == "specific_sheet" and not sheet_name:
            QMessageBox.warning(self, "警告", "请输入要合并的工作表名称")
            return
        
        self.merge_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在准备合并...")
        
        self.worker = MergeWorker(
            file_paths=self.selected_files,
            output_path=output_path,
            merge_type=merge_type,
            sheet_name=sheet_name,
            add_source_column=self.add_source_col_cb.isChecked(),
            skip_rows=self.skip_rows_spin.value(),
            header_row=self.header_row_spin.value()
        )
        
        self.worker.progress_signal.connect(self._on_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.start()
        
    def _on_progress(self, progress: int, message: str):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        
    def _on_finished(self, success: bool, message: str):
        self.merge_btn.setEnabled(True)
        
        options = {
            'merge_type': self._get_merge_type(),
            'add_source_column': self.add_source_col_cb.isChecked(),
            'skip_rows': self.skip_rows_spin.value(),
            'header_row': self.header_row_spin.value()
        }
        
        self.db_manager.add_history_record(
            operation_type="Excel合并",
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
            
        self.status_label.setText(message if success else "合并失败")
        
    def _open_output_dir(self):
        output_path = self.output_edit.text().strip()
        if output_path and os.path.exists(os.path.dirname(output_path)):
            os.startfile(os.path.dirname(output_path))
        else:
            QMessageBox.information(self, "提示", "请先选择输出文件路径")
