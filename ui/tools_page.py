import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem,
    QLabel, QGroupBox, QCheckBox, QLineEdit,
    QFileDialog, QMessageBox, QProgressBar,
    QComboBox, QSpinBox, QTabWidget, QTextEdit,
    QScrollArea, QFrame
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

from core.additional_features import AdditionalFeatures
from database import DatabaseManager


class ToolsWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)
    
    def __init__(
        self,
        tool_type,
        file_paths,
        output_path,
        options
    ):
        super().__init__()
        self.tool_type = tool_type
        self.file_paths = file_paths
        self.output_path = output_path
        self.options = options
        self.tools = AdditionalFeatures()
        self.tools.set_progress_callback(self._on_progress)
        self.tools.set_error_callback(self._on_error)
        self.error_message = ""
        
    def _on_progress(self, progress: int, message: str):
        self.progress_signal.emit(progress, message)
        
    def _on_error(self, message: str):
        self.error_message = message
        
    def run(self):
        try:
            success = False
            
            if self.tool_type == "data_cleaning":
                success = self.tools.data_cleaning(
                    file_path=self.file_paths[0] if self.file_paths else "",
                    output_path=self.output_path,
                    options=self.options
                )
            elif self.tool_type == "excel_to_csv":
                success = self.tools.excel_to_csv(
                    file_path=self.file_paths[0] if self.file_paths else "",
                    output_dir=self.output_path,
                    encoding=self.options.get('encoding', 'utf-8-sig')
                )
            elif self.tool_type == "csv_to_excel":
                success = self.tools.csv_to_excel(
                    csv_files=self.file_paths,
                    output_path=self.output_path,
                    encoding=self.options.get('encoding', 'utf-8')
                )
            
            if success:
                self.finished_signal.emit(True, f"操作完成！结果已保存至: {self.output_path}")
            else:
                self.finished_signal.emit(False, self.error_message or "操作失败")
                
        except Exception as e:
            self.finished_signal.emit(False, f"操作过程中发生错误: {str(e)}")


class ToolsPage(QWidget):
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
        
        self.tab_widget = QTabWidget()
        self.tab_widget.setFont(QFont("Microsoft YaHei", 11))
        
        self._create_data_cleaning_tab()
        self._create_conversion_tab()
        
        main_layout.addWidget(self.tab_widget)
        
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
        
        self.execute_btn = QPushButton("执行操作")
        self.execute_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        self.execute_btn.setMinimumHeight(45)
        self.execute_btn.setMinimumWidth(150)
        self.execute_btn.clicked.connect(self._execute_tool)
        action_layout.addWidget(self.execute_btn)
        
        self.open_output_btn = QPushButton("打开输出目录")
        self.open_output_btn.setFont(QFont("Microsoft YaHei", 10))
        self.open_output_btn.setMinimumHeight(45)
        self.open_output_btn.clicked.connect(self._open_output_dir)
        action_layout.addWidget(self.open_output_btn)
        
        action_layout.addStretch()
        main_layout.addLayout(action_layout)
        
    def _create_data_cleaning_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(15)
        
        file_group = QGroupBox("选择Excel文件")
        file_group.setFont(QFont("Microsoft YaHei", 11))
        file_layout = QHBoxLayout(file_group)
        
        file_label = QLabel("源文件:")
        file_label.setFont(QFont("Microsoft YaHei", 10))
        file_layout.addWidget(file_label)
        
        self.cleaning_file_edit = QLineEdit()
        self.cleaning_file_edit.setFont(QFont("Microsoft YaHei", 10))
        self.cleaning_file_edit.setPlaceholderText("选择要清洗的Excel文件...")
        self.cleaning_file_edit.setReadOnly(True)
        file_layout.addWidget(self.cleaning_file_edit)
        
        self.cleaning_browse_btn = QPushButton("浏览...")
        self.cleaning_browse_btn.setFont(QFont("Microsoft YaHei", 10))
        self.cleaning_browse_btn.clicked.connect(self._browse_cleaning_file)
        file_layout.addWidget(self.cleaning_browse_btn)
        
        layout.addWidget(file_group)
        
        options_group = QGroupBox("清洗选项")
        options_group.setFont(QFont("Microsoft YaHei", 11))
        options_layout = QVBoxLayout(options_group)
        
        row1 = QHBoxLayout()
        
        self.cleaning_duplicates_cb = QCheckBox("删除重复行")
        self.cleaning_duplicates_cb.setFont(QFont("Microsoft YaHei", 10))
        row1.addWidget(self.cleaning_duplicates_cb)
        
        self.cleaning_empty_rows_cb = QCheckBox("删除空行")
        self.cleaning_empty_rows_cb.setFont(QFont("Microsoft YaHei", 10))
        row1.addWidget(self.cleaning_empty_rows_cb)
        
        self.cleaning_empty_cols_cb = QCheckBox("删除空列")
        self.cleaning_empty_cols_cb.setFont(QFont("Microsoft YaHei", 10))
        row1.addWidget(self.cleaning_empty_cols_cb)
        
        row1.addStretch()
        options_layout.addLayout(row1)
        
        row2 = QHBoxLayout()
        
        self.cleaning_trim_cb = QCheckBox("清除多余空格")
        self.cleaning_trim_cb.setFont(QFont("Microsoft YaHei", 10))
        self.cleaning_trim_cb.setChecked(True)
        row2.addWidget(self.cleaning_trim_cb)
        
        self.cleaning_fill_cb = QCheckBox("填充缺失值为:")
        self.cleaning_fill_cb.setFont(QFont("Microsoft YaHei", 10))
        row2.addWidget(self.cleaning_fill_cb)
        
        self.cleaning_fill_edit = QLineEdit()
        self.cleaning_fill_edit.setFont(QFont("Microsoft YaHei", 10))
        self.cleaning_fill_edit.setPlaceholderText("0 或 空")
        self.cleaning_fill_edit.setMaximumWidth(100)
        self.cleaning_fill_edit.setEnabled(False)
        row2.addWidget(self.cleaning_fill_edit)
        
        self.cleaning_fill_cb.toggled.connect(lambda checked: self.cleaning_fill_edit.setEnabled(checked))
        
        row2.addStretch()
        options_layout.addLayout(row2)
        
        layout.addWidget(options_group)
        
        output_group = QGroupBox("输出设置")
        output_group.setFont(QFont("Microsoft YaHei", 11))
        output_layout = QHBoxLayout(output_group)
        
        output_label = QLabel("输出文件:")
        output_label.setFont(QFont("Microsoft YaHei", 10))
        output_layout.addWidget(output_label)
        
        self.cleaning_output_edit = QLineEdit()
        self.cleaning_output_edit.setFont(QFont("Microsoft YaHei", 10))
        self.cleaning_output_edit.setPlaceholderText("选择输出文件路径...")
        output_layout.addWidget(self.cleaning_output_edit)
        
        self.cleaning_output_browse_btn = QPushButton("浏览...")
        self.cleaning_output_browse_btn.setFont(QFont("Microsoft YaHei", 10))
        self.cleaning_output_browse_btn.clicked.connect(self._browse_cleaning_output)
        output_layout.addWidget(self.cleaning_output_browse_btn)
        
        layout.addWidget(output_group)
        
        layout.addStretch()
        
        self.tab_widget.addTab(tab, "🧹 数据清洗")
        
    def _create_conversion_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(15)
        
        type_group = QGroupBox("转换类型")
        type_group.setFont(QFont("Microsoft YaHei", 11))
        type_layout = QHBoxLayout(type_group)
        
        type_label = QLabel("转换方向:")
        type_label.setFont(QFont("Microsoft YaHei", 10))
        type_layout.addWidget(type_label)
        
        self.conversion_combo = QComboBox()
        self.conversion_combo.setFont(QFont("Microsoft YaHei", 10))
        self.conversion_combo.setMinimumWidth(200)
        self.conversion_combo.addItems(["Excel 转 CSV", "CSV 转 Excel"])
        self.conversion_combo.currentIndexChanged.connect(self._on_conversion_type_changed)
        type_layout.addWidget(self.conversion_combo)
        
        encoding_label = QLabel("编码:")
        encoding_label.setFont(QFont("Microsoft YaHei", 10))
        type_layout.addWidget(encoding_label)
        
        self.encoding_combo = QComboBox()
        self.encoding_combo.setFont(QFont("Microsoft YaHei", 10))
        self.encoding_combo.setMinimumWidth(120)
        self.encoding_combo.addItems(["utf-8", "utf-8-sig", "gbk", "gb2312"])
        self.encoding_combo.setCurrentIndex(1)
        type_layout.addWidget(self.encoding_combo)
        
        type_layout.addStretch()
        layout.addWidget(type_group)
        
        files_group = QGroupBox("选择文件")
        files_group.setFont(QFont("Microsoft YaHei", 11))
        files_layout = QVBoxLayout(files_group)
        
        files_btn_layout = QHBoxLayout()
        
        self.conversion_add_files_btn = QPushButton("添加文件")
        self.conversion_add_files_btn.setFont(QFont("Microsoft YaHei", 10))
        self.conversion_add_files_btn.setMinimumHeight(35)
        self.conversion_add_files_btn.clicked.connect(self._add_conversion_files)
        files_btn_layout.addWidget(self.conversion_add_files_btn)
        
        self.conversion_clear_btn = QPushButton("清空列表")
        self.conversion_clear_btn.setFont(QFont("Microsoft YaHei", 10))
        self.conversion_clear_btn.setMinimumHeight(35)
        self.conversion_clear_btn.clicked.connect(self._clear_conversion_files)
        files_btn_layout.addWidget(self.conversion_clear_btn)
        
        files_btn_layout.addStretch()
        files_layout.addLayout(files_btn_layout)
        
        self.conversion_files_list = QListWidget()
        self.conversion_files_list.setFont(QFont("Microsoft YaHei", 9))
        self.conversion_files_list.setMinimumHeight(100)
        files_layout.addWidget(self.conversion_files_list)
        
        self.conversion_files_count_label = QLabel("已选择: 0 个文件")
        self.conversion_files_count_label.setFont(QFont("Microsoft YaHei", 9))
        files_layout.addWidget(self.conversion_files_count_label)
        
        layout.addWidget(files_group)
        
        output_group = QGroupBox("输出设置")
        output_group.setFont(QFont("Microsoft YaHei", 11))
        output_layout = QHBoxLayout(output_group)
        
        output_label = QLabel("输出:")
        output_label.setFont(QFont("Microsoft YaHei", 10))
        output_layout.addWidget(output_label)
        
        self.conversion_output_edit = QLineEdit()
        self.conversion_output_edit.setFont(QFont("Microsoft YaHei", 10))
        self.conversion_output_edit.setPlaceholderText("选择输出路径...")
        output_layout.addWidget(self.conversion_output_edit)
        
        self.conversion_output_browse_btn = QPushButton("浏览...")
        self.conversion_output_browse_btn.setFont(QFont("Microsoft YaHei", 10))
        self.conversion_output_browse_btn.clicked.connect(self._browse_conversion_output)
        output_layout.addWidget(self.conversion_output_browse_btn)
        
        layout.addWidget(output_group)
        
        layout.addStretch()
        
        self.tab_widget.addTab(tab, "🔄 格式转换")
        
    def _on_conversion_type_changed(self, index):
        if index == 0:
            self.conversion_add_files_btn.setText("添加Excel文件")
        else:
            self.conversion_add_files_btn.setText("添加CSV文件")
        
    def _browse_cleaning_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls *.xlsm);;所有文件 (*.*)"
        )
        
        if file_path:
            self.cleaning_file_edit.setText(file_path)
            
            if not self.cleaning_output_edit.text():
                base, ext = os.path.splitext(file_path)
                self.cleaning_output_edit.setText(f"{base}_cleaned{ext}")
                
    def _browse_cleaning_output(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存清洗后的文件",
            "",
            "Excel文件 (*.xlsx);;所有文件 (*.*)"
        )
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.cleaning_output_edit.setText(file_path)
            
    def _add_conversion_files(self):
        conversion_type = self.conversion_combo.currentIndex()
        
        if conversion_type == 0:
            filter_str = "Excel文件 (*.xlsx *.xls *.xlsm);;所有文件 (*.*)"
        else:
            filter_str = "CSV文件 (*.csv);;所有文件 (*.*)"
        
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文件",
            "",
            filter_str
        )
        
        if files:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    item = QListWidgetItem(file)
                    item.setToolTip(file)
                    self.conversion_files_list.addItem(item)
            
            self._update_conversion_files_count()
            
    def _clear_conversion_files(self):
        self.selected_files.clear()
        self.conversion_files_list.clear()
        self._update_conversion_files_count()
        
    def _update_conversion_files_count(self):
        self.conversion_files_count_label.setText(f"已选择: {len(self.selected_files)} 个文件")
        
    def _browse_conversion_output(self):
        conversion_type = self.conversion_combo.currentIndex()
        
        if conversion_type == 0:
            folder = QFileDialog.getExistingDirectory(
                self,
                "选择输出目录"
            )
            if folder:
                self.conversion_output_edit.setText(folder)
        else:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存Excel文件",
                "",
                "Excel文件 (*.xlsx);;所有文件 (*.*)"
            )
            if file_path:
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                self.conversion_output_edit.setText(file_path)
                
    def _get_current_tool(self):
        current_tab = self.tab_widget.currentIndex()
        
        if current_tab == 0:
            return "data_cleaning"
        else:
            conversion_type = self.conversion_combo.currentIndex()
            if conversion_type == 0:
                return "excel_to_csv"
            else:
                return "csv_to_excel"
                
    def _execute_tool(self):
        tool_type = self._get_current_tool()
        
        if tool_type == "data_cleaning":
            source_file = self.cleaning_file_edit.text().strip()
            if not source_file:
                QMessageBox.warning(self, "警告", "请先选择要清洗的Excel文件")
                return
            
            if not os.path.exists(source_file):
                QMessageBox.warning(self, "警告", "所选文件不存在")
                return
            
            output_path = self.cleaning_output_edit.text().strip()
            if not output_path:
                QMessageBox.warning(self, "警告", "请选择输出文件路径")
                return
            
            files_to_process = [source_file]
            
            options = {
                'remove_duplicates': self.cleaning_duplicates_cb.isChecked(),
                'drop_empty_rows': self.cleaning_empty_rows_cb.isChecked(),
                'drop_empty_columns': self.cleaning_empty_cols_cb.isChecked(),
                'trim_spaces': self.cleaning_trim_cb.isChecked(),
                'fill_missing': self.cleaning_fill_cb.isChecked(),
                'fill_value': self.cleaning_fill_edit.text().strip() if self.cleaning_fill_cb.isChecked() else ''
            }
            
        else:
            if not self.selected_files:
                QMessageBox.warning(self, "警告", "请先选择要转换的文件")
                return
            
            output_path = self.conversion_output_edit.text().strip()
            if not output_path:
                QMessageBox.warning(self, "警告", "请选择输出路径")
                return
            
            files_to_process = self.selected_files
            
            options = {
                'encoding': self.encoding_combo.currentText()
            }
        
        self.execute_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在准备执行...")
        
        self.worker = ToolsWorker(
            tool_type=tool_type,
            file_paths=files_to_process,
            output_path=output_path,
            options=options
        )
        
        self.worker.progress_signal.connect(self._on_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.start()
        
    def _on_progress(self, progress: int, message: str):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        
    def _on_finished(self, success: bool, message: str):
        self.execute_btn.setEnabled(True)
        
        tool_type = self._get_current_tool()
        
        if tool_type == "data_cleaning":
            source_files = [self.cleaning_file_edit.text().strip()]
            target_file = self.cleaning_output_edit.text().strip()
            operation_name = "数据清洗"
        else:
            source_files = self.selected_files
            target_file = self.conversion_output_edit.text().strip()
            if tool_type == "excel_to_csv":
                operation_name = "Excel转CSV"
            else:
                operation_name = "CSV转Excel"
        
        self.db_manager.add_history_record(
            operation_type=operation_name,
            source_files=source_files,
            target_file=target_file,
            options={'tool_type': tool_type},
            status="completed" if success else "failed",
            message=message
        )
        
        if success:
            QMessageBox.information(self, "成功", message)
            self.progress_bar.setValue(100)
        else:
            QMessageBox.critical(self, "失败", message)
            self.progress_bar.setValue(0)
            
        self.status_label.setText(message if success else "操作失败")
        
    def _open_output_dir(self):
        tool_type = self._get_current_tool()
        
        if tool_type == "data_cleaning":
            output_path = self.cleaning_output_edit.text().strip()
        else:
            output_path = self.conversion_output_edit.text().strip()
        
        if output_path:
            if os.path.isfile(output_path):
                dir_path = os.path.dirname(output_path)
            else:
                dir_path = output_path
            
            if os.path.exists(dir_path):
                os.startfile(dir_path)
                return
        
        QMessageBox.information(self, "提示", "请先选择输出路径")
