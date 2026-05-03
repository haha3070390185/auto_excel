import os
from datetime import datetime
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem,
    QLabel, QGroupBox, QMessageBox, QHeaderView,
    QComboBox, QDateEdit, QSplitter, QTextEdit,
    QFrame, QSpinBox
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont, QColor

from database import DatabaseManager, HistoryRecord


class HistoryPage(QWidget):
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.current_records = []
        
        self._init_ui()
        self.refresh_history()
        
    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        filter_group = QGroupBox("筛选条件")
        filter_group.setFont(QFont("Microsoft YaHei", 11))
        filter_layout = QHBoxLayout(filter_group)
        
        type_label = QLabel("操作类型:")
        type_label.setFont(QFont("Microsoft YaHei", 10))
        filter_layout.addWidget(type_label)
        
        self.type_combo = QComboBox()
        self.type_combo.setFont(QFont("Microsoft YaHei", 10))
        self.type_combo.setMinimumWidth(150)
        self.type_combo.addItems(["全部", "Excel合并", "格式规整", "表格拆分", "数据清洗", "Excel转CSV", "CSV转Excel"])
        filter_layout.addWidget(self.type_combo)
        
        date_label = QLabel("日期范围:")
        date_label.setFont(QFont("Microsoft YaHei", 10))
        filter_layout.addWidget(date_label)
        
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setFont(QFont("Microsoft YaHei", 10))
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate().addDays(-30))
        filter_layout.addWidget(self.start_date_edit)
        
        to_label = QLabel("至")
        to_label.setFont(QFont("Microsoft YaHei", 10))
        filter_layout.addWidget(to_label)
        
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setFont(QFont("Microsoft YaHei", 10))
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate())
        filter_layout.addWidget(self.end_date_edit)
        
        limit_label = QLabel("显示数量:")
        limit_label.setFont(QFont("Microsoft YaHei", 10))
        filter_layout.addWidget(limit_label)
        
        self.limit_spin = QSpinBox()
        self.limit_spin.setFont(QFont("Microsoft YaHei", 10))
        self.limit_spin.setRange(10, 1000)
        self.limit_spin.setValue(100)
        filter_layout.addWidget(self.limit_spin)
        
        self.filter_btn = QPushButton("筛选")
        self.filter_btn.setFont(QFont("Microsoft YaHei", 10))
        self.filter_btn.setMinimumHeight(30)
        self.filter_btn.clicked.connect(self._filter_history)
        filter_layout.addWidget(self.filter_btn)
        
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.setFont(QFont("Microsoft YaHei", 10))
        self.refresh_btn.setMinimumHeight(30)
        self.refresh_btn.clicked.connect(self.refresh_history)
        filter_layout.addWidget(self.refresh_btn)
        
        filter_layout.addStretch()
        main_layout.addWidget(filter_group)
        
        stats_group = QGroupBox("统计信息")
        stats_group.setFont(QFont("Microsoft YaHei", 11))
        stats_layout = QHBoxLayout(stats_group)
        
        self.total_label = QLabel("总操作次数: 0")
        self.total_label.setFont(QFont("Microsoft YaHei", 10))
        stats_layout.addWidget(self.total_label)
        
        self.merge_label = QLabel("Excel合并: 0")
        self.merge_label.setFont(QFont("Microsoft YaHei", 10))
        stats_layout.addWidget(self.merge_label)
        
        self.format_label = QLabel("格式规整: 0")
        self.format_label.setFont(QFont("Microsoft YaHei", 10))
        stats_layout.addWidget(self.format_label)
        
        self.split_label = QLabel("表格拆分: 0")
        self.split_label.setFont(QFont("Microsoft YaHei", 10))
        stats_layout.addWidget(self.split_label)
        
        stats_layout.addStretch()
        main_layout.addWidget(stats_group)
        
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        table_group = QGroupBox("历史记录列表")
        table_group.setFont(QFont("Microsoft YaHei", 11))
        table_layout = QVBoxLayout(table_group)
        
        self.history_table = QTableWidget()
        self.history_table.setFont(QFont("Microsoft YaHei", 9))
        self.history_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.history_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.setColumnCount(7)
        self.history_table.setHorizontalHeaderLabels([
            "ID", "操作类型", "源文件数量", "目标文件", "时间", "状态", "消息"
        ])
        
        header = self.history_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
        
        self.history_table.setColumnWidth(0, 50)
        self.history_table.setColumnWidth(1, 100)
        self.history_table.setColumnWidth(2, 90)
        self.history_table.setColumnWidth(4, 150)
        self.history_table.setColumnWidth(5, 70)
        
        self.history_table.itemSelectionChanged.connect(self._on_record_selected)
        
        table_layout.addWidget(self.history_table)
        
        splitter.addWidget(table_group)
        
        detail_group = QGroupBox("记录详情")
        detail_group.setFont(QFont("Microsoft YaHei", 11))
        detail_layout = QVBoxLayout(detail_group)
        
        self.detail_text = QTextEdit()
        self.detail_text.setFont(QFont("Microsoft YaHei", 9))
        self.detail_text.setReadOnly(True)
        detail_layout.addWidget(self.detail_text)
        
        splitter.addWidget(detail_group)
        
        splitter.setSizes([500, 200])
        main_layout.addWidget(splitter)
        
        action_layout = QHBoxLayout()
        action_layout.addStretch()
        
        self.delete_btn = QPushButton("删除选中记录")
        self.delete_btn.setFont(QFont("Microsoft YaHei", 10))
        self.delete_btn.setMinimumHeight(35)
        self.delete_btn.clicked.connect(self._delete_selected_record)
        action_layout.addWidget(self.delete_btn)
        
        self.clear_all_btn = QPushButton("清空所有记录")
        self.clear_all_btn.setFont(QFont("Microsoft YaHei", 10))
        self.clear_all_btn.setMinimumHeight(35)
        self.clear_all_btn.clicked.connect(self._clear_all_records)
        action_layout.addWidget(self.clear_all_btn)
        
        self.open_dir_btn = QPushButton("打开目标文件夹")
        self.open_dir_btn.setFont(QFont("Microsoft YaHei", 10))
        self.open_dir_btn.setMinimumHeight(35)
        self.open_dir_btn.clicked.connect(self._open_target_directory)
        action_layout.addWidget(self.open_dir_btn)
        
        action_layout.addStretch()
        main_layout.addLayout(action_layout)
        
    def refresh_history(self):
        limit = self.limit_spin.value()
        records = self.db_manager.get_all_history(limit)
        self._display_records(records)
        self._update_statistics()
        
    def _filter_history(self):
        operation_type = self.type_combo.currentText()
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()
        
        start_datetime = datetime.combine(start_date, datetime.min.time())
        end_datetime = datetime.combine(end_date, datetime.max.time())
        
        if operation_type == "全部":
            records = self.db_manager.get_history_by_date_range(start_datetime, end_datetime)
        else:
            all_records = self.db_manager.get_history_by_type(operation_type, self.limit_spin.value())
            records = [
                r for r in all_records
                if start_datetime <= r.timestamp <= end_datetime
            ]
        
        self._display_records(records)
        
    def _display_records(self, records):
        self.current_records = records
        self.history_table.setRowCount(len(records))
        
        for row, record in enumerate(records):
            id_item = QTableWidgetItem(str(record.id))
            id_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.history_table.setItem(row, 0, id_item)
            
            type_item = QTableWidgetItem(record.operation_type)
            type_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.history_table.setItem(row, 1, type_item)
            
            source_files = self.db_manager.parse_source_files(record)
            count_item = QTableWidgetItem(str(len(source_files)))
            count_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.history_table.setItem(row, 2, count_item)
            
            target_item = QTableWidgetItem(record.target_file or "-")
            self.history_table.setItem(row, 3, target_item)
            
            time_str = record.timestamp.strftime("%Y-%m-%d %H:%M:%S")
            time_item = QTableWidgetItem(time_str)
            time_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.history_table.setItem(row, 4, time_item)
            
            status_item = QTableWidgetItem(record.status)
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if record.status == "completed":
                status_item.setForeground(QColor(0, 128, 0))
            else:
                status_item.setForeground(QColor(255, 0, 0))
            self.history_table.setItem(row, 5, status_item)
            
            msg_item = QTableWidgetItem(record.message or "-")
            self.history_table.setItem(row, 6, msg_item)
            
    def _update_statistics(self):
        stats = self.db_manager.get_statistics()
        
        self.total_label.setText(f"总操作次数: {stats['total_count']}")
        self.merge_label.setText(f"Excel合并: {stats['by_type'].get('Excel合并', 0)}")
        self.format_label.setText(f"格式规整: {stats['by_type'].get('格式规整', 0)}")
        self.split_label.setText(f"表格拆分: {stats['by_type'].get('表格拆分', 0)}")
        
    def _on_record_selected(self):
        selected_rows = self.history_table.selectedItems()
        if not selected_rows:
            self.detail_text.clear()
            return
        
        row = selected_rows[0].row()
        if row >= len(self.current_records):
            return
        
        record = self.current_records[row]
        source_files = self.db_manager.parse_source_files(record)
        options = self.db_manager.parse_options(record)
        
        detail_text = f"""操作类型: {record.operation_type}
操作时间: {record.timestamp.strftime('%Y-%m-%d %H:%M:%S')}
状态: {record.status}
消息: {record.message or '-'}

源文件列表 ({len(source_files)} 个文件):
"""
        for i, f in enumerate(source_files, 1):
            detail_text += f"  {i}. {f}\n"
        
        detail_text += f"\n目标文件: {record.target_file or '-'}\n"
        
        if options:
            detail_text += "\n操作选项:\n"
            for key, value in options.items():
                detail_text += f"  {key}: {value}\n"
        
        self.detail_text.setText(detail_text)
        
    def _delete_selected_record(self):
        selected_rows = self.history_table.selectedItems()
        if not selected_rows:
            QMessageBox.information(self, "提示", "请先选择要删除的记录")
            return
        
        row = selected_rows[0].row()
        if row >= len(self.current_records):
            return
        
        record = self.current_records[row]
        
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除记录 ID={record.id} 吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            if self.db_manager.delete_history_record(record.id):
                QMessageBox.information(self, "成功", "记录已删除")
                self.refresh_history()
            else:
                QMessageBox.warning(self, "失败", "删除记录失败")
                
    def _clear_all_records(self):
        reply = QMessageBox.question(
            self,
            "确认清空",
            "确定要清空所有历史记录吗？此操作不可恢复！",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            if self.db_manager.clear_all_history():
                QMessageBox.information(self, "成功", "所有记录已清空")
                self.refresh_history()
            else:
                QMessageBox.warning(self, "失败", "清空记录失败")
                
    def _open_target_directory(self):
        selected_rows = self.history_table.selectedItems()
        if not selected_rows:
            QMessageBox.information(self, "提示", "请先选择一条记录")
            return
        
        row = selected_rows[0].row()
        if row >= len(self.current_records):
            return
        
        record = self.current_records[row]
        
        target_file = record.target_file
        if not target_file:
            QMessageBox.information(self, "提示", "该记录没有目标文件路径")
            return
        
        if os.path.isfile(target_file):
            dir_path = os.path.dirname(target_file)
            if os.path.exists(dir_path):
                os.startfile(dir_path)
                return
        elif os.path.isdir(target_file):
            if os.path.exists(target_file):
                os.startfile(target_file)
                return
        
        QMessageBox.information(self, "提示", "目标路径不存在或已被删除")
