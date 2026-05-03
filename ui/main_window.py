import sys
import os
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QMessageBox, QStatusBar, QProgressBar,
    QLabel, QFrame
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QFont, QIcon, QAction

from ui.merge_page import MergePage
from ui.format_page import FormatPage
from ui.split_page import SplitPage
from ui.tools_page import ToolsPage
from ui.history_page import HistoryPage

from database import DatabaseManager


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.db_manager = DatabaseManager()
        
        self.setWindowTitle("Excel办公自动化工具")
        self.setMinimumSize(1200, 800)
        self.resize(1400, 900)
        
        self._init_ui()
        self._create_menu_bar()
        self._create_status_bar()
        
    def _init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        header_frame = QFrame()
        header_frame.setFrameShape(QFrame.Shape.StyledPanel)
        header_frame.setMinimumHeight(60)
        
        header_layout = QHBoxLayout(header_frame)
        
        title_label = QLabel("Excel办公自动化工具")
        title_font = QFont("Microsoft YaHei", 18, QFont.Weight.Bold)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        header_layout.addWidget(title_label)
        
        main_layout.addWidget(header_frame)
        
        self.tab_widget = QTabWidget()
        self.tab_widget.setFont(QFont("Microsoft YaHei", 11))
        self.tab_widget.setDocumentMode(True)
        self.tab_widget.setMovable(True)
        
        self.merge_page = MergePage(self.db_manager)
        self.tab_widget.addTab(self.merge_page, "📂 Excel合并")
        
        self.format_page = FormatPage(self.db_manager)
        self.tab_widget.addTab(self.format_page, "🎨 格式规整")
        
        self.split_page = SplitPage(self.db_manager)
        self.tab_widget.addTab(self.split_page, "✂️ 表格拆分")
        
        self.tools_page = ToolsPage(self.db_manager)
        self.tab_widget.addTab(self.tools_page, "🔧 更多工具")
        
        self.history_page = HistoryPage(self.db_manager)
        self.tab_widget.addTab(self.history_page, "📋 操作历史")
        
        main_layout.addWidget(self.tab_widget)
        
    def _create_menu_bar(self):
        menubar = self.menuBar()
        menubar.setFont(QFont("Microsoft YaHei", 10))
        
        file_menu = menubar.addMenu("文件(&F)")
        
        exit_action = QAction("退出(&X)", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        edit_menu = menubar.addMenu("编辑(&E)")
        
        clear_history_action = QAction("清空历史记录", self)
        clear_history_action.triggered.connect(self._clear_history)
        edit_menu.addAction(clear_history_action)
        
        help_menu = menubar.addMenu("帮助(&H)")
        
        about_action = QAction("关于(&A)", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)
        
    def _create_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        
        self.status_progress = QProgressBar()
        self.status_progress.setMaximumWidth(200)
        self.status_progress.hide()
        
        self.status_bar.addPermanentWidget(self.status_progress)
        
        self.status_label = QLabel("就绪")
        self.status_bar.addWidget(self.status_label)
        
    def update_status(self, message: str):
        self.status_label.setText(message)
        
    def update_progress(self, value: int):
        if value < 0 or value > 100:
            self.status_progress.hide()
        else:
            self.status_progress.show()
            self.status_progress.setValue(value)
            
    def _clear_history(self):
        reply = QMessageBox.question(
            self,
            "确认清空",
            "确定要清空所有历史记录吗？此操作不可恢复。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            if self.db_manager.clear_all_history():
                QMessageBox.information(self, "成功", "历史记录已清空")
                self.history_page.refresh_history()
            else:
                QMessageBox.warning(self, "失败", "清空历史记录失败")
                
    def _show_about(self):
        QMessageBox.about(
            self,
            "关于 Excel办公自动化工具",
            """
            <h2>Excel办公自动化工具</h2>
            <p>版本: 1.0.0</p>
            <p>基于 PyQt6 开发的Excel办公自动化工具</p>
            <h3>主要功能:</h3>
            <ul>
                <li>Excel文件批量合并</li>
                <li>批量格式规整</li>
                <li>表格按条件拆分</li>
                <li>数据清洗与转换</li>
                <li>操作历史记录</li>
            </ul>
            <p>© 2024 Excel Automation Tool</p>
            """
        )
        
    def closeEvent(self, event):
        reply = QMessageBox.question(
            self,
            "确认退出",
            "确定要退出程序吗？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()
