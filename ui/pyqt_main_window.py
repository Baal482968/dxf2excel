#!/usr/bin/env python3
"""
PyQt6 主視窗 - 解決執行緒問題
"""

import sys
import os
import threading
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QLineEdit, QPushButton, QProgressBar, 
    QFileDialog, QMessageBox, QFrame, QGroupBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QPalette, QColor
from core.cad_reader import CADReader
from core.excel_writer import ExcelWriter
from utils.helpers import get_file_info

class ConversionWorker(QThread):
    """轉換工作執行緒"""
    progress_updated = pyqtSignal(int, str)  # 進度, 狀態
    conversion_completed = pyqtSignal(dict)  # 結果資料
    error_occurred = pyqtSignal(str)  # 錯誤訊息
    
    def __init__(self, cad_file_path, excel_file_path):
        super().__init__()
        self.cad_file_path = cad_file_path
        self.excel_file_path = excel_file_path
        self.cad_reader = CADReader()
        self.excel_writer = ExcelWriter()
    
    def run(self):
        """執行轉換"""
        try:
            # 開啟 CAD 檔案
            self.progress_updated.emit(20, "正在開啟 CAD 檔案...")
            if not self.cad_reader.open_file(self.cad_file_path):
                self.error_occurred.emit("無法開啟 CAD 檔案")
                return
            
            # 處理圖面
            self.progress_updated.emit(40, "正在處理圖面...")
            rebar_data = self.cad_reader.process_drawing()
            if not rebar_data:
                self.error_occurred.emit("處理圖面失敗")
                return
            
            # 生成 Excel
            self.progress_updated.emit(60, "正在生成 Excel 檔案...")
            self.excel_writer.create_workbook()
            self.excel_writer.write_multi_sheet_rebar_data(rebar_data)
            self.excel_writer.save_workbook(self.excel_file_path)
            
            # 清理資源
            self.cad_reader.close_file()
            
            # 完成
            self.progress_updated.emit(100, "轉換完成！")
            self.conversion_completed.emit(rebar_data)
            
        except Exception as e:
            self.error_occurred.emit(f"轉換過程發生錯誤：{str(e)}")

class PyQtMainWindow(QMainWindow):
    """PyQt6 主視窗類別"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CAD 鋼筋計料轉換工具")
        self.setGeometry(100, 100, 800, 600)
        self.setup_ui()
        self.setup_styles()
        
        # 初始化變數
        self.cad_file_path = ""
        self.excel_file_path = ""
        self.conversion_worker = None
    
    def setup_ui(self):
        """建立使用者介面"""
        # 中央元件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主佈局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)
        
        # 標題
        title_label = QLabel("CAD 鋼筋計料轉換工具")
        title_label.setFont(QFont("Arial", 24, QFont.Weight.Bold))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 檔案選擇區域
        file_group = QGroupBox("檔案選擇")
        file_layout = QVBoxLayout(file_group)
        
        # CAD 檔案選擇
        cad_layout = QHBoxLayout()
        cad_layout.addWidget(QLabel("CAD 檔案："))
        self.cad_path_edit = QLineEdit()
        self.cad_path_edit.setPlaceholderText("請選擇 CAD 檔案 (.dxf)")
        cad_layout.addWidget(self.cad_path_edit)
        self.browse_cad_btn = QPushButton("瀏覽")
        self.browse_cad_btn.clicked.connect(self.browse_cad_file)
        cad_layout.addWidget(self.browse_cad_btn)
        file_layout.addLayout(cad_layout)
        
        # 檔案資訊
        self.file_info_label = QLabel()
        self.file_info_label.setWordWrap(True)
        file_layout.addWidget(self.file_info_label)
        
        # Excel 檔案選擇
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("Excel 檔案："))
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setPlaceholderText("請選擇 Excel 輸出檔案 (.xlsx)")
        excel_layout.addWidget(self.excel_path_edit)
        self.browse_excel_btn = QPushButton("瀏覽")
        self.browse_excel_btn.clicked.connect(self.browse_excel_file)
        excel_layout.addWidget(self.browse_excel_btn)
        file_layout.addLayout(excel_layout)
        
        main_layout.addWidget(file_group)
        
        # 進度區域
        progress_group = QGroupBox("處理進度")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("等待開始...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.status_label)
        
        main_layout.addWidget(progress_group)
        
        # 按鈕區域
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("開始轉換")
        self.start_button.setMinimumHeight(50)
        self.start_button.clicked.connect(self.start_conversion)
        button_layout.addWidget(self.start_button)
        
        self.reset_button = QPushButton("重置")
        self.reset_button.setMinimumHeight(50)
        self.reset_button.clicked.connect(self.reset_form)
        button_layout.addWidget(self.reset_button)
        
        # 開啟檔案按鈕（預設隱藏）
        self.open_file_btn = QPushButton("開啟檔案")
        self.open_file_btn.setMinimumHeight(50)
        self.open_file_btn.clicked.connect(self.open_excel_file)
        self.open_file_btn.setVisible(False)
        button_layout.addWidget(self.open_file_btn)
        
        main_layout.addLayout(button_layout)
        
        # Footer
        footer_label = QLabel("Powered by dxf2excel | © 2024 BaalWu")
        footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(footer_label)
    
    def setup_styles(self):
        """設定樣式"""
        # 設定深色主題
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #555555;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QPushButton {
                background-color: #0078d4;
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #555555;
            }
            QLineEdit {
                padding: 8px;
                border: 2px solid #555555;
                border-radius: 5px;
                background-color: #3c3c3c;
                color: white;
            }
            QProgressBar {
                border: 2px solid #555555;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 3px;
            }
        """)
    
    def browse_cad_file(self):
        """瀏覽 CAD 檔案"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "選擇 CAD 檔案", "", "DXF 檔案 (*.dxf);;所有檔案 (*.*)"
        )
        if file_path:
            self.cad_file_path = file_path
            self.cad_path_edit.setText(file_path)
            
            # 自動設定 Excel 檔案路徑
            excel_path = os.path.splitext(file_path)[0] + '.xlsx'
            self.excel_file_path = excel_path
            self.excel_path_edit.setText(excel_path)
            
            # 顯示檔案資訊
            self.show_file_info(file_path)
    
    def browse_excel_file(self):
        """瀏覽 Excel 檔案"""
        initial_dir = ""
        initial_file = ""
        if self.cad_file_path:
            initial_dir = os.path.dirname(self.cad_file_path)
            base = os.path.splitext(os.path.basename(self.cad_file_path))[0]
            initial_file = base + '.xlsx'
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "儲存 Excel 檔案", 
            os.path.join(initial_dir, initial_file),
            "Excel 檔案 (*.xlsx);;所有檔案 (*.*)"
        )
        if file_path:
            self.excel_file_path = file_path
            self.excel_path_edit.setText(file_path)
    
    def show_file_info(self, file_path):
        """顯示檔案資訊"""
        info = get_file_info(file_path)
        if 'error' not in info:
            self.file_info_label.setText(
                f"檔案大小：{info['size']}    建立時間：{info['created']}    修改時間：{info['modified']}"
            )
        else:
            self.file_info_label.setText('')
    
    def reset_form(self):
        """重置表單"""
        self.cad_file_path = ""
        self.excel_file_path = ""
        self.cad_path_edit.clear()
        self.excel_path_edit.clear()
        self.file_info_label.clear()
        self.progress_bar.setValue(0)
        self.status_label.setText("等待開始...")
        self.open_file_btn.setVisible(False)
        self.start_button.setEnabled(True)
    
    def start_conversion(self):
        """開始轉換"""
        # 檢查檔案路徑
        if not self.cad_file_path:
            QMessageBox.critical(self, "錯誤", "請選擇 CAD 檔案")
            return
        
        if not self.excel_file_path:
            QMessageBox.critical(self, "錯誤", "請選擇 Excel 檔案")
            return
        
        # 更新 UI 狀態
        self.start_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在啟動轉換...")
        
        # 創建並啟動轉換執行緒
        self.conversion_worker = ConversionWorker(self.cad_file_path, self.excel_file_path)
        self.conversion_worker.progress_updated.connect(self.update_progress)
        self.conversion_worker.conversion_completed.connect(self.conversion_completed)
        self.conversion_worker.error_occurred.connect(self.conversion_error)
        self.conversion_worker.start()
    
    def update_progress(self, value, status):
        """更新進度"""
        self.progress_bar.setValue(value)
        self.status_label.setText(status)
    
    def conversion_completed(self, rebar_data):
        """轉換完成"""
        self.status_label.setText("轉換完成！")
        self.progress_bar.setValue(100)
        self.open_file_btn.setVisible(True)
        self.start_button.setEnabled(True)
        
        QMessageBox.information(self, "完成", f"轉換完成！共處理 {len(rebar_data)} 個區塊。")
    
    def conversion_error(self, error_message):
        """轉換錯誤"""
        self.status_label.setText("轉換失敗")
        self.start_button.setEnabled(True)
        QMessageBox.critical(self, "錯誤", error_message)
    
    def open_excel_file(self):
        """開啟 Excel 檔案"""
        if self.excel_file_path and os.path.exists(self.excel_file_path):
            import subprocess
            subprocess.Popen(['open', self.excel_file_path])
