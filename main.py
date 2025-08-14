#!/usr/bin/env python3
"""
CAD 鋼筋計料轉換工具 Pro v2.1 - PyQt6 版本
=====================================

使用 PyQt6 重寫，解決執行緒問題

功能說明:
- 讀取 DXF 格式的 CAD 檔案
- 自動識別鋼筋文字標記並提取數據
- 生成專業的鋼筋彎曲圖示
- 輸出格式化的 Excel 鋼筋計料表
- 支援多種鋼筋規格和材料等級

作者: BaalWu
版本: v2.1 PyQt6 Edition
建立時間: 2025-06-03
最後更新: 2025-06-03
"""

import sys
from PyQt6.QtWidgets import QApplication
from ui.pyqt_main_window import PyQtMainWindow

def main():
    """主程式入口點"""
    print("正在啟動 PyQt6 應用程式...")
    
    # 創建應用程式
    app = QApplication(sys.argv)
    app.setApplicationName("CAD 鋼筋計料轉換工具")
    print("PyQt6 應用程式創建成功")
    
    try:
        # 創建主視窗
        print("正在創建主視窗...")
        window = PyQtMainWindow()
        print("主視窗創建成功")
        
        # 顯示視窗
        print("正在顯示視窗...")
        window.show()
        print("視窗顯示成功")
        
        print("啟動事件循環...")
        # 啟動事件循環
        sys.exit(app.exec())
        
    except Exception as e:
        print(f"創建主視窗時發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
