"""
CAD 鋼筋計料轉換工具 Pro v2.1
=====================================

功能說明:
- 讀取 DXF 格式的 CAD 檔案
- 自動識別鋼筋文字標記並提取數據
- 生成專業的鋼筋彎曲圖示
- 輸出格式化的 Excel 鋼筋計料表
- 支援多種鋼筋規格和材料等級

作者: BaalWu
版本: v2.1 Professional Edition
建立時間: 2025-06-03
最後更新: 2025-06-03
"""

import tkinter as tk
from ui.main_window import MainWindow
from config import WINDOW_TITLE, WINDOW_SIZE

def main():
    """主程式入口點"""
    # 建立主視窗
    root = tk.Tk()
    root.title(WINDOW_TITLE)
    root.geometry(WINDOW_SIZE)
    
    # 建立應用程式主視窗
    app = MainWindow(root)
    
    # 設定視窗關閉事件
    def on_closing():
        if tk.messagebox.askokcancel("確認", "確定要關閉程式嗎？"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 啟動主迴圈
    root.mainloop()

if __name__ == "__main__":
    main()