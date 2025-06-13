"""
主視窗相關功能模組
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from config import WINDOW_TITLE, WINDOW_SIZE, COLORS
from ui.styles import StyleManager, ModernButton, CardFrame, ModernEntry, ModernLabel
from core.cad_reader import CADReader
from core.excel_writer import ExcelWriter
from utils.helpers import get_file_info, create_progress_tracker, update_progress
from utils.graphics import GraphicsManager

class MainWindow:
    """主視窗類別"""
    
    def __init__(self, root):
        """初始化主視窗"""
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.configure(bg=COLORS['background'])
        
        # 初始化元件
        self.cad_reader = CADReader()
        self.excel_writer = ExcelWriter()
        self.graphics_manager = GraphicsManager()
        
        # 初始化變數
        self.cad_file_path = tk.StringVar()
        self.excel_file_path = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar()
        self.cad_file_info_var = tk.StringVar()
        
        # 設定樣式
        StyleManager.setup_modern_styles()
        
        # 建立介面
        self.setup_ui()
        
        # 設定快捷鍵
        self.setup_shortcuts()
    
    def setup_ui(self):
        """建立使用者介面"""
        # 建立主框架（內容集中，不撐滿整個 window）
        main_frame = CardFrame(self.root, padding=20)
        main_frame.pack(padx=40, pady=40)
        
        # 建立檔案選擇區域
        file_frame = CardFrame(main_frame, title="檔案選擇", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        # CAD 檔案選擇
        cad_frame = tk.Frame(file_frame)
        cad_frame.pack(fill=tk.X, pady=5)
        
        ModernLabel(cad_frame, text="CAD 檔案：").pack(side=tk.LEFT, padx=5)
        ModernEntry(cad_frame, textvariable=self.cad_file_path, width=50).pack(side=tk.LEFT, padx=5)
        ModernButton(cad_frame, text="瀏覽", command=self.browse_cad_file).pack(side=tk.LEFT, padx=5)
        
        # 新增檔案資訊顯示區
        self.cad_file_info_label = ModernLabel(file_frame, textvariable=self.cad_file_info_var, anchor='w')
        self.cad_file_info_label.pack(fill=tk.X, padx=5, pady=(0, 10))
        
        # Excel 檔案選擇
        excel_frame = tk.Frame(file_frame)
        excel_frame.pack(fill=tk.X, pady=5)
        
        ModernLabel(excel_frame, text="Excel 檔案：").pack(side=tk.LEFT, padx=5)
        ModernEntry(excel_frame, textvariable=self.excel_file_path, width=50).pack(side=tk.LEFT, padx=5)
        ModernButton(excel_frame, text="瀏覽", command=self.browse_excel_file).pack(side=tk.LEFT, padx=5)
        
        # 建立進度區域
        progress_frame = CardFrame(main_frame, title="處理進度", padding=10)
        progress_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            style='Modern.Horizontal.TProgressbar',
            variable=self.progress_var,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        self.status_label = ModernLabel(progress_frame, textvariable=self.status_var)
        self.status_label.pack(pady=5)
        self.status_label.pack_forget()  # 預設隱藏
        
        # 建立按鈕區域
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 20))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(2, weight=1)
        # 開始轉換按鈕（主色）
        ModernButton(
            button_frame,
            text="開始轉換",
            command=self.start_conversion,
            width=16
        ).grid(row=0, column=0, padx=16, ipadx=8, ipady=4, sticky="ew")
        # 重置按鈕（次要色）
        ModernButton(
            button_frame,
            text="重置",
            command=self.reset_form,
            width=16,
            style='Secondary.TButton'
        ).grid(row=0, column=1, padx=16, ipadx=8, ipady=4, sticky="ew")
        # 「開啟檔案」按鈕預設不顯示
        self.open_file_btn = ModernButton(
            button_frame,
            text="開啟檔案",
            command=self.open_excel_file,
            width=16,
            style='Secondary.TButton'
        )
        
        # Footer
        footer_frame = tk.Frame(self.root)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 8))
        ModernLabel(
            footer_frame,
            text="Powered by dxf2excel | © 2024 BaalWu",
            anchor='center',
            font=('Calibri', 10)
        ).pack(pady=2)
    
    def setup_shortcuts(self):
        """設定快捷鍵"""
        self.root.bind('<Control-o>', lambda e: self.browse_cad_file())
        self.root.bind('<Control-s>', lambda e: self.browse_excel_file())
        self.root.bind('<F5>', lambda e: self.start_conversion())
        self.root.bind('<Escape>', lambda e: self.reset_form())
    
    def browse_cad_file(self):
        """瀏覽 CAD 檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇 CAD 檔案",
            filetypes=[("DXF 檔案", "*.dxf"), ("所有檔案", "*.*")]
        )
        if file_path:
            self.cad_file_path.set(file_path)
            # 自動帶入 Excel 檔案預設路徑
            excel_path = os.path.splitext(file_path)[0] + '.xlsx'
            self.excel_file_path.set(excel_path)
            self.show_file_info(file_path)
    
    def browse_excel_file(self):
        """瀏覽 Excel 檔案（預設在 CAD 檔案同資料夾）"""
        initialdir = ''
        initialfile = ''
        cad_path = self.cad_file_path.get()
        if cad_path:
            initialdir = os.path.dirname(cad_path)
            base = os.path.splitext(os.path.basename(cad_path))[0]
            initialfile = base + '.xlsx'
        file_path = filedialog.asksaveasfilename(
            title="儲存 Excel 檔案",
            defaultextension=".xlsx",
            filetypes=[("Excel 檔案", "*.xlsx"), ("所有檔案", "*.*")],
            initialdir=initialdir,
            initialfile=initialfile
        )
        if file_path:
            self.excel_file_path.set(file_path)
    
    def show_file_info(self, file_path):
        """顯示檔案資訊於 UI 下方"""
        info = get_file_info(file_path)
        if 'error' not in info:
            self.cad_file_info_var.set(
                f"檔案大小：{info['size']}    建立時間：{info['created']}    修改時間：{info['modified']}"
            )
        else:
            self.cad_file_info_var.set('')
    
    def reset_form(self):
        """重置表單"""
        self.cad_file_path.set("")
        self.excel_file_path.set("")
        self.progress_var.set(0)
        self.status_var.set("")
        self.status_label.pack_forget()  # 重置時隱藏
        # 重置時隱藏「開啟檔案」按鈕
        self.open_file_btn.grid_remove()
    
    def start_conversion(self):
        """開始轉換程序"""
        # 檢查檔案路徑
        if not self.cad_file_path.get():
            messagebox.showerror("錯誤", "請選擇 CAD 檔案")
            return
        
        if not self.excel_file_path.get():
            messagebox.showerror("錯誤", "請選擇 Excel 檔案")
            return
        
        # 檢查圖形功能
        if not self.graphics_manager.check_dependencies():
            if not messagebox.askyesno(
                "缺少套件",
                "缺少圖形繪製所需的套件，是否要自動安裝？"
            ):
                return
            
            # 自動安裝套件
            if not self.graphics_manager.install_dependencies():
                messagebox.showerror("錯誤", "安裝套件失敗，請手動安裝")
                return
        
        # 建立進度追蹤器
        self.progress_tracker = create_progress_tracker(5)
        
        # 開始轉換
        threading.Thread(target=self.convert_process, daemon=True).start()
        self.status_label.pack(pady=5)  # 開始處理時顯示
    
    def write_excel_on_main_thread(self, grouped_data):
        self.excel_writer.create_workbook()
        self.excel_writer.write_multi_sheet_rebar_data(grouped_data)
        self.excel_writer.save_workbook(self.excel_file_path.get())
        self.status_var.set("轉換完成！")
        self.progress_var.set(100)
        self.open_file_btn.grid(row=0, column=2, padx=16, ipadx=8, ipady=4, sticky="ew")

    def convert_process(self):
        """轉換處理程序"""
        try:
            # 處理過程中確保顯示（直接主執行緒呼叫）
            self.status_label.pack(pady=5)
            # 更新狀態
            self.root.after(0, lambda: self.status_var.set("正在開啟 CAD 檔案..."))
            self.root.after(0, lambda: self.progress_var.set(20))
            
            # 開啟 CAD 檔案
            if not self.cad_reader.open_file(self.cad_file_path.get()):
                raise Exception("無法開啟 CAD 檔案")
            
            # 更新狀態
            self.root.after(0, lambda: self.status_var.set("正在處理圖面..."))
            self.root.after(0, lambda: self.progress_var.set(40))
            
            # 處理圖面
            rebar_data = self.cad_reader.process_drawing()
            if not rebar_data:
                raise Exception("處理圖面失敗")
            
            # 更新狀態
            self.root.after(0, lambda: self.status_var.set("正在生成 Excel 檔案..."))
            self.root.after(0, lambda: self.progress_var.set(60))
            # 這裡用 after 回主執行緒寫 Excel（含圖形）
            self.root.after(0, lambda: self.write_excel_on_main_thread(rebar_data))
        except Exception as e:
            self.root.after(0, lambda err=e: messagebox.showerror("錯誤", f"轉換過程發生錯誤：\n{str(err)}"))
            self.root.after(0, lambda: self.status_var.set("轉換失敗"))
        finally:
            self.cad_reader.close_file() 

    def open_excel_file(self):
        """開啟 Excel 檔案"""
        import subprocess, os
        path = self.excel_file_path.get()
        if path and os.path.exists(path):
            subprocess.Popen(['open', path]) 