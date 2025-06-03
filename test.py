import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import ezdxf
import math
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import threading
import time
import re

class ModernCADtoExcelConverter:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("CAD 鋼筋計料轉換工具 Pro v2.0")
            self.root.geometry("900x750")
            self.root.resizable(True, True)
            
            # 設定現代化配色主題
            self.colors = {
                'primary': '#4A90E2',      # 主要藍色
                'secondary': '#E24A90',    # 次要紫色
                'success': '#F5A623',      # 成功橙色
                'background': '#1E1E1E',   # 背景色（深灰）
                'surface': '#2D2D2D',      # 表面色（稍淺的深灰）
                'text_primary': '#FFFFFF', # 主要文字（白色）
                'text_secondary': '#B0B0B0', # 次要文字（淺灰）
                'border': '#404040',       # 邊框色（中灰）
                'accent': '#3D3D3D'        # 強調色（深灰）
            }
            
            # 設定根視窗背景
            self.root.configure(bg=self.colors['background'])
            
            # 材質密度設定 (kg/m³)
            self.material_density = {
                "鋼筋": 7850,
                "鋁": 2700,
                "銅": 8960,
                "不鏽鋼": 8000
            }
            
            # 預設材質
            self.default_material = "鋼筋"
            
            # 鋼筋單位重量 (kg/m)
            self.rebar_unit_weight = {
                "#2": 0.249, "#3": 0.561, "#4": 0.996, "#5": 1.552, "#6": 2.235,
                "#7": 3.042, "#8": 3.973, "#9": 5.026, "#10": 6.404, "#11": 7.906,
                "#12": 11.38, "#13": 13.87, "#14": 14.59, "#15": 20.24, "#16": 25.00,
                "#17": 31.20, "#18": 39.70
            }
            
            # 處理進度相關變數
            self.current_step = 0
            self.total_steps = 0
            self.step_descriptions = {}
            self.processing_start_time = 0
            
            self.setup_modern_styles()
            self.setup_ui()
            
        except Exception as e:
            print(f"初始化錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"程式初始化時發生錯誤:\n{str(e)}")
    
    def setup_modern_styles(self):
        """設定現代化樣式"""
        self.style = ttk.Style()
        
        # 設定主題
        self.style.theme_use('clam')
        
        # 配置樣式
        self.style.configure("Modern.TFrame", 
                           background=self.colors['surface'],
                           relief='flat',
                           borderwidth=1)
        
        self.style.configure("Card.TFrame",
                           background=self.colors['surface'],
                           relief='solid',
                           borderwidth=1,
                           bordercolor=self.colors['border'])
        
        self.style.configure("Header.TLabel",
                           background=self.colors['surface'],
                           foreground=self.colors['primary'],
                           font=("Segoe UI", 18, "bold"))
        
        self.style.configure("Title.TLabel",
                           background=self.colors['surface'],
                           foreground=self.colors['text_primary'],
                           font=("Segoe UI", 12, "bold"))
        
        self.style.configure("Body.TLabel",
                           background=self.colors['surface'],
                           foreground=self.colors['text_primary'],
                           font=("Segoe UI", 10))
        
        self.style.configure("Caption.TLabel",
                           background=self.colors['surface'],
                           foreground=self.colors['text_secondary'],
                           font=("Segoe UI", 9))
        
        self.style.configure("Success.TLabel",
                           background=self.colors['surface'],
                           foreground=self.colors['success'],
                           font=("Segoe UI", 10, "bold"))
        
        self.style.configure("Primary.TButton",
                           background=self.colors['primary'],
                           foreground='white',
                           font=("Segoe UI", 14, "bold"),  # 放大按鈕文字
                           relief='flat',
                           borderwidth=0,
                           focuscolor='none')
        
        self.style.configure("Secondary.TButton",
                           background=self.colors['accent'],
                           foreground=self.colors['text_primary'],
                           font=("Segoe UI", 12),  # 放大次要按鈕文字
                           relief='flat',
                           borderwidth=1,
                           focuscolor='none')
        
        self.style.configure("Modern.TEntry",
                           fieldbackground='white',
                           borderwidth=1,
                           relief='solid',
                           font=("Segoe UI", 10))
        
        # 進度條樣式
        self.style.configure("Horizontal.TProgressbar",
                           background=self.colors['primary'],
                           troughcolor=self.colors['accent'],
                           borderwidth=0,
                           lightcolor=self.colors['primary'],
                           darkcolor=self.colors['primary'])
        
        # 配置 hover 效果
        self.style.map("Primary.TButton",
                      background=[('active', self.colors['secondary'])])
        
        self.style.map("Secondary.TButton",
                      background=[('active', self.colors['border'])])
    
    def create_card_frame(self, parent, title="", padding=20):
        """創建卡片式框架"""
        # 外層容器（陰影效果）
        shadow_frame = tk.Frame(parent, bg=self.colors['border'], height=2)
        shadow_frame.pack(fill=tk.X, pady=(5, 0), padx=2)
        
        # 主卡片框架
        card_frame = tk.Frame(parent, bg=self.colors['surface'], relief='solid', bd=1)
        card_frame.pack(fill=tk.X, pady=(0, 15), padx=5)
        
        if title:
            title_frame = tk.Frame(card_frame, bg=self.colors['surface'])
            title_frame.pack(fill=tk.X, pady=(15, 5), padx=padding)
            
            title_label = tk.Label(title_frame, text=title,
                                 bg=self.colors['surface'],
                                 fg=self.colors['text_primary'],
                                 font=("Segoe UI", 12, "bold"))
            title_label.pack(side=tk.LEFT)
        
        content_frame = tk.Frame(card_frame, bg=self.colors['surface'])
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 15), padx=padding)
        
        return content_frame
    
    def setup_ui(self):
        """設定現代化使用者介面"""
        # 主容器
        main_container = tk.Frame(self.root, bg=self.colors['background'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 標題區域
        header_frame = tk.Frame(main_container, bg=self.colors['background'])
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 主標題
        title_label = tk.Label(header_frame,
                             text="🏗️ CAD 鋼筋計料轉換工具 Pro",
                             bg=self.colors['background'],
                             fg=self.colors['primary'],
                             font=("Segoe UI", 20, "bold"))
        title_label.pack()
        
        # 副標題
        subtitle_label = tk.Label(header_frame,
                                text="專業級 DXF 檔案鋼筋數據分析與 Excel 報表生成工具",
                                bg=self.colors['background'],
                                fg=self.colors['text_secondary'],
                                font=("Segoe UI", 11))
        subtitle_label.pack(pady=(5, 0))
        
        # 版本標籤
        version_label = tk.Label(header_frame,
                               text="v2.0 Professional Edition",
                               bg=self.colors['background'],
                               fg=self.colors['success'],
                               font=("Segoe UI", 9, "bold"))
        version_label.pack(pady=(5, 0))
        
        # 輸入檔案卡片
        input_card = self.create_card_frame(main_container, "📁 輸入檔案設定")
        
        # CAD 檔案選擇區域
        cad_section = tk.Frame(input_card, bg=self.colors['surface'])
        cad_section.pack(fill=tk.X, pady=(0, 15))
        
        cad_label = tk.Label(cad_section, text="CAD 檔案",
                           bg=self.colors['surface'],
                           fg=self.colors['text_primary'],
                           font=("Segoe UI", 10, "bold"))
        cad_label.pack(anchor='w')
        
        cad_input_frame = tk.Frame(cad_section, bg=self.colors['surface'])
        cad_input_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.cad_path = tk.StringVar()
        self.cad_entry = tk.Entry(cad_input_frame, textvariable=self.cad_path,
                                font=("Segoe UI", 11),  # 放大字體
                                bg=self.colors['surface'],  # 深色背景
                                fg=self.colors['text_primary'],  # 白色文字
                                insertbackground=self.colors['primary'],  # 游標顏色
                                relief='solid', bd=1,
                                selectbackground=self.colors['primary'],  # 選取背景色
                                selectforeground='white')  # 選取文字顏色
        self.cad_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        browse_cad_btn = tk.Button(cad_input_frame, text="📂 瀏覽",
                                 command=self.browse_cad_file,
                                 bg=self.colors['primary'], fg='white',
                                 font=("Segoe UI", 12, "bold"),  # 放大瀏覽按鈕文字
                                 relief='flat', bd=0,
                                 cursor='hand2', padx=20)
        browse_cad_btn.pack(side=tk.RIGHT, padx=(10, 0), ipady=8)
        
        # 檔案資訊顯示
        self.file_info_label = tk.Label(cad_section, text="",
                                      bg=self.colors['surface'],
                                      fg=self.colors['text_secondary'],
                                      font=("Segoe UI", 9))
        self.file_info_label.pack(anchor='w', pady=(5, 0))
        
        # 輸出設定卡片
        output_card = self.create_card_frame(main_container, "💾 輸出檔案設定")
        
        # Excel 檔案選擇區域
        excel_section = tk.Frame(output_card, bg=self.colors['surface'])
        excel_section.pack(fill=tk.X)
        
        excel_label = tk.Label(excel_section, text="Excel 輸出檔案",
                             bg=self.colors['surface'],
                             fg=self.colors['text_primary'],
                             font=("Segoe UI", 10, "bold"))
        excel_label.pack(anchor='w')
        
        excel_input_frame = tk.Frame(excel_section, bg=self.colors['surface'])
        excel_input_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.excel_path = tk.StringVar()
        excel_entry = tk.Entry(excel_input_frame, textvariable=self.excel_path,
                             font=("Segoe UI", 11),  # 放大字體
                             bg=self.colors['surface'],  # 深色背景
                             fg=self.colors['text_primary'],  # 白色文字
                             insertbackground=self.colors['primary'],  # 游標顏色
                             relief='solid', bd=1,
                             selectbackground=self.colors['primary'],  # 選取背景色
                             selectforeground='white')  # 選取文字顏色
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        browse_excel_btn = tk.Button(excel_input_frame, text="💾 另存新檔",
                                   command=self.browse_excel_file,
                                   bg=self.colors['secondary'], fg='white',
                                   font=("Segoe UI", 12, "bold"),  # 放大另存新檔按鈕文字
                                   relief='flat', bd=0,
                                   cursor='hand2', padx=20)
        browse_excel_btn.pack(side=tk.RIGHT, padx=(10, 0), ipady=8)
        
        # 控制按鈕區域
        control_card = self.create_card_frame(main_container, "🎮 執行控制")
        
        button_frame = tk.Frame(control_card, bg=self.colors['surface'])
        button_frame.pack(fill=tk.X)
        
        # 主要按鈕
        self.convert_button = tk.Button(button_frame, text="🚀 開始轉換",
                                      command=self.start_conversion,
                                      bg=self.colors['success'], fg='white',
                                      font=("Segoe UI", 16, "bold"),  # 放大主要按鈕文字
                                      relief='flat', bd=0,
                                      cursor='hand2', padx=30, pady=10)
        self.convert_button.pack(side=tk.RIGHT)
        
        # 次要按鈕
        reset_btn = tk.Button(button_frame, text="🔄 重置",
                            command=self.reset_form,
                            bg=self.colors['accent'], fg=self.colors['text_primary'],
                            font=("Segoe UI", 14),  # 放大重置按鈕文字
                            relief='flat', bd=1,
                            cursor='hand2', padx=20, pady=8)
        reset_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
        # 快捷鍵提示
        shortcut_label = tk.Label(button_frame,
                                text="💡 快捷鍵: Ctrl+O(開檔) | Ctrl+S(存檔) | F5(轉換) | Esc(重置)",
                                bg=self.colors['surface'],
                                fg=self.colors['text_secondary'],
                                font=("Segoe UI", 9))
        shortcut_label.pack(side=tk.LEFT)
        
        # 處理狀態卡片
        status_card = self.create_card_frame(main_container, "📊 處理狀態與進度")
        
        # 進度資訊區域
        progress_info_frame = tk.Frame(status_card, bg=self.colors['surface'])
        progress_info_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 當前步驟顯示
        self.current_step_label = tk.Label(progress_info_frame, text="⭐ 準備就緒",
                                         bg=self.colors['surface'],
                                         fg=self.colors['text_primary'],
                                         font=("Segoe UI", 11, "bold"))
        self.current_step_label.pack(side=tk.LEFT)
        
        # 時間和百分比顯示
        self.time_label = tk.Label(progress_info_frame, text="",
                                 bg=self.colors['surface'],
                                 fg=self.colors['text_secondary'],
                                 font=("Segoe UI", 10))
        self.time_label.pack(side=tk.RIGHT)
        
        # 現代化進度條
        progress_frame = tk.Frame(status_card, bg=self.colors['surface'])
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress = ttk.Progressbar(progress_frame,
                                      orient="horizontal",
                                      mode="determinate",
                                      style="Horizontal.TProgressbar")
        self.progress.pack(fill=tk.X, ipady=8)
        
        # 詳細進度標籤
        self.detail_progress_label = tk.Label(status_card, text="",
                                            bg=self.colors['surface'],
                                            fg=self.colors['text_secondary'],
                                            font=("Segoe UI", 9))
        self.detail_progress_label.pack(fill=tk.X, pady=(0, 15))
        
        # 狀態文字區域
        text_container = tk.Frame(status_card, bg=self.colors['surface'])
        text_container.pack(fill=tk.BOTH, expand=True)
        
        # 狀態文字標題
        text_title = tk.Label(text_container, text="📝 詳細處理日誌",
                            bg=self.colors['surface'],
                            fg=self.colors['text_primary'],
                            font=("Segoe UI", 10, "bold"))
        text_title.pack(anchor='w', pady=(0, 5))
        
        # 文字框架
        text_frame = tk.Frame(text_container, bg=self.colors['surface'])
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.status_text = tk.Text(text_frame, height=10,
                                 bg=self.colors['surface'],  # 修改文字框背景色
                                 fg=self.colors['text_primary'],  # 修改文字框文字顏色
                                 font=("Consolas", 10),
                                 relief='solid', bd=1,
                                 wrap=tk.WORD)
        
        scrollbar = tk.Scrollbar(text_frame, command=self.status_text.yview,
                               bg=self.colors['accent'])
        
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # 初始狀態訊息
        self.log_message("🎉 程式已啟動！請選擇 CAD 檔案開始轉換流程。")
        
        # 設定鍵盤快捷鍵
        self.setup_keyboard_shortcuts()
        
        # 添加 hover 效果
        self.add_hover_effects()
    
    def add_hover_effects(self):
        """添加按鈕 hover 效果"""
        def on_enter(event, widget, hover_color):
            widget.configure(bg=hover_color)
        
        def on_leave(event, widget, normal_color):
            widget.configure(bg=normal_color)
        
        # 為所有按鈕添加 hover 效果（這裡可以進一步擴展）
        pass
    
    def setup_keyboard_shortcuts(self):
        """設定鍵盤快捷鍵"""
        self.root.bind('<Control-o>', lambda e: self.browse_cad_file())
        self.root.bind('<Control-s>', lambda e: self.browse_excel_file())
        self.root.bind('<F5>', lambda e: self.start_conversion())
        self.root.bind('<Escape>', lambda e: self.reset_form())
    
    def show_file_info(self, file_path):
        """顯示檔案資訊"""
        try:
            if os.path.isfile(file_path):
                file_size = os.path.getsize(file_path)
                size_text = self.format_file_size(file_size)
                info_text = f"📄 {size_text}"
                
                # 嘗試快速檢查 DXF 檔案
                if file_path.lower().endswith('.dxf'):
                    try:
                        doc = ezdxf.readfile(file_path)
                        text_count = len(list(doc.modelspace().query("TEXT")))
                        info_text += f" | 📝 {text_count} 個文字實體"
                    except:
                        info_text += " | ⚠️ 檔案格式驗證中..."
                
                self.file_info_label.config(text=info_text)
            else:
                self.file_info_label.config(text="❌ 檔案不存在")
        except Exception as e:
            self.file_info_label.config(text=f"❌ 讀取錯誤: {str(e)}")
    
    def format_file_size(self, size_bytes):
        """格式化檔案大小顯示"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024**2:
            return f"{size_bytes/1024:.1f} KB"
        elif size_bytes < 1024**3:
            return f"{size_bytes/(1024**2):.1f} MB"
        else:
            return f"{size_bytes/(1024**3):.1f} GB"
    
    def init_progress(self, steps_config):
        """初始化進度追蹤"""
        self.current_step = 0
        self.total_steps = len(steps_config)
        self.step_descriptions = {i: desc for i, desc in enumerate(steps_config)}
        self.processing_start_time = time.time()
        self.progress["maximum"] = 100
        self.progress["value"] = 0
    
    def update_progress(self, step=None, detail="", percentage=None):
        """更新進度顯示"""
        try:
            # 更新步驟
            if step is not None:
                self.current_step = step
            
            # 計算百分比
            if percentage is not None:
                progress_value = percentage
            else:
                progress_value = (self.current_step / self.total_steps) * 100 if self.total_steps > 0 else 0
            
            # 更新進度條
            self.progress["value"] = progress_value
            
            # 更新當前步驟顯示
            if self.current_step < len(self.step_descriptions):
                step_icons = ["🔍", "📂", "📊", "⚙️", "📈", "📋", "🎨", "💾"]
                icon = step_icons[self.current_step] if self.current_step < len(step_icons) else "⭐"
                step_text = f"{icon} 步驟 {self.current_step + 1}/{self.total_steps}: {self.step_descriptions[self.current_step]}"
                self.current_step_label.config(text=step_text)
            
            # 更新詳細進度
            if detail:
                self.detail_progress_label.config(text=f"🔄 {detail}")
            
            # 計算並顯示時間資訊
            elapsed_time = time.time() - self.processing_start_time
            if progress_value > 0:
                estimated_total = elapsed_time * (100 / progress_value)
                remaining_time = estimated_total - elapsed_time
                time_text = f"⏱️ 已用時: {self.format_time(elapsed_time)} | ⏳ 預估剩餘: {self.format_time(remaining_time)} | 📊 {progress_value:.1f}%"
            else:
                time_text = f"⏱️ 已用時: {self.format_time(elapsed_time)} | 📊 0.0%"
            
            self.time_label.config(text=time_text)
            
            # 強制更新 UI
            self.root.update_idletasks()
            
        except Exception as e:
            print(f"更新進度時發生錯誤: {str(e)}")
    
    def format_time(self, seconds):
        """格式化時間顯示"""
        if seconds < 60:
            return f"{seconds:.1f}秒"
        elif seconds < 3600:
            return f"{int(seconds // 60)}分{int(seconds % 60)}秒"
        else:
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            return f"{hours}小時{minutes}分"
    
    def browse_cad_file(self):
        filetypes = (
            ("DXF 檔案", "*.dxf"),
            ("DWG 檔案", "*.dwg"),
            ("所有檔案", "*.*")
        )
        filename = filedialog.askopenfilename(
            title="選擇 CAD 檔案",
            filetypes=filetypes
        )
        if filename:
            self.cad_path.set(filename)
            # 自動設定 Excel 檔案路徑
            default_excel = os.path.splitext(filename)[0] + "_鋼筋計料.xlsx"
            self.excel_path.set(default_excel)
            # 顯示檔案資訊
            self.show_file_info(filename)
            self.log_message(f"✅ 已選擇檔案: {os.path.basename(filename)}")
    
    def browse_excel_file(self):
        filetypes = (
            ("Excel 檔案", "*.xlsx"),
            ("所有檔案", "*.*")
        )
        filename = filedialog.asksaveasfilename(
            title="儲存 Excel 檔案",
            filetypes=filetypes,
            defaultextension=".xlsx",
            confirmoverwrite=True
        )
        if filename:
            self.excel_path.set(filename)
            self.log_message(f"📁 輸出路徑: {os.path.basename(filename)}")
    
    def reset_form(self):
        self.cad_path.set("")
        self.excel_path.set("")
        self.file_info_label.config(text="")
        self.status_text.delete(1.0, tk.END)
        self.progress["value"] = 0
        self.current_step_label.config(text="⭐ 準備就緒")
        self.detail_progress_label.config(text="")
        self.time_label.config(text="")
        self.convert_button.config(state="normal")
        self.log_message("🔄 表單已重置，請重新選擇檔案。")
    
    def log_message(self, message):
        try:
            timestamp = time.strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] {message}"
            
            if not formatted_message.endswith('\n'):
                formatted_message += '\n'
            
            self.status_text.insert(tk.END, formatted_message)
            self.status_text.see(tk.END)
            self.root.update_idletasks()
            
            print(formatted_message.strip())
        except Exception as e:
            print(f"記錄訊息時發生錯誤: {str(e)}")
    
    def start_conversion(self):
        if not self.cad_path.get():
            messagebox.showerror("❌ 錯誤", "請先選擇 CAD 檔案！")
            return
        
        if not self.excel_path.get():
            messagebox.showerror("❌ 錯誤", "請先指定 Excel 輸出檔案！")
            return
        
        # 清空狀態文字
        self.status_text.delete(1.0, tk.END)
        
        # 禁用轉換按鈕
        self.convert_button.config(state="disabled", text="🔄 轉換中...", bg=self.colors['text_secondary'])
        
        # 在新線程中執行轉換，避免 UI 凍結
        conversion_thread = threading.Thread(target=self.convert_cad_to_excel)
        conversion_thread.daemon = True
        conversion_thread.start()
    
    # [這裡包含所有原始的轉換邏輯函數]
    def calculate_line_length(self, start_point, end_point):
        """計算兩點之間的距離"""
        return math.sqrt(
            (end_point[0] - start_point[0])**2 + 
            (end_point[1] - start_point[1])**2 + 
            (end_point[2] - start_point[2])**2 if len(start_point) > 2 and len(end_point) > 2 
            else (end_point[2] if len(end_point) > 2 else 0)
        )
    
    def calculate_polyline_length(self, points):
        """計算多段線的長度"""
        total_length = 0
        for i in range(len(points) - 1):
            total_length += self.calculate_line_length(points[i], points[i+1])
        return total_length
    
    def extract_rebar_info(self, text):
        """從文字中提取鋼筋信息"""
        if not text:
            return None, None, None, None
            
        number = ""
        count = 1
        length = None
        segments = []
        
        # 尋找鋼筋號數 (支援多種格式)
        number_match = re.search(r'#(\d+)', text)
        if number_match:
            number = "#" + number_match.group(1)
        
        if not number:
            number_match = re.search(r'D(\d+)', text)
            if number_match:
                diameter = float(number_match.group(1))
                for num, dia in self.rebar_unit_weight.items():
                    if abs(dia - diameter) < 0.1:
                        number = num
                        break
        
        if not number:
            number_match = re.search(r'(\d+(?:\.\d+)?)\s*mm', text)
            if number_match:
                diameter = float(number_match.group(1))
                for num, dia in self.rebar_unit_weight.items():
                    if abs(dia - diameter) < 0.1:
                        number = num
                        break
        
        # 尋找長度和數量
        length_count_match = re.search(r'[#D]?\d+[-_](?:(\d+(?:\+\d+)*))[xX×*-](\d+)', text)
        if length_count_match:
            try:
                length_parts = length_count_match.group(1).split('+')
                segments = [float(part) for part in length_parts]
                total_length = sum(segments)
                length = total_length
                count = int(length_count_match.group(2))
            except:
                length = None
                count = 1
        else:
            count_match = re.search(r'[xX×*-](\d+)', text)
            if count_match:
                try:
                    count = int(count_match.group(1))
                except:
                    count = 1
        
        return number, count, length, segments
    
    def get_rebar_diameter(self, number):
        """根據鋼筋號數獲取直徑(mm)"""
        rebar_diameter = {
            "#2": 6.4, "#3": 9.5, "#4": 12.7, "#5": 15.9, "#6": 19.1,
            "#7": 22.2, "#8": 25.4, "#9": 28.7, "#10": 32.3, "#11": 35.8,
            "#12": 43.0, "#13": 43.0, "#14": 43.0, "#15": 43.0, "#16": 43.0,
            "#17": 43.0, "#18": 57.3
        }
        return rebar_diameter.get(number, "")
    
    def get_rebar_unit_weight(self, number):
        """根據鋼筋號數獲取單位重量(kg/m)"""
        return self.rebar_unit_weight.get(number, 0)
    
    def calculate_rebar_weight(self, number, length, count=1):
        """計算鋼筋總重量"""
        unit_weight = self.get_rebar_unit_weight(number)
        if unit_weight and length:
            length_m = length / 100.0
            return round(unit_weight * length_m * count, 2)
        return 0
    
    def draw_ascii_rebar(self, segments):
        """使用 ASCII 字元繪製彎折示意圖"""
        if not segments:
            return "─"
            
        if len(segments) == 1:
            length = str(int(segments[0]))
            line = "─" * 10
            total_width = max(len(line), len(length))
            length_spaces = (total_width - len(length)) // 2
            line_spaces = (total_width - len(line)) // 2
            return f"{' ' * length_spaces}{length}\n{' ' * line_spaces}{line}"
        
        lines = []
        middle_chars = 10
        
        line = ""
        start_num = str(int(segments[0]))
        line += f"{start_num} |"
        
        if len(segments) > 1:
            line += "─" * middle_chars
        
        if len(segments) > 2:
            end_num = str(int(segments[-1]))
            line += f"| {end_num}"
        elif len(segments) == 2:
            end_num = str(int(segments[-1]))
            line += f"| {end_num}"
        
        if len(segments) > 1:
            middle_num = str(int(sum(segments[1:-1] if len(segments) > 2 else segments[1:])))
            total_width = len(line)
            start_pos = len(start_num) + 2
            middle_section_width = middle_chars
            middle_start = start_pos + (middle_section_width - len(middle_num)) // 2
            
            first_line = " " * total_width
            first_line = first_line[:middle_start] + middle_num + first_line[middle_start + len(middle_num):]
            lines.append(first_line)
        
        lines.append(line)
        return "\n".join(lines)
    
    def convert_cad_to_excel(self):
        try:
            # 定義處理步驟
            steps = [
                "驗證檔案",
                "載入 CAD 檔案",
                "分析文字實體",
                "處理鋼筋資料",
                "生成統計資料",
                "建立 Excel 工作簿",
                "格式化 Excel",
                "儲存檔案"
            ]
            
            self.init_progress(steps)
            self.log_message("🚀 開始轉換流程...")
            
            # 步驟 1: 驗證檔案
            self.update_progress(0, "正在驗證檔案...")
            self.log_message(f"📂 檢查檔案: {os.path.basename(self.cad_path.get())}")
            
            if not os.path.isfile(self.cad_path.get()):
                raise FileNotFoundError("找不到指定的 CAD 檔案")
                
            file_ext = os.path.splitext(self.cad_path.get())[1].lower()
            if file_ext != '.dxf':
                self.log_message("⚠️ 警告: 非 DXF 檔案可能無法正確讀取")
            
            # 步驟 2: 載入 CAD 檔案
            self.update_progress(1, "正在載入 CAD 檔案...")
            
            try:
                doc = ezdxf.readfile(self.cad_path.get())
                msp = doc.modelspace()
                self.log_message("✅ CAD 檔案載入成功")
            except Exception as e:
                raise Exception(f"無法讀取 CAD 檔案: {str(e)}")
            
            # 步驟 3: 分析文字實體
            self.update_progress(2, "正在分析文字實體...")
            
            text_entities = list(msp.query("TEXT"))
            entity_count = len(text_entities)
            self.log_message(f"📊 找到 {entity_count} 個文字實體")
            
            if entity_count == 0:
                raise Exception("未找到任何文字實體")
            
            # 步驟 4: 處理鋼筋資料
            self.update_progress(3, "正在處理鋼筋資料...")
            
            rebar_data = []
            processed_count = 0
            valid_rebar_count = 0
            
            for i, text in enumerate(text_entities):
                # 更新子進度
                sub_progress = 30 + (i / entity_count) * 20  # 30-50%
                detail = f"處理文字實體 {i+1}/{entity_count}"
                self.update_progress(percentage=sub_progress, detail=detail)
                
                text_content = text.dxf.text
                if text_content:
                    number, count, length, segments = self.extract_rebar_info(text_content)
                    if number and length is not None:
                        valid_rebar_count += 1
                        
                        # 計算重量
                        length_cm = length
                        unit_weight = self.get_rebar_unit_weight(number) if number.startswith("#") else 0
                        weight = self.calculate_rebar_weight(number, length_cm, count) if number.startswith("#") else 0
                        
                        # 建立資料
                        data = {
                            "編號": number,
                            "長度(cm)": round(length_cm, 2),
                            "數量": count,
                            "單位重(kg/m)": unit_weight,
                            "總重量(kg)": weight,
                            "圖層": text.dxf.layer,
                            "備註": text_content
                        }
                        
                        # 添加分段長度欄位
                        if segments:
                            for j, segment in enumerate(segments):
                                letter = chr(65 + j)  # A, B, C, ...
                                data[f"{letter}(cm)"] = round(segment, 2)
                        
                        rebar_data.append(data)
                        
                        if valid_rebar_count <= 5:  # 只顯示前5個的詳細資訊
                            self.log_message(f"✅ 鋼筋 #{valid_rebar_count}: {number}, {length_cm}cm, {count}支")
                
                processed_count += 1
            
            self.log_message(f"📈 處理完成: 找到 {valid_rebar_count} 個有效鋼筋標記")
            
            if not rebar_data:
                raise Exception("沒有找到任何可轉換的鋼筋數據")
            
            # 步驟 5: 生成統計資料
            self.update_progress(4, "正在生成統計資料...")
            
            # 統計數據
            total_quantity = sum(item["數量"] for item in rebar_data)
            total_length = sum(item["長度(cm)"] * item["數量"] for item in rebar_data)
            total_weight = sum(item["總重量(kg)"] for item in rebar_data)
            
            # 按號數統計
            rebar_types = {}
            for item in rebar_data:
                rebar_num = item["編號"]
                if rebar_num not in rebar_types:
                    rebar_types[rebar_num] = {"count": 0, "weight": 0}
                rebar_types[rebar_num]["count"] += item["數量"]
                rebar_types[rebar_num]["weight"] += item["總重量(kg)"]
            
            self.log_message(f"📊 統計結果: 總數量 {total_quantity}支, 總重量 {total_weight:.2f}kg")
            self.log_message(f"🔧 鋼筋類型: {len(rebar_types)} 種")
            
            # 根據號數排序
            sorted_data = sorted(rebar_data, key=lambda x: x["編號"] if "#" in x["編號"] else "z" + x["編號"])
            
            # 步驟 6: 建立 Excel 工作簿
            self.update_progress(5, "正在建立 Excel 工作簿...")
            
            # 重新排列欄位
            base_columns = ["編號", "長度(cm)", "數量", "總重量(kg)", "圖示", "備註"]
            
            # 找出所有可能的分段長度欄位
            segment_columns = set()
            for item in sorted_data:
                for key in item.keys():
                    if key.endswith("(cm)") and key != "長度(cm)":
                        segment_columns.add(key)
            
            # 按字母順序排序分段長度欄位
            segment_columns = sorted(segment_columns)
            
            # 在長度欄位後插入分段長度欄位
            columns = base_columns[:2] + segment_columns + base_columns[2:]
            
            df = pd.DataFrame(sorted_data)
            
            # 添加圖示欄位
            self.update_progress(percentage=75, detail="正在生成鋼筋圖示...")
            for index, row in df.iterrows():
                segments = []
                for key in sorted([k for k in row.keys() if k.endswith("(cm)") and k != "長度(cm)"]):
                    if not pd.isna(row[key]):
                        segments.append(row[key])
                df.at[index, "圖示"] = self.draw_ascii_rebar(segments)
            
            df = df.reindex(columns=columns)
            
            # 步驟 7: 格式化 Excel
            self.update_progress(6, "正在格式化 Excel...")
            
            try:
                # 創建 Excel 工作簿
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "鋼筋計料"
                
                # 設定標題
                project_name = os.path.basename(self.cad_path.get())
                ws['A1'] = f"鋼筋計料表"
                ws['A2'] = f"專案名稱: {project_name}"
                ws['A3'] = f"生成時間: {time.strftime('%Y-%m-%d %H:%M:%S')}"
                
                # 合併儲存格
                ws.merge_cells('A1:H1')
                ws.merge_cells('A2:H2')
                ws.merge_cells('A3:H3')
                
                # 設定標題樣式
                title_font = Font(bold=True, size=16)
                subtitle_font = Font(bold=True, size=12)
                info_font = Font(size=10)
                
                ws['A1'].font = title_font
                ws['A2'].font = subtitle_font
                ws['A3'].font = info_font
                
                # 居中對齊
                title_align = Alignment(horizontal='center', vertical='center')
                ws['A1'].alignment = title_align
                ws['A2'].alignment = title_align
                ws['A3'].alignment = title_align
                
                # 設定表頭 (從第5行開始)
                headers = columns
                header_row = 5
                
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=header_row, column=col_num)
                    cell.value = header
                    
                    # 設定表頭樣式
                    header_font = Font(bold=True)
                    header_align = Alignment(horizontal='center', vertical='center')
                    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
                    header_border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    
                    cell.font = header_font
                    cell.alignment = header_align
                    cell.fill = header_fill
                    cell.border = header_border
                
                # 寫入資料 (從第6行開始)
                data_start_row = 6
                row_num = data_start_row
                
                self.update_progress(percentage=80, detail="正在寫入資料...")
                
                for i, (_, row) in enumerate(df.iterrows()):
                    # 更新子進度
                    if len(df) > 0:
                        sub_progress = 80 + (i / len(df)) * 10  # 80-90%
                        if i % 10 == 0:  # 每10行更新一次進度
                            self.update_progress(percentage=sub_progress, detail=f"寫入資料行 {i+1}/{len(df)}")
                    
                    for col_num, col_name in enumerate(headers, 1):
                        cell = ws.cell(row=row_num, column=col_num)
                        
                        # 獲取對應的值
                        if col_name in row:
                            cell.value = row[col_name]
                        
                        # 設定資料樣式
                        data_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        data_border = Border(
                            left=Side(style='thin'),
                            right=Side(style='thin'),
                            top=Side(style='thin'),
                            bottom=Side(style='thin')
                        )
                        
                        cell.alignment = data_align
                        cell.border = data_border
                        
                        # 根據編號設定不同顏色
                        if col_num == 1 and row["編號"].startswith("#"):
                            cell.font = Font(bold=True)
                    
                    # 調整行高以適應圖示
                    ws.row_dimensions[row_num].height = 60
                    row_num += 1
                
                # 添加統計行
                summary_row = row_num + 1
                ws.cell(row=summary_row, column=1).value = "總計"
                ws.cell(row=summary_row, column=1).font = Font(bold=True)
                
                # 總數量
                quantity_col = headers.index("數量") + 1
                ws.cell(row=summary_row, column=quantity_col).value = df["數量"].sum()
                ws.cell(row=summary_row, column=quantity_col).font = Font(bold=True)
                
                # 總重量
                weight_col = headers.index("總重量(kg)") + 1
                ws.cell(row=summary_row, column=weight_col).value = round(df["總重量(kg)"].sum(), 2)
                ws.cell(row=summary_row, column=weight_col).font = Font(bold=True)
                
                # 為統計行設定樣式
                for col in range(1, len(headers) + 1):
                    cell = ws.cell(row=summary_row, column=col)
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='double')
                    )
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # 設定欄位寬度
                column_widths = {
                    "編號": 8,
                    "長度(cm)": 10,
                    "數量": 8,
                    "總重量(kg)": 12,
                    "圖示": 60,
                    "備註": 60
                }
                
                # 設定分段長度欄位的寬度
                for col in segment_columns:
                    column_widths[col] = 8
                
                # 根據欄位名稱設定寬度
                for col_num, header in enumerate(headers, 1):
                    if header in column_widths:
                        column_letter = openpyxl.utils.get_column_letter(col_num)
                        ws.column_dimensions[column_letter].width = column_widths[header]
                
                self.log_message("🎨 Excel 格式設定完成")
                
            except Exception as e:
                self.log_message(f"⚠️ Excel 格式化時出現問題: {str(e)}")
                self.log_message("🔄 使用基本儲存方式...")
                # 如果格式化失敗，至少確保數據被保存
                df.to_excel(self.excel_path.get(), sheet_name='鋼筋計料', index=False)
            
            # 步驟 8: 儲存檔案
            self.update_progress(7, "正在儲存檔案...")
            
            try:
                wb.save(self.excel_path.get())
                self.log_message(f"💾 檔案已儲存: {os.path.basename(self.excel_path.get())}")
            except Exception as e:
                raise Exception(f"儲存檔案時發生錯誤: {str(e)}")
            
            # 完成
            self.update_progress(percentage=100, detail="轉換完成!")
            
            # 顯示完成統計
            elapsed_time = time.time() - self.processing_start_time
            self.log_message("=" * 60)
            self.log_message("🎉 轉換完成!")
            self.log_message(f"📊 處理統計:")
            self.log_message(f"   • 處理檔案: {os.path.basename(self.cad_path.get())}")
            self.log_message(f"   • 文字實體: {entity_count} 個")
            self.log_message(f"   • 有效鋼筋: {valid_rebar_count} 個")
            self.log_message(f"   • 鋼筋類型: {len(rebar_types)} 種")
            self.log_message(f"   • 總數量: {total_quantity} 支")
            self.log_message(f"   • 總長度: {total_length:.2f} cm")
            self.log_message(f"   • 總重量: {total_weight:.2f} kg")
            self.log_message(f"   • 處理時間: {self.format_time(elapsed_time)}")
            self.log_message(f"   • 輸出檔案: {os.path.basename(self.excel_path.get())}")
            self.log_message("=" * 60)
            
            # 恢復按鈕狀態
            self.convert_button.config(state="normal", text="🚀 開始轉換", bg=self.colors['success'])
            
            # 顯示完成對話框
            result_message = f"""🎉 轉換完成！

📊 處理結果:
• 有效鋼筋: {valid_rebar_count} 個
• 鋼筋類型: {len(rebar_types)} 種  
• 總數量: {total_quantity} 支
• 總重量: {total_weight:.2f} kg
• 處理時間: {self.format_time(elapsed_time)}

💾 檔案已儲存至:
{self.excel_path.get()}

感謝您使用 CAD 鋼筋計料轉換工具 Pro！"""
            
            messagebox.showinfo("✅ 轉換完成", result_message)
            
        except Exception as e:
            self.log_message(f"❌ 錯誤: {str(e)}")
            self.update_progress(percentage=0, detail="轉換失敗")
            self.convert_button.config(state="normal", text="🚀 開始轉換", bg=self.colors['success'])
            messagebox.showerror("❌ 轉換錯誤", f"轉換過程中發生錯誤:\n\n{str(e)}\n\n請檢查檔案格式和內容後重試。")

def main():
    """主程式入口"""
    root = tk.Tk()
    
    # 設定視窗圖示（如果有的話）
    try:
        # root.iconbitmap('icon.ico')  # 如果有圖示檔案
        pass
    except:
        pass
    
    # 設定視窗居中
    def center_window(window, width=900, height=750):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")
    
    center_window(root)
    
    app = ModernCADtoExcelConverter(root)
    
    # 設定關閉事件
    def on_closing():
        if messagebox.askokcancel("退出確認", "確定要退出 CAD 鋼筋計料轉換工具嗎？"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 啟動程式
    root.mainloop()

if __name__ == "__main__":
    main()