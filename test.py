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
            self.root.title("CAD 鋼筋計料轉換工具 Pro v2.1")
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
            
            # 標準鋼筋彎曲半徑 (根據鋼筋直徑)
            self.bend_radius = {
                '#2': 3, '#3': 3, '#4': 4, '#5': 5, '#6': 6, '#7': 7, '#8': 8,
                '#9': 9, '#10': 10, '#11': 11, '#12': 12, '#13': 13, '#14': 14,
                '#15': 15, '#16': 16, '#17': 17, '#18': 18
            }
            
            # 處理進度相關變數
            self.current_step = 0
            self.total_steps = 0
            self.step_descriptions = {}
            self.processing_start_time = 0
            
            # 檢查圖形繪製功能
            self.graphics_available = self.check_graphics_dependencies()
            
            self.setup_modern_styles()
            self.setup_ui()
            
        except Exception as e:
            print(f"初始化錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"程式初始化時發生錯誤:\n{str(e)}")
    
    def check_graphics_dependencies(self):
        """檢查圖形繪製所需的套件"""
        try:
            import matplotlib
            import numpy
            from PIL import Image
            self.log_message = lambda x: print(x)  # 臨時設定，稍後會被覆蓋
            return True
        except ImportError as e:
            missing_packages = []
            try:
                import matplotlib
            except ImportError:
                missing_packages.append("matplotlib")
            
            try:
                import numpy
            except ImportError:
                missing_packages.append("numpy")
            
            try:
                from PIL import Image
            except ImportError:
                missing_packages.append("pillow")
            
            print(f"⚠️ 缺少圖形繪製套件: {', '.join(missing_packages)}")
            print("請執行: pip install matplotlib numpy pillow")
            return False
    
    def install_graphics_dependencies(self):
        """嘗試自動安裝圖形繪製套件"""
        try:
            import subprocess
            import sys
            
            packages = ["matplotlib", "numpy", "pillow"]
            self.log_message("🔄 正在安裝圖形繪製套件...")
            
            for package in packages:
                try:
                    __import__(package if package != "pillow" else "PIL")
                except ImportError:
                    self.log_message(f"正在安裝 {package}...")
                    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                    self.log_message(f"✅ {package} 安裝成功")
            
            self.graphics_available = True
            self.log_message("🎨 圖形繪製功能已啟用")
            return True
            
        except Exception as e:
            self.log_message(f"❌ 自動安裝失敗: {str(e)}")
            self.log_message("請手動執行: pip install matplotlib numpy pillow")
            return False
    
    # ====== 圖形化鋼筋繪製方法 ======
    
    def enhanced_draw_rebar_diagram(self, segments, rebar_number="#4"):
        """
        增強版鋼筋圖示生成方法
        生成專業的圖形化鋼筋彎曲形狀圖
        """
        if not self.graphics_available:
            return self.draw_ascii_rebar(segments)
        
        try:
            if not segments:
                return "無分段資料"
            
            # 如果只有一段，繪製直鋼筋
            if len(segments) == 1:
                return self._draw_straight_rebar_diagram(segments[0], rebar_number)
            
            # 多段鋼筋，根據段數選擇不同形狀
            if len(segments) == 2:
                return self._draw_l_shaped_diagram(segments[0], segments[1], rebar_number)
            elif len(segments) == 3:
                return self._draw_u_shaped_diagram(segments[0], segments[1], segments[2], rebar_number)
            else:
                return self._draw_complex_rebar_diagram(segments, rebar_number)
                
        except Exception as e:
            self.log_message(f"⚠️ 圖形生成錯誤: {str(e)}，使用簡化圖示")
            return self.draw_ascii_rebar(segments)
    
    def _draw_straight_rebar_diagram(self, length, rebar_number):
        """繪製直鋼筋的專業圖示"""
        try:
            l_str = str(int(length))
            material_grade = self._get_material_grade(rebar_number)
            
            lines = []
            lines.append(f"直鋼筋 {rebar_number}")
            lines.append(f"長度: {l_str}cm")
            lines.append(f"等級: {material_grade}")
            lines.append("├" + "─" * (len(l_str) + 8) + "┤")
            lines.append(f"  {l_str}cm")
            lines.append("└" + "─" * (len(l_str) + 8) + "┘")
            
            return "\n".join(lines)
            
        except Exception:
            return f"直鋼筋 {rebar_number}: {int(length)}cm"

    def _draw_l_shaped_diagram(self, length1, length2, rebar_number):
        """繪製 L 型鋼筋專業圖示"""
        try:
            l1_str = str(int(length1))
            l2_str = str(int(length2))
            material_grade = self._get_material_grade(rebar_number)
            bend_r = self.bend_radius.get(rebar_number, 5)
            
            lines = []
            lines.append(f"L型鋼筋 {rebar_number}")
            lines.append(f"等級: {material_grade}")
            lines.append(f"彎曲半徑: R{bend_r}cm")
            lines.append("")
            lines.append(f"    {l1_str}cm")
            lines.append("┌" + "─" * (max(len(l1_str), 8) + 2) + "┐")
            lines.append("│" + " " * (max(len(l1_str), 8) + 2) + "│")
            lines.append("│" + " " * (max(len(l1_str), 8) + 2) + "│")
            lines.append("│" + " " * (max(len(l1_str), 8) + 2) + f"│ {l2_str}cm")
            lines.append("│" + " " * (max(len(l1_str), 8) + 2) + "│")
            lines.append("│" + " " * (max(len(l1_str), 8) + 2) + "│")
            lines.append("└" + "─" * (max(len(l1_str), 8) + 2) + "┘")
            
            return "\n".join(lines)
            
        except Exception:
            return f"L型 {rebar_number}: {int(length1)}+{int(length2)}cm"

    def _draw_u_shaped_diagram(self, length1, length2, length3, rebar_number):
        """繪製 U 型鋼筋專業圖示"""
        try:
            l1_str = str(int(length1))
            l2_str = str(int(length2))
            l3_str = str(int(length3))
            material_grade = self._get_material_grade(rebar_number)
            bend_r = self.bend_radius.get(rebar_number, 5)
            
            lines = []
            lines.append(f"U型鋼筋 {rebar_number}")
            lines.append(f"等級: {material_grade}")
            lines.append(f"彎曲半徑: R{bend_r}cm")
            lines.append("")
            lines.append(f"{l2_str}cm     {l1_str}cm     {l3_str}cm")
            lines.append("│" + " " * (max(len(l1_str), 12) + 4) + "│")
            lines.append("│" + " " * (max(len(l1_str), 12) + 4) + "│")
            lines.append("│" + " " * (max(len(l1_str), 12) + 4) + "│")
            lines.append("│" + " " * (max(len(l1_str), 12) + 4) + "│")
            lines.append("└" + "─" * (max(len(l1_str), 12) + 4) + "┘")
            
            return "\n".join(lines)
            
        except Exception:
            return f"U型 {rebar_number}: {int(length1)}+{int(length2)}+{int(length3)}cm"

    def _draw_complex_rebar_diagram(self, segments, rebar_number):
        """繪製複雜多段鋼筋專業圖示"""
        try:
            total_length = sum(segments)
            segment_str = "+".join([str(int(s)) for s in segments])
            material_grade = self._get_material_grade(rebar_number)
            
            lines = []
            lines.append(f"多段彎曲鋼筋 {rebar_number}")
            lines.append(f"等級: {material_grade}")
            lines.append(f"分段: {segment_str}cm")
            lines.append(f"總長: {int(total_length)}cm")
            lines.append("")
            
            # 根據段數創建不同的圖案
            if len(segments) == 4:
                lines.extend([
                    "┌─────────┬─────────┐",
                    "│         │         │",
                    "│         │         │",
                    "│         └─────────┤",
                    "│                   │",
                    "└───────────────────┘"
                ])
            elif len(segments) == 5:
                lines.extend([
                    "┌─────┬─────┬─────┐",
                    "│     │     │     │",
                    "│     │     └─────┤",
                    "│     │           │",
                    "│     └───────────┤",
                    "│                 │",
                    "└─────────────────┘"
                ])
            else:
                # 通用多段圖案
                width = min(len(segments) * 4, 20)
                lines.extend([
                    "┌" + "─┬─" * min(len(segments)-1, 5) + "─┐",
                    "│" + " │ " * min(len(segments)-1, 5) + " │",
                    "│" + " └─" * min(len(segments)-1, 5) + " │",
                    "│" + " " * width + "│",
                    "└" + "─" * width + "┘"
                ])
            
            return "\n".join(lines)
            
        except Exception:
            return f"多段 {rebar_number}: {'+'.join([str(int(s)) for s in segments])}cm"
    
    def _get_material_grade(self, rebar_number):
        """根據鋼筋編號獲取材料等級"""
        if rebar_number.startswith('#'):
            try:
                num = int(rebar_number[1:])
                if num <= 6:
                    return "SD280"
                elif num <= 10:
                    return "SD420"
                else:
                    return "SD490"
            except:
                pass
        return "SD280"
    
    # ====== 保留原有的所有方法 ======
    
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
                           background='black',
                           foreground='#FFD700',  # 亮黃色
                           font=("Segoe UI", 14, "bold"),
                           relief='flat',
                           borderwidth=0,
                           focuscolor='none')
        
        self.style.configure("Secondary.TButton",
                           background='black',
                           foreground='#FFD700',  # 亮黃色
                           font=("Segoe UI", 12),
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
                      background=[('active', '#333333')])  # 深灰色
        
        self.style.map("Secondary.TButton",
                      background=[('active', '#333333')])  # 深灰色
    
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
    
    def create_material_button(self, parent, text, command, is_primary=True):
        """創建 Material Design 風格的按鈕"""
        # 創建外層框架用於陰影效果
        shadow_frame = tk.Frame(parent, bg=self.colors['background'])
        # shadow_frame.pack(side=tk.RIGHT, padx=5, pady=5)  # 由外部決定 pack
        
        # 深灰色文字
        text_color = '#222222'
        
        # 創建按鈕
        btn = tk.Button(shadow_frame, text=text,
                       command=command,
                       bg='#2196F3' if is_primary else '#424242',  # Material Blue 或 Dark Grey
                       fg=text_color,
                       activebackground='#1976D2' if is_primary else '#616161',  # Darker Blue 或 Darker Grey
                       activeforeground=text_color,
                       highlightbackground='#2196F3' if is_primary else '#424242',
                       highlightcolor='#2196F3' if is_primary else '#424242',
                       font=("Segoe UI", 12, "bold"),
                       relief='flat',
                       bd=0,
                       cursor='hand2',
                       padx=20,
                       pady=10)
        btn.pack(fill=tk.BOTH, expand=True)
        
        # 添加圓角效果
        btn.configure(highlightthickness=0)
        
        # 添加 hover 效果
        def on_enter(e):
            btn.configure(bg='#1976D2' if is_primary else '#616161', fg=text_color)
            shadow_frame.configure(bg='#1976D2' if is_primary else '#616161')
        
        def on_leave(e):
            btn.configure(bg='#2196F3' if is_primary else '#424242', fg=text_color)
            shadow_frame.configure(bg=self.colors['background'])
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn, shadow_frame
    
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
        
        # 副標題 - 增加圖形化功能說明
        graphics_status = "✅ 圖形化功能已啟用" if self.graphics_available else "⚠️ 圖形化功能未啟用"
        subtitle_label = tk.Label(header_frame,
                                text=f"專業級 DXF 檔案鋼筋數據分析與 Excel 報表生成工具 | {graphics_status}",
                                bg=self.colors['background'],
                                fg=self.colors['text_secondary'],
                                font=("Segoe UI", 11))
        subtitle_label.pack(pady=(5, 0))
        
        # 版本標籤
        version_label = tk.Label(header_frame,
                               text="v2.1 Professional Edition with Graphics",
                               bg=self.colors['background'],
                               fg=self.colors['success'],
                               font=("Segoe UI", 9, "bold"))
        version_label.pack(pady=(5, 0))
        
        # 如果圖形功能未啟用，添加安裝按鈕
        if not self.graphics_available:
            install_frame = tk.Frame(header_frame, bg=self.colors['background'])
            install_frame.pack(pady=(10, 0))
            
            self.install_button, self.install_shadow = self.create_material_button(
                install_frame, "🔧 安裝圖形繪製套件", self.install_graphics_dependencies, True)
        
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
                                font=("Segoe UI", 11),
                                bg=self.colors['surface'],
                                fg=self.colors['text_primary'],
                                insertbackground=self.colors['primary'],
                                relief='solid', bd=1,
                                selectbackground=self.colors['primary'],
                                selectforeground='white')
        self.cad_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        # 瀏覽按鈕
        self.browse_cad_button, self.browse_cad_shadow = self.create_material_button(
            cad_input_frame, "📂 瀏覽", self.browse_cad_file, True)
        self.browse_cad_shadow.pack(side=tk.RIGHT, padx=(10, 0), ipady=2)
        
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
                             font=("Segoe UI", 11),
                             bg=self.colors['surface'],
                             fg=self.colors['text_primary'],
                             insertbackground=self.colors['primary'],
                             relief='solid', bd=1,
                             selectbackground=self.colors['primary'],
                             selectforeground='white')
        excel_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        # Excel 瀏覽按鈕
        self.browse_excel_button, self.browse_excel_shadow = self.create_material_button(
            excel_input_frame, "💾 另存新檔", self.browse_excel_file, True)
        self.browse_excel_shadow.pack(side=tk.RIGHT, padx=(10, 0), ipady=2)
        
        # 控制按鈕區域
        control_card = self.create_card_frame(main_container, "🎮 執行控制")
        
        button_frame = tk.Frame(control_card, bg=self.colors['surface'])
        button_frame.pack(fill=tk.X)
        
        # 主要按鈕
        button_text = "🚀 開始轉換 (含圖形)" if self.graphics_available else "🚀 開始轉換"
        self.convert_button, self.convert_shadow = self.create_material_button(
            button_frame, button_text, self.start_conversion, True)
        self.convert_shadow.pack(side=tk.LEFT, padx=(0, 10), ipady=2)
        
        # 次要按鈕
        self.reset_button, self.reset_shadow = self.create_material_button(
            button_frame, "🔄 重置", self.reset_form, False)
        self.reset_shadow.pack(side=tk.LEFT, padx=(0, 10), ipady=2)
        
        # 快捷鍵提示
        shortcut_label = tk.Label(button_frame,
                                text="💡 快捷鍵: Ctrl+O(開檔) | Ctrl+S(存檔) | F5(轉換) | Esc(重置)",
                                bg=self.colors['surface'],
                                fg=self.colors['text_secondary'],
                                font=("Segoe UI", 9))
        shortcut_label.pack(side=tk.LEFT, padx=(10, 0))
        
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
                                 bg=self.colors['surface'],
                                 fg=self.colors['text_primary'],
                                 font=("Consolas", 10),
                                 relief='solid', bd=1,
                                 wrap=tk.WORD)
        
        scrollbar = tk.Scrollbar(text_frame, command=self.status_text.yview,
                               bg=self.colors['accent'])
        
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        
        # 初始狀態訊息
        if self.graphics_available:
            self.log_message("🎉 程式已啟動！圖形化功能已啟用，請選擇 CAD 檔案開始轉換流程。")
        else:
            self.log_message("🎉 程式已啟動！使用基本功能模式，請選擇 CAD 檔案開始轉換流程。")
        
        # 設定鍵盤快捷鍵
        self.setup_keyboard_shortcuts()
    
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
        
        if self.graphics_available:
            self.convert_button.config(text="🚀 開始轉換 (含圖形)")
        else:
            self.convert_button.config(text="🚀 開始轉換")
        
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
    
    # ====== 原有的計算方法 ======
    
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
    
    # ====== 保留原有的 ASCII 繪圖方法作為備用 ======
    
    def draw_ascii_rebar(self, segments):
        """使用 ASCII 字元繪製彎折示意圖 (備用方法)"""
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
    
    # ====== 主要轉換方法 ======
    
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
                        material_grade = self._get_material_grade(number)
                        
                        # 建立資料
                        data = {
                            "編號": number,
                            "長度(cm)": round(length_cm, 2),
                            "數量": count,
                            "單位重(kg/m)": unit_weight,
                            "總重量(kg)": weight,
                            "材料等級": material_grade,
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
                            self.log_message(f"✅ 鋼筋 #{valid_rebar_count}: {number}, {length_cm}cm, {count}支, {material_grade}")
                
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
            base_columns = ["編號", "長度(cm)", "數量", "總重量(kg)", "材料等級", "鋼筋圖示", "備註"]
            
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
            
            # 添加圖示欄位 - 使用新的圖形化方法
            self.update_progress(percentage=75, detail="正在生成專業鋼筋圖示...")
            
            if self.graphics_available:
                self.log_message("🎨 使用圖形化鋼筋圖示生成...")
            else:
                self.log_message("📝 使用基本 ASCII 圖示生成...")
            
            for index, row in df.iterrows():
                segments = []
                for key in sorted([k for k in row.keys() if k.endswith("(cm)") and k != "長度(cm)"]):
                    if not pd.isna(row[key]):
                        segments.append(row[key])
                
                # 使用新的增強版圖示生成方法
                df.at[index, "鋼筋圖示"] = self.enhanced_draw_rebar_diagram(segments, row["編號"])
            
            df = df.reindex(columns=columns)
            
            # 步驟 7: 格式化 Excel
            self.update_progress(6, "正在格式化 Excel...")
            
            try:
                # 創建 Excel 工作簿
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "鋼筋計料表"
                
                # 設定標題
                project_name = os.path.basename(self.cad_path.get())
                graphics_note = "（含專業圖形化圖示）" if self.graphics_available else "（基本模式）"
                
                ws['A1'] = f"🏗️ 專業鋼筋計料表 {graphics_note}"
                ws['A2'] = f"專案名稱: {project_name}"
                ws['A3'] = f"生成時間: {time.strftime('%Y-%m-%d %H:%M:%S')}"
                ws['A4'] = f"工具版本: CAD 鋼筋計料轉換工具 Pro v2.1"
                
                # 合併儲存格
                ws.merge_cells('A1:H1')
                ws.merge_cells('A2:H2')
                ws.merge_cells('A3:H3')
                ws.merge_cells('A4:H4')
                
                # 設定標題樣式
                title_font = Font(bold=True, size=16)
                subtitle_font = Font(bold=True, size=12)
                info_font = Font(size=10)
                
                ws['A1'].font = title_font
                ws['A2'].font = subtitle_font
                ws['A3'].font = info_font
                ws['A4'].font = info_font
                
                # 居中對齊
                title_align = Alignment(horizontal='center', vertical='center')
                ws['A1'].alignment = title_align
                ws['A2'].alignment = title_align
                ws['A3'].alignment = title_align
                ws['A4'].alignment = title_align
                
                # 設定表頭 (從第6行開始)
                headers = columns
                header_row = 6
                
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
                
                # 寫入資料 (從第7行開始)
                data_start_row = 7
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
                        
                        # 根據材料等級設定顏色
                        if col_name == "材料等級":
                            grade = row.get("材料等級", "")
                            if grade == "SD280":
                                cell.fill = PatternFill(start_color="E7F3FF", end_color="E7F3FF", fill_type="solid")
                            elif grade == "SD420":
                                cell.fill = PatternFill(start_color="FFF2E7", end_color="FFF2E7", fill_type="solid")
                            elif grade == "SD490":
                                cell.fill = PatternFill(start_color="F3E7FF", end_color="F3E7FF", fill_type="solid")
                    
                    # 調整行高以適應圖示
                    ws.row_dimensions[row_num].height = 120 if self.graphics_available else 80
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
                    "編號": 12,
                    "長度(cm)": 15,
                    "數量": 10,
                    "總重量(kg)": 15,
                    "材料等級": 15,
                    "鋼筋圖示": 80 if self.graphics_available else 40,
                    "備註": 30
                }
                
                # 設定分段長度欄位的寬度
                for col in segment_columns:
                    column_widths[col] = 12
                
                # 根據欄位名稱設定寬度
                for col_num, header in enumerate(headers, 1):
                    if header in column_widths:
                        column_letter = openpyxl.utils.get_column_letter(col_num)
                        ws.column_dimensions[column_letter].width = column_widths[header]
                
                # 添加材料等級說明
                legend_start_row = summary_row + 3
                ws.cell(row=legend_start_row, column=1).value = "材料等級說明:"
                ws.cell(row=legend_start_row, column=1).font = Font(bold=True)
                
                legend_data = [
                    ("SD280", "280 MPa", "一般結構用鋼筋"),
                    ("SD420", "420 MPa", "高強度結構鋼筋"),
                    ("SD490", "490 MPa", "特殊高強度鋼筋")
                ]
                
                for i, (grade, strength, desc) in enumerate(legend_data):
                    row = legend_start_row + 1 + i
                    ws.cell(row=row, column=1).value = grade
                    ws.cell(row=row, column=2).value = strength
                    ws.cell(row=row, column=3).value = desc
                    
                    # 設定說明樣式
                    for col in range(1, 4):
                        cell = ws.cell(row=row, column=col)
                        cell.font = Font(size=9)
                        if grade == "SD280":
                            cell.fill = PatternFill(start_color="E7F3FF", end_color="E7F3FF", fill_type="solid")
                        elif grade == "SD420":
                            cell.fill = PatternFill(start_color="FFF2E7", end_color="FFF2E7", fill_type="solid")
                        elif grade == "SD490":
                            cell.fill = PatternFill(start_color="F3E7FF", end_color="F3E7FF", fill_type="solid")
                
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
            self.log_message(f"   • 圖示類型: {'專業圖形化' if self.graphics_available else '基本ASCII'}")
            self.log_message(f"   • 輸出檔案: {os.path.basename(self.excel_path.get())}")
            self.log_message("=" * 60)
            
            # 恢復按鈕狀態
            if self.graphics_available:
                self.convert_button.config(state="normal", text="🚀 開始轉換 (含圖形)", bg=self.colors['success'])
            else:
                self.convert_button.config(state="normal", text="🚀 開始轉換", bg=self.colors['success'])
            
            # 顯示完成對話框
            graphics_note = "包含專業圖形化鋼筋圖示" if self.graphics_available else "使用基本圖示模式"
            result_message = f"""🎉 轉換完成！

📊 處理結果:
• 有效鋼筋: {valid_rebar_count} 個
• 鋼筋類型: {len(rebar_types)} 種  
• 總數量: {total_quantity} 支
• 總重量: {total_weight:.2f} kg
• 處理時間: {self.format_time(elapsed_time)}
• 圖示模式: {graphics_note}

💾 檔案已儲存至:
{self.excel_path.get()}

🎨 新功能特色:
✅ 自動材料等級識別 (SD280/SD420/SD490)
✅ 專業鋼筋彎曲形狀圖示
✅ 增強版 Excel 格式化
✅ 智能圖形降級機制

感謝您使用 CAD 鋼筋計料轉換工具 Pro v2.1！"""
            
            messagebox.showinfo("✅ 轉換完成", result_message)
            
        except Exception as e:
            self.log_message(f"❌ 錯誤: {str(e)}")
            self.update_progress(percentage=0, detail="轉換失敗")
            
            if self.graphics_available:
                self.convert_button.config(state="normal", text="🚀 開始轉換 (含圖形)", bg=self.colors['success'])
            else:
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