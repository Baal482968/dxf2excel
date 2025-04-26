import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import ezdxf
import math
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import threading
import time
import re

class CADtoExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("CAD 資料轉換 Excel 工具")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # 設定樣式
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("TLabel", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))
        
        # 材質密度設定 (kg/m³)
        self.material_density = {
            "鋼鐵": 7850,
            "鋁": 2700,
            "銅": 8960,
            "不鏽鋼": 8000
        }
        
        # 預設材質
        self.default_material = "鋼鐵"
        
        # 建立主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題
        header_label = ttk.Label(main_frame, text="CAD 資料轉換 Excel 工具", style="Header.TLabel")
        header_label.pack(pady=10)
        
        # 輸入檔案框架
        input_frame = ttk.LabelFrame(main_frame, text="輸入檔案", padding="10")
        input_frame.pack(fill=tk.X, pady=10)
        
        # CAD 檔案選擇
        cad_frame = ttk.Frame(input_frame)
        cad_frame.pack(fill=tk.X, pady=5)
        
        self.cad_path = tk.StringVar()
        ttk.Label(cad_frame, text="CAD 檔案:").pack(side=tk.LEFT)
        ttk.Entry(cad_frame, textvariable=self.cad_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(cad_frame, text="瀏覽...", command=self.browse_cad_file).pack(side=tk.LEFT)
        
        # 輸出檔案框架
        output_frame = ttk.LabelFrame(main_frame, text="輸出設定", padding="10")
        output_frame.pack(fill=tk.X, pady=10)
        
        # Excel 檔案選擇
        excel_frame = ttk.Frame(output_frame)
        excel_frame.pack(fill=tk.X, pady=5)
        
        self.excel_path = tk.StringVar()
        ttk.Label(excel_frame, text="Excel 檔案:").pack(side=tk.LEFT)
        ttk.Entry(excel_frame, textvariable=self.excel_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(excel_frame, text="瀏覽...", command=self.browse_excel_file).pack(side=tk.LEFT)
        
        # 轉換設定框架
        settings_frame = ttk.LabelFrame(main_frame, text="轉換設定", padding="10")
        settings_frame.pack(fill=tk.X, pady=10)
        
        # 材質選擇
        material_frame = ttk.Frame(settings_frame)
        material_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(material_frame, text="選擇材質:").pack(side=tk.LEFT)
        self.material_var = tk.StringVar(value=self.default_material)
        material_combo = ttk.Combobox(material_frame, textvariable=self.material_var, values=list(self.material_density.keys()))
        material_combo.pack(side=tk.LEFT, padx=5)
        
        # 要轉換的實體類型
        entity_frame = ttk.Frame(settings_frame)
        entity_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(entity_frame, text="選擇實體類型:").pack(side=tk.LEFT)
        
        self.entity_types = {
            "TEXT": tk.BooleanVar(value=True),
            "LINE": tk.BooleanVar(value=True),
            "CIRCLE": tk.BooleanVar(value=True),
            "ARC": tk.BooleanVar(value=True),
            "POLYLINE": tk.BooleanVar(value=True),
            "LWPOLYLINE": tk.BooleanVar(value=True),
            "BLOCK": tk.BooleanVar(value=True)
        }
        
        entity_check_frame = ttk.Frame(settings_frame)
        entity_check_frame.pack(fill=tk.X, pady=5)
        
        col = 0
        for entity, var in self.entity_types.items():
            ttk.Checkbutton(entity_check_frame, text=entity, variable=var).grid(row=0, column=col, padx=5)
            col += 1
            if col > 3:
                col = 0
        
        # 屬性選擇
        attr_frame = ttk.Frame(settings_frame)
        attr_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(attr_frame, text="選擇屬性:").pack(side=tk.LEFT)
        
        self.attributes = {
            "圖層": tk.BooleanVar(value=True),
            "顏色": tk.BooleanVar(value=True),
            "文字內容": tk.BooleanVar(value=True),
            "尺寸": tk.BooleanVar(value=True),
            "線型": tk.BooleanVar(value=True),
            "號數": tk.BooleanVar(value=True)  # 新增號數屬性選項
        }
        
        attr_check_frame = ttk.Frame(settings_frame)
        attr_check_frame.pack(fill=tk.X, pady=5)
        
        col = 0
        for attr, var in self.attributes.items():
            ttk.Checkbutton(attr_check_frame, text=attr, variable=var).grid(row=0, column=col, padx=5)
            col += 1
            if col > 3:
                col = 0
        
        # 執行按鈕
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        ttk.Button(button_frame, text="開始轉換", command=self.start_conversion).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="重置", command=self.reset_form).pack(side=tk.RIGHT, padx=5)
        
        # 狀態框架
        status_frame = ttk.LabelFrame(main_frame, text="處理狀態", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 進度條
        self.progress = ttk.Progressbar(status_frame, orient="horizontal", length=100, mode="determinate")
        self.progress.pack(fill=tk.X, pady=10)
        
        # 狀態文字
        self.status_text = tk.Text(status_frame, height=10, width=50)
        self.status_text.pack(fill=tk.BOTH, expand=True)
        
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
            default_excel = os.path.splitext(filename)[0] + ".xlsx"
            self.excel_path.set(default_excel)
    
    def browse_excel_file(self):
        filetypes = (
            ("Excel 檔案", "*.xlsx"),
            ("所有檔案", "*.*")
        )
        filename = filedialog.asksaveasfilename(
            title="儲存 Excel 檔案",
            filetypes=filetypes,
            defaultextension=".xlsx",
            confirmoverwrite=True  # 啟用覆蓋確認
        )
        if filename:
            self.excel_path.set(filename)
    
    def reset_form(self):
        self.cad_path.set("")
        self.excel_path.set("")
        for var in self.entity_types.values():
            var.set(True)
        for var in self.attributes.values():
            var.set(True)
        self.status_text.delete(1.0, tk.END)
        self.progress["value"] = 0
    
    def log_message(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def start_conversion(self):
        if not self.cad_path.get():
            messagebox.showerror("錯誤", "請選擇 CAD 檔案")
            return
        
        if not self.excel_path.get():
            messagebox.showerror("錯誤", "請指定 Excel 輸出檔案")
            return
        
        # 清空狀態文字
        self.status_text.delete(1.0, tk.END)
        
        # 在新線程中執行轉換，避免 UI 凍結
        conversion_thread = threading.Thread(target=self.convert_cad_to_excel)
        conversion_thread.start()
    
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
    
    def extract_number_from_text(self, text):
        """從帶有#的文字中提取號數"""
        if text and text.startswith('#'):
            # 去除#符號並清理空白
            text_without_hash = text[1:].strip()
            # 從文字中提取數字部分
            number_match = re.search(r'^\d+', text_without_hash)
            if number_match:
                return number_match.group()
        return ""
    
    def calculate_weight(self, length, diameter=None, thickness=None):
        """計算重量 (kg)"""
        try:
            density = self.material_density[self.material_var.get()]
            if diameter and thickness:  # 圓管
                outer_radius = diameter / 2
                inner_radius = outer_radius - thickness
                cross_section = math.pi * (outer_radius**2 - inner_radius**2)
                volume = cross_section * length
            else:  # 實心圓棒
                cross_section = math.pi * (diameter/2)**2 if diameter else 0.0001  # 預設直徑1mm
                volume = cross_section * length
            
            weight = volume * density / 1000000  # 轉換為kg
            return round(weight, 3)
        except:
            return 0
    
    def convert_cad_to_excel(self):
        try:
            self.progress["value"] = 10
            self.log_message(f"開始處理 CAD 檔案: {self.cad_path.get()}")
            
            # 檢查檔案是否存在
            if not os.path.isfile(self.cad_path.get()):
                self.log_message("錯誤: 找不到指定的 CAD 檔案")
                messagebox.showerror("錯誤", "找不到指定的 CAD 檔案")
                self.progress["value"] = 0
                return
                
            # 檢查檔案副檔名
            file_ext = os.path.splitext(self.cad_path.get())[1].lower()
            if file_ext != '.dxf':
                self.log_message("警告: 非 DXF 檔案可能無法正確讀取")
                messagebox.showwarning("警告", "此程式主要支援 DXF 檔案格式，其他格式可能無法正確讀取。")
            
            # 讀取 DXF 檔案
            try:
                doc = ezdxf.readfile(self.cad_path.get())
                msp = doc.modelspace()
                self.log_message("CAD 檔案載入成功")
            except Exception as e:
                self.log_message(f"錯誤: 無法讀取 CAD 檔案: {str(e)}")
                messagebox.showerror("錯誤", f"無法讀取 CAD 檔案: {str(e)}")
                self.progress["value"] = 0
                return
            
            self.progress["value"] = 20
            
            # 收集要轉換的實體類型
            selected_entities = [entity for entity, var in self.entity_types.items() if var.get()]
            self.log_message(f"選定的實體類型: {', '.join(selected_entities)}")
            
            # 準備資料表
            data = []
            
            # 獲取實體
            self.log_message("正在處理 CAD 實體...")
            
            # 計算實體總數
            entity_count = 0
            for entity_type in selected_entities:
                entity_count += len(msp.query(entity_type))
            
            # 如果沒有實體
            if entity_count == 0:
                self.log_message("未找到符合條件的實體")
                self.progress["value"] = 100
                messagebox.showinfo("完成", "處理完成，但未找到符合條件的實體")
                return
            
            processed_count = 0
            
            # 統計資料
            stats = {
                "總數量": 0,
                "總長度": 0,
                "總重量": 0,
                "各類型數量": {}
            }
            
            # 處理 TEXT 實體
            if "TEXT" in selected_entities:
                text_count = len(msp.query("TEXT"))
                stats["各類型數量"]["文字"] = text_count
                stats["總數量"] += text_count
                
                for text in msp.query("TEXT"):
                    text_content = text.dxf.text if self.attributes["文字內容"].get() else ""
                    number = ""
                    if self.attributes["號數"].get() and text_content.startswith('#'):
                        number = self.extract_number_from_text(text_content)
                    
                    row = {
                        "實體類型": "文字",
                        "內容": text_content,
                        "號數": number,
                        "圖層": text.dxf.layer if self.attributes["圖層"].get() else "",
                        "顏色": text.dxf.color if self.attributes["顏色"].get() else "",
                        "高度": text.dxf.height if self.attributes["尺寸"].get() else "",
                        "線型": text.dxf.linetype if self.attributes["線型"].get() else ""
                    }
                    data.append(row)
                    processed_count += 1
                    if processed_count % 100 == 0:
                        progress = 20 + int(60 * processed_count / entity_count)
                        self.progress["value"] = min(80, progress)
                        self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            # 處理 LINE 實體
            if "LINE" in selected_entities:
                line_count = len(msp.query("LINE"))
                stats["各類型數量"]["直線"] = line_count
                stats["總數量"] += line_count
                
                for line in msp.query("LINE"):
                    start_point = (line.dxf.start.x, line.dxf.start.y, line.dxf.start.z)
                    end_point = (line.dxf.end.x, line.dxf.end.y, line.dxf.end.z)
                    line_length = self.calculate_line_length(start_point, end_point)
                    stats["總長度"] += line_length
                    
                    # 估算重量 (假設直徑10mm)
                    weight = self.calculate_weight(line_length, diameter=10)
                    stats["總重量"] += weight
                    
                    row = {
                        "實體類型": "直線",
                        "內容": "",
                        "號數": "",
                        "長度": line_length if self.attributes["尺寸"].get() else "",
                        "重量": weight if self.attributes["尺寸"].get() else "",
                        "圖層": line.dxf.layer if self.attributes["圖層"].get() else "",
                        "顏色": line.dxf.color if self.attributes["顏色"].get() else "",
                        "線型": line.dxf.linetype if self.attributes["線型"].get() else ""
                    }
                    data.append(row)
                    processed_count += 1
                    if processed_count % 100 == 0:
                        progress = 20 + int(60 * processed_count / entity_count)
                        self.progress["value"] = min(80, progress)
                        self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            # 處理 CIRCLE 實體
            if "CIRCLE" in selected_entities:
                circle_count = len(msp.query("CIRCLE"))
                stats["各類型數量"]["圓"] = circle_count
                stats["總數量"] += circle_count
                
                for circle in msp.query("CIRCLE"):
                    try:
                        radius = circle.dxf.radius
                        circumference = 2 * math.pi * radius
                        stats["總長度"] += circumference
                        
                        # 估算重量 (假設直徑10mm)
                        weight = self.calculate_weight(circumference, diameter=10)
                        stats["總重量"] += weight
                        
                        row = {
                            "實體類型": "圓",
                            "內容": "",
                            "號數": "",
                            "半徑": radius if self.attributes["尺寸"].get() else "",
                            "周長": circumference if self.attributes["尺寸"].get() else "",
                            "重量": weight if self.attributes["尺寸"].get() else "",
                            "圖層": circle.dxf.layer if self.attributes["圖層"].get() else "",
                            "顏色": circle.dxf.color if self.attributes["顏色"].get() else "",
                            "線型": circle.dxf.linetype if self.attributes["線型"].get() else ""
                        }
                        data.append(row)
                    except AttributeError as e:
                        self.log_message(f"警告: 處理圓形實體時發生問題: {str(e)}")
                    
                    processed_count += 1
                    if processed_count % 100 == 0:
                        progress = 20 + int(60 * processed_count / entity_count)
                        self.progress["value"] = min(80, progress)
                        self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            # 處理 ARC 實體
            if "ARC" in selected_entities:
                arc_count = len(msp.query("ARC"))
                stats["各類型數量"]["弧"] = arc_count
                stats["總數量"] += arc_count
                
                for arc in msp.query("ARC"):
                    try:
                        radius = arc.dxf.radius
                        start_angle = math.radians(arc.dxf.start_angle)
                        end_angle = math.radians(arc.dxf.end_angle)
                        arc_length = radius * abs(end_angle - start_angle)
                        stats["總長度"] += arc_length
                        
                        # 估算重量 (假設直徑10mm)
                        weight = self.calculate_weight(arc_length, diameter=10)
                        stats["總重量"] += weight
                        
                        row = {
                            "實體類型": "弧",
                            "內容": "",
                            "號數": "",
                            "半徑": radius if self.attributes["尺寸"].get() else "",
                            "弧長": arc_length if self.attributes["尺寸"].get() else "",
                            "重量": weight if self.attributes["尺寸"].get() else "",
                            "起始角度": arc.dxf.start_angle if self.attributes["尺寸"].get() else "",
                            "結束角度": arc.dxf.end_angle if self.attributes["尺寸"].get() else "",
                            "圖層": arc.dxf.layer if self.attributes["圖層"].get() else "",
                            "顏色": arc.dxf.color if self.attributes["顏色"].get() else "",
                            "線型": arc.dxf.linetype if self.attributes["線型"].get() else ""
                        }
                        data.append(row)
                    except AttributeError as e:
                        self.log_message(f"警告: 處理弧形實體時發生問題: {str(e)}")
                    
                    processed_count += 1
                    if processed_count % 100 == 0:
                        progress = 20 + int(60 * processed_count / entity_count)
                        self.progress["value"] = min(80, progress)
                        self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            # 處理 POLYLINE 和 LWPOLYLINE 實體
            for entity_type in ["POLYLINE", "LWPOLYLINE"]:
                if entity_type in selected_entities:
                    polyline_count = len(msp.query(entity_type))
                    stats["各類型數量"]["多段線"] = polyline_count
                    stats["總數量"] += polyline_count
                    
                    for polyline in msp.query(entity_type):
                        try:
                            if entity_type == "POLYLINE":
                                vertices = list(polyline.vertices())
                                points = [(v.dxf.location.x, v.dxf.location.y, v.dxf.location.z) for v in vertices]
                            else:  # LWPOLYLINE
                                points = list(polyline.points())
                                points = [(p[0], p[1], 0) if len(p) >= 2 else (p[0], 0, 0) for p in points]
                            
                            try:
                                if hasattr(polyline, 'length') and callable(getattr(polyline, 'length')):
                                    length = polyline.length()
                                else:
                                    length = self.calculate_polyline_length(points)
                            except:
                                length = 0
                                self.log_message(f"警告: 無法計算多段線長度")
                            
                            stats["總長度"] += length
                            
                            # 估算重量 (假設直徑10mm)
                            weight = self.calculate_weight(length, diameter=10)
                            stats["總重量"] += weight
                            
                            is_closed = "未知"
                            try:
                                is_closed = "是" if polyline.closed else "否"
                            except:
                                pass
                                
                            row = {
                                "實體類型": "多段線",
                                "內容": "",
                                "號數": "",
                                "點數": len(points) if self.attributes["尺寸"].get() else "",
                                "長度": length if self.attributes["尺寸"].get() else "",
                                "重量": weight if self.attributes["尺寸"].get() else "",
                                "是否閉合": is_closed if self.attributes["尺寸"].get() else "",
                                "圖層": polyline.dxf.layer if self.attributes["圖層"].get() else "",
                                "顏色": polyline.dxf.color if self.attributes["顏色"].get() else "",
                                "線型": polyline.dxf.linetype if self.attributes["線型"].get() else ""
                            }
                            data.append(row)
                        except Exception as e:
                            self.log_message(f"警告: 處理多段線實體時發生問題: {str(e)}")
                        
                        processed_count += 1
                        if processed_count % 100 == 0:
                            progress = 20 + int(60 * processed_count / entity_count)
                            self.progress["value"] = min(80, progress)
                            self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            # 處理 BLOCK 實體
            if "BLOCK" in selected_entities:
                block_count = len(msp.query("INSERT"))
                stats["各類型數量"]["圖塊"] = block_count
                stats["總數量"] += block_count
                
                for block_ref in msp.query("INSERT"):
                    try:
                        row = {
                            "實體類型": "圖塊",
                            "內容": block_ref.dxf.name,
                            "號數": "",
                            "X縮放": block_ref.dxf.xscale if self.attributes["尺寸"].get() else "",
                            "Y縮放": block_ref.dxf.yscale if self.attributes["尺寸"].get() else "",
                            "旋轉角度": block_ref.dxf.rotation if self.attributes["尺寸"].get() else "",
                            "圖層": block_ref.dxf.layer if self.attributes["圖層"].get() else "",
                            "顏色": block_ref.dxf.color if self.attributes["顏色"].get() else "",
                        }
                        data.append(row)
                    except AttributeError as e:
                        self.log_message(f"警告: 處理圖塊實體時發生問題: {str(e)}")
                    
                    processed_count += 1
                    if processed_count % 100 == 0:
                        progress = 20 + int(60 * processed_count / entity_count)
                        self.progress["value"] = min(80, progress)
                        self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            self.progress["value"] = 80
            self.log_message(f"已處理完成 {processed_count} 個實體")
            
            # 添加統計資料
            stats_row = {
                "實體類型": "統計",
                "總數量": stats["總數量"],
                "總長度": round(stats["總長度"], 2),
                "總重量": round(stats["總重量"], 2),
                "各類型數量": ", ".join([f"{k}: {v}" for k, v in stats["各類型數量"].items()])
            }
            data.append(stats_row)
            
            # 檢查數據是否為空
            if not data:
                self.log_message("警告: 沒有找到任何可轉換的數據")
                messagebox.showwarning("警告", "沒有找到任何可轉換的數據")
                self.progress["value"] = 0
                return
            
            # 創建 DataFrame
            df = pd.DataFrame(data)
            
            # 寫入 Excel
            self.log_message(f"正在寫入資料到 Excel: {self.excel_path.get()}")
            
            try:
                # 使用 pandas 直接寫入，更為穩定
                df.to_excel(self.excel_path.get(), sheet_name='CAD數據', index=False)
                self.log_message("Excel 已基本寫入，正在套用格式...")
                
                # 使用 openpyxl 進行格式化
                wb = openpyxl.load_workbook(self.excel_path.get())
                ws = wb['CAD數據']
                
                # 設定標題樣式
                header_font = Font(bold=True)
                header_alignment = Alignment(horizontal='center', vertical='center')
                header_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                # 應用標題樣式
                for col_num in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.font = header_font
                    cell.alignment = header_alignment
                    cell.border = header_border
                
                # 調整列寬
                for column in ws.columns:
                    max_length = 0
                    column_letter = openpyxl.utils.get_column_letter(column[0].column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = min(len(str(cell.value)), 50)  # 限制最大寬度
                        except:
                            pass
                    adjusted_width = (max_length + 2) * 1.2
                    ws.column_dimensions[column_letter].width = min(adjusted_width, 50)  # 限制最大寬度
                
                # 儲存檔案
                wb.save(self.excel_path.get())
                self.log_message("Excel 格式套用完成")
                
            except Exception as e:
                self.log_message(f"警告: Excel 格式化時出現問題: {str(e)}")
                self.log_message("嘗試使用基本儲存方式...")
                # 如果格式化失敗，至少確保數據被保存
                df.to_excel(self.excel_path.get(), sheet_name='CAD數據', index=False)
            
            self.progress["value"] = 100
            self.log_message("轉換完成!")
            messagebox.showinfo("成功", f"CAD資料已成功轉換為Excel檔案:\n{self.excel_path.get()}")
            
        except Exception as e:
            self.log_message(f"錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"轉換過程中發生錯誤:\n{str(e)}")
            self.progress["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = CADtoExcelConverter(root)
    root.mainloop()