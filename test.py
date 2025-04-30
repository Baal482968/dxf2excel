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

class CADtoExcelConverter:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("CAD 鋼筋計料轉換 Excel 工具")
            self.root.geometry("700x600")
            self.root.resizable(True, True)
            
            # 設定樣式
            self.style = ttk.Style()
            self.style.configure("TButton", font=("Arial", 10))
            self.style.configure("TLabel", font=("Arial", 10))
            self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))
            
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
                "#2": 0.249,
                "#3": 0.561,
                "#4": 0.996,
                "#5": 1.552,
                "#6": 2.235,
                "#7": 3.042,
                "#8": 3.973,
                "#9": 5.026,
                "#10": 6.404,
                "#11": 7.906,
                "#12": 11.38,
                "#13": 13.87,
                "#14": 14.59,
                "#15": 20.24,
                "#16": 25.00,
                "#17": 31.20,
                "#18": 39.70
            }
            
            # 建立主框架
            main_frame = ttk.Frame(root, padding="20")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # 標題
            header_label = ttk.Label(main_frame, text="CAD 鋼筋計料轉換 Excel 工具", style="Header.TLabel")
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
                "CIRCLE": tk.BooleanVar(value=False),
                "ARC": tk.BooleanVar(value=False),
                "POLYLINE": tk.BooleanVar(value=True),
                "LWPOLYLINE": tk.BooleanVar(value=True),
                "BLOCK": tk.BooleanVar(value=False)
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
                "線型": tk.BooleanVar(value=False),
                "號數": tk.BooleanVar(value=True)
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
            
            # 加入滾動條
            scrollbar = ttk.Scrollbar(self.status_text, command=self.status_text.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.status_text.config(yscrollcommand=scrollbar.set)
            
            # 初始狀態訊息
            self.log_message("程式已啟動，請選擇 CAD 檔案並設定轉換選項。")
            
        except Exception as e:
            print(f"初始化錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"程式初始化時發生錯誤:\n{str(e)}")
    
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
        self.entity_types["TEXT"].set(True)
        self.entity_types["LINE"].set(True)
        self.entity_types["POLYLINE"].set(True)
        self.entity_types["LWPOLYLINE"].set(True)
        self.entity_types["CIRCLE"].set(False)
        self.entity_types["ARC"].set(False)
        self.entity_types["BLOCK"].set(False)
        
        for var in self.attributes.values():
            var.set(True)
        self.attributes["線型"].set(False)
        
        self.status_text.delete(1.0, tk.END)
        self.progress["value"] = 0
    
    def log_message(self, message):
        try:
            # 確保訊息以換行符結尾
            if not message.endswith('\n'):
                message += '\n'
            
            # 插入訊息並滾動到底部
            self.status_text.insert(tk.END, message)
            self.status_text.see(tk.END)
            
            # 強制更新 UI
            self.root.update_idletasks()
            
            # 同時輸出到控制台，方便除錯
            print(message.strip())
        except Exception as e:
            print(f"記錄訊息時發生錯誤: {str(e)}")
    
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
    
    def extract_rebar_info(self, text):
        """從文字中提取鋼筋信息"""
        if not text:
            return None, None, None, None
            
        number = ""
        diameter = ""
        count = 1
        length = None
        
        # 尋找鋼筋號數 (支援多種格式)
        # 格式1: #3, #4 等
        number_match = re.search(r'#(\d+)', text)
        if number_match:
            number = "#" + number_match.group(1)
            diameter = self.get_rebar_diameter(number)
        
        # 格式2: D13, D16 等
        if not number:
            number_match = re.search(r'D(\d+)', text)
            if number_match:
                diameter = float(number_match.group(1))
                # 根據直徑反推號數
                for num, dia in self.rebar_diameter.items():
                    if abs(dia - diameter) < 0.1:
                        number = num
                        break
        
        # 格式3: 直接寫直徑，如 13mm, 16mm 等
        if not number:
            number_match = re.search(r'(\d+(?:\.\d+)?)\s*mm', text)
            if number_match:
                diameter = float(number_match.group(1))
                # 根據直徑反推號數
                for num, dia in self.rebar_diameter.items():
                    if abs(dia - diameter) < 0.1:
                        number = num
                        break
        
        # 尋找長度和數量 (支援多種格式)
        # 格式1: #8-100+1200x3 (多段長度)
        length_count_match = re.search(r'[#D]?\d+[-_](?:(\d+(?:\+\d+)*))[xX×*-](\d+)', text)
        if length_count_match:
            try:
                # 解析多段長度
                length_parts = length_count_match.group(1).split('+')
                total_length = sum(float(part) for part in length_parts)
                length = total_length
                count = int(length_count_match.group(2))
            except:
                length = None
                count = 1
        else:
            # 格式2: 只有數量
            count_match = re.search(r'[xX×*-](\d+)', text)
            if count_match:
                try:
                    count = int(count_match.group(1))
                except:
                    count = 1
        
        return number, diameter, count, length
    
    def get_rebar_diameter(self, number):
        """根據鋼筋號數獲取直徑(mm)"""
        rebar_diameter = {
            "#2": 6.4,   # 1/4"
            "#3": 9.5,   # 3/8"
            "#4": 12.7,  # 1/2"
            "#5": 15.9,  # 5/8"
            "#6": 19.1,  # 3/4"
            "#7": 22.2,  # 7/8"
            "#8": 25.4,  # 1"
            "#9": 28.7,  # 1-1/8"
            "#10": 32.3, # 1-1/4"
            "#11": 35.8, # 1-3/8"
            "#12": 43.0, # 1-1/2"
            "#13": 43.0, # 上面開始就不確定了
            "#14": 43.0,
            "#15": 43.0,
            "#16": 43.0,
            "#17": 43.0,
            "#18": 57.3  # 2-1/4"
        }
        
        return rebar_diameter.get(number, "")
    
    def get_rebar_unit_weight(self, number):
        """根據鋼筋號數獲取單位重量(kg/m)"""
        return self.rebar_unit_weight.get(number, 0)
    
    def calculate_rebar_weight(self, number, length, count=1):
        """計算鋼筋總重量"""
        unit_weight = self.get_rebar_unit_weight(number)
        if unit_weight and length:
            # 將長度從cm轉換為m
            length_m = length / 100.0
            return round(unit_weight * length_m * count, 2)
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
            
            # 準備鋼筋資料表
            rebar_data = []
            
            # 編號計數器 (用於沒有號數的實體)
            unnamed_counter = 1
            
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
            
            # 鋼筋統計數據
            rebar_stats = {
                "總數量": 0,
                "總長度": 0,
                "總重量": 0,
                "各類型數量": {}
            }
            
            # 處理 TEXT 實體 (通常包含鋼筋標記)
            text_entities = {}
            if "TEXT" in selected_entities:
                for text in msp.query("TEXT"):
                    text_content = text.dxf.text if self.attributes["文字內容"].get() else ""
                    if text_content:
                        # 儲存文本位置和內容，稍後用於匹配線條
                        position = (text.dxf.insert.x, text.dxf.insert.y)
                        text_entities[position] = text_content
                        
                        # 直接從文字提取鋼筋資訊
                        number, diameter, count, length = self.extract_rebar_info(text_content)
                        if number:
                            rebar_stats["各類型數量"][number] = rebar_stats["各類型數量"].get(number, 0) + count
                            rebar_stats["總數量"] += count
                            
                            # 如果有長度資訊，直接加入資料表
                            if length is not None:
                                # 將長度從mm轉換為cm
                                length_cm = length / 10.0
                                # 計算重量
                                unit_weight = self.get_rebar_unit_weight(number) if number.startswith("#") else 0
                                weight = self.calculate_rebar_weight(number, length_cm, count) if number.startswith("#") else 0
                                
                                # 新增到資料表
                                rebar_data.append({
                                    "編號": number,
                                    "直徑": diameter,
                                    "長度(cm)": round(length_cm, 2),
                                    "數量": count,
                                    "單位重(kg/m)": unit_weight,
                                    "總重量(kg)": weight,
                                    "圖層": text.dxf.layer if self.attributes["圖層"].get() else "",
                                    "備註": text_content
                                })
                                
                                # 更新統計數據
                                rebar_stats["總長度"] += length_cm * count
                                rebar_stats["總重量"] += weight
                                
                                # 記錄詳細資訊
                                self.log_message(f"找到鋼筋標記: {text_content}")
                                self.log_message(f"  號數: {number}")
                                self.log_message(f"  直徑: {diameter}mm")
                                self.log_message(f"  長度: {length_cm}cm")
                                self.log_message(f"  數量: {count}")
                                self.log_message(f"  重量: {weight}kg")
            
            # 處理線條和多段線實體 (鋼筋主體)
            if "LINE" in selected_entities:
                for line in msp.query("LINE"):
                    # 取得線條資訊
                    start_point = (line.dxf.start.x, line.dxf.start.y, line.dxf.start.z)
                    end_point = (line.dxf.end.x, line.dxf.end.y, line.dxf.end.z)
                    line_length = self.calculate_line_length(start_point, end_point)
                    # 將長度從mm轉換為cm
                    line_length_cm = line_length / 10.0
                    layer = line.dxf.layer if self.attributes["圖層"].get() else ""
                    
                    # 嘗試尋找鄰近的文字標註
                    number = None
                    diameter = None
                    count = 1
                    remark = ""
                    
                    # 查找最接近線條端點的文字
                    min_distance = float('inf')
                    nearest_text = None
                    
                    for text_pos, text_content in text_entities.items():
                        # 計算文字到線條兩端的距離
                        d1 = math.sqrt((text_pos[0] - start_point[0])**2 + (text_pos[1] - start_point[1])**2)
                        d2 = math.sqrt((text_pos[0] - end_point[0])**2 + (text_pos[1] - end_point[1])**2)
                        min_d = min(d1, d2)
                        
                        # 如果距離小於當前最小距離且在閾值內，更新最近文字
                        if min_d < min_distance and min_d < 50:  # 閾值可調整
                            min_distance = min_d
                            nearest_text = text_content
                    
                    # 如果找到鄰近文字，提取鋼筋信息
                    if nearest_text:
                        number, diameter, count, length = self.extract_rebar_info(nearest_text)
                        remark = nearest_text
                    
                    # 如果沒有找到鋼筋號數，給予默認編號
                    if not number:
                        number = f"未標註-{unnamed_counter}"
                        unnamed_counter += 1
                        count = 1
                    
                    # 計算重量
                    unit_weight = self.get_rebar_unit_weight(number) if number.startswith("#") else 0
                    weight = self.calculate_rebar_weight(number, line_length_cm, count) if number.startswith("#") else 0
                    
                    # 新增到資料表
                    rebar_data.append({
                        "編號": number,
                        "直徑": diameter,
                        "長度(cm)": round(line_length_cm, 2),
                        "數量": count,
                        "單位重(kg/m)": unit_weight,
                        "總重量(kg)": weight,
                        "圖層": layer,
                        "備註": remark
                    })
                    
                    # 更新統計數據
                    rebar_stats["總長度"] += line_length_cm * count
                    rebar_stats["總重量"] += weight
                    
                    processed_count += 1
                    if processed_count % 100 == 0:
                        progress = 20 + int(60 * processed_count / entity_count)
                        self.progress["value"] = min(80, progress)
                        self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            # 處理多段線
            for polyline_type in ["POLYLINE", "LWPOLYLINE"]:
                if polyline_type in selected_entities:
                    for polyline in msp.query(polyline_type):
                        try:
                            # 取得多段線點集
                            if polyline_type == "POLYLINE":
                                vertices = list(polyline.vertices())
                                points = [(v.dxf.location.x, v.dxf.location.y, v.dxf.location.z) for v in vertices]
                            else:  # LWPOLYLINE
                                points = list(polyline.points())
                                points = [(p[0], p[1], 0) if len(p) >= 2 else (p[0], 0, 0) for p in points]
                            
                            # 計算長度
                            try:
                                if hasattr(polyline, 'length') and callable(getattr(polyline, 'length')):
                                    polyline_length = polyline.length()
                                else:
                                    polyline_length = self.calculate_polyline_length(points)
                            except:
                                polyline_length = 0
                                self.log_message(f"警告: 無法計算多段線長度")
                            
                            # 將長度從mm轉換為cm
                            polyline_length_cm = polyline_length / 10.0
                            layer = polyline.dxf.layer if self.attributes["圖層"].get() else ""
                            
                            # 嘗試尋找鄰近的文字標註
                            number = None
                            diameter = None
                            count = 1
                            remark = ""
                            
                            # 如果有點集，以第一個點作為參考
                            if points:
                                # 查找最接近多段線起點的文字
                                min_distance = float('inf')
                                nearest_text = None
                                
                                for text_pos, text_content in text_entities.items():
                                    # 計算文字到多段線起點的距離
                                    d = math.sqrt((text_pos[0] - points[0][0])**2 + (text_pos[1] - points[0][1])**2)
                                    
                                    # 如果距離小於當前最小距離且在閾值內，更新最近文字
                                    if d < min_distance and d < 50:  # 閾值可調整
                                        min_distance = d
                                        nearest_text = text_content
                                
                                # 如果找到鄰近文字，提取鋼筋信息
                                if nearest_text:
                                    number, diameter, count, length = self.extract_rebar_info(nearest_text)
                                    remark = nearest_text
                            
                            # 如果沒有找到鋼筋號數，給予默認編號
                            if not number:
                                number = f"未標註-{unnamed_counter}"
                                unnamed_counter += 1
                                count = 1
                            
                            # 計算重量
                            unit_weight = self.get_rebar_unit_weight(number) if number.startswith("#") else 0
                            weight = self.calculate_rebar_weight(number, polyline_length_cm, count) if number.startswith("#") else 0
                            
                            # 新增到資料表
                            rebar_data.append({
                                "編號": number,
                                "直徑": diameter,
                                "長度(cm)": round(polyline_length_cm, 2),
                                "數量": count,
                                "單位重(kg/m)": unit_weight,
                                "總重量(kg)": weight,
                                "圖層": layer,
                                "備註": remark
                            })
                            
                            # 更新統計數據
                            rebar_stats["總長度"] += polyline_length_cm * count
                            rebar_stats["總重量"] += weight
                            
                        except Exception as e:
                            self.log_message(f"警告: 處理多段線實體時發生問題: {str(e)}")
                        
                        processed_count += 1
                        if processed_count % 100 == 0:
                            progress = 20 + int(60 * processed_count / entity_count)
                            self.progress["value"] = min(80, progress)
                            self.log_message(f"已處理 {processed_count}/{entity_count} 個實體")
            
            self.progress["value"] = 80
            self.log_message(f"已處理完成 {processed_count} 個實體")
            
            # 根據號數排序
            sorted_data = sorted(rebar_data, key=lambda x: x["編號"] if "#" in x["編號"] else "z" + x["編號"])
            
            # 檢查數據是否為空
            if not sorted_data:
                self.log_message("警告: 沒有找到任何可轉換的鋼筋數據")
                messagebox.showwarning("警告", "沒有找到任何可轉換的鋼筋數據")
                self.progress["value"] = 0
                return
            
            # 創建 DataFrame
            df = pd.DataFrame(sorted_data)
            
            # 重新排列欄位
            columns = ["編號", "直徑", "長度(cm)", "數量", "單位重(kg/m)", "總重量(kg)", "圖層", "備註"]
            df = df.reindex(columns=columns)
            
            # 寫入 Excel
            self.log_message(f"正在寫入資料到 Excel: {self.excel_path.get()}")
            
            try:
                # 創建 Excel 工作簿
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "鋼筋計料"
                
                # 設定標題
                project_name = os.path.basename(self.cad_path.get())
                ws['A1'] = f"鋼筋計料表"
                ws['A2'] = f"專案名稱: {project_name}"
                
                # 合併儲存格
                ws.merge_cells('A1:H1')
                ws.merge_cells('A2:H2')
                
                # 設定標題樣式
                title_font = Font(bold=True, size=16)
                subtitle_font = Font(bold=True, size=12)
                ws['A1'].font = title_font
                ws['A2'].font = subtitle_font
                
                # 居中對齊
                title_align = Alignment(horizontal='center', vertical='center')
                ws['A1'].alignment = title_align
                ws['A2'].alignment = title_align
                
                # 設定表頭 (從第3行開始)
                headers = ["編號", "直徑", "長度(cm)", "數量", "單位重(kg/m)", "總重量(kg)", "圖層", "備註"]
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=3, column=col_num)
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
                
                # 寫入資料 (從第4行開始)
                row_num = 4
                for _, row in df.iterrows():
                    for col_num, col_name in enumerate(headers, 1):
                        cell = ws.cell(row=row_num, column=col_num)
                        
                        # 獲取對應的值
                        if col_name in row:
                            cell.value = row[col_name]
                        
                        # 設定資料樣式
                        data_align = Alignment(horizontal='center', vertical='center')
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
                    
                    row_num += 1
                
                # 添加統計行
                summary_row = row_num
                ws.cell(row=summary_row, column=1).value = "總計"
                ws.cell(row=summary_row, column=1).font = Font(bold=True)
                
                # 總數量
                ws.cell(row=summary_row, column=4).value = df["數量"].sum()
                ws.cell(row=summary_row, column=4).font = Font(bold=True)
                
                # 總重量
                ws.cell(row=summary_row, column=6).value = round(df["總重量(kg)"].sum(), 2)
                ws.cell(row=summary_row, column=6).font = Font(bold=True)
                
                # 為統計行設定樣式
                for col in range(1, 9):
                    cell = ws.cell(row=summary_row, column=col)
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='double')
                    )
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # 調整列寬
                column_widths = {
                    1: 12,  # 編號
                    2: 10,  # 直徑
                    3: 15,  # 長度
                    4: 10,  # 數量
                    5: 15,  # 單位重
                    6: 15,  # 總重量
                    7: 15,  # 圖層
                    8: 30,  # 備註
                }
                
                for col_num, width in column_widths.items():
                    column_letter = openpyxl.utils.get_column_letter(col_num)
                    ws.column_dimensions[column_letter].width = width
                
                # 儲存檔案
                wb.save(self.excel_path.get())
                
                self.log_message("Excel 格式套用完成")
                
            except Exception as e:
                self.log_message(f"警告: Excel 格式化時出現問題: {str(e)}")
                self.log_message("嘗試使用基本儲存方式...")
                # 如果格式化失敗，至少確保數據被保存
                df.to_excel(self.excel_path.get(), sheet_name='鋼筋計料', index=False)
            
            self.progress["value"] = 100
            self.log_message("轉換完成!")
            messagebox.showinfo("成功", f"鋼筋計料已成功轉換為Excel檔案:\n{self.excel_path.get()}")
            
        except Exception as e:
            self.log_message(f"錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"轉換過程中發生錯誤:\n{str(e)}")
            self.progress["value"] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = CADtoExcelConverter(root)
    root.mainloop()