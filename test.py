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
        count = 1
        length = None
        segments = []  # 新增：儲存分段長度
        
        # 尋找鋼筋號數 (支援多種格式)
        # 格式1: #3, #4 等
        number_match = re.search(r'#(\d+)', text)
        if number_match:
            number = "#" + number_match.group(1)
        
        # 格式2: D13, D16 等
        if not number:
            number_match = re.search(r'D(\d+)', text)
            if number_match:
                diameter = float(number_match.group(1))
                # 根據直徑反推號數
                for num, dia in self.rebar_unit_weight.items():
                    if abs(dia - diameter) < 0.1:
                        number = num
                        break
        
        # 格式3: 直接寫直徑，如 13mm, 16mm 等
        if not number:
            number_match = re.search(r'(\d+(?:\.\d+)?)\s*mm', text)
            if number_match:
                diameter = float(number_match.group(1))
                # 根據直徑反推號數
                for num, dia in self.rebar_unit_weight.items():
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
                segments = [float(part) for part in length_parts]  # 儲存分段長度
                total_length = sum(segments)
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
        
        return number, count, length, segments
    
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
    
    def draw_ascii_rebar(self, segments):
        """使用 ASCII 字元繪製彎折示意圖"""
        if not segments:
            return "─"
            
        # 如果只有一段，直接畫直線
        if len(segments) == 1:
            length = str(int(segments[0]))
            # 固定橫線長度為 10 個字元
            line = "─" * 10
            # 計算需要的空格數來置中數字
            total_width = max(len(line), len(length))
            length_spaces = (total_width - len(length)) // 2
            line_spaces = (total_width - len(line)) // 2
            return f"{' ' * length_spaces}{length}\n{' ' * line_spaces}{line}"
        
        # 如果有多段，畫彎折圖
        lines = []  # 用於存儲多行文字
        
        # 固定中間段的長度
        middle_chars = 10  # 固定為 10 個字元
        
        # 構建第二行（線條）
        line = ""
        # 起始段
        start_num = str(int(segments[0]))
        line += f"{start_num} |"
        
        # 中間段
        if len(segments) > 1:
            line += "─" * middle_chars
        
        # 結束段
        if len(segments) > 2:
            end_num = str(int(segments[-1]))
            line += f"| {end_num}"
        elif len(segments) == 2:
            end_num = str(int(segments[-1]))
            line += f"| {end_num}"
        
        # 第一行：中間段的長度（置中對齊）
        if len(segments) > 1:
            middle_num = str(int(sum(segments[1:-1] if len(segments) > 2 else segments[1:])))
            # 計算整個圖形的寬度
            total_width = len(line)
            # 計算中間數字的起始位置
            # 中間數字應該在中間段的中心位置
            start_pos = len(start_num) + 2  # 起始數字加上 " |" 的長度
            middle_section_width = middle_chars  # 中間段的寬度
            middle_start = start_pos + (middle_section_width - len(middle_num)) // 2
            
            # 創建第一行，先填充空格
            first_line = " " * total_width
            # 在正確位置插入中間數字
            first_line = first_line[:middle_start] + middle_num + first_line[middle_start + len(middle_num):]
            lines.append(first_line)
        
        lines.append(line)
        
        return "\n".join(lines)
    
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
            
            # 使用固定的實體類型
            selected_entities = ["TEXT"]
            self.log_message(f"處理實體類型: {', '.join(selected_entities)}")
            
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
            for text in msp.query("TEXT"):
                text_content = text.dxf.text
                if text_content:
                    # 儲存文本位置和內容，稍後用於匹配線條
                    position = (text.dxf.insert.x, text.dxf.insert.y)
                    text_entities[position] = text_content
                    
                    # 直接從文字提取鋼筋資訊
                    number, count, length, segments = self.extract_rebar_info(text_content)
                    if number:
                        rebar_stats["各類型數量"][number] = rebar_stats["各類型數量"].get(number, 0) + count
                        rebar_stats["總數量"] += count
                        
                        # 如果有長度資訊，直接加入資料表
                        if length is not None:
                            # 將長度從mm轉換為cm（移除除以10的轉換）
                            length_cm = length
                            # 計算重量
                            unit_weight = self.get_rebar_unit_weight(number) if number.startswith("#") else 0
                            weight = self.calculate_rebar_weight(number, length_cm, count) if number.startswith("#") else 0
                            
                            # 新增到資料表
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
                                for i, segment in enumerate(segments):
                                    letter = chr(65 + i)  # A, B, C, ...
                                    data[f"{letter}(cm)"] = round(segment, 2)
                            
                            rebar_data.append(data)
                            
                            # 更新統計數據
                            rebar_stats["總長度"] += length_cm * count
                            rebar_stats["總重量"] += weight
                            
                            # 記錄詳細資訊
                            self.log_message(f"找到鋼筋標記: {text_content}")
                            self.log_message(f"  號數: {number}")
                            self.log_message(f"  長度: {length_cm}cm")
                            self.log_message(f"  數量: {count}")
                            self.log_message(f"  重量: {weight}kg")
            
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
            for index, row in df.iterrows():
                segments = []
                for key in sorted([k for k in row.keys() if k.endswith("(cm)") and k != "長度(cm)"]):
                    if not pd.isna(row[key]):
                        segments.append(row[key])
                df.at[index, "圖示"] = self.draw_ascii_rebar(segments)
            
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
                headers = columns
                
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
                summary_row = row_num
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