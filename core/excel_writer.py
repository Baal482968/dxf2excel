"""
Excel 輸出相關功能模組
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
from datetime import datetime
import tempfile
import base64
import os
from utils.graphics import GraphicsManager

class ExcelWriter:
    """Excel 檔案寫入器"""
    
    def __init__(self):
        self.workbook = None
        self.worksheet = None
        self.temp_image_files = []  # 新增暫存圖檔清單
        
        # 定義樣式
        self.styles = {
            'header_font': Font(name='Calibri', size=12, bold=True, color='FFFFFF'),
            'normal_font': Font(name='Calibri', size=11),
            'title_font': Font(name='Calibri', size=14, bold=True),
            'header_fill': PatternFill(start_color='4A90E2', end_color='4A90E2', fill_type='solid'),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
    
    def create_workbook(self):
        """創建新的工作簿"""
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "鋼筋計料表"
    
    def save_workbook(self, file_path):
        """儲存工作簿，並在儲存後刪除暫存圖檔"""
        if self.workbook:
            self.workbook.save(file_path)
        # 儲存後刪除暫存圖檔
        for f in getattr(self, 'temp_image_files', []):
            try:
                os.remove(f)
            except Exception:
                pass
        self.temp_image_files = []
    
    def write_header(self):
        """寫入表頭（第 2 列，含圖示欄與讀取CAD文字）"""
        headers = [
            "編號", "號數", "圖示", "長度(cm)", "數量", "重量(kg)", "備註", "讀取CAD文字"
        ]
        for col, header in enumerate(headers, 1):
            cell = self.worksheet.cell(row=2, column=col)
            cell.value = header
            cell.font = self.styles['header_font']
            cell.fill = self.styles['header_fill']
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = self.styles['border']
        for col in range(1, len(headers) + 1):
            self.worksheet.column_dimensions[get_column_letter(col)].width = 18 if col == 3 else 12
    
    def write_title(self, title):
        """寫入標題"""
        # 合併儲存格
        self.worksheet.merge_cells('A1:H1')
        # 只對左上角寫入 value
        cell = self.worksheet.cell(row=1, column=1)
        cell.value = title
        cell.font = self.styles['title_font']
        cell.alignment = Alignment(horizontal='center', vertical='center')
        self.worksheet.row_dimensions[1].height = 30
    
    def write_rebar_data(self, rebar_data, start_row=3):
        """寫入鋼筋資料（含圖示與CAD原文）"""
        current_row = start_row
        for idx, rebar in enumerate(rebar_data, 1):
            self.worksheet.cell(row=current_row, column=1).value = idx
            self.worksheet.cell(row=current_row, column=2).value = rebar.get('rebar_number', '')
            img_b64 = GraphicsManager.draw_straight_rebar(rebar.get('length', 0), rebar.get('rebar_number', ''))
            img_data = base64.b64decode(img_b64)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                tmp.write(img_data)
                tmp_path = tmp.name
                self.temp_image_files.append(tmp_path)
            img = ExcelImage(tmp_path)
            img.width = 120
            img.height = 32
            img.anchor = f'C{current_row}'
            self.worksheet.add_image(img)
            self.worksheet.cell(row=current_row, column=4).value = rebar.get('length', 0)
            self.worksheet.cell(row=current_row, column=5).value = rebar.get('count', 1)
            self.worksheet.cell(row=current_row, column=6).value = round(rebar.get('weight', 0))
            self.worksheet.cell(row=current_row, column=7).value = rebar.get('note', '')
            self.worksheet.cell(row=current_row, column=8).value = rebar.get('raw_text', '')
            for col in range(1, 9):
                cell = self.worksheet.cell(row=current_row, column=col)
                cell.font = self.styles['normal_font']
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = self.styles['border']
            current_row += 1
        return current_row
    
    def write_footer(self, row):
        """寫入頁尾"""
        # 合併儲存格，只對左上角寫入 value
        self.worksheet.merge_cells(f'A{row}:H{row}')
        cell = self.worksheet.cell(row=row, column=1)
        cell.value = f"生成時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        cell.font = self.styles['normal_font']
        cell.alignment = Alignment(horizontal='right', vertical='center')
    
    def format_worksheet(self):
        """格式化工作表"""
        # 設定列印區域
        self.worksheet.print_area = self.worksheet.dimensions
        
        # 設定頁面方向為橫向
        self.worksheet.page_setup.orientation = self.worksheet.ORIENTATION_LANDSCAPE
        
        # 設定頁面邊距
        self.worksheet.page_margins.left = 0.5
        self.worksheet.page_margins.right = 0.5
        self.worksheet.page_margins.top = 0.5
        self.worksheet.page_margins.bottom = 0.5
        
        # 設定頁首頁尾
        self.worksheet.oddHeader.center.text = "鋼筋計料表"
        self.worksheet.oddFooter.right.text = "第 &P 頁，共 &N 頁" 