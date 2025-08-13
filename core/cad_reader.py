"""
CAD 檔案讀取相關功能模組
"""

import ezdxf
from core.rebar_processor import RebarProcessor

class CADReader:
    """CAD 檔案讀取器"""
    
    def __init__(self):
        self.dxf_file = None
        self.modelspace = None
        self.rebar_processor = RebarProcessor()
    
    def open_file(self, file_path):
        """開啟 DXF 檔案"""
        try:
            self.dxf_file = ezdxf.readfile(file_path)
            self.modelspace = self.dxf_file.modelspace()
            return True
        except Exception as e:
            print(f"開啟檔案錯誤: {str(e)}")
            return False
    
    def close_file(self):
        """關閉 DXF 檔案"""
        if self.dxf_file:
            self.dxf_file = None
            self.modelspace = None
    
    def extract_rebar_texts(self):
        """提取圖面中的鋼筋文字標記"""
        if not self.modelspace:
            return []
        
        rebar_texts = []
        
        try:
            # 遍歷所有文字實體
            for text in self.modelspace.query('TEXT'):
                print(f"[DEBUG][TEXT] {text.dxf.text}")
                rebar_info = self.rebar_processor.parse_rebar_text(text.dxf.text)
                if rebar_info:
                    rebar_info['position'] = text.dxf.insert
                    rebar_info['rotation'] = text.dxf.rotation
                    rebar_info['raw_text'] = text.dxf.text
                    rebar_texts.append(rebar_info)
            
            # 遍歷所有多行文字實體
            for mtext in self.modelspace.query('MTEXT'):
                text_content = mtext.text
                print(f"[DEBUG][MTEXT] {text_content}")
                # 分割多行文字
                for line in text_content.split('\n'):
                    rebar_info = self.rebar_processor.parse_rebar_text(line)
                    if rebar_info:
                        rebar_info['position'] = mtext.dxf.insert
                        rebar_info['rotation'] = mtext.dxf.rotation
                        rebar_info['raw_text'] = line
                        rebar_texts.append(rebar_info)
        
        except Exception as e:
            print(f"提取鋼筋文字錯誤: {str(e)}")
        
        return rebar_texts
    
    def get_rebar_tables(self):
        """取得所有 $P- 開頭的 LWPOLYLINE 框線及名稱與多邊形座標"""
        if not self.modelspace:
            return []
        tables = []
        for polyline in self.modelspace.query('LWPOLYLINE'):
            layer = polyline.dxf.layer if hasattr(polyline.dxf, 'layer') else ''
            if layer.startswith('$P-'):
                name = layer[3:] if len(layer) > 3 else layer
                points = [(point[0], point[1]) for point in polyline.get_points()]
                tables.append({'name': name, 'points': points})
        return tables

    @staticmethod
    def point_in_polygon(x, y, polygon):
        """判斷點 (x, y) 是否在多邊形 polygon 內 (射線法)"""
        num = len(polygon)
        j = num - 1
        inside = False
        for i in range(num):
            xi, yi = polygon[i]
            xj, yj = polygon[j]
            intersect = ((yi > y) != (yj > y)) and \
                        (x < (xj - xi) * (y - yi) / (yj - yi + 1e-12) + xi)
            if intersect:
                inside = not inside
            j = i
        return inside

    def process_drawing(self):
        """處理整個圖面，依據框線分組回傳 dict: {區塊名稱: [rebar list]}"""
        if not self.modelspace:
            return None
        try:
            rebar_texts = self.extract_rebar_texts()
            tables = self.get_rebar_tables()
            
            # 預設分組: {區塊名稱: [rebar list]}
            grouped = {tb['name']: [] for tb in tables}
            
            # 若沒框線，全部歸入 '全部'
            if not tables:
                grouped = {'全部': []}
                tables = [{'name': '全部', 'points': None}]
            
            # 處理每個鋼筋文字
            for rebar_text in rebar_texts:
                pos = rebar_text.get('position')
                target_name = tables[0]['name']  # 預設歸入第一個區塊
                
                if pos:
                    x, y = pos[0], pos[1]
                    for tb in tables:
                        if tb['points'] and self.point_in_polygon(x, y, tb['points']):
                            target_name = tb['name']
                            break
                
                # 建立鋼筋條目（不包含線條相關資訊）
                rebar_entry = dict(rebar_text)
                rebar_entry.update({
                    'diameter': self.rebar_processor.get_rebar_diameter(rebar_entry['rebar_number']),
                    'unit_weight': self.rebar_processor.get_rebar_unit_weight(rebar_entry['rebar_number']),
                    'grade': self.rebar_processor.get_rebar_grade(rebar_entry['rebar_number']),
                    'position': rebar_text.get('position'),
                })
                
                grouped[target_name].append(rebar_entry)
            
            return grouped
            
        except Exception as e:
            print(f"處理圖面錯誤: {str(e)}")
            return None 