"""
CAD 檔案讀取相關功能模組
"""

import ezdxf
from utils.helpers import calculate_line_length, calculate_polyline_length
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
    
    def get_drawing_info(self):
        """獲取圖面基本資訊"""
        if not self.dxf_file:
            return None
        
        try:
            return {
                'filename': self.dxf_file.filename,
                'dxfversion': self.dxf_file.dxfversion,
                'encoding': self.dxf_file.encoding,
                'header': dict(self.dxf_file.header),
                'layers': [layer.dxf.name for layer in self.dxf_file.layers],
                'blocks': [block.dxf.name for block in self.dxf_file.blocks],
                'entities_count': len(self.modelspace)
            }
        except Exception as e:
            print(f"獲取圖面資訊錯誤: {str(e)}")
            return None
    
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
    
    def extract_rebar_lines(self):
        """提取圖面中的鋼筋線段"""
        if not self.modelspace:
            return []
        
        rebar_lines = []
        
        try:
            # 遍歷所有線段實體
            for line in self.modelspace.query('LINE'):
                length = calculate_line_length(
                    (line.dxf.start.x, line.dxf.start.y),
                    (line.dxf.end.x, line.dxf.end.y)
                )
                rebar_lines.append({
                    'type': 'LINE',
                    'start': (line.dxf.start.x, line.dxf.start.y),
                    'end': (line.dxf.end.x, line.dxf.end.y),
                    'length': length
                })
            
            # 遍歷所有多段線實體
            for polyline in self.modelspace.query('LWPOLYLINE'):
                points = [(point[0], point[1]) for point in polyline.get_points()]
                length = calculate_polyline_length(points)
                rebar_lines.append({
                    'type': 'POLYLINE',
                    'points': points,
                    'length': length
                })
        
        except Exception as e:
            print(f"提取鋼筋線段錯誤: {str(e)}")
        
        return rebar_lines
    
    def find_associated_lines(self, rebar_text, rebar_lines, max_distance=100):
        """尋找與鋼筋文字相關的線段"""
        if not rebar_text or not rebar_lines:
            return []
        
        associated_lines = []
        text_pos = rebar_text['position']
        
        for line in rebar_lines:
            if line['type'] == 'LINE':
                # 計算線段中點到文字的距離
                mid_x = (line['start'][0] + line['end'][0]) / 2
                mid_y = (line['start'][1] + line['end'][1]) / 2
                distance = calculate_line_length(
                    (mid_x, mid_y),
                    (text_pos[0], text_pos[1])
                )
                
                if distance <= max_distance:
                    line['rebar_info'] = rebar_text
                    associated_lines.append(line)
            
            elif line['type'] == 'POLYLINE':
                # 計算多段線中點到文字的距離
                points = line['points']
                mid_x = sum(p[0] for p in points) / len(points)
                mid_y = sum(p[1] for p in points) / len(points)
                distance = calculate_line_length(
                    (mid_x, mid_y),
                    (text_pos[0], text_pos[1])
                )
                
                if distance <= max_distance:
                    line['rebar_info'] = rebar_text
                    associated_lines.append(line)
        
        return associated_lines
    
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
            rebar_lines = self.extract_rebar_lines()
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
                assigned = False
                if pos:
                    x, y = pos[0], pos[1]
                    for tb in tables:
                        if tb['points'] and self.point_in_polygon(x, y, tb['points']):
                            assigned = True
                            target_name = tb['name']
                            break
                    else:
                        target_name = tables[0]['name']  # 若沒命中，歸第一個
                else:
                    target_name = tables[0]['name']
                associated_lines = self.find_associated_lines(rebar_text, rebar_lines)
                segments = rebar_text.get('segments', [])
                if segments:
                    total_length = sum(segments)
                else:
                    total_length = sum(line['length'] for line in associated_lines)
                rebar_entry = dict(rebar_text)
                rebar_entry.update({
                    'diameter': self.rebar_processor.get_rebar_diameter(rebar_entry['rebar_number']),
                    'unit_weight': self.rebar_processor.get_rebar_unit_weight(rebar_entry['rebar_number']),
                    'grade': self.rebar_processor.get_rebar_grade(rebar_entry['rebar_number']),
                    'length': round(total_length),
                    'weight': round(self.rebar_processor.calculate_rebar_weight(
                        rebar_entry['rebar_number'],
                        total_length
                    )),
                    'position': rebar_text.get('position'),
                    'associated_lines': associated_lines,
                })
                grouped[target_name].append(rebar_entry)
            return grouped
        except Exception as e:
            print(f"處理圖面錯誤: {str(e)}")
            return None 