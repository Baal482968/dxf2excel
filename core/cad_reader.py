"""
CAD 檔案讀取相關功能模組
"""

import ezdxf
from utils.helpers import calculate_line_length, calculate_polyline_length
from core.rebar_processor import RebarProcessor
from utils.image_ocr import create_image_ocr_processor
from ezdxf.math import BoundingBox
from itertools import chain
import logging

class CADReader:
    """CAD 檔案讀取器"""
    
    def __init__(self):
        self.dxf_file = None
        self.modelspace = None
        self.rebar_processor = RebarProcessor()
        self.image_ocr_processor = create_image_ocr_processor()
    
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
    
    def extract_image_numbers(self):
        """提取圖面中圖片實體的數字"""
        if not self.modelspace:
            return []
        
        try:
            print("[DEBUG] 開始提取圖片中的數字...")
            image_numbers = self.image_ocr_processor.process_dxf_images(self.modelspace)
            print(f"[DEBUG] 從圖片中提取到 {len(image_numbers)} 個數字")
            return image_numbers
        except Exception as e:
            print(f"提取圖片數字錯誤: {str(e)}")
            return []
    
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
        line_count = 0
        polyline_count = 0
        
        try:
            print("[DEBUG] 開始提取線段實體...")
            # 遍歷所有線段實體
            for line in self.modelspace.query('LINE'):
                line_count += 1
                length = calculate_line_length(
                    (line.dxf.start.x, line.dxf.start.y),
                    (line.dxf.end.x, line.dxf.end.y)
                )
                print(f"[DEBUG][LINE] 線段 {line_count}: 起點({line.dxf.start.x:.2f}, {line.dxf.start.y:.2f}) -> 終點({line.dxf.end.x:.2f}, {line.dxf.end.y:.2f}) 長度={length:.2f}")
                rebar_lines.append({
                    'type': 'LINE',
                    'start': (line.dxf.start.x, line.dxf.start.y),
                    'end': (line.dxf.end.x, line.dxf.end.y),
                    'length': length
                })
            
            print(f"[DEBUG] 找到 {line_count} 個線段實體")
            
            print("[DEBUG] 開始提取多段線實體...")
            # 遍歷所有多段線實體
            for polyline in self.modelspace.query('LWPOLYLINE'):
                polyline_count += 1
                points = [(point[0], point[1]) for point in polyline.get_points()]
                length = calculate_polyline_length(points)
                print(f"[DEBUG][POLYLINE] 多段線 {polyline_count}: 點數={len(points)}, 長度={length:.2f}")
                print(f"[DEBUG][POLYLINE] 點座標: {points}")
                rebar_lines.append({
                    'type': 'POLYLINE',
                    'points': points,
                    'length': length
                })
            
            print(f"[DEBUG] 找到 {polyline_count} 個多段線實體")
            print(f"[DEBUG] 總共找到 {len(rebar_lines)} 個線段元件")
        
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

    def group_entities(self, tolerance=30.0):
        """
        將 DXF 圖面中的實體（線、多段線、文字）根據距離分組。
        使用 DSU (Disjoint Set Union) 演算法。
        """
        if not self.modelspace:
            return []

        entities = list(self.modelspace.query('LINE LWPOLYLINE TEXT MTEXT'))
        if not entities:
            return []

        # 1. 計算每個實體的邊界框
        entity_bboxes = {}
        for i, e in enumerate(entities):
            try:
                if e.dxftype() in ('LINE', 'LWPOLYLINE'):
                    points = list(e.vertices())
                    if points:
                        entity_bboxes[i] = BoundingBox(points)
                elif e.dxftype() in ('TEXT', 'MTEXT'):
                    p1 = e.dxf.insert
                    h = e.dxf.height
                    text = e.text if e.dxftype() == 'MTEXT' else e.dxf.text
                    w = len(text) * h * 0.6  # 估算文字寬度
                    p2 = p1 + (w, h)
                    entity_bboxes[i] = BoundingBox([p1, p2])
            except (AttributeError, ValueError):
                continue  # 忽略無法計算邊界框的實體

        # 2. DSU 初始化
        parent = list(range(len(entities)))
        def find(i):
            if parent[i] == i:
                return i
            parent[i] = find(parent[i])
            return parent[i]

        def union(i, j):
            root_i = find(i)
            root_j = find(j)
            if root_i != root_j:
                parent[root_j] = root_i

        # 3. 根據邊界框重疊進行分組
        entity_indices = list(entity_bboxes.keys())
        for i in range(len(entity_indices)):
            for j in range(i + 1, len(entity_indices)):
                idx1 = entity_indices[i]
                idx2 = entity_indices[j]
                
                bbox1 = entity_bboxes[idx1].copy()
                bbox1.grow(tolerance)  # 擴大邊界框以尋找鄰近實體

                if bbox1.has_overlap(entity_bboxes[idx2]):
                    union(idx1, idx2)
        
        # 4. 提取分組結果
        groups = {}
        for i in entity_indices:
            root = find(i)
            if root not in groups:
                groups[root] = []
            groups[root].append(entities[i])
        
        return list(groups.values())

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
            print("[DEBUG] ===== 開始處理圖面 =====")
            
            # 使用新的分組邏輯
            print("[DEBUG] 1. 根據實體位置進行分組...")
            entity_groups = self.group_entities()
            print(f"[DEBUG] 找到 {len(entity_groups)} 個實體群組")

            drawings = []
            for i, group in enumerate(entity_groups):
                rebar_number = ''
                count = 1
                found_count_text = False
                other_texts = []
                geometric_entities = []

                # 1. 將群組內的實體分類，並解析文字內容
                for e in group:
                    if e.dxftype() in ('TEXT', 'MTEXT'):
                        text = (e.text if e.dxftype() == 'MTEXT' else e.dxf.text).strip()
                        if not text:
                            continue
                        
                        # 檢查是否為鋼筋號數 (e.g., #4)
                        if text.startswith('#') and text[1:].isdigit():
                            rebar_number = text
                        # 檢查是否為數量 (e.g., x10)
                        elif text.lower().startswith('x') and text[1:].isdigit():
                            count = int(text[1:])
                            found_count_text = True
                        else:
                            other_texts.append(text)
                    elif e.dxftype() in ('LINE', 'LWPOLYLINE'):
                        geometric_entities.append(e)

                # 如果群組中沒有任何線條，可能只是游離的文字，跳過此群組
                if not geometric_entities:
                    continue

                # 組合原始文字和備註
                # 如果沒有解析到鋼筋號數，則使用預設名稱
                display_rebar_number = rebar_number if rebar_number else f"圖形 {i+1}"
                
                note = ' '.join(other_texts)
                
                raw_text_parts = [rebar_number]
                if found_count_text:
                    raw_text_parts.append(f'x{count}')
                raw_text_parts.append(note)
                
                raw_text = ' '.join(filter(None, raw_text_parts))

                # 計算總長度
                total_length = 0
                for e in geometric_entities:
                    if e.dxftype() == 'LINE':
                        total_length += e.dxf.start.distance(e.dxf.end)
                    elif e.dxftype() == 'LWPOLYLINE':
                        points = [(p[0], p[1]) for p in e.get_points()]
                        total_length += calculate_polyline_length(points)

                drawings.append({
                    'rebar_number': display_rebar_number,
                    'count': count,
                    'note': note,
                    'length': round(total_length) if total_length > 0 else 0,
                    'weight': '',
                    'raw_text': raw_text.strip(),
                    'entities_to_draw': group, # 將整個群組傳遞給繪圖器
                })

            print("[DEBUG] 2. 提取表格框線...")
            tables = self.get_rebar_tables()
            print(f"[DEBUG] 找到 {len(tables)} 個表格框線")
            
            # 將所有圖形放入分組
            grouped = {tb['name']: [] for tb in tables}
            if not tables:
                print("[DEBUG] 沒有找到表格框線，使用預設分組 '全部'")
                grouped = {'全部': drawings}
            else:
                # 這裡可以實作更複雜的邏輯，判斷每個群組屬於哪個表格
                # 目前為簡化，全放入第一個表格
                grouped[tables[0]['name']] = drawings
            
            print("[DEBUG] ===== 處理完成 =====")
            for group_name, items in grouped.items():
                print(f"[DEBUG] 分組 '{group_name}': {len(items)} 個圖形")
            
            return grouped
            
        except Exception as e:
            logging.error(f"處理圖面錯誤: {e}", exc_info=True)
            return None 