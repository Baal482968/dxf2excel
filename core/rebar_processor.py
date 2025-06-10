"""
鋼筋處理相關功能模組
"""

import re
from config import REBAR_UNIT_WEIGHT, REBAR_DIAMETERS, REBAR_GRADES

class RebarProcessor:
    """鋼筋處理器"""
    
    @staticmethod
    def extract_rebar_info(text):
        """
        從文字中提取鋼筋資訊
        
        格式範例：
        - #4@20
        - #4@20c/c
        - #4@20cm
        - #4@200mm
        """
        # 移除空白字元
        text = text.strip()
        
        # 基本格式檢查
        if not text or not text.startswith('#'):
            return None
        
        try:
            # 提取鋼筋編號
            rebar_number = re.match(r'#\d+', text).group()
            
            # 提取間距
            spacing_match = re.search(r'@(\d+)(?:c/c|cm|mm)?', text)
            if spacing_match:
                spacing = int(spacing_match.group(1))
                # 如果沒有單位，預設為公分
                if not re.search(r'@\d+(?:c/c|cm|mm)', text):
                    spacing *= 10  # 轉換為公釐
            else:
                spacing = None
            
            return {
                'rebar_number': rebar_number,
                'spacing': spacing,
                'diameter': RebarProcessor.get_rebar_diameter(rebar_number),
                'unit_weight': RebarProcessor.get_rebar_unit_weight(rebar_number),
                'grade': RebarProcessor.get_rebar_grade(rebar_number)
            }
        except Exception:
            return None

    @staticmethod
    def get_rebar_diameter(number):
        """獲取鋼筋直徑（公釐）"""
        return REBAR_DIAMETERS.get(number, 0)

    @staticmethod
    def get_rebar_unit_weight(number):
        """獲取鋼筋單位重量（kg/m）"""
        return REBAR_UNIT_WEIGHT.get(number, 0)

    @staticmethod
    def get_rebar_grade(number):
        """獲取鋼筋材質等級"""
        return REBAR_GRADES.get(number, "未知")

    @staticmethod
    def calculate_rebar_weight(number, length, count=1):
        """計算鋼筋重量（kg）"""
        unit_weight = RebarProcessor.get_rebar_unit_weight(number)
        return unit_weight * length * count

    @staticmethod
    def parse_rebar_text(text):
        """
        支援多種鋼筋標記格式的解析，允許 # 前有任意字元，支援全形乘號、逗號、cm、括號等符號，並自動解析角度。
        
        支援格式：
        1. #10-1000+1000+1000x31
        2. #10-21+550x6
        3. #10-1400x1
        4. 折140#10-1000+1200x20
        5. #10-510.5x11
        6. #10-550+21x5
        """
        import re
        text = text.strip()
        
        # Debug: 印出原始文字
        # print(f"[DEBUG] 解析鋼筋文字: {text}")
        
        # 解析角度（如「折135度」「135°」）
        angles = []
        angle_match = re.search(r'(?:折)?(\d{2,3})[°度]', text)
        if angle_match:
            try:
                angle = int(angle_match.group(1))
                angles.append(angle)
                # print(f"[DEBUG] 解析到角度: {angle}°")
            except:
                pass
        
        # 將全形乘號、逗號、空白等轉成半形
        text = text.replace('×', 'x').replace('，', ',').replace(' ', '')
        
        # 找到第一個 # 開頭
        m = re.search(r'#\d+', text)
        if not m:
            # print(f"[DEBUG] 未找到鋼筋編號")
            return None
            
        start = m.start()
        text = text[start:]  # 只取 # 開頭到結尾
        
        # 強化正則：#號-分段x數量
        # 支援：
        # 1. 整數：1000
        # 2. 小數：510.5
        # 3. 多段：1000+1000+1000
        pattern = r'#(\d+)-([\d\.]+(?:\+[\d\.]+)*)x(\d+)'
        m = re.match(pattern, text)
        
        if m:
            number = f"#{m.group(1)}"
            segments_str = m.group(2)
            count = int(m.group(3))
            
            # 解析分段
            segments = []
            for seg in segments_str.split('+'):
                try:
                    if '.' in seg:
                        segments.append(float(seg))
                    else:
                        segments.append(int(seg))
                except ValueError:
                    # print(f"[DEBUG] 無法解析分段: {seg}")
                    continue
            
            # 確保 segments 不為空
            if not segments:
                # print(f"[DEBUG] segments 為空，無法解析")
                return None
            
            result = {
                'rebar_number': number,
                'segments': segments,
                'angles': angles,
                'count': count,
                'raw_text': text,
                'length': sum(segments)  # 添加總長度
            }
            
            # print(f"[DEBUG] 解析結果: {result}")
            
            return result
            
        # print(f"[DEBUG] 無法解析鋼筋文字格式: {text}")
        return None

    @staticmethod
    def validate_rebar_number(number):
        """驗證鋼筋編號是否有效"""
        return number in REBAR_UNIT_WEIGHT

    @staticmethod
    def get_rebar_summary(rebar_list):
        """生成鋼筋統計摘要"""
        summary = {}
        
        for rebar in rebar_list:
            number = rebar['rebar_number']
            if number not in summary:
                summary[number] = {
                    'count': 0,
                    'total_length': 0,
                    'total_weight': 0,
                    'diameter': RebarProcessor.get_rebar_diameter(number),
                    'grade': RebarProcessor.get_rebar_grade(number)
                }
            
            summary[number]['count'] += 1
            summary[number]['total_length'] += rebar.get('length', 0)
            summary[number]['total_weight'] += rebar.get('weight', 0)
        
        return summary 