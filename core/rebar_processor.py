"""
鋼筋處理相關功能模組
"""

import re
from config import REBAR_UNIT_WEIGHT, REBAR_DIAMETERS, REBAR_GRADES
# 圖形相關模組已移除，改為使用 assets/materials/ 資料夾中的圖示檔案

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
        解析鋼筋文字格式
        
        支援格式：
        - #3-700x99 (type10 單段直料)
        - 安#3-390x40 (type11 安全彎鉤直)
        """
        import re
        text = text.strip()
        
        # 處理 type10 直料鋼筋格式
        # 格式: #3-700x99 (單段直料)
        type10_pattern = r'(#\d+)-([\d\.]+)x(\d+)'
        type10_match = re.match(type10_pattern, text)

        if type10_match:
            rebar_number = type10_match.group(1)
            length = float(type10_match.group(2))
            count = int(type10_match.group(3))
            
            # 計算重量
            unit_weight = RebarProcessor.get_rebar_unit_weight(rebar_number)
            weight = unit_weight * length * count / 100  # 轉換為 kg

            return {
                'rebar_number': rebar_number,
                'segments': [length],
                'angles': [],
                'count': count,
                'raw_text': text,
                'length': length,
                'weight': weight,
                'type': 'type10',
                'note': '直料'
            }
        
        # 處理 type11 安全彎鉤直鋼筋格式
        # 格式: 安#3-390x40 (安全彎鉤直)
        type11_pattern = r'安(#\d+)-([\d\.]+)x(\d+)'
        type11_match = re.match(type11_pattern, text)

        if type11_match:
            rebar_number = type11_match.group(1)
            length = float(type11_match.group(2))
            count = int(type11_match.group(3))
            
            # 計算重量
            unit_weight = RebarProcessor.get_rebar_unit_weight(rebar_number)
            weight = unit_weight * length * count / 100  # 轉換為 kg

            return {
                'rebar_number': rebar_number,
                'segments': [length],
                'angles': [],
                'count': count,
                'raw_text': text,
                'length': length,
                'weight': weight,
                'type': 'type11',
                'note': '安全彎鉤直'
            }
        
        # 處理 type12 折料鋼筋格式
        # 格式: V113°#10-900+200x2 (折料)
        print(f"🔍 type12 文字: {text}")
        type12_pattern = r'V(\d+)°(#\d+)-([\d\.]+)\+([\d\.]+)x(\d+)'
        type12_match = re.match(type12_pattern, text)
        print(f"🔍 type12 正則匹配結果: {type12_match}")

        if type12_match:
            angle = int(type12_match.group(1))
            rebar_number = type12_match.group(2)
            length1 = float(type12_match.group(3))
            length2 = float(type12_match.group(4))
            count = int(type12_match.group(5))
            
            # 計算總長度
            total_length = length1 + length2
            
            # 計算重量
            unit_weight = RebarProcessor.get_rebar_unit_weight(rebar_number)
            weight = unit_weight * total_length * count / 100  # 轉換為 kg

            return {
                'rebar_number': rebar_number,
                'segments': [length1, length2],
                'angles': [angle],
                'count': count,
                'raw_text': text,
                'length': total_length,
                'weight': weight,
                'type': 'type12',
                'note': f'折料 {angle}°'
            }
        
        # 無法解析的格式
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