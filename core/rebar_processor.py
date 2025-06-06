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
        """解析鋼筋文字標記"""
        # 移除空白字元
        text = text.strip()
        
        # 檢查是否為鋼筋標記
        if not text.startswith('#'):
            return None
        
        # 提取鋼筋編號和間距
        match = re.match(r'#(\d+)@(\d+)(?:c/c|cm|mm)?', text)
        if not match:
            return None
        
        rebar_number = f"#{match.group(1)}"
        spacing = int(match.group(2))
        
        # 如果沒有單位，預設為公分
        if not re.search(r'@\d+(?:c/c|cm|mm)', text):
            spacing *= 10  # 轉換為公釐
        
        return {
            'rebar_number': rebar_number,
            'spacing': spacing,
            'diameter': RebarProcessor.get_rebar_diameter(rebar_number),
            'unit_weight': RebarProcessor.get_rebar_unit_weight(rebar_number),
            'grade': RebarProcessor.get_rebar_grade(rebar_number)
        }

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