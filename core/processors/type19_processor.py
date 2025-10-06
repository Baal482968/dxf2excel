
"""
Type19 直段+弧段鋼筋處理器
"""

from .base_processor import BaseRebarProcessor

class Type19Processor(BaseRebarProcessor):
    """Type19 直段+弧段鋼筋處理器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type19'
    
    def get_pattern(self):
        """獲取正則表達式模式"""
        return r'直弧(\d+)(#\d+)-([\d\.]+)\+([\d\.]+)x(\d+)'
    
    def parse_match(self, match, text):
        """解析匹配結果"""
        radius = int(match.group(1))  # 半徑
        rebar_number = match.group(2)  # 鋼筋號數
        straight_length = float(match.group(3))  # 直段長度
        arc_length = float(match.group(4))  # 弧段長度
        count = int(match.group(5))  # 數量
        
        # 計算總長度
        total_length = straight_length + arc_length
        
        # 計算重量
        weight = self.calculate_weight(rebar_number, total_length, count)
        
        return {
            'rebar_number': rebar_number,
            'segments': [straight_length, arc_length],
            'angles': [],
            'radius': radius,  # 半徑
            'count': count,
            'raw_text': text,
            'length': total_length,
            'weight': weight,
            'type': 'type19',
            'note': f'直段+弧段 R{radius}'
        }
