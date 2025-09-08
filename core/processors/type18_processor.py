"""
Type18 直料圓弧鋼筋處理器
"""

from .base_processor import BaseRebarProcessor

class Type18Processor(BaseRebarProcessor):
    """Type18 直料圓弧鋼筋處理器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type18'
    
    def get_pattern(self):
        """獲取正則表達式模式"""
        return r'弧(\d+)(#\d+)-([\d\.]+)x(\d+)'
    
    def parse_match(self, match, text):
        """解析匹配結果"""
        radius = int(match.group(1))  # 半徑
        rebar_number = match.group(2)  # 鋼筋號數
        length = float(match.group(3))  # 長度
        count = int(match.group(4))  # 數量
        
        # 計算重量
        weight = self.calculate_weight(rebar_number, length, count)
        
        return {
            'rebar_number': rebar_number,
            'segments': [length],
            'angles': [],
            'radius': radius,  # 新增半徑欄位
            'count': count,
            'raw_text': text,
            'length': length,
            'weight': weight,
            'type': 'type18',
            'note': f'直料圓弧 R{radius}'
        }
