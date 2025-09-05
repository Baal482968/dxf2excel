"""
Type12 折料鋼筋處理器
"""

from .base_processor import BaseRebarProcessor

class Type12Processor(BaseRebarProcessor):
    """Type12 折料鋼筋處理器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type12'
    
    def get_pattern(self):
        """獲取正則表達式模式"""
        return r'V(\d+)°(#\d+)-([\d\.]+)\+([\d\.]+)x(\d+)'
    
    def parse_match(self, match, text):
        """解析匹配結果"""
        angle = int(match.group(1))
        rebar_number = match.group(2)
        length1 = float(match.group(3))
        length2 = float(match.group(4))
        count = int(match.group(5))
        
        # 計算總長度
        total_length = length1 + length2
        
        # 計算重量
        weight = self.calculate_weight(rebar_number, total_length, count)
        
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
