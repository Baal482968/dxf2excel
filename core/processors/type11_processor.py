"""
Type11 安全彎鉤直鋼筋處理器
"""

from .base_processor import BaseRebarProcessor

class Type11Processor(BaseRebarProcessor):
    """Type11 安全彎鉤直鋼筋處理器"""
    
    def __init__(self):
        super().__init__()
        self.rebar_type = 'type11'
    
    def get_pattern(self):
        """獲取正則表達式模式"""
        return r'安(#\d+)-([\d\.]+)x(\d+)'
    
    def parse_match(self, match, text):
        """解析匹配結果"""
        rebar_number = match.group(1)
        length = float(match.group(2))
        count = int(match.group(3))
        
        # 計算重量
        weight = self.calculate_weight(rebar_number, length, count)
        
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
