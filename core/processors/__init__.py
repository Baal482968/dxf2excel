"""
鋼筋處理器模組
"""

from .base_processor import BaseRebarProcessor
from .type10_processor import Type10Processor
from .type11_processor import Type11Processor
from .type12_processor import Type12Processor

# 註冊所有處理器
PROCESSORS = {
    'type10': Type10Processor(),
    'type11': Type11Processor(),
    'type12': Type12Processor(),
}

def get_processor(rebar_type):
    """根據鋼筋類型獲取對應的處理器"""
    return PROCESSORS.get(rebar_type)

def get_all_processors():
    """獲取所有處理器"""
    return PROCESSORS

__all__ = ['BaseRebarProcessor', 'get_processor', 'get_all_processors']
