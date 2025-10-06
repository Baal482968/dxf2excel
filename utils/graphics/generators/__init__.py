"""
圖形生成器模組
"""

from .base_generator import BaseImageGenerator
from .type10_generator import Type10ImageGenerator
from .type11_generator import Type11ImageGenerator
from .type12_generator import Type12ImageGenerator
from .type18_generator import Type18ImageGenerator
from .type19_generator import Type19ImageGenerator

# 註冊所有生成器
GENERATORS = {
    'type10': Type10ImageGenerator(),
    'type11': Type11ImageGenerator(),
    'type12': Type12ImageGenerator(),
    'type18': Type18ImageGenerator(),
    'type19': Type19ImageGenerator(),
}

def get_generator(rebar_type):
    """根據鋼筋類型獲取對應的生成器"""
    return GENERATORS.get(rebar_type)

def get_all_generators():
    """獲取所有生成器"""
    return GENERATORS

__all__ = ['BaseImageGenerator', 'get_generator', 'get_all_generators']
