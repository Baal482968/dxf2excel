"""
Excel 寫入器模組
"""

from .base_excel_writer import BaseExcelWriter
from .type10_excel_writer import Type10ExcelWriter
from .type11_excel_writer import Type11ExcelWriter
from .type12_excel_writer import Type12ExcelWriter
from .type18_excel_writer import Type18ExcelWriter
from .type19_excel_writer import Type19ExcelWriter

# 註冊所有 Excel 寫入器
EXCEL_WRITERS = {
    'type10': Type10ExcelWriter,
    'type11': Type11ExcelWriter,
    'type12': Type12ExcelWriter,
    'type18': Type18ExcelWriter,
    'type19': Type19ExcelWriter,
}

def get_excel_writer(rebar_type, graphics_manager=None):
    """根據鋼筋類型獲取對應的 Excel 寫入器"""
    writer_class = EXCEL_WRITERS.get(rebar_type)
    if writer_class:
        return writer_class(graphics_manager)
    return None

def get_all_excel_writers():
    """獲取所有 Excel 寫入器"""
    return EXCEL_WRITERS

def create_excel_writer_for_rebar(rebar, graphics_manager=None):
    """根據鋼筋資料創建對應的 Excel 寫入器"""
    rebar_type = rebar.get('type')
    if rebar_type:
        return get_excel_writer(rebar_type, graphics_manager)
    return None

__all__ = ['BaseExcelWriter', 'get_excel_writer', 'get_all_excel_writers', 'create_excel_writer_for_rebar']
