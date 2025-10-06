#!/usr/bin/env python3
"""
圖形管理器 - 負責生成鋼筋圖形
"""

import os
from pathlib import Path
from .generators import get_generator, get_all_generators


class GraphicsManager:
    """圖形管理器類別"""
    
    def __init__(self):
        """初始化圖形管理器"""
        self.materials_dir = Path("assets/materials")
        self.available_materials = self._scan_materials()
        self.generators = get_all_generators()
        print(f"📁 找到 {len(self.available_materials)} 種材料類型")
        print(f"🔧 載入 {len(self.generators)} 個圖形生成器")
    
    def _scan_materials(self):
        """掃描材料目錄"""
        materials = []
        if self.materials_dir.exists():
            for item in self.materials_dir.iterdir():
                if item.is_dir():
                    # 檢查是否有必要的檔案
                    svg_file = item / "graphic-material.svg"
                    text_file = item / "text.dxf"
                    if svg_file.exists() and text_file.exists():
                        materials.append(item.name)
        return materials
    
    def generate_type10_rebar_image(self, length, rebar_number, output_path=None):
        """生成 type10 鋼筋圖片"""
        print(f"🔍 開始生成 type10 鋼筋圖片，長度: {length}, 號數: {rebar_number}")
        generator = get_generator('type10')
        if generator:
            return generator.generate_image(length, rebar_number, self.available_materials)
        else:
            print(f"❌ 找不到 type10 生成器")
            return None

    def generate_type11_rebar_image(self, length, rebar_number, output_path=None):
        """生成 type11 鋼筋（安全彎鉤直）圖片"""
        print(f"🔍 開始生成 type11 鋼筋圖片，長度: {length}, 號數: {rebar_number}")
        generator = get_generator('type11')
        if generator:
            return generator.generate_image(length, rebar_number, self.available_materials)
        else:
            print(f"❌ 找不到 type11 生成器")
            return None

    def generate_type12_rebar_image(self, segments, angles, rebar_number, output_path=None):
        """生成 type12 鋼筋（折料）圖片"""
        print(f"🔍 開始生成 type12 鋼筋圖片，段長: {segments}, 角度: {angles}, 號數: {rebar_number}")
        generator = get_generator('type12')
        if generator:
            return generator.generate_image(segments, angles, rebar_number, self.available_materials)
        else:
            print(f"❌ 找不到 type12 生成器")
            return None

    def generate_type18_rebar_image(self, length, radius, rebar_number, output_path=None):
        """生成 type18 鋼筋（直料圓弧）圖片"""
        print(f"🔍 開始生成 type18 鋼筋圖片，長度: {length}, 半徑: {radius}, 號數: {rebar_number}")
        generator = get_generator('type18')
        if generator:
            return generator.generate_image(length, radius, rebar_number, self.available_materials)
        else:
            print(f"❌ 找不到 type18 生成器")
            return None

    def generate_type19_rebar_image(self, straight_length, arc_length, radius, rebar_number, output_path=None):
        """生成 type19 鋼筋（直段+弧段）圖片"""
        print(f"🔍 開始生成 type19 鋼筋圖片，直段: {straight_length}, 弧段: {arc_length}, 半徑: {radius}, 號數: {rebar_number}")
        generator = get_generator('type19')
        if generator:
            return generator.generate_image(straight_length, arc_length, radius, rebar_number, self.available_materials)
        else:
            print(f"❌ 找不到 type19 生成器")
            return None