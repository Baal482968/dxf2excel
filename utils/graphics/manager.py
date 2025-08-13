#!/usr/bin/env python3
"""
圖形管理器 - 負責生成鋼筋圖形
"""

import os
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path


class GraphicsManager:
    """圖形管理器類別"""
    
    def __init__(self):
        """初始化圖形管理器"""
        self.materials_dir = Path("assets/materials")
        self.available_materials = self._scan_materials()
        print(f"📁 找到 {len(self.available_materials)} 種材料類型")
    
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
        try:
            # 尋找 type10 材料
            type10_material = None
            for material in self.available_materials:
                if material.startswith("10-"):
                    type10_material = material
                    break
            
            if not type10_material:
                print(f"❌ 找不到 type10 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.materials_dir / type10_material / "graphic-material.svg"
            
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            return self._create_type10_rebar_image_from_svg(svg_path, length, rebar_number)
            
        except Exception as e:
            print(f"❌ 生成 type10 鋼筋圖片失敗: {e}")
            return None
    
    def _create_type10_rebar_image_from_svg(self, svg_path, length, rebar_number):
        """從 SVG 創建 type10 鋼筋圖片"""
        try:
            # 解析 SVG
            tree = ET.parse(svg_path)
            root = tree.getroot()
            
            # 尋找 line 元素
            line_element = root.find(".//{http://www.w3.org/2000/svg}line")
            
            if line_element is None:
                print(f"❌ SVG 中找不到 line 元素")
                return None
            
            # 獲取 line 的座標
            x1 = float(line_element.get('x1', 350))
            y1 = float(line_element.get('y1', 200))
            x2 = float(line_element.get('x2', 450))  # 固定長度
            y2 = float(line_element.get('y2', 200))
            
            # 創建圖片
            img_width = 800  # 固定寬度
            img_height = 400
            image = Image.new('RGB', (img_width, img_height), color='white')
            draw = ImageDraw.Draw(image)
            
            # 繪製鋼筋線條（固定長度）
            line_start_x = 350
            line_end_x = 450  # 固定長度
            line_y = 200
            
            # 繪製線條
            draw.line([(line_start_x, line_y), (line_end_x, line_y)], fill='black', width=3)
            
            # 繪製長度文字（在線條上方，不帶單位和小數點）
            text = str(int(length))  # 轉換為整數並移除小數點
            try:
                # 嘗試使用系統字體
                font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 24)
            except:
                # 回退到預設字體
                font = ImageFont.load_default()
            
            # 計算文字位置（置中於線條上方）
            text_bbox = draw.textbbox((0, 0), text, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            text_x = line_start_x + (line_end_x - line_start_x) // 2 - text_width // 2
            text_y = line_y - 40  # 在線條上方 40 像素
            
            # 繪製文字
            draw.text((text_x, text_y), text, fill='black', font=font)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建圖片失敗: {e}")
            return None
    
    def save_image(self, image, output_path):
        """保存圖片"""
        try:
            if image:
                image.save(output_path)
                print(f"✅ 圖片已儲存: {output_path}")
                return True
            return False
        except Exception as e:
            print(f"❌ 保存圖片失敗: {e}")
            return False
    
    def list_available_materials(self):
        """列出可用的材料"""
        return self.available_materials.copy()
