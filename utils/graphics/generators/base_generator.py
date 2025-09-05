"""
基礎圖形生成器抽象類
"""

from abc import ABC, abstractmethod
from pathlib import Path
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont

class BaseImageGenerator(ABC):
    """圖形生成器基礎類"""
    
    def __init__(self, materials_dir="assets/materials"):
        self.materials_dir = Path(materials_dir)
        self.rebar_type = None
    
    @abstractmethod
    def get_material_prefix(self):
        """獲取材料目錄前綴"""
        pass
    
    @abstractmethod
    def generate_image(self, *args, **kwargs):
        """生成圖片"""
        pass
    
    def find_material(self, available_materials):
        """尋找對應的材料目錄"""
        prefix = self.get_material_prefix()
        for material in available_materials:
            if material.startswith(prefix):
                return material
        return None
    
    def get_svg_path(self, material_name):
        """獲取 SVG 檔案路徑"""
        return self.materials_dir / material_name / "graphic-material.svg"
    
    def parse_svg(self, svg_path):
        """解析 SVG 檔案"""
        try:
            tree = ET.parse(svg_path)
            root = tree.getroot()
            return root
        except Exception as e:
            print(f"❌ SVG 解析失敗: {e}")
            return None
    
    def create_base_image(self, width=800, height=400):
        """創建基礎圖片"""
        image = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(image)
        return image, draw
    
    def get_font(self, size=32):
        """獲取字體"""
        try:
            return ImageFont.truetype("/System/Library/Fonts/Arial.ttf", size)
        except:
            return ImageFont.load_default()
    
    def draw_text_centered(self, draw, text, x, y, font, fill='black'):
        """繪製置中文字"""
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_x = x - text_width // 2
        draw.text((text_x, y), text, fill=fill, font=font)
