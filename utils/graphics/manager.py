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

    def generate_type11_rebar_image(self, length, rebar_number, output_path=None):
        """生成 type11 鋼筋（安全彎鉤直）圖片"""
        try:
            # 尋找 type11 材料
            type11_material = None
            for material in self.available_materials:
                if material.startswith("11-"):
                    type11_material = material
                    break
            
            if not type11_material:
                print(f"❌ 找不到 type11 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.materials_dir / type11_material / "graphic-material.svg"
            
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            return self._create_type11_rebar_image_from_svg(svg_path, length, rebar_number)
            
        except Exception as e:
            print(f"❌ 生成 type11 鋼筋圖片失敗: {e}")
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
            
            # 創建圖片 - 更大的尺寸以支援更好的縮放
            img_width = 1200  # 增加寬度
            img_height = 600
            image = Image.new('RGB', (img_width, img_height), color='white')
            draw = ImageDraw.Draw(image)
            
            # 繪製鋼筋線條（固定長度）- 調整位置和寬度，接近圖片寬度
            padding = 100  # 左右 padding
            line_start_x = padding
            line_end_x = img_width - padding  # 從左邊 padding 到右邊 padding
            line_y = img_height // 2  # 垂直置中
            
            # 繪製線條 - 增加寬度
            draw.line([(line_start_x, line_y), (line_end_x, line_y)], fill='black', width=12)
            
            # 繪製長度文字（在線條上方，不帶單位和小數點）- 調整字體大小和位置
            text = str(int(length))  # 轉換為整數並移除小數點
            try:
                # 嘗試使用系統字體 - 增加字體大小，讓數字更明顯
                font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 72)
            except:
                # 回退到預設字體
                font = ImageFont.load_default()
            
            # 計算文字位置（置中於線條上方）- 調整位置
            text_bbox = draw.textbbox((0, 0), text, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            text_x = line_start_x + (line_end_x - line_start_x) // 2 - text_width // 2
            text_y = line_y - 120  # 在線條上方 120 像素，給數字更多空間
            
            # 繪製文字
            draw.text((text_x, text_y), text, fill='black', font=font)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建圖片失敗: {e}")
            return None

    def _create_type11_rebar_image_from_svg(self, svg_path, length, rebar_number):
        """從 SVG 創建 type11 鋼筋（安全彎鉤直）圖片"""
        try:
            # 解析 SVG
            tree = ET.parse(svg_path)
            root = tree.getroot()
            
            # 創建圖片 - 更大的尺寸以支援標註文字
            img_width = 1200
            img_height = 800  # 增加高度以容納上方標註
            image = Image.new('RGB', (img_width, img_height), color='white')
            draw = ImageDraw.Draw(image)
            
            # 繪製基礎鋼筋線條（從 SVG 的 path 資訊）
            # 根據 SVG 中的 path，繪製安全彎鉤直的形狀
            padding = 100
            line_start_x = padding
            line_end_x = img_width - padding
            line_y = img_height // 2
            
            # 繪製主要直線段
            draw.line([(line_start_x, line_y), (line_end_x - 150, line_y)], fill='black', width=12)
            
            # 繪製彎鉤部分（向右彎曲）
            hook_start_x = line_end_x - 150
            hook_end_x = line_end_x - 50
            hook_radius = 50
            
            # 繪製彎鉤的圓弧（簡化為直線）
            draw.line([(hook_start_x, line_y), (hook_start_x, line_y - hook_radius)], fill='black', width=12)
            draw.line([(hook_start_x, line_y - hook_radius), (hook_end_x, line_y - hook_radius)], fill='black', width=12)
            
            # 嘗試載入字體
            try:
                # 主要字體（用於長度數字）
                main_font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 60)
                # 標註字體（用於角度和說明文字）
                label_font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 48)
                # 標題字體（用於"安全彎鉤"）
                title_font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 56)
            except:
                # 回退到預設字體
                main_font = ImageFont.load_default()
                label_font = ImageFont.load_default()
                title_font = ImageFont.load_default()
            
            # 1. 左上方：根據文字長度放入長度 398
            length_text = str(int(length))
            length_bbox = draw.textbbox((0, 0), length_text, main_font)
            length_width = length_bbox[2] - length_bbox[0]
            length_x = line_start_x + 50
            length_y = line_y - 200
            draw.text((length_x, length_y), length_text, fill='black', font=main_font)
            
            # 2. 正右邊：固定是 180'
            angle_text = "180°"
            angle_bbox = draw.textbbox((0, 0), angle_text, label_font)
            angle_width = angle_bbox[2] - angle_bbox[0]
            angle_x = line_end_x + 20
            angle_y = line_y - 20
            draw.text((angle_x, angle_y), angle_text, fill='black', font=label_font)
            
            # 3. 右上方：固定是 10
            size_text = "10"
            size_bbox = draw.textbbox((0, 0), size_text, label_font)
            size_width = size_bbox[2] - size_bbox[0]
            size_x = line_end_x - 50
            size_y = line_y - 250
            draw.text((size_x, size_y), size_text, fill='black', font=label_font)
            
            # 4. 正上方：固定文字"安全彎鉤"
            title_text = "安全彎鉤"
            title_bbox = draw.textbbox((0, 0), title_text, title_font)
            title_width = title_bbox[2] - title_bbox[0]
            title_x = (img_width - title_width) // 2  # 水平置中
            title_y = 80  # 靠近頂部
            draw.text((title_x, title_y), title_text, fill='black', font=title_font)
            
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type11 圖片失敗: {e}")
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
