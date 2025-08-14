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
        print(f"🔍 開始生成 type11 鋼筋圖片，長度: {length}, 號數: {rebar_number}")
        try:
            # 尋找 type11 材料
            type11_material = None
            for material in self.available_materials:
                if material.startswith("11-"):
                    type11_material = material
                    break
            
            print(f"🔍 找到 type11 材料: {type11_material}")
            
            if not type11_material:
                print(f"❌ 找不到 type11 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.materials_dir / type11_material / "graphic-material.svg"
            print(f"🔍 SVG 檔案路徑: {svg_path}")
            
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            print(f"🔍 開始調用 _create_type11_rebar_image_from_svg")
            result = self._create_type11_rebar_image_from_svg(svg_path, length, rebar_number)
            print(f"🔍 _create_type11_rebar_image_from_svg 返回: {result}")
            return result
            
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
        print(f"🔍 _create_type11_rebar_image_from_svg 開始執行")
        try:
            # 解析 SVG
            tree = ET.parse(svg_path)
            root = tree.getroot()
            print(f"🔍 SVG 解析成功")
            
            # 創建圖片 - 簡潔的尺寸
            img_width = 800
            img_height = 400
            image = Image.new('RGB', (img_width, img_height), color='white')
            draw = ImageDraw.Draw(image)
            print(f"🔍 圖片創建成功，尺寸: {img_width}x{img_height}")
            
            # 解析 SVG 中的 path 數據
            path_element = root.find(".//{http://www.w3.org/2000/svg}path")
            if path_element is not None:
                path_data = path_element.get('d', '')
                print(f"🔍 找到 path 數據: {path_data}")
                
                # 解析 path 數據：M50.00,336.89 L750.00,336.89 L750.00,263.11 L661.46,263.11
                # 這表示：起點 -> 水平線 -> 垂直線 -> 短水平線
                
                # 計算縮放比例（SVG 800x600 -> 圖片 800x400）
                scale_x = img_width / 800
                scale_y = img_height / 600
                
                # 繪製鋼筋線條
                line_width = 8
                
                # 1. 主要水平線段 (50,336.89 -> 750,336.89)
                x1 = int(50 * scale_x)
                y1 = int(336.89 * scale_y)
                x2 = int(750 * scale_x)
                y2 = int(336.89 * scale_y)
                draw.line([(x1, y1), (x2, y2)], fill='black', width=line_width)
                
                # 2. 垂直彎折線段 (750,336.89 -> 750,263.11)
                x3 = int(750 * scale_x)
                y3 = int(336.89 * scale_y)
                x4 = int(750 * scale_x)
                y4 = int(263.11 * scale_y)
                draw.line([(x3, y3), (x4, y4)], fill='black', width=line_width)
                
                # 3. 短水平線段 (750,263.11 -> 661.46,263.11)
                x5 = int(750 * scale_x)
                y5 = int(263.11 * scale_y)
                x6 = int(661.46 * scale_x)
                y6 = int(263.11 * scale_y)
                draw.line([(x5, y5), (x6, y6)], fill='black', width=line_width)
                
                print(f"🔍 鋼筋線條繪製完成")
                
                # 添加長度標註（簡潔的）
                try:
                    font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 36)
                except:
                    font = ImageFont.load_default()
                
                # 在主要水平線上方顯示長度
                length_text = str(int(length))
                text_bbox = draw.textbbox((0, 0), length_text, font)
                text_width = text_bbox[2] - text_bbox[0]
                text_x = (x1 + x2) // 2 - text_width // 2
                text_y = y1 - 50
                draw.text((text_x, text_y), length_text, fill='black', font=font)
                
            else:
                print(f"⚠️ 找不到 path 元素，使用預設繪製")
                # 如果找不到 path，使用預設的簡單繪製
                padding = 100
                line_start_x = padding
                line_end_x = img_width - padding
                line_y = img_height // 2
                
                # 繪製主要直線段
                draw.line([(line_start_x, line_y), (line_end_x - 100, line_y)], fill='black', width=8)
                
                # 繪製彎鉤
                hook_start_x = line_end_x - 100
                hook_end_x = line_end_x - 50
                hook_height = 50
                
                draw.line([(hook_start_x, line_y), (hook_start_x, line_y - hook_height)], fill='black', width=8)
                draw.line([(hook_start_x, line_y - hook_height), (hook_end_x, line_y - hook_height)], fill='black', width=8)
                
                # 添加長度標註
                try:
                    font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 36)
                except:
                    font = ImageFont.load_default()
                
                length_text = str(int(length))
                text_bbox = draw.textbbox((0, 0), length_text, font)
                text_width = text_bbox[2] - text_bbox[0]
                text_x = (line_start_x + line_end_x - 100) // 2 - text_width // 2
                text_y = line_y - 50
                draw.text((text_x, text_y), length_text, fill='black', font=font)
            
            print(f"🔍 type11 圖片生成完成")
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
