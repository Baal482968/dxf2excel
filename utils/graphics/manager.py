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

    def generate_type12_rebar_image(self, segments, angles, rebar_number, output_path=None):
        """生成 type12 鋼筋（折料）圖片"""
        print(f"🔍 開始生成 type12 鋼筋圖片，段長: {segments}, 角度: {angles}, 號數: {rebar_number}")
        try:
            # 尋找 type12 材料
            type12_material = None
            for material in self.available_materials:
                if material.startswith("12-"):
                    type12_material = material
                    break
            
            print(f"🔍 找到 type12 材料: {type12_material}")
            
            if not type12_material:
                print(f"❌ 找不到 type12 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.materials_dir / type12_material / "graphic-material.svg"
            print(f"🔍 SVG 檔案路徑: {svg_path}")
            
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            print(f"🔍 開始調用 _create_type12_rebar_image_from_svg")
            result = self._create_type12_rebar_image_from_svg(svg_path, segments, angles, rebar_number)
            print(f"🔍 _create_type12_rebar_image_from_svg 返回: {result}")
            return result
            
        except Exception as e:
            print(f"❌ 生成 type12 鋼筋圖片失敗: {e}")
            return None

    def generate_type18_rebar_image(self, length, radius, rebar_number, output_path=None):
        """生成 type18 鋼筋（直料圓弧）圖片"""
        print(f"🔍 開始生成 type18 鋼筋圖片，長度: {length}, 半徑: {radius}, 號數: {rebar_number}")
        try:
            # 尋找 type18 材料
            type18_material = None
            for material in self.available_materials:
                if material.startswith("18-"):
                    type18_material = material
                    break
            
            print(f"🔍 找到 type18 材料: {type18_material}")
            
            if not type18_material:
                print(f"❌ 找不到 type18 材料")
                return None
            
            # 構建 SVG 檔案路徑
            svg_path = self.materials_dir / type18_material / "graphic-material.svg"
            print(f"🔍 SVG 檔案路徑: {svg_path}")
            
            if not svg_path.exists():
                print(f"❌ SVG 檔案不存在: {svg_path}")
                return None
            
            # 解析 SVG 並生成圖片
            print(f"🔍 開始調用 _create_type18_rebar_image_from_svg")
            result = self._create_type18_rebar_image_from_svg(svg_path, length, radius, rebar_number)
            print(f"🔍 _create_type18_rebar_image_from_svg 返回: {result}")
            return result
            
        except Exception as e:
            print(f"❌ 生成 type18 鋼筋圖片失敗: {e}")
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

    def _create_type12_rebar_image_from_svg(self, svg_path, segments, angles, rebar_number):
        """從 SVG 創建 type12 鋼筋（折料）圖片"""
        print(f"🔍 _create_type12_rebar_image_from_svg 開始執行")
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
            
            # 解析 SVG 中的 line 元素
            line_elements = root.findall(".//{http://www.w3.org/2000/svg}line")
            print(f"🔍 找到 {len(line_elements)} 個 line 元素")
            
            if len(line_elements) >= 2:
                # 計算縮放比例（SVG 800x600 -> 圖片 800x400）
                scale_x = img_width / 800
                scale_y = img_height / 600
                
                # 繪製鋼筋線條
                line_width = 8
                
                # 第一條線：(550,400) -> (50,400) - 水平線
                line1 = line_elements[0]
                x1 = int(float(line1.get('x1', 550)) * scale_x)
                y1 = int(float(line1.get('y1', 400)) * scale_y)
                x2 = int(float(line1.get('x2', 50)) * scale_x)
                y2 = int(float(line1.get('y2', 400)) * scale_y)
                draw.line([(x1, y1), (x2, y2)], fill='black', width=line_width)
                
                # 第二條線：(550,400) -> (750,200) - 斜線
                line2 = line_elements[1]
                x3 = int(float(line2.get('x1', 550)) * scale_x)
                y3 = int(float(line2.get('y1', 400)) * scale_y)
                x4 = int(float(line2.get('x2', 750)) * scale_x)
                y4 = int(float(line2.get('y2', 200)) * scale_y)
                draw.line([(x3, y3), (x4, y4)], fill='black', width=line_width)
                
                print(f"🔍 鋼筋線條繪製完成")
                
                # 添加標註（長度和角度）
                try:
                    font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 32)
                    small_font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 24)
                except:
                    font = ImageFont.load_default()
                    small_font = ImageFont.load_default()
                
                # 在水平線上方顯示第一段長度
                if len(segments) >= 1:
                    length1_text = str(int(segments[0]))
                    text_bbox = draw.textbbox((0, 0), length1_text, font)
                    text_width = text_bbox[2] - text_bbox[0]
                    text_x = (x1 + x2) // 2 - text_width // 2
                    text_y = y1 - 50
                    draw.text((text_x, text_y), length1_text, fill='black', font=font)
                
                # 在斜線旁邊顯示第二段長度
                if len(segments) >= 2:
                    length2_text = str(int(segments[1]))
                    text_bbox = draw.textbbox((0, 0), length2_text, small_font)
                    text_width = text_bbox[2] - text_bbox[0]
                    text_x = x4 + 10
                    text_y = (y3 + y4) // 2 - 10
                    draw.text((text_x, text_y), length2_text, fill='black', font=small_font)
                
                # 在斜線上方顯示角度
                if len(angles) >= 1:
                    angle_text = f"{angles[0]}°"
                    text_bbox = draw.textbbox((0, 0), angle_text, small_font)
                    text_width = text_bbox[2] - text_bbox[0]
                    text_x = (x3 + x4) // 2 - text_width // 2
                    text_y = min(y3, y4) - 30
                    draw.text((text_x, text_y), angle_text, fill='black', font=small_font)
                
            else:
                print(f"⚠️ 找不到足夠的 line 元素，使用預設繪製")
                # 如果找不到 line 元素，使用預設的簡單繪製
                padding = 100
                line_start_x = padding
                line_end_x = img_width - padding
                line_y = img_height // 2
                
                # 繪製水平線段
                draw.line([(line_start_x, line_y), (line_end_x - 100, line_y)], fill='black', width=8)
                
                # 繪製斜線段（折料）
                hook_start_x = line_end_x - 100
                hook_end_x = line_end_x - 50
                hook_height = 100
                
                draw.line([(hook_start_x, line_y), (hook_end_x, line_y - hook_height)], fill='black', width=8)
                
                # 添加標註
                try:
                    font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 32)
                    small_font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 24)
                except:
                    font = ImageFont.load_default()
                    small_font = ImageFont.load_default()
                
                # 顯示長度
                if len(segments) >= 1:
                    length_text = str(int(segments[0]))
                    text_bbox = draw.textbbox((0, 0), length_text, font)
                    text_width = text_bbox[2] - text_bbox[0]
                    text_x = (line_start_x + line_end_x - 100) // 2 - text_width // 2
                    text_y = line_y - 50
                    draw.text((text_x, text_y), length_text, fill='black', font=font)
                
                # 顯示角度
                if len(angles) >= 1:
                    angle_text = f"{angles[0]}°"
                    text_bbox = draw.textbbox((0, 0), angle_text, small_font)
                    text_width = text_bbox[2] - text_bbox[0]
                    text_x = (hook_start_x + hook_end_x) // 2 - text_width // 2
                    text_y = line_y - hook_height - 30
                    draw.text((text_x, text_y), angle_text, fill='black', font=small_font)
            
            print(f"🔍 type12 圖片生成完成")
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type12 圖片失敗: {e}")
            return None

    def _create_type18_rebar_image_from_svg(self, svg_path, length, radius, rebar_number):
        """從 SVG 創建 type18 鋼筋（直料圓弧）圖片"""
        print(f"🔍 _create_type18_rebar_image_from_svg 開始執行")
        try:
            # 解析 SVG
            tree = ET.parse(svg_path)
            root = tree.getroot()
            print(f"🔍 SVG 解析成功")
            
            # 創建圖片
            img_width = 800
            img_height = 400
            image = Image.new('RGB', (img_width, img_height), color='white')
            draw = ImageDraw.Draw(image)
            print(f"🔍 圖片創建成功，尺寸: {img_width}x{img_height}")
            
            # 解析 SVG 中的 path 數據（圓弧）
            path_element = root.find(".//{http://www.w3.org/2000/svg}path")
            if path_element is not None:
                # 計算縮放比例
                scale_x = img_width / 800
                scale_y = img_height / 600
                
                # 繪製圓弧鋼筋
                self._draw_arc_rebar(draw, length, radius, img_width, img_height, scale_x, scale_y)
                
            else:
                # 使用預設繪製
                self._draw_default_arc(draw, length, radius, img_width, img_height)
            
            print(f"🔍 type18 圖片生成完成")
            return image
            
        except Exception as e:
            print(f"❌ 從 SVG 創建 type18 圖片失敗: {e}")
            return None
    
    def _draw_arc_rebar(self, draw, length, radius, img_width, img_height, scale_x, scale_y):
        """繪製圓弧鋼筋"""
        import math
        
        # 計算圓弧參數 - 調整位置讓弧線在中央偏上
        center_x = img_width // 2
        center_y = img_height // 2  # 放在中央
        
        # 根據長度計算圓弧角度（假設圓弧長度對應角度）
        # 圓弧長度 = 半徑 × 角度（弧度）
        # 角度 = 圓弧長度 / 半徑
        arc_angle_rad = length / radius if radius > 0 else math.pi / 2
        arc_angle_deg = math.degrees(arc_angle_rad)
        
        # 限制角度範圍（0-180度）
        arc_angle_deg = min(arc_angle_deg, 180)
        
        # 計算圓弧的起始和結束角度 - 開口向下
        # PIL 的角度系統：0度在右側，90度在上方，180度在左側，270度在下方
        # 要開口向下，需要轉 180 度
        start_angle = 270 - arc_angle_deg / 2  # 從左側開始
        end_angle = 270 + arc_angle_deg / 2    # 到右側結束
        
        # 調整半徑以適應圖片大小
        max_radius = min(img_width, img_height) // 3
        actual_radius = min(radius, max_radius)
        
        # 繪製圓弧
        line_width = 8
        bbox = [
            center_x - actual_radius,
            center_y - actual_radius,
            center_x + actual_radius,
            center_y + actual_radius
        ]
        
        # 使用 PIL 的 arc 方法繪製圓弧
        draw.arc(bbox, start_angle, end_angle, fill='black', width=line_width)
        
        # 添加標註
        self._add_arc_annotations(draw, length, radius, center_x, center_y, img_width, img_height)
    
    def _draw_default_arc(self, draw, length, radius, img_width, img_height):
        """繪製預設圓弧"""
        import math
        
        center_x = img_width // 2
        center_y = img_height // 2  # 放在中央
        
        # 計算圓弧角度
        arc_angle_rad = length / radius if radius > 0 else math.pi / 2
        arc_angle_deg = math.degrees(arc_angle_rad)
        arc_angle_deg = min(arc_angle_deg, 180)
        
        # 調整半徑以適應圖片大小
        max_radius = min(img_width, img_height) // 3
        actual_radius = min(radius, max_radius)
        
        # 繪製圓弧 - 開口向下
        line_width = 8
        bbox = [
            center_x - actual_radius,
            center_y - actual_radius,
            center_x + actual_radius,
            center_y + actual_radius
        ]
        
        start_angle = 270 - arc_angle_deg / 2  # 從左側開始
        end_angle = 270 + arc_angle_deg / 2    # 到右側結束
        
        draw.arc(bbox, start_angle, end_angle, fill='black', width=line_width)
        
        # 添加標註
        self._add_arc_annotations(draw, length, radius, center_x, center_y, img_width, img_height)
    
    def _add_arc_annotations(self, draw, length, radius, center_x, center_y, img_width, img_height):
        """添加圓弧標註"""
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 32)
            small_font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", 24)
        except:
            font = ImageFont.load_default()
            small_font = ImageFont.load_default()
        
        # 計算實際使用的半徑
        max_radius = min(img_width, img_height) // 3
        actual_radius = min(radius, max_radius)
        
        # 在弧線開口處顯示長度（弧線下方）
        length_text = str(int(length))
        text_bbox = draw.textbbox((0, 0), length_text, font)
        text_width = text_bbox[2] - text_bbox[0]
        text_x = center_x - text_width // 2
        text_y = center_y + actual_radius + 20  # 在弧線下方
        draw.text((text_x, text_y), length_text, fill='black', font=font)
        
        # 在更下方顯示半徑
        radius_text = f"半徑={radius}"
        text_bbox = draw.textbbox((0, 0), radius_text, small_font)
        text_width = text_bbox[2] - text_bbox[0]
        text_x = center_x - text_width // 2
        text_y = center_y + actual_radius + 60  # 在長度下方
        draw.text((text_x, text_y), radius_text, fill='black', font=small_font)
    
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
