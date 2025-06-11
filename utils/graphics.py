"""
圖形繪製相關功能模組 - 增強版
整合專業鋼筋圖示功能，符合範例1的需求
"""

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="matplotlib")

# 設定中文字體支援
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'Heiti TC', 'Microsoft JhengHei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

import matplotlib.patches as patches
from matplotlib.patches import Arc
import numpy as np
from PIL import Image
import io
import base64
import os
import re

class GraphicsManager:
    """圖形繪製管理器 - 增強版"""
    
    def __init__(self):
        """初始化圖形管理器"""
        # 標準鋼筋直徑對照表 (mm) - 符合CNS 560
        self.rebar_diameters = {
            "#2": 6.4, "#3": 9.5, "#4": 12.7, "#5": 15.9, "#6": 19.1,
            "#7": 22.2, "#8": 25.4, "#9": 28.7, "#10": 32.3, "#11": 35.8,
            "#12": 43.0, "#13": 50.0, "#14": 57.0
        }
        
        # 標準彎曲半徑倍數 - 依據建築技術規則
        self.bend_radius_multiplier = {
            "#2": 3, "#3": 3, "#4": 4, "#5": 5, "#6": 6, "#7": 7,
            "#8": 8, "#9": 9, "#10": 10, "#11": 11, "#12": 12,
            "#13": 13, "#14": 14
        }
        
        # 專業模式圖形參數
        self.professional_settings = {
            'line_width': 4,
            'font_size': 12,
            'dimension_offset': 25,
            'margin': 30,
            'colors': {
                'rebar': '#2C3E50',      # 鋼筋顏色（深藍灰）
                'dimension': '#E74C3C',   # 尺寸線顏色（紅色）
                'text': '#2C3E50',       # 文字顏色
                'radius': '#27AE60'      # 半徑標註顏色（綠色）
            }
        }
        
        # 基本模式圖形參數
        self.basic_settings = {
            'line_width': 2,
            'font_size': 10,
            'colors': {
                'rebar': 'blue',
                'text': 'black'
            }
        }
    
    @staticmethod
    def check_dependencies():
        """檢查圖形繪製所需的套件是否已安裝"""
        try:
            import matplotlib
            import numpy
            from PIL import Image
            return True, []
        except ImportError as e:
            missing_packages = []
            
            try:
                import matplotlib
            except ImportError:
                missing_packages.append("matplotlib")
            
            try:
                import numpy
            except ImportError:
                missing_packages.append("numpy")
            
            try:
                from PIL import Image
            except ImportError:
                missing_packages.append("pillow")
            
            return False, missing_packages
    
    def get_bend_radius(self, rebar_number):
        """計算彎曲半徑（符合建築技術規則）"""
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        multiplier = self.bend_radius_multiplier.get(rebar_number, 4)
        return diameter * multiplier
    
    def get_material_grade(self, rebar_number):
        """根據鋼筋編號判定材料等級"""
        if rebar_number in ["#2", "#3", "#4", "#5", "#6"]:
            return "SD280"
        elif rebar_number in ["#7", "#8", "#9", "#10"]:
            return "SD420"
        else:
            return "SD490"
    
    def draw_dimension_line(self, ax, start_point, end_point, offset, text, 
                           horizontal=True, settings=None):
        """繪製專業尺寸標註線"""
        if settings is None:
            settings = self.professional_settings
            
        x1, y1 = start_point
        x2, y2 = end_point
        color = settings['colors']['dimension']
        font_size = settings['font_size']
        
        if horizontal:
            # 水平尺寸線
            dim_y = y1 + offset
            # 主尺寸線
            ax.plot([x1, x2], [dim_y, dim_y], color=color, linewidth=1.5)
            # 引出線
            ax.plot([x1, x1], [y1, dim_y], color=color, linewidth=1)
            ax.plot([x2, x2], [y2, dim_y], color=color, linewidth=1)
            # 箭頭
            arrow_size = 3
            ax.plot([x1, x1+arrow_size], [dim_y, dim_y-arrow_size/2], color=color, linewidth=1.5)
            ax.plot([x1, x1+arrow_size], [dim_y, dim_y+arrow_size/2], color=color, linewidth=1.5)
            ax.plot([x2, x2-arrow_size], [dim_y, dim_y-arrow_size/2], color=color, linewidth=1.5)
            ax.plot([x2, x2-arrow_size], [dim_y, dim_y+arrow_size/2], color=color, linewidth=1.5)
            # 尺寸文字
            mid_x = (x1 + x2) / 2
            ax.text(mid_x, dim_y + 8, text, ha='center', va='bottom',
                   fontsize=font_size, fontweight='bold', color=settings['colors']['text'])
        else:
            # 垂直尺寸線
            dim_x = x1 + offset
            ax.plot([dim_x, dim_x], [y1, y2], color=color, linewidth=1.5)
            ax.plot([x1, dim_x], [y1, y1], color=color, linewidth=1)
            ax.plot([x2, dim_x], [y2, y2], color=color, linewidth=1)
            # 箭頭
            arrow_size = 3
            ax.plot([dim_x, dim_x-arrow_size/2], [y1, y1+arrow_size], color=color, linewidth=1.5)
            ax.plot([dim_x, dim_x+arrow_size/2], [y1, y1+arrow_size], color=color, linewidth=1.5)
            ax.plot([dim_x, dim_x-arrow_size/2], [y2, y2-arrow_size], color=color, linewidth=1.5)
            ax.plot([dim_x, dim_x+arrow_size/2], [y2, y2-arrow_size], color=color, linewidth=1.5)
            # 尺寸文字
            mid_y = (y1 + y2) / 2
            ax.text(dim_x + 12, mid_y, text, ha='left', va='center',
                   fontsize=font_size, fontweight='bold', color=settings['colors']['text'], rotation=90)
    
    def draw_straight_rebar(self, length, rebar_number, professional=True, width=600, height=180):
        """繪製直鋼筋圖示"""
        if professional:
            return self._draw_professional_straight_rebar(length, rebar_number, width, height)
        else:
            return self._draw_basic_straight_rebar(length, rebar_number, width, height)
    
    def _draw_professional_straight_rebar(self, length, rebar_number, width=240, height=80):
        """極簡直鋼筋圖示（放大線條與字體，內容置中）"""
        fig, ax = plt.subplots(figsize=(width/100, height/100))
        ax.set_aspect('equal')
        scale = 0.5  # 放大比例
        length_scaled = length * scale
        # 置中計算
        start_x = (width - length_scaled) / 2
        start_y = height / 2
        end_x = start_x + length_scaled
        # 主線
        ax.plot([start_x, end_x], [start_y, start_y], color='#2C3E50', linewidth=4, solid_capstyle='round')
        # 長度標註（線上方，黑色，粗體，白色底框，明顯在線上方）
        ax.text((start_x+end_x)/2, start_y-25, f'{int(length)}', ha='center', va='top', fontsize=16, color='black', fontweight='bold', bbox=dict(boxstyle="square,pad=0.2", facecolor='white', edgecolor='none', alpha=1.0))
        # 不顯示編號
        ax.set_xlim(0, width)
        ax.set_ylim(0, height)
        ax.axis('off')
        plt.tight_layout(pad=0.2)
        return self._figure_to_base64(fig)
    
    def _draw_basic_straight_rebar(self, length, rebar_number, width=600, height=180):
        """繪製基本直鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(width/300, height/300))
        ax.plot([0, length], [0, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.set_xlim(-1, length + 1)
        ax.set_ylim(-1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_l_shaped_rebar(self, length1, length2, rebar_number, professional=True, width=600, height=180):
        """繪製 L 型鋼筋圖示（線段長度固定，標註數字依 DXF 文字順序）"""
        return self._draw_professional_l_shaped_rebar(length1, length2, rebar_number, width, height) if professional else self._draw_basic_l_shaped_rebar(length1, length2, rebar_number, width, height)
    
    def _draw_professional_l_shaped_rebar(self, length1, length2, rebar_number, width=360, height=180):
        """極簡L型鋼筋圖示，線段長度固定，標註數字依 DXF 文字順序"""
        fig, ax = plt.subplots(figsize=(width/100, height/100))
        ax.set_aspect('equal')
        # 固定圖形長度
        hor_len = 220  # 橫線長度(px)
        ver_len = 80   # 直線長度(px)
        center_x = width / 2
        center_y = height / 2
        # 左側直線
        start_x = center_x - hor_len / 2
        start_y = center_y + ver_len / 2
        # 畫直線（左側）
        ax.plot([start_x, start_x], [start_y, start_y - ver_len], color='#2C3E50', linewidth=5, solid_capstyle='round')
        # 畫橫線（下方）
        ax.plot([start_x, start_x + hor_len], [start_y - ver_len, start_y - ver_len], color='#2C3E50', linewidth=5, solid_capstyle='round')
        # segments[0] 標在直線中央
        ax.text(start_x - 20, start_y - ver_len/2, f'{int(length1)}', ha='right', va='center', fontsize=32, color='black', fontweight='bold', bbox=dict(boxstyle="square,pad=0.3", facecolor='white', edgecolor='none', alpha=1.0))
        # segments[1] 標在橫線下方中央
        ax.text(start_x + hor_len/2, start_y - ver_len - 20, f'{int(length2)}', ha='center', va='top', fontsize=28, color='black', fontweight='bold', bbox=dict(boxstyle="square,pad=0.3", facecolor='white', edgecolor='none', alpha=1.0))
        ax.set_xlim(0, width)
        ax.set_ylim(0, height)
        ax.axis('off')
        plt.tight_layout(pad=0.2)
        return self._figure_to_base64(fig)
    
    def _draw_basic_l_shaped_rebar(self, length1, length2, rebar_number, width=600, height=180):
        """繪製基本L型鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(width/300, height/300))
        
        # 繪製 L 型鋼筋
        ax.plot([0, length1], [0, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([length1, length1], [0, -length2], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        
        ax.set_xlim(-1, length1 + 1)
        ax.set_ylim(-length2 - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_u_shaped_rebar(self, length1, length2, length3, rebar_number, professional=True, width=600, height=180):
        """繪製 U 型或階梯型鋼筋圖示（自動判斷對稱或階梯）"""
        # 若兩側長度相等，畫對稱U型，否則畫階梯型
        if abs(length1 - length3) < 1e-3:
            return self._draw_professional_u_shaped_rebar(length1, length2, length3, rebar_number, width, height) if professional else self._draw_basic_u_shaped_rebar(length1, length2, length3, rebar_number, width, height)
        else:
            # 階梯型視為複雜型
            return self._draw_professional_complex_rebar([length1, length2, length3], rebar_number, angles=None, width=width, height=height) if professional else self._draw_basic_complex_rebar([length1, length2, length3], rebar_number, width, height)
    
    def _draw_professional_u_shaped_rebar(self, length1, length2, length3, rebar_number, width=240, height=80):
        """極簡U型鋼筋圖示（橫線加長，標註分散，不顯示編號）"""
        fig, ax = plt.subplots(figsize=(width/100, height/100))
        ax.set_aspect('equal')
        # 預留邊界
        margin_x = 18
        margin_y = 16
        # 橫線加長倍率
        hor_scale = 2.6
        # 計算縮放比例
        l1 = float(length1)
        l2 = float(length2) * hor_scale
        l3 = float(length3)
        total_w = l2
        total_h = l1 + l3
        scale_x = (width - 2*margin_x) / total_w if total_w > 0 else 1
        scale_y = (height - 2*margin_y) / total_h if total_h > 0 else 1
        scale = min(scale_x, scale_y)
        # 重新計算各段長度
        l1s = float(length1) * scale
        l2s = float(length2) * hor_scale * scale
        l3s = float(length3) * scale
        # U 字形：橫線在下，兩豎線朝下
        # 起點設在左上角
        p1 = (margin_x, margin_y)  # 左上
        p2 = (margin_x, margin_y + l1s)  # 左下
        p3 = (margin_x + l2s, margin_y + l1s)  # 右下
        p4 = (margin_x + l2s, margin_y)  # 右上
        # 主線（黑色，細線）
        ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color='black', linewidth=1.2, solid_capstyle='butt')  # 左豎
        ax.plot([p2[0], p3[0]], [p2[1], p3[1]], color='black', linewidth=1.2, solid_capstyle='butt')  # 橫（下）
        ax.plot([p3[0], p4[0]], [p3[1], p4[1]], color='black', linewidth=1.2, solid_capstyle='butt')  # 右豎
        # 長度標註（左、下、右，黑色，無粗體，字體小）
        ax.text(p1[0]-10, (p1[1]+p2[1])/2, f'{int(length1)}', ha='right', va='center', fontsize=13, color='black')
        ax.text((p2[0]+p3[0])/2, p2[1]-10, f'{int(length2)}', ha='center', va='bottom', fontsize=13, color='black')
        ax.text(p4[0]+10, (p4[1]+p3[1])/2, f'{int(length3)}', ha='left', va='center', fontsize=13, color='black')
        # 不顯示編號
        ax.set_xlim(0, width)
        ax.set_ylim(height, 0)  # y軸反轉，讓原點在左上
        ax.axis('off')
        plt.tight_layout(pad=0.2)
        return self._figure_to_base64(fig)
    
    def _draw_basic_u_shaped_rebar(self, length1, length2, length3, rebar_number, width=600, height=180):
        """繪製基本U型鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(width/300, height/300))
        
        # 繪製 U 型鋼筋
        ax.plot([0, length1], [0, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([length1, length1], [0, -length2], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([length1, 0], [-length2, -length2], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        ax.plot([0, 0], [-length2, 0], 
                color=self.basic_settings['colors']['rebar'], 
                linewidth=self.basic_settings['line_width'])
        
        ax.set_xlim(-1, length1 + 1)
        ax.set_ylim(-length2 - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(length1/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_complex_rebar(self, segments, rebar_number, professional=True, angles=None, width=600, height=180):
        """繪製複雜多段鋼筋圖示，支援自訂彎曲角度"""
        if professional:
            return self._draw_professional_complex_rebar(segments, rebar_number, angles, width, height)
        else:
            return self._draw_basic_complex_rebar(segments, rebar_number, width, height)
    
    def _draw_professional_complex_rebar(self, segments, rebar_number, angles=None, width=600, height=180):
        """優化：多段階梯式排列，標註每一段長度，角度標註於折點"""
        fig, ax = plt.subplots(figsize=(width/300, height/300))
        ax.set_aspect('equal')
        settings = self.professional_settings
        scale = 2
        bend_radius = self.get_bend_radius(rebar_number) / 10 * scale
        # 起始點
        current_x, current_y = 100, 150
        start_x, start_y = current_x, current_y
        # 方向序列：右、下，交錯排列
        directions = [(1, 0), (0, -1)]
        current_direction = 0
        points = [(current_x, current_y)]
        segment_midpoints = []
        # 預設角度序列
        if angles is None or len(angles) < len(segments)-1:
            angles = (angles or []) + [90] * (len(segments)-1 - (len(angles) or 0))
        # 繪製各段
        for i, length in enumerate(segments):
            length_scaled = length * scale
            dx, dy = directions[current_direction % 2]
            # 計算段終點
            segment_end_x = current_x + length_scaled * dx
            segment_end_y = current_y + length_scaled * dy
            # 繪製直線段
            ax.plot([current_x, segment_end_x], [current_y, segment_end_y], 
                    color=settings['colors']['rebar'], 
                    linewidth=settings['line_width'], 
                    solid_capstyle='round')
            # 標註長度
            if dx != 0:  # 水平線段
                mid_x = (current_x + segment_end_x) / 2
                mid_y = current_y
                ax.text(mid_x, mid_y + settings['dimension_offset'], f'{int(length)}', 
                        ha='center', va='bottom', fontsize=settings['font_size'], fontweight='bold', 
                        color=settings['colors']['text'], bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.8))
            else:  # 垂直線段
                mid_x = current_x
                mid_y = (current_y + segment_end_y) / 2
                ax.text(mid_x + settings['dimension_offset'], mid_y, f'{int(length)}', 
                        ha='left', va='center', fontsize=settings['font_size'], fontweight='bold', 
                        color=settings['colors']['text'], rotation=90, bbox=dict(boxstyle="round,pad=0.2", facecolor='white', alpha=0.8))
            # 角度標註
            if i > 0 and angles and len(angles) >= i:
                angle = angles[i-1]
                if angle != 90:
                    ax.text(current_x, current_y, f'{angle}°', ha='right', va='bottom', fontsize=settings['font_size'], color='red')
            # 更新當前位置
            current_x, current_y = segment_end_x, segment_end_y
            points.append((current_x, current_y))
            current_direction += 1
        # 鋼筋編號標註
        ax.text(start_x - 60, start_y, rebar_number, 
                ha='center', va='center', 
                fontsize=settings['font_size'] + 4, 
                fontweight='bold', 
                color=settings['colors']['text'],
                bbox=dict(boxstyle="round,pad=0.4", facecolor='white', edgecolor='gray'))
        # 技術資訊
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        material_grade = self.get_material_grade(rebar_number)
        total_length = sum(segments)
        segment_text = ' + '.join([str(int(s)) for s in segments])
        all_x = [p[0] for p in points]
        all_y = [p[1] for p in points]
        info_x = max(all_x) + 30
        info_y = max(all_y)
        ax.text(info_x, info_y, 
                f'鋼筋規格: D{diameter}mm\n'
                f'材料等級: {material_grade}\n'
                f'總長度: {int(total_length)}cm\n'
                f'分段: {segment_text}cm\n'
                f'段數: {len(segments)}段\n'
                f'形狀: 階梯', 
                ha='left', va='top', 
                fontsize=settings['font_size'] - 1,
                color=settings['colors']['text'], 
                linespacing=1.5,
                bbox=dict(boxstyle="round,pad=0.5", facecolor='lightblue', alpha=0.3))
        margin = 80
        ax.set_xlim(min(all_x) - margin, max(all_x) + margin + 150)
        ax.set_ylim(min(all_y) - margin, max(all_y) + margin)
        ax.axis('off')
        plt.tight_layout()
        return self._figure_to_base64(fig)
    
    def _draw_basic_complex_rebar(self, segments, rebar_number, width=600, height=180):
        """繪製基本複雜鋼筋圖示（保留原有功能）"""
        fig, ax = plt.subplots(figsize=(width/300, height/300))
        
        # 計算總長度和高度
        total_length = sum(segments)
        max_height = max(segments)
        
        # 繪製多段鋼筋
        x, y = 0, 0
        for i, length in enumerate(segments):
            if i % 2 == 0:
                ax.plot([x, x + length], [y, y], 
                        color=self.basic_settings['colors']['rebar'], 
                        linewidth=self.basic_settings['line_width'])
                x += length
            else:
                ax.plot([x, x], [y, y - length], 
                        color=self.basic_settings['colors']['rebar'], 
                        linewidth=self.basic_settings['line_width'])
                y -= length
        
        ax.set_xlim(-1, total_length + 1)
        ax.set_ylim(-max_height - 1, 1)
        ax.axis('off')
        
        # 加入鋼筋編號標記
        ax.text(total_length/2, 0.3, rebar_number, ha='center', va='center')
        
        return self._figure_to_base64(fig)

    def draw_ascii_rebar(self, segments):
        """生成 ASCII 格式的鋼筋圖示（保留原有功能）"""
        if not segments:
            return ""
            
        # 根據段數生成不同的ASCII圖示
        if len(segments) == 1:
            # 直鋼筋
            length = int(segments[0])
            return f"{'─' * min(length//10, 20)}"
        
        elif len(segments) == 2:
            # L型鋼筋
            return """┌─────────
│
│
│
└"""
        
        elif len(segments) == 3:
            # U型鋼筋
            return """│     │
│     │
└─────┘"""
        
        else:
            # 複雜鋼筋
            return """┌─┬─┐
│ │ │
└─┴─┘"""

    def create_detailed_description(self, segments, rebar_number):
        """創建詳細的鋼筋描述文字"""
        # 基本資訊計算
        total_length = sum(segments)
        diameter = self.rebar_diameters.get(rebar_number, 12.7)
        bend_radius = self.get_bend_radius(rebar_number)
        material_grade = self.get_material_grade(rebar_number)
        
        # 確定鋼筋類型
        if len(segments) == 1:
            shape_type = "直鋼筋"
            bend_info = "無彎曲"
        elif len(segments) == 2:
            shape_type = "L型鋼筋"
            bend_info = "90°彎曲 × 1"
        elif len(segments) == 3:
            shape_type = "U型鋼筋"
            bend_info = "90°彎曲 × 2"
        else:
            shape_type = f"{len(segments)}段複合鋼筋"
            bend_info = f"90°彎曲 × {len(segments)-1}"
        
        # 分段資訊
        segment_info = ""
        total_with_letters = ""
        for i, length in enumerate(segments):
            letter = chr(65 + i)  # A, B, C, ...
            segment_info += f"    {letter}段: {int(length):>4} cm\n"
            if i == 0:
                total_with_letters = f"{int(length)}"
            else:
                total_with_letters += f" + {int(length)}"
        
        # 組合完整描述
        description = f"""╔══════════════════════════════════════════════════════════╗
║                    專業鋼筋規格圖                        ║
╚══════════════════════════════════════════════════════════╝

🔧 基本資訊:
  鋼筋編號: {rebar_number}
  鋼筋類型: {shape_type}
  材料等級: {material_grade}

📏 尺寸規格:
  總長度: {int(total_length):>4} cm
  計算式: {total_with_letters} = {int(total_length)} cm
{segment_info}
⚙️ 技術參數:
  鋼筋直徑: D{diameter} mm
  彎曲半徑: R{int(bend_radius)} mm
  彎曲資訊: {bend_info}
  最小保護層: {int(diameter + 10)} mm

📋 設計依據:
  • CNS 560 鋼筋混凝土用鋼筋
  • 建築技術規則施工編
  • 結構混凝土設計規範
  • 最小彎曲半徑 = {bend_radius//diameter:.0f} × D

⚠️  施工注意事項:
  • 彎曲時須使用標準彎曲機具
  • 彎曲半徑不得小於規範要求
  • 彎曲角度誤差不得超過±2°
  • 鋼筋表面不得有裂縫或損傷"""
        
        return description

    def draw_n_shaped_rebar(self, length1, length2, length3, rebar_number, width=240, height=80):
        """繪製 N 型鋼筋圖示（左上→左中，左中→右中，右中→右下，兩豎線都朝下，橫線水平）"""
        fig, ax = plt.subplots(figsize=(width/100, height/100))
        ax.set_aspect('equal')
        # 固定長度
        ver_len = 32   # 豎線長度(px)
        hor_len = 80   # 橫線長度(px)
        margin_x = 30
        margin_y = 20
        # 座標
        p1 = (margin_x, margin_y)  # 左上
        p2 = (margin_x, margin_y + ver_len)  # 左中
        p3 = (margin_x + hor_len, margin_y + ver_len)  # 右中
        p4 = (margin_x + hor_len, margin_y + ver_len + ver_len)  # 右下
        # 主線
        ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color='black', linewidth=1.2)  # 左豎
        ax.plot([p2[0], p3[0]], [p2[1], p3[1]], color='black', linewidth=1.2)  # 水平橫線
        ax.plot([p3[0], p4[0]], [p3[1], p4[1]], color='black', linewidth=1.2)  # 右豎
        # 長度標註
        ax.text(p1[0]-10, (p1[1]+p2[1])/2, f'{int(length1)}', ha='right', va='center', fontsize=13, color='black')
        ax.text((p2[0]+p3[0])/2, p2[1]-10, f'{int(length2)}', ha='center', va='top', fontsize=13, color='black')
        ax.text(p4[0]+10, (p3[1]+p4[1])/2, f'{int(length3)}', ha='left', va='center', fontsize=13, color='black')
        ax.set_xlim(0, width)
        ax.set_ylim(0, height)
        ax.axis('off')
        plt.tight_layout(pad=0.2)
        return self._figure_to_base64(fig)

    def draw_bent_rebar_diagram(self, angle, length1, length2, rebar_number, width=240, height=80):
        """繪製折彎鋼筋圖示（工程圖標準：左水平→折點→右水平，標註位置正確，置中顯示）"""
        fig, ax = plt.subplots(figsize=(width/100, height/100))
        ax.set_aspect('equal')
        # 固定顯示長度
        L1 = 80
        L2 = 80
        margin = 30
        # 起點設在左側中間
        x0, y0 = margin, height/2
        # 第一段：水平線
        x1, y1 = x0 + L1, y0
        # 第二段：折點後，逆時針 angle
        theta = np.radians(180 - angle)  # 工程圖標準，逆時針
        x2 = x1 + L2 * np.cos(theta)
        y2 = y1 + L2 * np.sin(theta)
        # 畫線
        ax.plot([x0, x1], [y0, y1], color='black', linewidth=1.8)
        ax.plot([x1, x2], [y1, y2], color='black', linewidth=1.8)
        # 長度標註
        ax.text((x0+x1)/2, y0-12, f'{int(length1)}', ha='center', va='bottom', fontsize=12, fontweight='bold', color='black', bbox=dict(boxstyle="square,pad=0.2", facecolor='white', edgecolor='none', alpha=1.0))
        ax.text((x1+x2)/2, (y1+y2)/2-12, f'{int(length2)}', ha='center', va='bottom', fontsize=12, fontweight='bold', color='black', bbox=dict(boxstyle="square,pad=0.2", facecolor='white', edgecolor='none', alpha=1.0))
        # 角度標註在折點下方中央
        ax.text(x1, y1+18, f'{int(angle)}°', ha='center', va='top', fontsize=12, color='black', bbox=dict(boxstyle="round,pad=0.1", facecolor='white', edgecolor='none', alpha=0.8))
        # 邊界與顯示
        ax.set_xlim(0, width)
        ax.set_ylim(0, height)
        ax.axis('off')
        plt.tight_layout(pad=0.2)
        return self._figure_to_base64(fig)

    def _parse_bent_rebar(self, rebar_number):
        """解析折彎鋼筋字串，回傳 (角度, 長度1, 長度2)"""
        # 支援格式：折140#10-1000+1200 或 折140#10-1000+1200x20
        # 先抓角度
        m = re.match(r'^折(\d+)', rebar_number)
        angle = int(m.group(1)) if m else 0
        # 再抓號數後的長度們
        m2 = re.search(r'#\d+-([\d\.]+)\+([\d\.]+)', rebar_number)
        if m2:
            length1 = int(float(m2.group(1)))
            length2 = int(float(m2.group(2)))
            return angle, length1, length2
        return None

    def generate_rebar_diagram(self, segments, rebar_number, mode="professional", angles=None, width=600, height=180):
        """
        主要入口函數，根據段數和模式生成對應的鋼筋圖示，支援 angles
        """
        try:
            if not segments:
                return None
            # 折彎鋼筋判斷
            if isinstance(rebar_number, str) and rebar_number.startswith('折'):
                parsed = self._parse_bent_rebar(rebar_number)
                if parsed:
                    angle, length1, length2 = parsed
                    return self.draw_bent_rebar_diagram(angle, length1, length2, rebar_number, 240, 80)
            # 過濾有效分段
            valid_segments = [s for s in segments if s and s > 0]
            if not valid_segments:
                return None
            # 確保 segments 是數字列表
            segments = [float(s) for s in segments]
            # N 型判斷
            if isinstance(rebar_number, str) and rebar_number.strip().upper().startswith('N#') and len(valid_segments) == 3:
                return self.draw_n_shaped_rebar(valid_segments[0], valid_segments[1], valid_segments[2], rebar_number, 240, 80)
            # 根據鋼筋類型選擇繪圖函數
            if mode == "ascii":
                return self.draw_ascii_rebar(valid_segments)
            professional = (mode == "professional")
            if len(valid_segments) == 1:
                return self.draw_straight_rebar(valid_segments[0], rebar_number, professional, width, height)
            elif len(valid_segments) == 2:
                return self.draw_l_shaped_rebar(valid_segments[0], valid_segments[1], rebar_number, professional, width, height)
            elif len(valid_segments) == 3:
                # 無論是否對稱，三段都畫成U型
                return self._draw_professional_u_shaped_rebar(valid_segments[0], valid_segments[1], valid_segments[2], rebar_number, 240, 80)
            else:
                return self.draw_complex_rebar(valid_segments, rebar_number, professional, angles, width, height)
        except Exception as e:
            print(f"[ERROR] generate_rebar_diagram 發生錯誤: {str(e)}")
            ascii_diagram = self.draw_ascii_rebar(valid_segments)
            return ascii_diagram

    def save_figure_as_file(self, base64_data, filename, dpi=300):
        """
        將base64圖片數據儲存為檔案
        
        Args:
            base64_data: base64編碼的圖片數據
            filename: 儲存檔名
            dpi: 解析度（實際由原圖決定）
        
        Returns:
            str: 儲存的檔案路徑
        """
        try:
            # 解碼base64數據
            image_data = base64.b64decode(base64_data)
            
            # 確保目錄存在
            directory = os.path.dirname(filename)
            if directory and not os.path.exists(directory):
                os.makedirs(directory)
            
            # 儲存檔案
            with open(filename, 'wb') as f:
                f.write(image_data)
            
            return filename
        except Exception as e:
            raise Exception(f"圖檔儲存失敗: {str(e)}")

    @staticmethod
    def _figure_to_base64(fig):
        """將 matplotlib figure 轉換為 base64 字串"""
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        img_str = base64.b64encode(buf.getvalue()).decode('ascii')
        plt.close(fig)
        return img_str


# 便利函數，讓使用更簡單
def create_graphics_manager():
    """創建圖形管理器實例"""
    return GraphicsManager()

def quick_draw_rebar(segments, rebar_number="#4", mode="professional"):
    """
    快速繪製鋼筋圖示的便利函數
    
    Args:
        segments: 分段長度列表
        rebar_number: 鋼筋編號
        mode: 繪圖模式
    
    Returns:
        tuple: (base64圖片數據, 詳細描述文字)
    """
    gm = GraphicsManager()
    
    # 檢查依賴
    deps_ok, missing = gm.check_dependencies()
    if not deps_ok:
        print(f"⚠️ 缺少套件: {missing}")
        mode = "ascii"
    
    # 生成圖示
    image_data = gm.generate_rebar_diagram(segments, rebar_number, mode)
    
    # 生成描述
    if mode != "ascii":
        description = gm.create_detailed_description(segments, rebar_number)
    else:
        description = f"ASCII 鋼筋圖示 {rebar_number}"
    
    return image_data, description


# 測試函數
def test_graphics_manager():
    """測試圖形管理器的各種功能"""
    gm = GraphicsManager()
    
    # 檢查依賴
    deps_ok, missing = gm.check_dependencies()
    print(f"圖形依賴檢查: {'✅ 通過' if deps_ok else '❌ 失敗'}")
    if missing:
        print(f"缺少套件: {missing}")
    
    # 測試案例
    test_cases = [
        ([300], "#4", "直鋼筋"),
        ([150, 200], "#5", "L型鋼筋"),
        ([120, 300, 120], "#6", "U型鋼筋"),
        ([100, 150, 200, 100], "#7", "複合鋼筋")
    ]
    
    print("\n🎨 圖形管理器測試")
    print("=" * 50)
    
    for i, (segments, rebar_num, desc) in enumerate(test_cases, 1):
        print(f"\n測試 {i}: {desc} {rebar_num}")
        print(f"分段: {segments}")
        
        try:
            # 測試專業模式
            if deps_ok:
                image_data = gm.generate_rebar_diagram(segments, rebar_num, "professional")
                description = gm.create_detailed_description(segments, rebar_num)
                
                # 儲存圖檔
                filename = f"test_rebar_{i}_{rebar_num.replace('#', '')}.png"
                gm.save_figure_as_file(image_data, filename)
                print(f"✅ 專業圖檔已儲存: {filename}")
                
                # 顯示部分描述
                lines = description.split('\n')
                print("📝 描述預覽:")
                for line in lines[:8]:  # 只顯示前8行
                    print(f"  {line}")
                print("  ...")
            else:
                # 測試ASCII模式
                ascii_art = gm.generate_rebar_diagram(segments, rebar_num, "ascii")
                print(f"📝 ASCII圖示: {ascii_art}")
                
        except Exception as e:
            print(f"❌ 錯誤: {str(e)}")
        
        print("-" * 30)


if __name__ == "__main__":
    """當檔案被直接執行時進行測試"""
    test_graphics_manager()