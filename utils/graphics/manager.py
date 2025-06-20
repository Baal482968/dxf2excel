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

import numpy as np
from PIL import Image
import io
import base64
import os
import re

from .shapes import (
    draw_straight_rebar,
    draw_l_shaped_rebar,
    draw_u_shaped_rebar,
    draw_n_shaped_rebar,
    draw_bent_rebar,
    parse_bent_rebar_string,
    draw_complex_rebar,
    draw_stirrup,
)
from .shapes.common import figure_to_base64

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
            'line_width': 2,
            'font_size': 12,
            'dimension_offset': 25,
            'margin': 30,
            'colors': {
                'rebar': '#000000',      # 鋼筋顏色（黑色）
                'dimension': '#000000',   # 尺寸線顏色（黑色）
                'text': '#000000',       # 文字顏色 (黑色)
                'radius': '#000000'      # 半徑標註顏色（黑色）
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

    def generate_rebar_diagram(self, segments, rebar_number, mode="professional", angles=None, width=700, height=260, shape_type=None):
        """
        主要入口函數，根據段數和模式生成對應的鋼筋圖示，支援 angles 和 shape_type
        """
        try:
            # 箍筋判斷
            if shape_type and '箍' in shape_type:
                if len(segments) >= 2:
                    w, h = segments[0], segments[1]
                    settings = self.professional_settings if mode == "professional" else self.basic_settings
                    return draw_stirrup(shape_type, w, h, rebar_number, settings, image_width=width, image_height=height)

            # 折彎鋼筋判斷
            if isinstance(rebar_number, str) and rebar_number.startswith('折'):
                parsed = parse_bent_rebar_string(rebar_number)
                if parsed:
                    angle, length1, length2 = parsed
                    return draw_bent_rebar(angle, length1, length2, rebar_number, 240, 80)

            if not segments:
                return None

            # 過濾有效分段
            valid_segments = [s for s in segments if s and s > 0]
            if not valid_segments:
                return None

            # 確保 segments 是數字列表
            segments_float = [float(s) for s in valid_segments]
            
            # N 型判斷
            if isinstance(rebar_number, str) and rebar_number.strip().upper().startswith('N#') and len(segments_float) == 3:
                return draw_n_shaped_rebar(segments_float[0], segments_float[1], segments_float[2], rebar_number, 240, 80)

            # 根據鋼筋類型選擇繪圖函數
            if mode == "ascii":
                return self.draw_ascii_rebar(segments_float)

            professional = (mode == "professional")
            settings = self.professional_settings if professional else self.basic_settings

            if len(segments_float) == 1:
                return draw_straight_rebar(segments_float[0], rebar_number, professional, width, height, settings)
            elif len(segments_float) == 2:
                return draw_l_shaped_rebar(segments_float[0], segments_float[1], rebar_number, professional, width, height, settings)
            elif len(segments_float) == 3:
                 # 呼叫U型繪圖函式，它內部會判斷是否為階梯形
                 # For non-symmetric U-shapes, it will internally call the complex drawer
                return draw_u_shaped_rebar(
                    segments_float[0], segments_float[1], segments_float[2],
                    rebar_number, professional, 240, 80, settings,
                    complex_drawer=self.draw_complex_rebar_wrapper # 傳入包裝後的函式
                )
            else:
                return self.draw_complex_rebar_wrapper(segments_float, rebar_number, professional, angles, width, height, settings)

        except Exception as e:
            print(f"[ERROR] generate_rebar_diagram 發生錯誤: {str(e)}")
            return self.draw_ascii_rebar(segments)

    def draw_complex_rebar_wrapper(self, segments, rebar_number, professional, angles, width, height, settings):
        """包裝 draw_complex_rebar 以傳入 instance-specific 輔助函式"""
        return draw_complex_rebar(
            segments, rebar_number, professional, angles, width, height, settings,
            get_bend_radius_func=self.get_bend_radius,
            get_material_grade_func=self.get_material_grade,
            rebar_diameters=self.rebar_diameters
        )
    
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
            # The base64 data is already a string, no need to decode from bytes
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

# 便利函數，讓使用更簡單
def create_graphics_manager():
    """創建圖形管理器實例"""
    return GraphicsManager()

def quick_draw_rebar(segments, rebar_number="#4", mode="professional", shape_type=None):
    """
    快速繪製鋼筋圖示的便利函數
    
    Args:
        segments: 分段長度列表
        rebar_number: 鋼筋編號
        mode: 繪圖模式
        shape_type: 圖形類型 (例如 '地箍')
    
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
    image_data = gm.generate_rebar_diagram(segments, rebar_number, mode, shape_type=shape_type)
    
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