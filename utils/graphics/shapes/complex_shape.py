"""
繪製複雜形狀鋼筋
"""
import matplotlib.pyplot as plt
from .common import figure_to_base64

def draw_complex_rebar(segments, rebar_number, professional=True, angles=None, width=700, height=260, settings=None, get_bend_radius_func=None, get_material_grade_func=None, rebar_diameters=None):
    """繪製複雜多段鋼筋圖示，支援自訂彎曲角度"""
    if professional:
        if not all([get_bend_radius_func, get_material_grade_func, rebar_diameters]):
            raise ValueError("Missing required functions for professional complex rebar drawing.")
        return _draw_professional_complex_rebar(segments, rebar_number, angles, width, height, settings, get_bend_radius_func, get_material_grade_func, rebar_diameters)
    else:
        return _draw_basic_complex_rebar(segments, rebar_number, width, height, settings)

def _draw_professional_complex_rebar(segments, rebar_number, angles=None, width=700, height=260, settings=None, get_bend_radius_func=None, get_material_grade_func=None, rebar_diameters=None):
    """優化：多段階梯式排列，標註每一段長度，角度標註於折點"""
    fig, ax = plt.subplots(figsize=(width/300, height/300))
    ax.set_aspect('equal')
    scale = 2
    bend_radius = get_bend_radius_func(rebar_number) / 10 * scale
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
    diameter = rebar_diameters.get(rebar_number, 12.7)
    material_grade = get_material_grade_func(rebar_number)
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
    return figure_to_base64(fig)

def _draw_basic_complex_rebar(segments, rebar_number, width=700, height=260, settings=None):
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
                    color=settings['colors']['rebar'], 
                    linewidth=settings['line_width'])
            x += length
        else:
            ax.plot([x, x], [y, y - length], 
                    color=settings['colors']['rebar'], 
                    linewidth=settings['line_width'])
            y -= length
    
    ax.set_xlim(-1, total_length + 1)
    ax.set_ylim(-max_height - 1, 1)
    ax.axis('off')
    
    # 加入鋼筋編號標記
    ax.text(total_length/2, 0.3, rebar_number, ha='center', va='center')
    
    return figure_to_base64(fig) 