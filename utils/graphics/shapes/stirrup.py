import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch
from .common import figure_to_base64, draw_dimension_line, set_plot_limits
from config import MAIN_REBAR_HOOK_LENGTHS, SUB_REBAR_HOOK_LENGTHS

def _get_hook_length(rebar_number, main=False):
    """查詢鋼筋彎鉤長度"""
    if main:
        return MAIN_REBAR_HOOK_LENGTHS['normal'].get(rebar_number, 15)
    return SUB_REBAR_HOOK_LENGTHS.get(rebar_number, 10)

def _draw_ground_stirrup_schematic(ax, width, height, rebar_number, settings):
    """繪製「地箍」的示意圖"""
    hook_len = _get_hook_length(rebar_number)
    color = settings['colors']['rebar']
    line_width = settings['line_width']
    font_size = settings['font_size'] - 2 # 示意圖文字稍小
    
    # 繪製外框
    box_width, box_height = 100, 60
    rect = FancyBboxPatch((0, 0), box_width, box_height,
                          boxstyle="round,pad=0,rounding_size=5",
                          edgecolor=color, facecolor='none', linewidth=line_width)
    ax.add_patch(rect)

    # 繪製尺寸
    ax.text(box_width / 2, -10, f'{width:.0f}', ha='center', va='top', fontsize=font_size, color=color)
    ax.text(box_width + 5, box_height / 2, f'{height:.0f}', ha='left', va='center', fontsize=font_size, color=color)

    # 繪製內部彎鉤和文字
    ax.text(box_width * 0.4, box_height * 0.7, f'{hook_len:.0f}', ha='center', va='center', fontsize=font_size, color=color)

    # 繪製角度文字
    ax.text(-10, box_height * 0.8, "180°", ha='right', va='center', fontsize=font_size, color=color)
    ax.text(-10, box_height * 0.6, "90°", ha='right', va='center', fontsize=font_size, color=color)

    # 繪製彎鉤符號
    u_x = [box_width * 0.35, box_width * 0.35, box_width * 0.45, box_width * 0.45]
    u_y = [box_height * 0.8, box_height * 0.7, box_height * 0.7, box_height * 0.8]
    ax.plot(u_x, u_y, color=color, linewidth=line_width)

    # 手動設定邊界以確保所有元素可見
    ax.set_xlim(-25, box_width + 25)
    ax.set_ylim(-25, box_height + 25)


def draw_stirrup(shape_type, width, height, rebar_number, settings, image_width=400, image_height=300):
    """
    繪製各種類型的箍筋（地箍、U箍、L箍等），符合範例圖示
    """
    fig, ax = plt.subplots(figsize=(image_width / 100, image_height / 100))
    
    # --- 「地箍」使用新的示意圖繪製邏輯 ---
    if shape_type == '地箍':
        _draw_ground_stirrup_schematic(ax, width, height, rebar_number, settings)
    else:
        # --- 其他箍筋暫時維持舊的繪製邏輯 ---
        color = settings['colors']['rebar']
        line_width = settings['line_width']
        font_size = settings['font_size']
        dim_offset = settings['dimension_offset']
        hook_len = _get_hook_length(rebar_number, main=True)

        if shape_type in ['柱箍', '牆箍']: # 封閉矩形
            x = [0, width, width, 0, 0]
            y = [0, 0, height, height, 0]
            ax.plot(x, y, color=color, linewidth=line_width, solid_capstyle='round')
            hook_angle_rad = np.deg2rad(45)
            if shape_type == '牆箍':
                 hook_angle_rad = np.deg2rad(90)
            ax.plot([0, -hook_len * np.cos(hook_angle_rad)], 
                    [height, height + hook_len * np.sin(hook_angle_rad)], 
                    color=color, linewidth=line_width, solid_capstyle='round')
            ax.plot([width, width + hook_len * np.cos(hook_angle_rad)], 
                    [height, height + hook_len * np.sin(hook_angle_rad)], 
                    color=color, linewidth=line_width, solid_capstyle='round')
            draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
            draw_dimension_line(ax, (width, 0), (width, height), f'{height:.0f}', settings, offset=dim_offset, horizontal=False)

        elif shape_type == 'U箍':
            x = [hook_len, 0, 0, width, width, width-hook_len]
            y = [height, height, 0, 0, height, height]
            ax.plot(x, y, color=color, linewidth=line_width, solid_capstyle='round')
            draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
            draw_dimension_line(ax, (0, 0), (0, height), f'{height:.0f}', settings, offset=-dim_offset, horizontal=False)

        elif shape_type == 'L箍':
            x = [0, 0, width]
            y = [height, 0, 0]
            ax.plot(x, y, color=color, linewidth=line_width, solid_capstyle='round')
            ax.plot([width, width+hook_len], [0,0], color=color, linewidth=line_width, solid_capstyle='round')
            hook_angle_rad = np.deg2rad(45)
            ax.plot([0, -hook_len * np.cos(hook_angle_rad)], 
                    [height, height + hook_len * np.sin(hook_angle_rad)], 
                    color=color, linewidth=line_width, solid_capstyle='round')
            draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
            draw_dimension_line(ax, (0, 0), (0, height), f'{height:.0f}', settings, offset=-dim_offset, horizontal=False)

        elif shape_type == '半箍': # Z字形
            x = [0, 0, width, width]
            y = [hook_len, 0, height, height-hook_len]
            ax.plot(x,y, color=color, linewidth=line_width)
            hook_angle_rad = np.deg2rad(45)
            ax.plot([0, -hook_len * np.cos(hook_angle_rad)], 
                    [0, -hook_len * np.sin(hook_angle_rad)], 
                    color=color, linewidth=line_width, solid_capstyle='round')
            ax.plot([width, width + hook_len * np.cos(hook_angle_rad)], 
                    [height, height + hook_len * np.sin(hook_angle_rad)], 
                    color=color, linewidth=line_width, solid_capstyle='round')
            draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
            draw_dimension_line(ax, (width, 0), (width, height), f'{height:.0f}', settings, offset=dim_offset, horizontal=False)

        else: # 預設畫一個問號
            ax.text(0.5, 0.5, '?', fontsize=font_size*5, ha='center', va='center', transform=ax.transAxes)
        
        # 僅為非地箍圖形設定自動邊界
        set_plot_limits(ax, settings['margin'])

    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')
    
    return figure_to_base64(fig) 