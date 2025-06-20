import matplotlib.pyplot as plt
import numpy as np
from .common import figure_to_base64, draw_dimension_line, set_plot_limits
from config import MAIN_REBAR_HOOK_LENGTHS

def _get_hook_length(rebar_number):
    """查詢鋼筋彎鉤長度"""
    return MAIN_REBAR_HOOK_LENGTHS['normal'].get(rebar_number, 15) # 預設15cm

def draw_stirrup(shape_type, width, height, rebar_number, settings, image_width=400, image_height=300):
    """
    繪製各種類型的箍筋（地箍、U箍、L箍等），符合範例圖示
    """
    fig, ax = plt.subplots(figsize=(image_width / 100, image_height / 100))
    color = settings['colors']['rebar']
    line_width = settings['line_width']
    font_size = settings['font_size']
    dim_offset = settings['dimension_offset']
    
    hook_len = _get_hook_length(rebar_number)
    
    # 根據圖片範例調整繪圖邏輯
    if shape_type in ['地箍', '柱箍', '牆箍']: # 封閉矩形
        # 繪製主體
        x = [0, width, width, 0, 0]
        y = [0, 0, height, height, 0]
        ax.plot(x, y, color=color, linewidth=line_width, solid_capstyle='round')

        # 繪製彎鉤 (模擬135度或90度彎鉤)
        hook_angle_rad = np.deg2rad(45) # 135度彎鉤
        if shape_type == '牆箍':
             hook_angle_rad = np.deg2rad(90) # 90度彎鉤
        
        # 左上角彎鉤
        ax.plot([0, -hook_len * np.cos(hook_angle_rad)], 
                [height, height + hook_len * np.sin(hook_angle_rad)], 
                color=color, linewidth=line_width, solid_capstyle='round')
        # 右上角彎鉤 (地箍為180度 + 90度)
        if shape_type == '地箍':
            ax.plot([width, width], [height, height + hook_len], color=color, linewidth=line_width) # 90度
            ax.plot([width, width-hook_len], [height+hook_len, height+hook_len], color=color, linewidth=line_width) # 180度
        else:
            ax.plot([width, width + hook_len * np.cos(hook_angle_rad)], 
                    [height, height + hook_len * np.sin(hook_angle_rad)], 
                    color=color, linewidth=line_width, solid_capstyle='round')

        # 標註尺寸
        draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
        draw_dimension_line(ax, (width, 0), (width, height), f'{height:.0f}', settings, offset=dim_offset, horizontal=False)

    elif shape_type == 'U箍':
        x = [hook_len, 0, 0, width, width, width-hook_len]
        y = [height, height, 0, 0, height, height]
        ax.plot(x, y, color=color, linewidth=line_width, solid_capstyle='round')
        
        # 標註尺寸
        draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
        draw_dimension_line(ax, (0, 0), (0, height), f'{height:.0f}', settings, offset=-dim_offset, horizontal=False)

    elif shape_type == 'L箍':
        x = [0, 0, width]
        y = [height, 0, 0]
        ax.plot(x, y, color=color, linewidth=line_width, solid_capstyle='round')

        # 90度彎鉤
        ax.plot([width, width+hook_len], [0,0], color=color, linewidth=line_width, solid_capstyle='round')
        # 135度彎鉤
        hook_angle_rad = np.deg2rad(45)
        ax.plot([0, -hook_len * np.cos(hook_angle_rad)], 
                [height, height + hook_len * np.sin(hook_angle_rad)], 
                color=color, linewidth=line_width, solid_capstyle='round')

        # 標註尺寸
        draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
        draw_dimension_line(ax, (0, 0), (0, height), f'{height:.0f}', settings, offset=-dim_offset, horizontal=False)

    elif shape_type == '半箍': # Z字形
        x = [0, 0, width, width]
        y = [hook_len, 0, height, height-hook_len]
        ax.plot(x,y, color=color, linewidth=line_width)
        # 135度彎鉤
        hook_angle_rad = np.deg2rad(45)
        ax.plot([0, -hook_len * np.cos(hook_angle_rad)], 
                [0, -hook_len * np.sin(hook_angle_rad)], 
                color=color, linewidth=line_width, solid_capstyle='round')
        ax.plot([width, width + hook_len * np.cos(hook_angle_rad)], 
                [height, height + hook_len * np.sin(hook_angle_rad)], 
                color=color, linewidth=line_width, solid_capstyle='round')

        # 標註尺寸
        draw_dimension_line(ax, (0, 0), (width, 0), f'{width:.0f}', settings, offset=-dim_offset)
        draw_dimension_line(ax, (width, 0), (width, height), f'{height:.0f}', settings, offset=dim_offset, horizontal=False)

    else: # 預設畫一個問號
        ax.text(0.5, 0.5, '?', fontsize=font_size*5, ha='center', va='center', transform=ax.transAxes)

    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')
    
    # 根據圖形內容自動調整邊界
    set_plot_limits(ax, settings['margin'])
    
    return figure_to_base64(fig) 