"""
圖形繪製的通用輔助函式
"""
import io
import base64
import matplotlib.pyplot as plt

def figure_to_base64(fig):
    """將 matplotlib figure 轉換為 base64 字串"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=100, bbox_inches='tight')
    buf.seek(0)
    img_str = base64.b64encode(buf.getvalue()).decode('ascii')
    plt.close(fig)
    return img_str

def draw_dimension_line(ax, start_point, end_point, offset, text, 
                       horizontal=True, settings=None):
    """繪製專業尺寸標註線"""
    if settings is None:
        raise ValueError("settings must be provided for draw_dimension_line")
        
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