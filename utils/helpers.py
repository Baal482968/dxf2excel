"""
通用輔助函數模組
"""

import os
import time
import math
from datetime import datetime

def format_file_size(size_bytes):
    """格式化檔案大小顯示"""
    if size_bytes == 0:
        return "0 B"
    
    size_names = ("B", "KB", "MB", "GB", "TB")
    i = 0
    while size_bytes >= 1024 and i < len(size_names) - 1:
        size_bytes /= 1024
        i += 1
    
    return f"{size_bytes:.2f} {size_names[i]}"

def format_time(seconds):
    """格式化時間顯示"""
    if seconds < 60:
        return f"{seconds:.1f} 秒"
    elif seconds < 3600:
        minutes = int(seconds / 60)
        seconds = seconds % 60
        return f"{minutes} 分 {seconds:.1f} 秒"
    else:
        hours = int(seconds / 3600)
        minutes = int((seconds % 3600) / 60)
        seconds = seconds % 60
        return f"{hours} 時 {minutes} 分 {seconds:.1f} 秒"

def calculate_line_length(start_point, end_point):
    """計算兩點之間的直線距離"""
    x1, y1 = start_point
    x2, y2 = end_point
    return math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)

def calculate_polyline_length(points):
    """計算多段線的總長度"""
    total_length = 0
    for i in range(len(points) - 1):
        total_length += calculate_line_length(points[i], points[i + 1])
    return total_length

def get_file_info(file_path):
    """獲取檔案資訊"""
    try:
        stat = os.stat(file_path)
        return {
            'size': format_file_size(stat.st_size),
            'created': datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
            'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
            'extension': os.path.splitext(file_path)[1].lower()
        }
    except Exception as e:
        return {
            'error': str(e)
        }

def create_progress_tracker(total_steps):
    """創建進度追蹤器"""
    return {
        'current_step': 0,
        'total_steps': total_steps,
        'start_time': time.time(),
        'step_descriptions': {}
    }

def update_progress(tracker, step=None, detail=""):
    """更新進度追蹤器"""
    if step is not None:
        tracker['current_step'] = step
    
    if detail:
        tracker['step_descriptions'][tracker['current_step']] = detail
    
    elapsed_time = time.time() - tracker['start_time']
    percentage = (tracker['current_step'] / tracker['total_steps']) * 100
    
    return {
        'step': tracker['current_step'],
        'total': tracker['total_steps'],
        'percentage': percentage,
        'elapsed_time': format_time(elapsed_time),
        'detail': tracker['step_descriptions'].get(tracker['current_step'], "")
    } 