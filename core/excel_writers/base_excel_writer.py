"""
基礎 Excel 寫入器抽象類
"""

from abc import ABC, abstractmethod
import os
import tempfile


class BaseExcelWriter(ABC):
    """Excel 寫入器基礎類"""
    
    def __init__(self, graphics_manager=None):
        self.graphics_manager = graphics_manager
        self.graphics_available = graphics_manager is not None
        self.temp_files = []
        self.rebar_type = None
    
    @abstractmethod
    def get_rebar_type(self):
        """獲取鋼筋類型"""
        pass
    
    @abstractmethod
    def generate_visual(self, rebar):
        """生成鋼筋視覺表示（圖片或文字描述）"""
        pass
    
    def generate_text_description(self, rebar):
        """生成文字描述"""
        segments = self._get_rebar_segments(rebar)
        rebar_id = rebar.get('rebar_number', '#4')
        
        if len(segments) == 1:
            return f"直鋼筋 {rebar_id}\n長度: {int(segments[0])}cm"
        elif len(segments) == 2:
            return f"L型鋼筋 {rebar_id}\n{int(segments[0])} + {int(segments[1])}cm"
        elif len(segments) == 3:
            return f"U型鋼筋 {rebar_id}\n{int(segments[0])} + {int(segments[1])} + {int(segments[2])}cm"
        else:
            return f"複雜鋼筋 {rebar_id}\n{' + '.join(str(int(s)) for s in segments)}cm"
    
    def _get_rebar_segments(self, rebar):
        """從鋼筋資料中提取分段長度"""
        # 直接檢查 segments 欄位
        if 'segments' in rebar and isinstance(rebar['segments'], list) and rebar['segments']:
            return rebar['segments']
        # 如果沒有 segments，嘗試其他欄位
        segments = []
        segment_keys = ['lengths', 'A', 'B', 'C', 'D', 'E']
        for key in segment_keys:
            if key in rebar and rebar[key] is not None:
                if key == 'lengths' and isinstance(rebar[key], list):
                    segments = rebar[key]
                    break
                elif key in ['A', 'B', 'C', 'D', 'E']:
                    if key == 'A' and segments == []:
                        segments = []
                    if rebar[key] > 0:
                        segments.append(rebar[key])
        # 如果還是沒有分段資料，使用總長度
        if not segments and 'length' in rebar and rebar['length'] > 0:
            segments = [rebar['length']]
        return segments
    
    def _save_image_to_temp(self, image):
        """將圖片保存到臨時檔案"""
        if image:
            temp_img_path = tempfile.mktemp(suffix='.png')
            image.save(temp_img_path)
            self.temp_files.append(temp_img_path)
            return temp_img_path
        return None
    
    def cleanup_temp_files(self):
        """清理暫存檔案"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    print(f"🗑️ 已清理暫存檔案: {temp_file}")
            except Exception as e:
                print(f"⚠️ 清理暫存檔案失敗: {e}")
        self.temp_files.clear()
