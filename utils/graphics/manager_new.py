"""
圖形管理器 - 重構版
使用模組化設計，支援多種鋼筋類型
"""

from pathlib import Path
from .generators import get_generator, get_all_generators

class GraphicsManager:
    """圖形管理器 - 主協調器"""
    
    def __init__(self):
        """初始化圖形管理器"""
        self.materials_dir = Path("assets/materials")
        self.available_materials = self._scan_materials()
        self.generators = get_all_generators()
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
    
    def generate_rebar_image(self, rebar_type, *args, **kwargs):
        """生成鋼筋圖片 - 統一入口"""
        generator = self.generators.get(rebar_type)
        if not generator:
            print(f"❌ 找不到 {rebar_type} 的圖形生成器")
            return None
        
        try:
            # 將 available_materials 添加到參數中
            kwargs['available_materials'] = self.available_materials
            return generator.generate_image(*args, **kwargs)
        except Exception as e:
            print(f"❌ 生成 {rebar_type} 鋼筋圖片失敗: {e}")
            return None
    
    # 保留舊的方法以維持向後相容性
    def generate_type10_rebar_image(self, length, rebar_number, output_path=None):
        """生成 type10 鋼筋圖片（向後相容）"""
        return self.generate_rebar_image('type10', length, rebar_number)
    
    def generate_type11_rebar_image(self, length, rebar_number, output_path=None):
        """生成 type11 鋼筋圖片（向後相容）"""
        return self.generate_rebar_image('type11', length, rebar_number)
    
    def generate_type12_rebar_image(self, segments, angles, rebar_number, output_path=None):
        """生成 type12 鋼筋圖片（向後相容）"""
        return self.generate_rebar_image('type12', segments, angles, rebar_number)
    
    def add_generator(self, rebar_type, generator):
        """添加新的生成器"""
        self.generators[rebar_type] = generator
    
    def remove_generator(self, rebar_type):
        """移除生成器"""
        if rebar_type in self.generators:
            del self.generators[rebar_type]
    
    def get_supported_types(self):
        """獲取支援的鋼筋類型"""
        return list(self.generators.keys())
    
    def list_available_materials(self):
        """列出可用的材料"""
        return self.available_materials.copy()
    
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
