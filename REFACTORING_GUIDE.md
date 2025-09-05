# 鋼筋處理系統重構指南

## 🎯 重構目標

將原本集中在三個檔案中的鋼筋處理邏輯，重構為模組化架構，提高可維護性和擴展性。

## 📁 新架構結構

### 1. 鋼筋處理器模組化
```
core/
├── rebar_processor.py          # 原始檔案（保留向後相容）
├── rebar_processor_new.py      # 新的主協調器
└── processors/
    ├── __init__.py
    ├── base_processor.py       # 基礎處理器抽象類
    ├── type10_processor.py     # 直料處理器
    ├── type11_processor.py     # 安全彎鉤直處理器
    ├── type12_processor.py     # 折料處理器
    └── type18_processor.py     # 未來擴展...
```

### 2. 圖形生成器模組化
```
utils/graphics/
├── manager.py                  # 原始檔案（保留向後相容）
├── manager_new.py              # 新的主協調器
└── generators/
    ├── __init__.py
    ├── base_generator.py       # 基礎生成器抽象類
    ├── type10_generator.py     # 直料圖形生成器
    ├── type11_generator.py     # 安全彎鉤直圖形生成器
    ├── type12_generator.py     # 折料圖形生成器
    └── type18_generator.py     # 未來擴展...
```

### 3. Excel 寫入器簡化
```
core/
├── excel_writer.py             # 原始檔案（保留向後相容）
└── excel_writer_new.py         # 新的簡化版本
```

## 🔄 遷移步驟

### 步驟 1：測試新架構
```python
# 使用新的處理器
from core.rebar_processor_new import RebarProcessor

processor = RebarProcessor()
result = processor.parse_rebar_text("V113°#10-900+200x2")
print(f"支援的類型: {processor.get_supported_types()}")
```

### 步驟 2：使用新的圖形管理器
```python
# 使用新的圖形管理器
from utils.graphics.manager_new import GraphicsManager

manager = GraphicsManager()
image = manager.generate_rebar_image('type12', segments, angles, rebar_number)
```

### 步驟 3：使用新的 Excel 寫入器
```python
# 使用新的 Excel 寫入器
from core.excel_writer_new import ExcelWriter

writer = ExcelWriter(image_mode="mixed")
writer.create_workbook()
# ... 其他操作
```

## ✨ 新架構優勢

### 1. **模組化設計**
- 每個鋼筋類型有獨立的處理器和生成器
- 易於維護和測試
- 支援動態添加新類型

### 2. **易於擴展**
```python
# 添加新的鋼筋類型只需：
# 1. 創建新的處理器
class Type18Processor(BaseRebarProcessor):
    def get_pattern(self):
        return r'新格式正則'
    
    def parse_match(self, match, text):
        return {...}

# 2. 創建新的生成器
class Type18ImageGenerator(BaseImageGenerator):
    def generate_image(self, ...):
        return image

# 3. 註冊到系統
PROCESSORS['type18'] = Type18Processor()
GENERATORS['type18'] = Type18ImageGenerator()
```

### 3. **向後相容**
- 保留原始檔案和 API
- 現有代碼無需修改
- 可以逐步遷移

### 4. **統一接口**
```python
# 統一的圖片生成接口
image = graphics_manager.generate_rebar_image(rebar_type, *args, **kwargs)

# 統一的文字解析接口
result = processor.parse_rebar_text(text)
```

## 🚀 使用範例

### 添加新的鋼筋類型（Type18）

1. **創建處理器** (`core/processors/type18_processor.py`):
```python
from .base_processor import BaseRebarProcessor

class Type18Processor(BaseRebarProcessor):
    def get_pattern(self):
        return r'新格式正則'
    
    def parse_match(self, match, text):
        # 解析邏輯
        return {...}
```

2. **創建生成器** (`utils/graphics/generators/type18_generator.py`):
```python
from .base_generator import BaseImageGenerator

class Type18ImageGenerator(BaseImageGenerator):
    def generate_image(self, ...):
        # 圖形生成邏輯
        return image
```

3. **註冊到系統**:
```python
# 在 __init__.py 中添加
PROCESSORS['type18'] = Type18Processor()
GENERATORS['type18'] = Type18ImageGenerator()
```

## 📊 性能對比

| 特性 | 原始架構 | 新架構 |
|------|----------|--------|
| 檔案數量 | 3個大檔案 | 多個小檔案 |
| 代碼行數 | 單檔案 500+ 行 | 單檔案 <200 行 |
| 維護性 | 困難 | 容易 |
| 擴展性 | 需要修改多個檔案 | 只需添加新檔案 |
| 測試性 | 困難 | 容易 |

## 🔧 配置選項

### 啟用新架構
```python
# 在 main.py 中
from core.rebar_processor_new import RebarProcessor
from utils.graphics.manager_new import GraphicsManager
from core.excel_writer_new import ExcelWriter

# 使用新架構
processor = RebarProcessor()
graphics_manager = GraphicsManager()
excel_writer = ExcelWriter()
```

### 保持向後相容
```python
# 繼續使用原始檔案
from core.rebar_processor import RebarProcessor
from utils.graphics.manager import GraphicsManager
from core.excel_writer import ExcelWriter
```

## 🎉 總結

新的模組化架構提供了：
- ✅ **更好的可維護性**：每個類型獨立管理
- ✅ **更強的擴展性**：添加新類型只需創建新檔案
- ✅ **更高的測試性**：可以獨立測試每個模組
- ✅ **向後相容性**：現有代碼無需修改
- ✅ **統一的接口**：簡化使用方式

建議逐步遷移到新架構，享受更好的開發體驗！
