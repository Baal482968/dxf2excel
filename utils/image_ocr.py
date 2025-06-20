"""
圖片 OCR 處理模組
用於識別 DXF 中圖片實體的數字
"""

import os
import sys
import cv2
import numpy as np
from PIL import Image
import pytesseract
from typing import List, Dict, Optional, Tuple

class ImageOCRProcessor:
    """圖片 OCR 處理器"""
    
    def __init__(self):
        """初始化 OCR 處理器"""
        self.tesseract_config = '--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789.'
        self.min_confidence = 60  # 最小信心度
        
        # Windows 環境下的 Tesseract 路徑設定
        self._setup_tesseract_path()
    
    def _setup_tesseract_path(self):
        """設定 Tesseract 路徑（主要針對 Windows 和 macOS）"""
        if sys.platform.startswith('win'):
            # Windows 常見的 Tesseract 安裝路徑
            possible_paths = [
                r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                r'C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'.format(os.getenv('USERNAME', '')),
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    print(f"[DEBUG] 設定 Tesseract 路徑: {path}")
                    return
            
            print("[WARNING] 未找到 Tesseract，請確保已正確安裝並設定 PATH")
        elif sys.platform.startswith('darwin'):
            # macOS 上 Homebrew 安裝的 Tesseract 路徑
            possible_paths = [
                '/opt/homebrew/bin/tesseract',
                '/usr/local/bin/tesseract',
                '/opt/homebrew/Cellar/tesseract/*/bin/tesseract'
            ]
            
            for path in possible_paths:
                if '*' in path:
                    # 處理通配符路徑
                    import glob
                    matches = glob.glob(path)
                    if matches:
                        # 使用最新版本
                        latest_path = sorted(matches)[-1]
                        pytesseract.pytesseract.tesseract_cmd = latest_path
                        print(f"[DEBUG] 設定 Tesseract 路徑: {latest_path}")
                        return
                elif os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    print(f"[DEBUG] 設定 Tesseract 路徑: {path}")
                    return
            
            print("[WARNING] 未找到 Tesseract，請確保已正確安裝並設定 PATH")
        else:
            # Linux 通常會自動找到
            print("[DEBUG] 使用系統預設的 Tesseract 路徑")
    
    def extract_images_from_dxf(self, modelspace) -> List[Dict]:
        """從 DXF 中提取圖片實體"""
        images = []
        
        try:
            # 遍歷所有圖片實體
            for image in modelspace.query('IMAGE'):
                image_info = {
                    'entity': image,
                    'path': image.dxf.image_def.dxf.filename,
                    'insert_point': (image.dxf.insert.x, image.dxf.insert.y),
                    'size': (image.dxf.size.x, image.dxf.size.y),
                    'rotation': image.dxf.rotation,
                    'scale': (image.dxf.scale.x, image.dxf.scale.y)
                }
                images.append(image_info)
                print(f"[DEBUG] 找到圖片實體: {image_info['path']}")
        
        except Exception as e:
            print(f"[ERROR] 提取圖片實體錯誤: {str(e)}")
        
        return images
    
    def load_image(self, image_path: str) -> Optional[np.ndarray]:
        """載入圖片檔案"""
        try:
            # 嘗試多種路徑
            possible_paths = [
                image_path,
                os.path.join(os.getcwd(), image_path),
                os.path.join(os.path.dirname(image_path), os.path.basename(image_path))
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    print(f"[DEBUG] 載入圖片: {path}")
                    return cv2.imread(path)
            
            print(f"[WARNING] 找不到圖片檔案: {image_path}")
            return None
            
        except Exception as e:
            print(f"[ERROR] 載入圖片錯誤: {str(e)}")
            return None
    
    def preprocess_image(self, image: np.ndarray) -> np.ndarray:
        """圖片預處理，提高 OCR 準確率"""
        try:
            # 轉為灰度圖
            if len(image.shape) == 3:
                gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            else:
                gray = image
            
            # 調整大小（放大以提高清晰度）
            height, width = gray.shape
            scale_factor = 2
            resized = cv2.resize(gray, (width * scale_factor, height * scale_factor), interpolation=cv2.INTER_CUBIC)
            
            # 降噪
            denoised = cv2.medianBlur(resized, 3)
            
            # 二值化
            _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # 形態學操作（去除小雜點）
            kernel = np.ones((2, 2), np.uint8)
            cleaned = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
            
            return cleaned
            
        except Exception as e:
            print(f"[ERROR] 圖片預處理錯誤: {str(e)}")
            return image
    
    def extract_numbers_from_image(self, image: np.ndarray) -> List[Dict]:
        """從圖片中提取數字"""
        results = []
        
        try:
            # 預處理圖片
            processed_image = self.preprocess_image(image)
            
            # 使用 Tesseract OCR
            ocr_data = pytesseract.image_to_data(
                processed_image, 
                config=self.tesseract_config,
                output_type=pytesseract.Output.DICT
            )
            
            # 解析 OCR 結果
            for i in range(len(ocr_data['text'])):
                text = ocr_data['text'][i].strip()
                confidence = int(ocr_data['conf'][i])
                
                # 只保留數字和符合信心度要求的結果
                if text and confidence > self.min_confidence:
                    # 過濾出純數字
                    if self.is_numeric_text(text):
                        result = {
                            'text': text,
                            'confidence': confidence,
                            'bbox': (
                                ocr_data['left'][i],
                                ocr_data['top'][i],
                                ocr_data['left'][i] + ocr_data['width'][i],
                                ocr_data['top'][i] + ocr_data['height'][i]
                            )
                        }
                        results.append(result)
                        print(f"[DEBUG] 識別到數字: {text} (信心度: {confidence}%)")
        
        except Exception as e:
            print(f"[ERROR] OCR 識別錯誤: {str(e)}")
        
        return results
    
    def is_numeric_text(self, text: str) -> bool:
        """判斷文字是否為數字"""
        # 允許的字符：數字、小數點、負號
        allowed_chars = set('0123456789.-')
        return all(char in allowed_chars for char in text) and len(text) > 0
    
    def process_dxf_images(self, modelspace) -> List[Dict]:
        """處理 DXF 中的所有圖片實體"""
        all_results = []
        
        try:
            # 提取圖片實體
            images = self.extract_images_from_dxf(modelspace)
            
            for image_info in images:
                # 載入圖片
                image = self.load_image(image_info['path'])
                if image is None:
                    continue
                
                # 提取數字
                numbers = self.extract_numbers_from_image(image)
                
                # 將結果與圖片資訊結合
                for number in numbers:
                    result = {
                        'image_path': image_info['path'],
                        'insert_point': image_info['insert_point'],
                        'image_size': image_info['size'],
                        'rotation': image_info['rotation'],
                        'number_text': number['text'],
                        'confidence': number['confidence'],
                        'bbox': number['bbox']
                    }
                    all_results.append(result)
        
        except Exception as e:
            print(f"[ERROR] 處理 DXF 圖片錯誤: {str(e)}")
        
        return all_results
    
    def save_processed_image(self, image: np.ndarray, output_path: str):
        """儲存處理後的圖片（用於除錯）"""
        try:
            cv2.imwrite(output_path, image)
            print(f"[DEBUG] 處理後的圖片已儲存: {output_path}")
        except Exception as e:
            print(f"[ERROR] 儲存圖片錯誤: {str(e)}")

def create_image_ocr_processor():
    """創建圖片 OCR 處理器實例"""
    return ImageOCRProcessor() 