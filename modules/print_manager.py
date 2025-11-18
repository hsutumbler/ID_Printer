# 列印管理模組
import os
import tempfile
import datetime
import time
import subprocess
import platform
import configparser
import io
from .logger import logger

try:
    # 修正 reportlab 與新版 Python 的相容性問題
    import hashlib
    # 如果 hashlib 沒有 usedforsecurity 參數支援，添加相容性處理
    if not hasattr(hashlib, '_usedforsecurity_supported'):
        original_md5 = hashlib.md5
        def patched_md5(*args, **kwargs):
            kwargs.pop('usedforsecurity', None)
            return original_md5(*args, **kwargs)
        hashlib.md5 = patched_md5
    
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
    logger.info("ReportLab 載入成功，已修正相容性問題")
except ImportError:
    REPORTLAB_AVAILABLE = False
    logger.warning("ReportLab 未安裝，將使用簡單文字檔案列印")
except Exception as e:
    REPORTLAB_AVAILABLE = False
    logger.error(f"ReportLab 載入失敗: {e}，將使用簡單文字檔案列印")

# 嘗試匯入條碼生成套件
try:
    import barcode
    from barcode.writer import ImageWriter
    from PIL import Image, ImageDraw, ImageFont
    
    # 修正 Pillow 10.0+ 版本相容性：添加 getsize() 方法到 FreeTypeFont
    # 新版本 Pillow (10.0+) 移除了 getsize()，改用 getbbox() 或 textbbox()
    # python-barcode 套件仍在使用舊的 getsize() API
    try:
        # 定義 getsize() 修補方法
        def _patched_getsize(self, text, direction=None):
            """為 FreeTypeFont 添加 getsize() 方法相容性修補"""
            try:
                # 優先使用 getbbox() (Pillow 10.0+ 推薦方法)
                if hasattr(self, 'getbbox'):
                    bbox = self.getbbox(text)
                    if bbox and len(bbox) >= 4:
                        return (bbox[2] - bbox[0], bbox[3] - bbox[1])
                # 備用：使用 textbbox() (需要提供起始位置)
                elif hasattr(self, 'textbbox'):
                    bbox = self.textbbox((0, 0), text)
                    if bbox and len(bbox) >= 4:
                        return (bbox[2] - bbox[0], bbox[3] - bbox[1])
            except Exception as e:
                logger.debug(f"getsize() 修補方法執行失敗: {e}")
            return (0, 0)
        
        # 為 FreeTypeFont 類別添加 getsize 方法（如果不存在）
        if hasattr(ImageFont, 'FreeTypeFont'):
            if not hasattr(ImageFont.FreeTypeFont, 'getsize'):
                ImageFont.FreeTypeFont.getsize = _patched_getsize
                logger.debug("已為 FreeTypeFont 添加 getsize() 相容性修補")
        
        # 也為 ImageFont 基類添加（如果存在）
        if hasattr(ImageFont, 'ImageFont') and not hasattr(ImageFont.ImageFont, 'getsize'):
            ImageFont.ImageFont.getsize = _patched_getsize
            logger.debug("已為 ImageFont 基類添加 getsize() 相容性修補")
            
    except Exception as e:
        logger.warning(f"字型相容性修補失敗: {e}")
    
    # 強制匯入條碼子模組，確保 PyInstaller 打包時包含它們
    # 使用正確的匯入方式：barcode.code128 而不是 from barcode import code128
    try:
        import barcode.code128
        import barcode.code39
        import barcode.ean13
        import barcode.ean8
        from barcode.writer import base
        from barcode.writer import image as writer_image
        from barcode.writer import svg
        logger.debug("條碼子模組匯入成功")
    except ImportError as e:
        logger.warning(f"部分條碼子模組匯入失敗（可能不影響功能）: {e}")
    
    BARCODE_AVAILABLE = True
    IMAGE_AVAILABLE = True
except ImportError:
    BARCODE_AVAILABLE = False
    IMAGE_AVAILABLE = False
    logger.warning("python-barcode 未安裝，將使用文字模擬條碼")

class PrintManagerError(Exception):
    """列印管理錯誤"""
    pass

class PrintManager:
    def __init__(self):
        # 讀取配置檔案
        self.config = configparser.ConfigParser()
        self.config.read('config.ini', encoding='utf-8')
        
        # 條碼設定
        self.use_barcode = self.config.getboolean('標籤設定', 'use_barcode', fallback=True)
        self.barcode_type = self.config.get('標籤設定', 'barcode_type', fallback='code128')
        
        # 從配置檔案獲取標籤尺寸
        label_width_mm = float(self.config.get('標籤設定', 'label_width', fallback='50'))
        label_height_mm = float(self.config.get('標籤設定', 'label_height', fallback='35'))
        
        # 標籤尺寸設定
        if REPORTLAB_AVAILABLE:
            self.label_width = label_width_mm * mm
            self.label_height = label_height_mm * mm
        else:
            self.label_width = label_width_mm
            self.label_height = label_height_mm
            
        # 印表機設定
        self.label_width_mm = label_width_mm
        self.label_height_mm = label_height_mm
        self.print_mode = self.config.get('印表機設定', 'print_mode', fallback='pdf')
        self.use_default_printer = self.config.getboolean('印表機設定', 'use_default_printer', fallback=True)
        self.show_print_dialog = self.config.getboolean('印表機設定', 'show_print_dialog', fallback=True)
        
        # 嘗試註冊中文字體 (如果有的話)
        self.font_registered = False
        if REPORTLAB_AVAILABLE:
            self._try_register_chinese_font()
        
        # ZPL 圖形模式設定
        self.zpl_chinese_font_path = None
        self.zpl_chinese_font_size = 22  # 點數，與 ZPL 字型大小對應
        self.zpl_fixed_graphics = {}  # 儲存固定文字的預定義圖形
        if IMAGE_AVAILABLE:
            self._load_chinese_font_for_zpl()
            self._generate_fixed_graphics()  # 生成固定文字的預定義圖形
        
        # 預先載入條碼類別，確保 PyInstaller 打包時包含必要的模組
        if BARCODE_AVAILABLE and self.use_barcode:
            try:
                # 預先載入常用的條碼類別，確保打包時包含
                _ = barcode.get_barcode_class('code128')
                _ = barcode.get_barcode_class('code39')
                logger.debug("條碼類別預載入成功")
            except Exception as e:
                logger.warning(f"條碼類別預載入失敗: {e}")
        
        logger.info(f"列印管理模組初始化完成 - 標籤尺寸: {label_width_mm}x{label_height_mm}mm, 列印模式: {self.print_mode}")

    def _try_register_chinese_font(self):
        """註冊微軟正黑體供 PDF 模式使用（統一使用微軟正黑體）"""
        try:
            # 優先使用標準微軟正黑體，如果找不到則嘗試粗體版本
            font_paths = [
                "C:/Windows/Fonts/msjh.ttc",   # 微軟正黑體（標準版）
                "C:/Windows/Fonts/msjhbd.ttc", # 微軟正黑體（粗體版，備用）
            ]
            
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                        self.font_registered = True
                        logger.info(f"成功註冊字型: {font_path} (微軟正黑體)")
                        return
                    except Exception as e:
                        logger.debug(f"無法註冊字型 {font_path}: {e}")
                        continue
            
            # 跨平台備用方案（如果 Windows 上找不到）
            if platform.system() == "Darwin":  # macOS
                mac_fonts = ["/System/Library/Fonts/PingFang.ttc"]
                for font_path in mac_fonts:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                            self.font_registered = True
                            logger.info(f"成功註冊字型: {font_path} (macOS 備用)")
                            return
                        except:
                            continue
            elif platform.system() == "Linux":
                linux_fonts = ["/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"]
                for font_path in linux_fonts:
                    if os.path.exists(font_path):
                        try:
                            pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                            self.font_registered = True
                            logger.info(f"成功註冊字型: {font_path} (Linux 備用)")
                            return
                        except:
                            continue
            
            if not self.font_registered:
                logger.warning("未找到微軟正黑體字型，PDF 模式將使用預設字體")
                logger.warning("請確認系統中已安裝微軟正黑體 (msjh.ttc 或 msjhbd.ttc)")
                
        except Exception as e:
            logger.warning(f"註冊微軟正黑體失敗: {e}")
    
    def _load_chinese_font_for_zpl(self):
        """載入微軟正黑體供 ZPL 圖形模式使用（統一使用微軟正黑體）"""
        try:
            # 優先使用標準微軟正黑體，如果找不到則嘗試粗體版本
            font_paths = [
                "C:/Windows/Fonts/msjh.ttc",   # 微軟正黑體（標準版）
                "C:/Windows/Fonts/msjhbd.ttc", # 微軟正黑體（粗體版，備用）
            ]
            
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        # 測試是否可以載入字型
                        test_font = ImageFont.truetype(font_path, self.zpl_chinese_font_size)
                        self.zpl_chinese_font_path = font_path
                        logger.info(f"ZPL 圖形模式使用字型: {font_path} (微軟正黑體)")
                        return
                    except Exception as e:
                        logger.debug(f"無法載入字型 {font_path}: {e}")
                        continue
            
            logger.error("未找到微軟正黑體字型，ZPL 圖形模式可能無法正確顯示文字")
            logger.error("請確認系統中已安裝微軟正黑體 (msjh.ttc 或 msjhbd.ttc)")
                
        except Exception as e:
            logger.error(f"載入微軟正黑體失敗: {e}")
            logger.error("ZPL 圖形模式可能無法正確顯示文字")
    
    def _generate_fixed_graphics(self):
        """生成固定文字的預定義圖形（使用微軟正黑體）"""
        if not IMAGE_AVAILABLE or not self.zpl_chinese_font_path:
            logger.warning("無法生成固定圖形：缺少 PIL 或微軟正黑體字型")
            return
        
        try:
            # 固定文字的列表（使用半形冒號）
            fixed_texts = {
                "CHART_NO": "病歷號:",
                "ID": "身分證字號:",
                "NAME": "姓名:",
                "BIRTH": "生日:",
                "TIME": "列印時間:",
                "NOTE": "備註:"
            }
            
            for key, text in fixed_texts.items():
                zpl_graphic = self._text_to_zpl_graphic(text, f"ITEM_{key}")
                if zpl_graphic:
                    self.zpl_fixed_graphics[key] = zpl_graphic
                    logger.debug(f"生成固定圖形: {key} = {text}")
            
            logger.info(f"成功生成 {len(self.zpl_fixed_graphics)} 個固定圖形")
            
        except Exception as e:
            logger.error(f"生成固定圖形失敗: {e}")
    
    def _text_to_zpl_graphic(self, text, item_name):
        """將文字轉換為 ZPL 圖形格式（~DGR 指令）
        
        Args:
            text: 要轉換的文字
            item_name: ZPL 圖形項目名稱（如 ITEM_NAME）
        
        Returns:
            ZPL 圖形指令字串（包含 ~DGR 定義），如果失敗則返回 None
        """
        if not IMAGE_AVAILABLE or not self.zpl_chinese_font_path:
            return None
        
        try:
            # 載入字型
            font = ImageFont.truetype(self.zpl_chinese_font_path, self.zpl_chinese_font_size)
            
            # 計算文字尺寸
            # 使用較大的臨時圖片來測量文字大小（包含上升和下降部分）
            temp_img = Image.new('RGB', (1000, 200), (255, 255, 255))
            temp_draw = ImageDraw.Draw(temp_img)
            bbox = temp_draw.textbbox((0, 0), text, font=font)
            
            # 邊界框：left, top, right, bottom
            # top 可能是負數（上升部分，如大寫字母），bottom 是正數
            text_left = bbox[0]
            text_top = bbox[1]  # 可能是負數
            text_right = bbox[2]
            text_bottom = bbox[3]
            
            text_width = text_right - text_left
            text_height = text_bottom - text_top
            
            # 增加邊距（從4增加到6，避免邊緣裁切）
            padding = 6
            img_width = text_width + padding * 2
            # 圖形高度需要包含完整的文字高度（包括上升和下降部分）
            img_height = text_height + padding * 2
            
            # 建立 1-bit 黑白圖片
            img = Image.new('1', (img_width, img_height), 1)  # 1 = 白色
            draw = ImageDraw.Draw(img)
            
            # 計算文字繪製的 Y 座標
            # 因為 text_top 可能是負數，需要調整 Y 座標
            # 將文字向上移動 |text_top| 的距離，再加上 padding
            # 這樣可以確保文字完整顯示，不會被裁切
            text_y = padding - text_top  # 這樣可以確保文字完整顯示
            
            # 繪製文字（0 = 黑色）
            draw.text((padding, text_y), text, font=font, fill=0)
            
            # 轉換為 ZPL HEX 格式
            # ZPL 使用 ASCII 85 編碼或壓縮的 HEX
            # 這裡使用壓縮的 HEX（ZPL 標準格式）
            hex_data = self._image_to_zpl_hex(img)
            
            if not hex_data:
                return None
            
            # 計算總位元組數和每行列數
            bytes_per_row = (img_width + 7) // 8  # 每行需要幾個位元組（8位=1位元組）
            total_bytes = bytes_per_row * img_height
            
            # 生成 ~DGR 指令
            # 格式：~DGR:名稱,總位元組數,每行列數,壓縮的HEX資料
            zpl_command = f"~DGR:{item_name},{total_bytes:05d},{bytes_per_row:03d},{hex_data}"
            
            return zpl_command
            
        except Exception as e:
            logger.error(f"將文字轉換為 ZPL 圖形失敗: {text}, 錯誤: {e}")
            return None
    
    def _image_to_zpl_hex(self, img):
        """將 PIL 圖片轉換為 ZPL HEX 格式
        
        ZPL ~DGR 指令使用的格式：
        - 每個位元組代表 8 個像素（水平方向，從左到右）
        - 位元 7（最高位）對應最左側的像素
        - 1 = 黑色，0 = 白色
        
        PIL '1' 模式：
        - 0 = 黑色，1 = 白色
        - 需要反轉
        """
        try:
            width, height = img.size
            bytes_per_row = (width + 7) // 8
            
            # 取得圖片的像素數據
            pixels = img.load()
            
            # 轉換為位元組陣列
            hex_chars = []
            for y in range(height):
                for x_byte in range(bytes_per_row):
                    byte_value = 0
                    for bit in range(8):
                        pixel_x = x_byte * 8 + bit
                        if pixel_x < width:
                            # PIL 的 '1' 模式：0=黑色, 1=白色
                            # ZPL 需要：0=白色, 1=黑色
                            # 所以需要反轉：如果 PIL 是黑色(0)，ZPL 設為 1
                            pixel_value = pixels[pixel_x, y]
                            if pixel_value == 0:  # PIL 黑色像素
                                # ZPL 設為 1（黑色）
                                # 位元順序：最高位（bit 7）對應最左側像素
                                byte_value |= (1 << (7 - bit))
                            # 如果 pixel_value == 1（PIL 白色），ZPL 設為 0（白色），不需要操作
                    
                    # 轉換為 HEX（兩位數，大寫）
                    hex_chars.append(f"{byte_value:02X}")
            
            # 合併所有 HEX 字元
            hex_string = ''.join(hex_chars)
            
            return hex_string
            
        except Exception as e:
            logger.error(f"圖片轉換為 ZPL HEX 失敗: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None

    def generate_barcode(self, patient_id):
        """生成條碼圖片"""
        if not BARCODE_AVAILABLE or not self.use_barcode:
            logger.debug("條碼功能不可用或已禁用")
            return None
            
        try:
            # 移除非數字和字母的字符
            clean_id = ''.join(c for c in patient_id if c.isalnum())
            if not clean_id:
                logger.warning(f"無效的身分證字號，無法生成條碼: {patient_id}")
                return None
            
            logger.debug(f"開始生成條碼: {clean_id}, 類型: {self.barcode_type}")
            
            # 創建條碼
            barcode_class = barcode.get_barcode_class(self.barcode_type)
            if barcode_class is None:
                logger.error(f"無法找到條碼類型: {self.barcode_type}")
                return None
            
            # 使用 ImageWriter 生成條碼
            # 注意：在打包環境中，可能需要確保所有依賴都正確載入
            writer = ImageWriter()
            barcode_instance = barcode_class(clean_id, writer=writer)
            
            # 生成條碼圖片到內存
            buffer = io.BytesIO()
            barcode_instance.write(buffer)
            
            # 檢查緩衝區是否有資料
            if buffer.tell() == 0:
                logger.error("條碼緩衝區為空，生成失敗")
                return None
            
            # 將緩衝區轉換為圖片對象
            buffer.seek(0)
            barcode_image = Image.open(buffer)
            
            # 轉換為 RGB 模式（確保與 ReportLab 相容）
            if barcode_image.mode != 'RGB':
                barcode_image = barcode_image.convert('RGB')
            
            logger.debug(f"條碼生成成功: {clean_id}, 尺寸: {barcode_image.size}")
            return barcode_image
        
        except ImportError as e:
            logger.error(f"條碼模組匯入失敗: {e}")
            logger.error("這可能是打包後的問題，請檢查 PyInstaller 隱藏匯入設定")
            logger.error(f"請確認 ID_Printer.spec 中包含所有必要的 barcode 子模組")
            import traceback
            logger.error(traceback.format_exc())
            return None
        except AttributeError as e:
            logger.error(f"條碼模組屬性錯誤: {e}")
            logger.error("可能是動態載入的模組在打包後無法正確載入")
            import traceback
            logger.error(traceback.format_exc())
            return None
        except Exception as e:
            logger.error(f"生成條碼失敗: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return None

    def _generate_label_pdf(self, patient_data, filename):
        """生成標籤 PDF (5cm x 3.5cm)"""
        if not REPORTLAB_AVAILABLE:
            raise PrintManagerError("ReportLab 未安裝，無法生成 PDF")
            
        try:
            logger.info(f"開始生成 PDF 標籤: {filename}")
            c = canvas.Canvas(filename, pagesize=(self.label_width, self.label_height))
            
            # 獲取當前列印時間
            print_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
            
            # 調整字體大小和排版以適應小標籤 (5cm x 3.5cm)
            font_size = 8  # 縮小字體以適應小標籤
            line_height = 4 * mm  # 縮小行距
            start_x = 2 * mm  # 縮小左邊距
            start_y = self.label_height - 4 * mm  # 縮小上邊距
            
            # 設定字體
            if self.font_registered:
                c.setFont('ChineseFont', font_size)
            else:
                c.setFont('Helvetica', font_size)

            # 標籤內容 - 按照要求的排版
            y_pos = start_y
            
            # 獲取病歷號和身分證字號
            chart_no = patient_data.get('chart_no', '').strip()
            patient_id = patient_data.get('id', 'N/A')
            
            # 生成條碼（優先使用病歷號，如果沒有病歷號則使用身分證字號）
            # 確保病歷號不為空且不是 'N/A'，才使用病歷號
            if chart_no and chart_no != 'N/A':
                barcode_value = chart_no
            else:
                barcode_value = patient_id if patient_id and patient_id != 'N/A' else 'N/A'
            if self.use_barcode and BARCODE_AVAILABLE:
                logger.info(f"開始生成條碼: barcode_value={barcode_value}, BARCODE_AVAILABLE={BARCODE_AVAILABLE}, use_barcode={self.use_barcode}")
                barcode_image = self.generate_barcode(barcode_value)
                
                if barcode_image:
                    try:
                        # 改用 ReportLab 的 ImageReader 直接從內存讀取，避免打包環境中臨時檔案路徑問題
                        from reportlab.lib.utils import ImageReader
                        
                        # 將 PIL Image 轉換為 ReportLab 可用的格式
                        # 先將圖片儲存到 BytesIO，然後用 ImageReader 讀取
                        img_buffer = io.BytesIO()
                        barcode_image.save(img_buffer, format='PNG')
                        img_buffer.seek(0)
                        
                        # 使用 ImageReader 從內存讀取圖片（避免臨時檔案問題）
                        img_reader = ImageReader(img_buffer)
                        
                        # 計算條碼尺寸和位置
                        barcode_width = self.label_width - 4 * mm  # 留出左右邊距
                        barcode_height = 10 * mm  # 條碼高度
                        
                        # 修正：條碼從標籤頂部開始繪製
                        # ReportLab 座標系統：Y 座標是圖片的左下角
                        # 如果要讓條碼貼齊標籤頂部，Y 座標應該是標籤高度減去條碼高度，再減去頂部邊距
                        barcode_y = self.label_height - barcode_height - 2 * mm  # 頂部留2mm邊距
                        
                        # 在 PDF 中插入條碼圖片（直接從內存讀取，避免臨時檔案問題）
                        logger.info(f"插入條碼: 位置=({start_x}, {barcode_y}), 尺寸=({barcode_width}, {barcode_height}), 圖片尺寸={barcode_image.size}")
                        c.drawImage(img_reader, start_x, barcode_y, width=barcode_width, height=barcode_height)
                        logger.info("條碼圖片已成功插入 PDF（使用內存圖片，避免臨時檔案問題）")
                        
                        # 移動到條碼下方（條碼本身帶一行文字，再加上一行間距）
                        y_pos = barcode_y - font_size - 2.5 * mm
                        
                    except Exception as e:
                        logger.error(f"處理條碼圖片失敗: {e}")
                        import traceback
                        logger.error(traceback.format_exc())
                        # 即使條碼失敗，也繼續生成 PDF
                else:
                    logger.warning("條碼生成返回 None，跳過條碼顯示")
                    # 診斷條碼生成失敗的原因
                    if not BARCODE_AVAILABLE:
                        logger.error("條碼模組不可用（可能在打包後環境中無法載入）")
                    elif not self.use_barcode:
                        logger.debug("條碼功能已在配置中禁用")
                    else:
                        logger.error("條碼生成失敗，但未拋出異常（可能是模組載入問題，請檢查 PyInstaller 隱藏導入設定）")
            
            # 設定統一的行距
            uniform_spacing = line_height + 0.5 * mm
            
            # 病歷號和身分證字號同行
            c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
            if chart_no:
                chart_and_id = f"病歷號：{chart_no}   {patient_id}"
            else:
                chart_and_id = f"病歷號：   {patient_id}"
            c.drawString(start_x, y_pos, chart_and_id)
            y_pos -= uniform_spacing
            
            # 姓名與生日同行
            c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
            name_and_dob = f"姓名：{patient_data.get('name', 'N/A')}    {patient_data.get('dob', 'N/A')}"
            c.drawString(start_x, y_pos, name_and_dob)
            y_pos -= uniform_spacing
            
            # 列印時間（格式：YYYY/MM/DD   HH:MM）
            c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
            # 將時間格式改為日期和時間分開顯示
            time_parts = print_time.split(' ')
            if len(time_parts) == 2:
                date_part, time_part = time_parts
                time_display = f"列印時間：{date_part}   {time_part}"
            else:
                time_display = f"列印時間：{print_time}"
            c.drawString(start_x, y_pos, time_display)
            y_pos -= uniform_spacing
            
            # 備註 - 只在有內容時顯示
            note = patient_data.get('note', '').strip()
            if note:
                c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
                c.drawString(start_x, y_pos, f"備註：{note}")
                y_pos -= uniform_spacing

            c.showPage()
            c.save()
            logger.info(f"生成 PDF 標籤: {filename} ({self.label_width_mm}x{self.label_height_mm}mm)")
            return filename
            
        except Exception as e:
            logger.error(f"生成 PDF 失敗: {e}")
            raise PrintManagerError(f"生成 PDF 失敗: {e}")

    def _generate_label_text(self, patient_data, filename):
        """生成文字標籤 (備用方案)"""
        try:
            print_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
            patient_id = patient_data.get('id', 'N/A')
            chart_no = patient_data.get('chart_no', '').strip()
            
            # 條碼值（優先使用病歷號，如果沒有病歷號則使用身分證字號）
            # 確保病歷號不為空且不是 'N/A'，才使用病歷號
            if chart_no and chart_no != 'N/A':
                barcode_value = chart_no
            else:
                barcode_value = patient_id if patient_id and patient_id != 'N/A' else 'N/A'
            
            # 使用 ASCII 字符模擬條碼
            barcode_ascii = "|||||||||||  ||||||||||||  ||||||" if self.use_barcode else ""
            barcode_line = f"{barcode_ascii} (條碼-{barcode_value})\n" if self.use_barcode else ""
            
            # 病歷號和身分證字號同行
            if chart_no:
                chart_and_id_line = f"病歷號：{chart_no}   {patient_id}\n"
            else:
                chart_and_id_line = f"病歷號：   {patient_id}\n"
            
            # 列印時間（格式：YYYY/MM/DD   HH:MM）
            time_parts = print_time.split(' ')
            if len(time_parts) == 2:
                date_part, time_part = time_parts
                time_line = f"列印時間：{date_part}   {time_part}\n"
            else:
                time_line = f"列印時間：{print_time}\n"
            
            # 備註 - 只在有內容時顯示
            note = patient_data.get('note', '').strip()
            note_line = f"備註：{note}\n" if note else ""
            
            content = f"""
==========================================
{barcode_line}{chart_and_id_line}姓名：{patient_data.get('name', 'N/A')}    {patient_data.get('dob', 'N/A')}
{time_line}{note_line}==========================================
"""
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            
            logger.info(f"生成文字標籤: {filename}")
            return filename
            
        except Exception as e:
            logger.error(f"生成文字標籤失敗: {e}")
            raise PrintManagerError(f"生成文字標籤失敗: {e}")

    def _generate_zpl_content(self, patient_data):
        """生成 ZPL 內容 (不寫入檔案) - 統一使用圖形模式（固定圖形 + 動態圖形）"""
        try:
            print_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
            patient_id = patient_data.get('id', 'N/A')
            patient_name = patient_data.get('name', 'N/A')
            dob = patient_data.get('dob', 'N/A')
            note = patient_data.get('note', '').strip()
            chart_no = patient_data.get('chart_no', '').strip()
            
            # 條碼值（優先使用病歷號，如果沒有病歷號則使用身分證字號）
            # 確保病歷號不為空且不是 'N/A'，才使用病歷號
            if chart_no and chart_no != 'N/A':
                barcode_value = chart_no
            else:
                barcode_value = patient_id if patient_id and patient_id != 'N/A' else 'N/A'
            
            # ZPL 指令開始
            zpl_content = "^XA\n"  # 開始標籤格式
            
            # 設定標籤尺寸 (以點為單位，203 DPI)
            # 50mm = 394點, 35mm = 276點
            zpl_content += "^PW394\n"  # 設定標籤寬度
            zpl_content += "^LL276\n"  # 設定標籤長度
            
            # ========== 統一圖形模式 ==========
            # 1. 先定義固定文字的預定義圖形（在初始化時已生成）
            for key, graphic_def in self.zpl_fixed_graphics.items():
                zpl_content += graphic_def + "\n"
            
            # 2. 動態生成所有變動文字的圖形（統一使用圖形模式，確保字型一致）
            dynamic_graphics = {}  # 儲存動態生成的圖形定義和顯示指令
            
            # 生成身分證字號的圖形
            if patient_id and patient_id != 'N/A':
                item_name = f"ITEM_ID_DYN_{hash(patient_id) % 10000}"
                id_graphic = self._text_to_zpl_graphic(patient_id, item_name)
                if id_graphic:
                    zpl_content += id_graphic + "\n"
                    dynamic_graphics['patient_id'] = item_name
            
            # 生成病人姓名的圖形
            if patient_name and patient_name != 'N/A':
                item_name = f"ITEM_NAME_DYN_{hash(patient_name) % 10000}"
                name_graphic = self._text_to_zpl_graphic(patient_name, item_name)
                if name_graphic:
                    zpl_content += name_graphic + "\n"
                    dynamic_graphics['patient_name'] = item_name
            
            # 生成生日的圖形
            if dob and dob != 'N/A':
                item_name = f"ITEM_DOB_DYN_{hash(dob) % 10000}"
                dob_graphic = self._text_to_zpl_graphic(dob, item_name)
                if dob_graphic:
                    zpl_content += dob_graphic + "\n"
                    dynamic_graphics['dob'] = item_name
            
            # 生成列印時間的圖形
            if print_time:
                item_name = f"ITEM_TIME_DYN_{hash(print_time) % 10000}"
                time_graphic = self._text_to_zpl_graphic(print_time, item_name)
                if time_graphic:
                    zpl_content += time_graphic + "\n"
                    dynamic_graphics['print_time'] = item_name
            
            # 生成病歷號的圖形
            chart_no = patient_data.get('chart_no', '').strip()
            if chart_no:
                item_name = f"ITEM_CHART_NO_DYN_{hash(chart_no) % 10000}"
                chart_no_graphic = self._text_to_zpl_graphic(chart_no, item_name)
                if chart_no_graphic:
                    zpl_content += chart_no_graphic + "\n"
                    dynamic_graphics['chart_no'] = item_name
            
            # 生成備註的圖形
            if note:
                item_name = f"ITEM_NOTE_DYN_{hash(note) % 10000}"
                note_graphic = self._text_to_zpl_graphic(note, item_name)
                if note_graphic:
                    zpl_content += note_graphic + "\n"
                    dynamic_graphics['note'] = item_name
            
            y_pos = 30  # 起始Y位置
            
            # 條碼 (如果啟用，使用病歷號或身分證字號)
            if self.use_barcode:
                # Code 128 條碼，高度50點
                zpl_content += f"^FO30,{y_pos}^BY2^BCN,50,Y,N,N^FD{barcode_value}^FS\n"
                y_pos += 70 + 12  # 條碼後移動位置（條碼本身帶一行文字，再加一行間距約12點）
            
            # 病歷號和身分證字號同行
            if "CHART_NO" in self.zpl_fixed_graphics:
                # 顯示「病歷號:」標籤
                zpl_content += f"^FO30,{y_pos}^XGITEM_CHART_NO^FS\n"
                chart_no_label_width = 60  # 「病歷號:」的寬度約60點
                
                # 顯示病歷號值（如果有）
                if chart_no:
                    if 'chart_no' in dynamic_graphics:
                        zpl_content += f"^FO{30 + chart_no_label_width},{y_pos}^XG{dynamic_graphics['chart_no']}^FS\n"
                    else:
                        zpl_content += f"^FO{30 + chart_no_label_width},{y_pos}^A0N,22,22^FD{chart_no}^FS\n"
                    # 病歷號後留空格，然後顯示身分證字號
                    chart_no_value_width = len(chart_no) * 12  # 估算病歷號寬度
                    id_x = 30 + chart_no_label_width + chart_no_value_width + 20  # 加20點間距
                else:
                    # 沒有病歷號，直接顯示身分證字號
                    id_x = 30 + chart_no_label_width + 20
                
                # 顯示身分證字號（不使用標籤，直接顯示值）
                if 'patient_id' in dynamic_graphics:
                    zpl_content += f"^FO{id_x},{y_pos}^XG{dynamic_graphics['patient_id']}^FS\n"
                else:
                    zpl_content += f"^FO{id_x},{y_pos}^A0N,22,22^FD{patient_id}^FS\n"
            else:
                # 如果沒有固定圖形，使用動態圖形或字型
                if chart_no:
                    chart_id_text = f"病歷號：{chart_no}   {patient_id}"
                else:
                    chart_id_text = f"病歷號：   {patient_id}"
                # 生成整行文字的圖形
                item_name = f"ITEM_CHART_ID_DYN_{hash(chart_id_text) % 10000}"
                chart_id_graphic = self._text_to_zpl_graphic(chart_id_text, item_name)
                if chart_id_graphic:
                    zpl_content += chart_id_graphic + "\n"
                    zpl_content += f"^FO30,{y_pos}^XG{item_name}^FS\n"
                else:
                    zpl_content += f"^FO30,{y_pos}^A0N,22,22^FD{chart_id_text}^FS\n"
            y_pos += 40
            
            # 姓名與生日
            if "NAME" in self.zpl_fixed_graphics:
                # 使用固定圖形顯示「姓名:」
                zpl_content += f"^FO30,{y_pos}^XGITEM_NAME^FS\n"
                name_label_width = 50  # 「姓名:」的寬度約50點
                
                # 使用動態圖形顯示姓名
                if 'patient_name' in dynamic_graphics:
                    zpl_content += f"^FO{30 + name_label_width},{y_pos}^XG{dynamic_graphics['patient_name']}^FS\n"
                else:
                    # 如果圖形生成失敗，使用備用方案（字型）
                    zpl_content += f"^FO{30 + name_label_width},{y_pos}^A0N,22,22^FD{patient_name}^FS\n"
                
                # 在同行的右側顯示生日
                if "BIRTH" in self.zpl_fixed_graphics:
                    birth_x = 200  # 生日標籤的 X 位置
                    zpl_content += f"^FO{birth_x},{y_pos}^XGITEM_BIRTH^FS\n"
                    birth_label_width = 50
                    # 使用動態圖形顯示生日
                    if 'dob' in dynamic_graphics:
                        zpl_content += f"^FO{birth_x + birth_label_width},{y_pos}^XG{dynamic_graphics['dob']}^FS\n"
                    else:
                        # 如果圖形生成失敗，使用備用方案（字型）
                        zpl_content += f"^FO{birth_x + birth_label_width},{y_pos}^A0N,22,22^FD{dob}^FS\n"
                else:
                    # 如果沒有固定圖形，直接顯示動態圖形
                    if 'dob' in dynamic_graphics:
                        zpl_content += f"^FO200,{y_pos}^XG{dynamic_graphics['dob']}^FS\n"
                    else:
                        zpl_content += f"^FO200,{y_pos}^A0N,22,22^FD{dob}^FS\n"
            else:
                # 如果沒有固定圖形，使用動態圖形
                if 'patient_name' in dynamic_graphics:
                    zpl_content += f"^FO30,{y_pos}^XG{dynamic_graphics['patient_name']}^FS\n"
                else:
                    zpl_content += f"^FO30,{y_pos}^A0N,22,22^FDName: {patient_name}^FS\n"
                # 顯示生日
                if 'dob' in dynamic_graphics:
                    zpl_content += f"^FO200,{y_pos}^XG{dynamic_graphics['dob']}^FS\n"
                else:
                    zpl_content += f"^FO200,{y_pos}^A0N,22,22^FDDOB: {dob}^FS\n"
            y_pos += 40
            
            # 列印時間（格式：YYYY/MM/DD   HH:MM）
            time_parts = print_time.split(' ')
            if len(time_parts) == 2:
                date_part, time_part = time_parts
                time_display = f"{date_part}   {time_part}"
            else:
                time_display = print_time
            
            # 生成新的時間顯示格式的圖形
            if time_display != print_time:
                item_name = f"ITEM_TIME_DYN_{hash(time_display) % 10000}"
                time_graphic = self._text_to_zpl_graphic(time_display, item_name)
                if time_graphic:
                    zpl_content += time_graphic + "\n"
                    dynamic_graphics['print_time_display'] = item_name
            
            if "TIME" in self.zpl_fixed_graphics:
                zpl_content += f"^FO30,{y_pos}^XGITEM_TIME^FS\n"
                time_label_width = 80  # 「列印時間:」的寬度約80點
                # 時間值往右移動
                time_value_x = 30 + time_label_width + 22
                # 使用動態圖形顯示列印時間
                if 'print_time_display' in dynamic_graphics:
                    zpl_content += f"^FO{time_value_x},{y_pos}^XG{dynamic_graphics['print_time_display']}^FS\n"
                elif 'print_time' in dynamic_graphics:
                    zpl_content += f"^FO{time_value_x},{y_pos}^XG{dynamic_graphics['print_time']}^FS\n"
                else:
                    # 如果圖形生成失敗，使用備用方案（字型）
                    zpl_content += f"^FO{time_value_x},{y_pos}^A0N,22,22^FD{time_display}^FS\n"
            else:
                # 如果沒有固定圖形，使用動態圖形
                if 'print_time_display' in dynamic_graphics:
                    zpl_content += f"^FO30,{y_pos}^XG{dynamic_graphics['print_time_display']}^FS\n"
                elif 'print_time' in dynamic_graphics:
                    zpl_content += f"^FO30,{y_pos}^XG{dynamic_graphics['print_time']}^FS\n"
                else:
                    zpl_content += f"^FO30,{y_pos}^A0N,22,22^FD列印時間: {time_display}^FS\n"
            y_pos += 40
            
            # 備註 (如果有)
            if note:
                if "NOTE" in self.zpl_fixed_graphics:
                    zpl_content += f"^FO30,{y_pos}^XGITEM_NOTE^FS\n"
                    note_label_width = 50  # 「備註:」的寬度約50點
                    # 使用動態圖形顯示備註
                    if 'note' in dynamic_graphics:
                        zpl_content += f"^FO{30 + note_label_width},{y_pos}^XG{dynamic_graphics['note']}^FS\n"
                    else:
                        # 如果圖形生成失敗，使用備用方案（字型）
                        zpl_content += f"^FO{30 + note_label_width},{y_pos}^A0N,22,22^FD{note}^FS\n"
                else:
                    # 如果沒有固定圖形，使用動態圖形
                    if 'note' in dynamic_graphics:
                        zpl_content += f"^FO30,{y_pos}^XG{dynamic_graphics['note']}^FS\n"
                    else:
                        zpl_content += f"^FO30,{y_pos}^A0N,22,22^FDNote: {note}^FS\n"
            
            # ZPL 指令結束
            zpl_content += "^XZ\n"
            
            return zpl_content
            
        except Exception as e:
            logger.error(f"生成 ZPL 內容失敗: {e}")
            raise PrintManagerError(f"生成 ZPL 內容失敗: {e}")

    def _generate_label_zpl(self, patient_data, filename):
        """生成 ZPL 標籤檔案 (向後相容)"""
        try:
            zpl_content = self._generate_zpl_content(patient_data)
            
            # 寫入 ZPL 檔案
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(zpl_content)
            
            logger.info(f"生成 ZPL 標籤: {filename}")
            return filename
            
        except Exception as e:
            logger.error(f"生成 ZPL 標籤失敗: {e}")
            raise PrintManagerError(f"生成 ZPL 標籤失敗: {e}")

    def _send_zpl_to_printer(self, zpl_content):
        """直接發送 ZPL 內容到 Zebra 印表機"""
        try:
            system = platform.system()
            
            if system == "Windows":
                # Windows 系統：使用 win32print 發送原始 ZPL 資料到印表機
                try:
                    import win32print
                    import win32api
                    
                    # 取得預設印表機
                    printer_name = win32print.GetDefaultPrinter()
                    logger.info(f"使用預設印表機: {printer_name}")
                    
                    # 開啟印表機
                    printer_handle = win32print.OpenPrinter(printer_name)
                    
                    try:
                        # 開始列印工作
                        job_info = ("ZPL Label", None, "RAW")
                        job_id = win32print.StartDocPrinter(printer_handle, 1, job_info)
                        win32print.StartPagePrinter(printer_handle)
                        
                        # 發送 ZPL 資料
                        win32print.WritePrinter(printer_handle, zpl_content.encode('utf-8'))
                        
                        # 結束列印
                        win32print.EndPagePrinter(printer_handle)
                        win32print.EndDocPrinter(printer_handle)
                        
                        logger.info(f"ZPL 已直接發送到印表機: {printer_name}")
                        
                    finally:
                        win32print.ClosePrinter(printer_handle)
                        
                except ImportError:
                    logger.warning("win32print 未安裝，改用檔案方式列印")
                    # 備用方案：建立臨時檔案並發送到印表機
                    temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.zpl', delete=False, encoding='utf-8')
                    temp_file.write(zpl_content)
                    temp_file.close()
                    
                    os.startfile(temp_file.name, "print")
                    logger.info(f"ZPL 透過檔案發送到印表機: {temp_file.name}")
                    
                    # 延遲刪除臨時檔案
                    import threading
                    def cleanup():
                        import time
                        time.sleep(5)
                        try:
                            os.remove(temp_file.name)
                        except:
                            pass
                    threading.Thread(target=cleanup).start()
                    
                except Exception as e:
                    logger.error(f"直接列印 ZPL 失敗: {e}")
                    raise PrintManagerError(f"直接列印 ZPL 失敗: {e}")
            else:
                # 非 Windows 系統：使用 lpr 命令
                try:
                    import subprocess
                    process = subprocess.Popen(['lpr', '-o', 'raw'], stdin=subprocess.PIPE)
                    process.communicate(input=zpl_content.encode('utf-8'))
                    logger.info("ZPL 已透過 lpr 發送到印表機")
                except Exception as e:
                    logger.error(f"lpr 列印失敗: {e}")
                    raise PrintManagerError(f"lpr 列印失敗: {e}")
                    
        except Exception as e:
            logger.error(f"發送 ZPL 到印表機失敗: {e}")
            raise PrintManagerError(f"發送 ZPL 到印表機失敗: {e}")

    def _send_to_printer(self, filename):
        """發送檔案到印表機"""
        try:
            system = platform.system()
            
            if system == "Windows":
                # 直接列印到標籤機，不顯示對話框
                logger.info(f"直接列印到標籤機: {filename}")
                os.startfile(filename, "print")
            elif system == "Darwin":  # macOS
                if self.show_print_dialog:
                    # 開啟檔案讓使用者手動列印
                    subprocess.run(["open", filename], check=True)
                else:
                    # 直接列印
                    subprocess.run(["lpr", filename], check=True)
            elif system == "Linux":
                if self.show_print_dialog:
                    # 開啟檔案讓使用者手動列印
                    subprocess.run(["xdg-open", filename], check=True)
                else:
                    # 直接列印
                    subprocess.run(["lp", filename], check=True)
            else:
                raise PrintManagerError(f"不支援的作業系統: {system}")
                
            logger.info(f"已發送檔案到印表機: {filename}, 顯示對話框: {self.show_print_dialog}")
            
        except Exception as e:
            logger.error(f"發送到印表機失敗: {e}")
            raise PrintManagerError(f"發送到印表機失敗: {e}")


    def print_labels(self, patient_data, count, printer_mode=None):
        """列印多張標籤 - 支援 PDF、ZPL 和文字檔模式"""
        if count <= 0:
            raise PrintManagerError("列印張數必須大於 0")
        
        # 如果沒有指定 printer_mode，使用配置檔案的設定
        if printer_mode is None:
            printer_mode = "pdf" if self.print_mode.lower() == 'pdf' else "text"
            
        logger.info(f"開始列印標籤，病人: {patient_data.get('name')}, 張數: {count}, 模式: {printer_mode}")
        
        try:
            # 根據模式選擇生成方法
            temp_files = []
            try:
                for i in range(count):
                    # 根據 printer_mode 選擇生成方法
                    if printer_mode == "zpl":
                        # ZPL 模式：直接生成 ZPL 內容並發送到印表機
                        zpl_content = self._generate_zpl_content(patient_data)
                        self._send_zpl_to_printer(zpl_content)
                        logger.info(f"ZPL 標籤已直接發送到印表機")
                    elif printer_mode == "pdf":
                        # PDF 模式
                        pdf_failed = False
                        if REPORTLAB_AVAILABLE:
                            try:
                                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                                temp_file.close()
                                self._generate_label_pdf(patient_data, temp_file.name)
                                logger.info(f"PDF 標籤生成成功: {temp_file.name}")
                            except Exception as e:
                                logger.warning(f"PDF 生成失敗，自動切換到文字模式: {e}")
                                pdf_failed = True
                        else:
                            pdf_failed = True
                        
                        # 如果 PDF 失敗，使用文字模式
                        if pdf_failed:
                            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8')
                            temp_file.close()
                            self._generate_label_text(patient_data, temp_file.name)
                            logger.info(f"文字標籤生成成功: {temp_file.name}")
                    else:
                        # 文字模式 (備用)
                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8')
                        temp_file.close()
                        self._generate_label_text(patient_data, temp_file.name)
                        logger.info(f"文字標籤生成成功: {temp_file.name}")
                    
                    # 只有非 ZPL 模式才需要處理檔案
                    if printer_mode != "zpl":
                        temp_files.append(temp_file.name)
                        # 發送到印表機
                        self._send_to_printer(temp_file.name)
                    
                    # 給印表機一些處理時間
                    if i < count - 1:  # 最後一張不需要等待
                        time.sleep(1)
                
                logger.info(f"成功透過 {printer_mode.upper()} 列印 {count} 張標籤")
                return True
                
            finally:
                # 清理臨時檔案
                for temp_file in temp_files:
                    try:
                        time.sleep(2)  # 等待列印完成
                        if os.path.exists(temp_file):
                            os.remove(temp_file)
                    except Exception as e:
                        logger.warning(f"清理臨時檔案失敗 {temp_file}: {e}")
            
        except Exception as e:
            logger.error(f"列印失敗: {e}")
            raise PrintManagerError(f"列印失敗: {e}")

    def test_printer(self, printer_mode="pdf"):
        """測試印表機連線"""
        try:
            test_data = {
                "id": "TEST123456",
                "name": "測試病人",
                "dob": "1990/01/01"
            }
            
            # 根據模式測試
            if printer_mode == "zpl":
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".zpl")
                temp_file.close()
                self._generate_label_zpl(test_data, temp_file.name)
            elif printer_mode == "pdf" and REPORTLAB_AVAILABLE:
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                temp_file.close()
                self._generate_label_pdf(test_data, temp_file.name)
            else:
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
                temp_file.close()
                self._generate_label_text(test_data, temp_file.name)
            
            logger.info(f"{printer_mode.upper()} 印表機測試完成")
            return True, temp_file.name
            
        except Exception as e:
            logger.error(f"印表機測試失敗: {e}")
            return False, str(e)
