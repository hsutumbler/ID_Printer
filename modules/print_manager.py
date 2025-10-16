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
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    logger.warning("ReportLab 未安裝，將使用簡單文字檔案列印")

# 嘗試匯入條碼生成套件
try:
    import barcode
    from barcode.writer import ImageWriter
    from PIL import Image
    BARCODE_AVAILABLE = True
except ImportError:
    BARCODE_AVAILABLE = False
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
        
        logger.info(f"列印管理模組初始化完成 - 標籤尺寸: {label_width_mm}x{label_height_mm}mm, 列印模式: {self.print_mode}")

    def _try_register_chinese_font(self):
        """嘗試註冊中文字體"""
        try:
            # 常見的中文字體路徑
            font_paths = [
                "C:/Windows/Fonts/msjh.ttc",  # 微軟正黑體
                "C:/Windows/Fonts/kaiu.ttf",   # 標楷體
                "/System/Library/Fonts/PingFang.ttc",  # macOS
                "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"  # Linux
            ]
            
            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                    self.font_registered = True
                    logger.info(f"成功註冊中文字體: {font_path}")
                    break
            
            if not self.font_registered:
                logger.warning("未找到中文字體，將使用預設字體")
                
        except Exception as e:
            logger.warning(f"註冊中文字體失敗: {e}")

    def generate_barcode(self, patient_id):
        """生成條碼圖片"""
        if not BARCODE_AVAILABLE or not self.use_barcode:
            return None
            
        try:
            # 移除非數字和字母的字符
            clean_id = ''.join(c for c in patient_id if c.isalnum())
            
            # 創建條碼
            barcode_class = barcode.get_barcode_class(self.barcode_type)
            barcode_instance = barcode_class(clean_id, writer=ImageWriter())
            
            # 生成條碼圖片到內存
            buffer = io.BytesIO()
            barcode_instance.write(buffer)
            
            # 將緩衝區轉換為圖片對象
            buffer.seek(0)
            barcode_image = Image.open(buffer)
            
            return barcode_image
        
        except Exception as e:
            logger.warning(f"生成條碼失敗: {e}")
            return None

    def _generate_label_pdf(self, patient_data, filename):
        """生成標籤 PDF (5cm x 3.5cm)"""
        if not REPORTLAB_AVAILABLE:
            raise PrintManagerError("ReportLab 未安裝，無法生成 PDF")
            
        try:
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
            
            # 生成條碼
            patient_id = patient_data.get('id', 'N/A')
            if self.use_barcode and BARCODE_AVAILABLE:
                barcode_image = self.generate_barcode(patient_id)
                
                if barcode_image:
                    # 儲存條碼圖片到臨時檔案
                    temp_barcode = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    barcode_image.save(temp_barcode.name)
                    temp_barcode.close()
                    
                    # 計算條碼尺寸和位置
                    barcode_width = self.label_width - 4 * mm  # 留出左右邊距
                    barcode_height = 10 * mm  # 條碼高度
                    
                    # 在 PDF 中插入條碼圖片
                    c.drawImage(temp_barcode.name, start_x, y_pos - barcode_height, width=barcode_width, height=barcode_height)
                    
                    # 刪除臨時檔案
                    try:
                        os.remove(temp_barcode.name)
                    except:
                        pass
                    
                    # 移動到條碼下方（加寬與 ID 之間的間距 0.5mm）
                    y_pos -= barcode_height + 2.5 * mm
            
            # 設定統一的行距
            uniform_spacing = line_height + 0.5 * mm
            
            # 身分證字號
            c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
            c.drawString(start_x, y_pos, f"ID：{patient_data.get('id', 'N/A')}")
            y_pos -= uniform_spacing
            
            # 姓名
            c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
            c.drawString(start_x, y_pos, f"姓名：{patient_data.get('name', 'N/A')}")
            y_pos -= uniform_spacing
            
            # 出生年月日
            c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
            c.drawString(start_x, y_pos, f"生日：{patient_data.get('dob', 'N/A')}")
            y_pos -= uniform_spacing
            
            # 列印時間
            c.setFont('ChineseFont' if self.font_registered else 'Helvetica', font_size)
            c.drawString(start_x, y_pos, f"列印時間：{print_time}")

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
            
            # 使用 ASCII 字符模擬條碼
            barcode_ascii = "|||||||||||  ||||||||||||  ||||||" if self.use_barcode else ""
            barcode_line = f"{barcode_ascii} (條碼)\n" if self.use_barcode else ""
            
            content = f"""
==========================================
{barcode_line}ID：{patient_data.get('id', 'N/A')}
姓名：{patient_data.get('name', 'N/A')}
生日：{patient_data.get('dob', 'N/A')}
列印時間：{print_time}
==========================================
"""
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content)
            
            logger.info(f"生成文字標籤: {filename}")
            return filename
            
        except Exception as e:
            logger.error(f"生成文字標籤失敗: {e}")
            raise PrintManagerError(f"生成文字標籤失敗: {e}")

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


    def print_labels(self, patient_data, count):
        """列印多張標籤 - 支援 PDF 和文字檔模式"""
        if count <= 0:
            raise PrintManagerError("列印張數必須大於 0")
            
        logger.info(f"開始列印標籤，病人: {patient_data.get('name')}, 張數: {count}, 模式: {self.print_mode}")
        
        try:
            # PDF 或文字檔模式
            temp_files = []
            try:
                for i in range(count):
                    # 建立臨時檔案
                    if self.print_mode.lower() == 'pdf' and REPORTLAB_AVAILABLE:
                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                        temp_file.close()
                        self._generate_label_pdf(patient_data, temp_file.name)
                    else:
                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8')
                        temp_file.close()
                        self._generate_label_text(patient_data, temp_file.name)
                    
                    temp_files.append(temp_file.name)
                    
                    # 發送到印表機
                    self._send_to_printer(temp_file.name)
                    
                    # 給印表機一些處理時間
                    if i < count - 1:  # 最後一張不需要等待
                        time.sleep(1)
                
                logger.info(f"成功透過 {self.print_mode.upper()} 列印 {count} 張標籤")
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

    def test_printer(self):
        """測試印表機連線"""
        try:
            test_data = {
                "id": "TEST123456",
                "name": "測試病人",
                "dob": "1990/01/01"
            }
            
            # 測試 PDF 或文字檔模式
            if self.print_mode.lower() == 'pdf' and REPORTLAB_AVAILABLE:
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                temp_file.close()
                self._generate_label_pdf(test_data, temp_file.name)
            else:
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
                temp_file.close()
                self._generate_label_text(test_data, temp_file.name)
            
            logger.info(f"{self.print_mode.upper()} 印表機測試完成")
            return True, temp_file.name
            
        except Exception as e:
            logger.error(f"印表機測試失敗: {e}")
            return False, str(e)
