# -*- coding: utf-8 -*-
"""
中央健康保險署讀卡機 DLL 整合模組
此模組用於包裝健保卡讀卡機的 DLL 函式庫
"""

import os
import ctypes
from ctypes import wintypes, c_char_p, c_int, c_bool, byref, create_string_buffer, c_char
import datetime
import configparser
from .logger import logger

# 嘗試匯入 COM 支援 (用於 GNT NhiCard.dll)
try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False
    logger.warning("win32com 未安裝，無法使用 GNT COM 介面")

class NHICardDLLError(Exception):
    """健保卡 DLL 呼叫錯誤"""
    pass

class NHICardDLL:
    """健保卡 DLL 包裝類別"""
    
    def __init__(self, dll_path=None):
        """
        初始化健保卡 DLL 包裝
        
        參數:
            dll_path: DLL 檔案路徑，如果為 None，則使用預設路徑
        """
        self.dll = None
        self.com_object = None
        self.dll_path = dll_path or self._get_default_dll_path()
        self.initialized = False
        self.is_gnt_dll = False
        self.com_port = self._get_com_port()
        
        try:
            # 載入 DLL
            self._load_dll()
            self.initialized = True
            logger.info(f"健保卡 DLL 載入成功: {self.dll_path}")
        except Exception as e:
            logger.error(f"健保卡 DLL 載入失敗: {e}")
            raise NHICardDLLError(f"無法載入健保卡 DLL: {e}")
    
    def _get_com_port(self):
        """從配置檔案讀取 COM 埠設定"""
        try:
            config = configparser.ConfigParser()
            config.read('config.ini', encoding='utf-8')
            com_port = config.getint('健保卡設定', 'com_port', fallback=3)  # DLL 固定在 COMX3
            logger.info(f"讀卡機 COM 埠設定: COM{com_port} (DLL 固定在 COMX3)")
            return com_port
        except Exception as e:
            logger.warning(f"讀取 COM 埠設定失敗: {e}，使用預設值 COM3")
            return 3
    
    def _get_default_dll_path(self):
        """取得預設的 DLL 路徑 - 優先使用健保署官方 DLL"""
        
        # 檢查配置檔案中指定的路徑（最高優先級）
        try:
            import configparser
            config = configparser.ConfigParser()
            config.read('config.ini', encoding='utf-8')
            config_path = config.get('健保卡設定', 'dll_path', fallback='').strip()
            if config_path and os.path.exists(config_path):
                logger.info(f"使用配置檔案中指定的 DLL 路徑: {config_path}")
                return config_path
        except Exception as e:
            logger.warning(f"讀取配置檔案 DLL 路徑失敗: {e}")
        
        # 檢查環境變數中的路徑（第二優先級）
        nhi_path = os.environ.get("NHI_CARD_DLL_PATH")
        if nhi_path and os.path.exists(nhi_path):
            logger.info(f"使用環境變數中的 DLL 路徑: {nhi_path}")
            return nhi_path
        
        # 健保署官方標準路徑（第三優先級）
        standard_paths = [
            r"C:\NHI\LIB\CsHis50.dll",                    # 健保署官方標準路徑（注意大小寫）
            r"C:\NHI\LIB\csHis50.dll",                    # 健保署官方標準路徑（小寫）
            r"C:\NHI\LIB\CSHIS.dll",                      # 健保署官方標準路徑（全大寫）
            r"C:\Program Files\NHI\LIB\CsHis50.dll",      # Program Files 路徑
            r"C:\Program Files (x86)\NHI\LIB\CsHis50.dll", # Program Files (x86) 路徑
        ]
        
        # 尋找第一個存在的路徑
        for path in standard_paths:
            if os.path.exists(path):
                logger.info(f"找到健保署官方健保卡 DLL: {path}")
                return path
        
        # 檢查是否有外部 GNT DLL（第四優先級）
        # 注意：GNT 資料夾已移除，但保留檢查邏輯以防有其他位置的 GNT DLL
        potential_gnt_paths = [
            os.path.join(os.getcwd(), "GNT", "HenCs", "NhiCard.dll"),      # 原 GNT 路徑（已移除）
            r"C:\GNT\HenCs\NhiCard.dll",                                   # 系統安裝的 GNT
            r"C:\Program Files\GNT\NhiCard.dll",                           # Program Files 中的 GNT
            r"C:\Program Files (x86)\GNT\NhiCard.dll",                     # Program Files (x86) 中的 GNT
        ]
        
        for path in potential_gnt_paths:
            if os.path.exists(path):
                logger.info(f"找到外部 GNT 健保卡 DLL: {path}")
                return path
        
        # 如果都找不到，返回標準健保署路徑（讓錯誤訊息更清楚）
        logger.warning("未找到健保卡 DLL，將使用標準健保署路徑")
        return r"C:\NHI\LIB\CsHis50.dll"
    
    def _load_dll(self):
        """載入 DLL 並設定函式簽名"""
        # 檢查是否為 GNT 的 NhiCard.dll 或標準的 csHis50.dll
        self.is_gnt_dll = "NhiCard.dll" in self.dll_path
        
        if self.is_gnt_dll:
            logger.info("偵測到 GNT NhiCard.dll，這是 .NET 組件，嘗試建立 COM 物件")
            # GNT NhiCard.dll 是 .NET 組件，不能用 ctypes.CDLL 載入
            # 需要透過 COM 介面使用
            if COM_AVAILABLE:
                try:
                    # 嘗試建立 GNT Patient COM 物件
                    self.com_object = win32com.client.Dispatch("NhiCard.Patient")
                    logger.info("成功建立 GNT Patient COM 物件")
                    # 對於 GNT DLL，不需要載入為 ctypes.CDLL
                    self.dll = None
                    return
                except Exception as e:
                    logger.warning(f"建立 GNT COM 物件失敗: {e}")
                    self.com_object = None
                    # 如果 COM 失敗，拋出錯誤而不是嘗試 ctypes 載入
                    raise NHICardDLLError(f"GNT NhiCard.dll 需要 COM 介面，但建立失敗: {e}")
            else:
                logger.warning("win32com 不可用，無法使用 GNT COM 介面")
                raise NHICardDLLError("GNT NhiCard.dll 需要 win32com 套件支援")
        else:
            # 標準健保署 DLL 使用 ctypes 載入
            try:
                self.dll = ctypes.CDLL(self.dll_path)
                logger.info("使用標準健保署 DLL 函式簽名")
                
                # 設定標準健保署 DLL 函式簽名 (使用 cs 系列函式)
                
                # 開啟讀卡機連結埠
                if hasattr(self.dll, 'csOpenCom'):
                    self.dll.csOpenCom.argtypes = [c_int]  # COM 埠號
                    self.dll.csOpenCom.restype = c_int
                
                # 關閉讀卡機連結埠
                if hasattr(self.dll, 'csCloseCom'):
                    self.dll.csCloseCom.argtypes = []
                    self.dll.csCloseCom.restype = c_int
                
                # 讀取健保卡基本資料
                if hasattr(self.dll, 'csReadCard'):
                    self.dll.csReadCard.argtypes = [c_char_p]  # 緩衝區
                    self.dll.csReadCard.restype = c_int
                
                # 備用：舊版函式簽名 (如果 cs 系列不存在)
                # 初始化函式
                if hasattr(self.dll, 'NHI_Initialize'):
                    self.dll.NHI_Initialize.argtypes = []
                    self.dll.NHI_Initialize.restype = c_bool
                
                # 讀卡函式
                if hasattr(self.dll, 'NHI_ReadCard'):
                    self.dll.NHI_ReadCard.argtypes = []
                    self.dll.NHI_ReadCard.restype = c_bool
                
                # 取得身分證字號
                if hasattr(self.dll, 'NHI_GetID'):
                    self.dll.NHI_GetID.argtypes = [c_char_p, c_int]
                    self.dll.NHI_GetID.restype = c_bool
                
                # 取得姓名
                if hasattr(self.dll, 'NHI_GetName'):
                    self.dll.NHI_GetName.argtypes = [c_char_p, c_int]
                    self.dll.NHI_GetName.restype = c_bool
                
                # 取得出生日期
                if hasattr(self.dll, 'NHI_GetBirthDate'):
                    self.dll.NHI_GetBirthDate.argtypes = [c_char_p, c_int]
                    self.dll.NHI_GetBirthDate.restype = c_bool
                
                # 取得錯誤訊息
                if hasattr(self.dll, 'NHI_GetLastError'):
                    self.dll.NHI_GetLastError.argtypes = [c_char_p, c_int]
                    self.dll.NHI_GetLastError.restype = c_int
                
                # 釋放資源
                if hasattr(self.dll, 'NHI_Release'):
                    self.dll.NHI_Release.argtypes = []
                    self.dll.NHI_Release.restype = c_bool
                    
            except Exception as e:
                logger.error(f"載入標準健保署 DLL 失敗: {e}")
                raise NHICardDLLError(f"載入標準健保署 DLL 失敗: {e}")
    
    def initialize(self):
        """初始化讀卡機"""
        if not self.initialized:
            raise NHICardDLLError("DLL 尚未載入")
        
        # 優先使用 csOpenCom 函式
        if hasattr(self.dll, 'csOpenCom'):
            logger.info(f"使用 csOpenCom 開啟 COM{self.com_port}")
            result = self.dll.csOpenCom(self.com_port)
            if result != 0:  # 假設 0 表示成功
                raise NHICardDLLError(f"開啟讀卡機連結埠失敗，錯誤碼: {result}")
            logger.info("讀卡機連結埠開啟成功")
            return True
        elif hasattr(self.dll, 'NHI_Initialize'):
            result = self.dll.NHI_Initialize()
            if not result:
                error = self.get_last_error()
                raise NHICardDLLError(f"初始化讀卡機失敗: {error}")
            return True
        else:
            logger.warning("DLL 中找不到 csOpenCom 或 NHI_Initialize 函式")
            return True  # 假設成功
    
    def read_card(self):
        """讀取健保卡資料"""
        if not self.initialized:
            raise NHICardDLLError("DLL 尚未載入")
        
        try:
            # 優先使用 csReadCard 函式
            if hasattr(self.dll, 'csReadCard'):
                logger.info("使用 csReadCard 讀取健保卡")
                return self._read_card_with_csreadcard()
            elif self.is_gnt_dll and self.com_object:
                # 使用 GNT COM 介面讀取
                return self._read_card_gnt_com()
            elif self.is_gnt_dll:
                # 使用 GNT DLL 直接呼叫 (備用方案)
                return self._read_card_gnt_dll()
            else:
                # 使用標準健保署 DLL
                return self._read_card_standard()
                
        except NHICardDLLError as e:
            logger.error(f"讀取健保卡失敗: {e}")
            raise
        except Exception as e:
            logger.error(f"讀取健保卡時發生未知錯誤: {e}")
            raise NHICardDLLError(f"讀取健保卡時發生未知錯誤: {e}")
        finally:
            # 釋放資源
            self.release()
    
    def _read_card_with_csreadcard(self):
        """使用 csReadCard 函式讀取健保卡"""
        try:
            # 建立緩衝區接收資料
            buffer_size = 1024
            buffer = create_string_buffer(buffer_size)
            
            # 呼叫 csReadCard
            logger.info("呼叫 csReadCard 函式")
            result = self.dll.csReadCard(buffer)
            
            if result == 0:  # 假設 0 表示成功
                # 取得回傳的字串資料
                raw_data = buffer.value.decode('big5', errors='ignore')
                logger.info(f"csReadCard 回傳原始資料: {repr(raw_data)}")
                
                # 解析資料
                parsed_data = self._parse_csreadcard_data(raw_data)
                return parsed_data
            else:
                raise NHICardDLLError(f"csReadCard 執行失敗，錯誤碼: {result}")
                
        except Exception as e:
            logger.error(f"csReadCard 呼叫失敗: {e}")
            raise NHICardDLLError(f"csReadCard 呼叫失敗: {e}")
    
    def _parse_csreadcard_data(self, raw_data):
        """解析 csReadCard 回傳的資料"""
        try:
            logger.info(f"開始解析 csReadCard 資料: {repr(raw_data)}")
            
            # 移除空白字符
            data = raw_data.strip()
            
            if not data:
                raise NHICardDLLError("csReadCard 回傳空資料")
            
            # 嘗試不同的解析方式
            parsed_data = None
            
            # 方法1: 嘗試分隔符解析
            for delimiter in ['|', ',', '\t', ';']:
                if delimiter in data:
                    parts = data.split(delimiter)
                    if len(parts) >= 3:  # 至少要有身分證、姓名、出生日期
                        parsed_data = self._parse_delimited_data(parts)
                        logger.info(f"使用分隔符 '{delimiter}' 解析成功")
                        break
            
            # 方法2: 如果沒有分隔符，嘗試固定長度解析
            if not parsed_data:
                parsed_data = self._parse_fixed_length_data(data)
                logger.info("使用固定長度解析")
            
            # 方法3: 如果都失敗，嘗試正則表達式解析
            if not parsed_data:
                parsed_data = self._parse_regex_data(data)
                logger.info("使用正則表達式解析")
            
            if not parsed_data:
                raise NHICardDLLError("無法解析 csReadCard 回傳的資料格式")
            
            logger.info(f"解析結果: {parsed_data}")
            return parsed_data
            
        except Exception as e:
            logger.error(f"解析 csReadCard 資料失敗: {e}")
            raise NHICardDLLError(f"解析健保卡資料失敗: {e}")
    
    def _parse_delimited_data(self, parts):
        """解析分隔符格式的資料"""
        try:
            # 假設格式: 身分證|姓名|出生日期|性別|健保卡號
            return {
                "ID_NUMBER": parts[0].strip() if len(parts) > 0 else "",
                "FULL_NAME": parts[1].strip() if len(parts) > 1 else "",
                "BIRTH_DATE": parts[2].strip() if len(parts) > 2 else "",
                "SEX": parts[3].strip() if len(parts) > 3 else "",
                "CARD_NUMBER": parts[4].strip() if len(parts) > 4 else ""
            }
        except Exception as e:
            logger.error(f"分隔符解析失敗: {e}")
            return None
    
    def _parse_fixed_length_data(self, data):
        """解析固定長度格式的資料"""
        try:
            # 假設格式: 身分證10位 + 姓名20位 + 出生日期8位 + 性別1位 + 健保卡號12位
            if len(data) >= 38:  # 最少需要的長度
                return {
                    "ID_NUMBER": data[0:10].strip(),
                    "FULL_NAME": data[10:30].strip(),
                    "BIRTH_DATE": data[30:38].strip(),
                    "SEX": data[38:39].strip() if len(data) > 38 else "",
                    "CARD_NUMBER": data[39:51].strip() if len(data) > 50 else ""
                }
        except Exception as e:
            logger.error(f"固定長度解析失敗: {e}")
        return None
    
    def _parse_regex_data(self, data):
        """使用正則表達式解析資料"""
        try:
            import re
            
            # 身分證字號模式 (1個英文字母 + 9個數字)
            id_match = re.search(r'[A-Z]\d{9}', data)
            
            # 出生日期模式 (8個數字，可能是民國年或西元年)
            birth_match = re.search(r'\d{7,8}', data)
            
            # 姓名模式 (中文字符)
            name_match = re.search(r'[\u4e00-\u9fff]+', data)
            
            if id_match and name_match:
                return {
                    "ID_NUMBER": id_match.group(),
                    "FULL_NAME": name_match.group(),
                    "BIRTH_DATE": birth_match.group() if birth_match else "",
                    "SEX": "",
                    "CARD_NUMBER": ""
                }
        except Exception as e:
            logger.error(f"正則表達式解析失敗: {e}")
        return None
    
    def _read_card_gnt_com(self):
        """使用 GNT COM 介面讀取健保卡"""
        logger.info("使用 GNT COM 介面讀取健保卡")
        
        try:
            # 開啟讀卡機連接埠
            if hasattr(self.com_object, 'Open'):
                result = self.com_object.Open()
                if not result:
                    raise NHICardDLLError("無法開啟讀卡機連接埠")
            
            # 取得病患資料
            if hasattr(self.com_object, 'GetPatientData'):
                result = self.com_object.GetPatientData()
                if not result:
                    raise NHICardDLLError("讀取病患資料失敗")
            
            # 檢查是否有卡片
            if hasattr(self.com_object, 'CardCheck') and not self.com_object.CardCheck:
                raise NHICardDLLError("未偵測到健保卡，請確認卡片已正確插入")
            
            # 取得病患資料
            card_data = {
                "ID_NUMBER": getattr(self.com_object, 'GetPatientIdCard', ''),
                "FULL_NAME": getattr(self.com_object, 'GetPatientName', ''),
                "BIRTH_DATE": '',  # GNT COM 介面可能不提供出生日期
                "SEX": getattr(self.com_object, 'GetPatientSex', '')
            }
            
            # 檢查必要資料
            if not card_data["ID_NUMBER"] or not card_data["FULL_NAME"]:
                raise NHICardDLLError("讀取的健保卡資料不完整")
            
            logger.info(f"成功讀取健保卡 (GNT COM)，病人: {card_data['FULL_NAME']}")
            return card_data
            
        except Exception as e:
            raise NHICardDLLError(f"GNT COM 讀取失敗: {e}")
    
    def _read_card_gnt_dll(self):
        """使用 GNT DLL 直接呼叫讀取健保卡 (備用方案)"""
        logger.info("使用 GNT DLL 直接呼叫讀取健保卡")
        # 這裡可以實作直接呼叫 GNT DLL 的方法
        # 由於 GNT DLL 主要設計為 COM 介面，直接呼叫較複雜
        raise NHICardDLLError("GNT DLL 直接呼叫尚未實作，請確保 win32com 可用")
    
    def _read_card_standard(self):
        """使用標準健保署 DLL 讀取健保卡"""
        logger.info("使用標準健保署 DLL 讀取健保卡")
        
        # 初始化讀卡機
        self.initialize()
        
        # 讀取卡片
        if hasattr(self.dll, 'NHI_ReadCard'):
            result = self.dll.NHI_ReadCard()
            if not result:
                error = self.get_last_error()
                raise NHICardDLLError(f"讀取健保卡失敗: {error}")
        
        # 取得身分證字號
        id_number = self._get_string_value('NHI_GetID', 20)
        
        # 取得姓名
        name = self._get_string_value('NHI_GetName', 50)
        
        # 取得出生日期
        birth_date = self._get_string_value('NHI_GetBirthDate', 20)
        
        # 整理資料
        card_data = {
            "ID_NUMBER": id_number,
            "FULL_NAME": name,
            "BIRTH_DATE": self._format_birth_date(birth_date)
        }
        
        logger.info(f"成功讀取健保卡 (標準 DLL)，病人: {name}")
        return card_data
    
    def _get_string_value(self, function_name, buffer_size):
        """從 DLL 函式取得字串值"""
        if not hasattr(self.dll, function_name):
            logger.warning(f"DLL 中找不到 {function_name} 函式")
            return ""
        
        buffer = create_string_buffer(buffer_size)
        result = getattr(self.dll, function_name)(buffer, buffer_size)
        
        if not result:
            error = self.get_last_error()
            logger.warning(f"{function_name} 失敗: {error}")
            return ""
        
        # 將 bytes 轉換為 string
        value = buffer.value.decode('utf-8', errors='ignore')
        return value.strip()
    
    def _format_birth_date(self, birth_date):
        """格式化出生日期"""
        # 處理各種可能的日期格式
        if not birth_date:
            return ""
        
        # 移除空白和特殊字元
        birth_date = birth_date.strip().replace('-', '').replace('/', '')
        
        # 如果是 YYYYMMDD 格式
        if len(birth_date) == 8:
            try:
                year = int(birth_date[0:4])
                month = int(birth_date[4:6])
                day = int(birth_date[6:8])
                
                # 檢查日期是否有效
                if 1900 <= year <= 2100 and 1 <= month <= 12 and 1 <= day <= 31:
                    return f"{year:04d}/{month:02d}/{day:02d}"
            except ValueError:
                pass
        
        # 如果是民國年格式 (YYYMMDD)
        elif len(birth_date) == 7:
            try:
                year = int(birth_date[0:3]) + 1911  # 民國年轉西元年
                month = int(birth_date[3:5])
                day = int(birth_date[5:7])
                
                # 檢查日期是否有效
                if 1900 <= year <= 2100 and 1 <= month <= 12 and 1 <= day <= 31:
                    return f"{year:04d}/{month:02d}/{day:02d}"
            except ValueError:
                pass
        
        # 如果無法解析，直接返回原始值
        return birth_date
    
    def get_last_error(self):
        """取得最後的錯誤訊息"""
        if hasattr(self.dll, 'NHI_GetLastError'):
            buffer = create_string_buffer(256)
            self.dll.NHI_GetLastError(buffer, 256)
            return buffer.value.decode('utf-8', errors='ignore').strip()
        else:
            return "無法取得錯誤訊息"
    
    def release(self):
        """釋放資源"""
        if self.initialized:
            try:
                # 優先使用 csCloseCom 函式
                if hasattr(self.dll, 'csCloseCom'):
                    logger.info("使用 csCloseCom 關閉讀卡機連結埠")
                    result = self.dll.csCloseCom()
                    if result == 0:
                        logger.info("讀卡機連結埠關閉成功")
                    else:
                        logger.warning(f"關閉讀卡機連結埠失敗，錯誤碼: {result}")
                elif hasattr(self.dll, 'NHI_Release'):
                    self.dll.NHI_Release()
                    logger.info("資源釋放成功")
                else:
                    logger.warning("DLL 中找不到 csCloseCom 或 NHI_Release 函式")
            except Exception as e:
                logger.warning(f"釋放資源失敗: {e}")

# 測試函式
def test_nhi_card_dll(dll_path=None):
    """測試健保卡 DLL 功能"""
    try:
        print("開始測試健保卡 DLL...")
        nhi_dll = NHICardDLL(dll_path)
        print(f"DLL 載入成功: {nhi_dll.dll_path}")
        
        print("請插入健保卡，然後按 Enter 繼續...")
        input()
        
        print("讀取健保卡中...")
        card_data = nhi_dll.read_card()
        
        print("\n讀取結果:")
        print(f"身分證字號: {card_data['ID_NUMBER']}")
        print(f"姓名: {card_data['FULL_NAME']}")
        print(f"出生日期: {card_data['BIRTH_DATE']}")
        
        print("\n測試成功!")
        return True
    except Exception as e:
        print(f"測試失敗: {e}")
        return False

if __name__ == "__main__":
    # 如果直接執行此模組，執行測試
    test_nhi_card_dll()
