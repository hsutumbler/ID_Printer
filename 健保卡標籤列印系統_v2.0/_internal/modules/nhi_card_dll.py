# -*- coding: utf-8 -*-
"""
中央健康保險署讀卡機 DLL 整合模組
此模組用於包裝健保卡讀卡機的 DLL 函式庫
"""

import os
import ctypes
from ctypes import wintypes, c_char_p, c_int, c_bool, byref, create_string_buffer, c_char
import datetime
from .logger import logger

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
        self.dll_path = dll_path or self._get_default_dll_path()
        self.initialized = False
        
        try:
            # 載入 DLL
            self._load_dll()
            self.initialized = True
            logger.info(f"健保卡 DLL 載入成功: {self.dll_path}")
        except Exception as e:
            logger.error(f"健保卡 DLL 載入失敗: {e}")
            raise NHICardDLLError(f"無法載入健保卡 DLL: {e}")
    
    def _get_default_dll_path(self):
        """取得預設的 DLL 路徑"""
        # 健保署提供的 DLL 路徑
        csHis50_dll_path = r"C:\NHI\LIB\csHis50.dll"
        
        # 常見的健保卡 DLL 安裝路徑
        possible_paths = [
            csHis50_dll_path,  # 健保署提供的路徑優先
            r"C:\Program Files\NHI\NHICardReader.dll",
            r"C:\Program Files (x86)\NHI\NHICardReader.dll",
            r"C:\Windows\System32\NHICardReader.dll",
            r"C:\Windows\SysWOW64\NHICardReader.dll",
            r"NHICardReader.dll"  # 當前目錄
        ]
        
        # 檢查環境變數中的路徑
        nhi_path = os.environ.get("NHI_CARD_DLL_PATH")
        if nhi_path and os.path.exists(nhi_path):
            possible_paths.insert(0, nhi_path)
        
        # 尋找第一個存在的路徑
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        # 如果都找不到，返回健保署提供的路徑
        return csHis50_dll_path
    
    def _load_dll(self):
        """載入 DLL 並設定函式簽名"""
        try:
            self.dll = ctypes.CDLL(self.dll_path)
            
            # 設定常見的健保卡 DLL 函式簽名
            # 注意: 實際的函式名稱和參數可能需要根據實際的 DLL 文件進行調整
            
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
            logger.error(f"設定 DLL 函式簽名失敗: {e}")
            raise NHICardDLLError(f"設定 DLL 函式簽名失敗: {e}")
    
    def initialize(self):
        """初始化讀卡機"""
        if not self.initialized:
            raise NHICardDLLError("DLL 尚未載入")
        
        if hasattr(self.dll, 'NHI_Initialize'):
            result = self.dll.NHI_Initialize()
            if not result:
                error = self.get_last_error()
                raise NHICardDLLError(f"初始化讀卡機失敗: {error}")
            return True
        else:
            logger.warning("DLL 中找不到 NHI_Initialize 函式")
            return True  # 假設成功
    
    def read_card(self):
        """讀取健保卡資料"""
        if not self.initialized:
            raise NHICardDLLError("DLL 尚未載入")
        
        try:
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
            
            logger.info(f"成功讀取健保卡，病人: {name}")
            return card_data
            
        except NHICardDLLError as e:
            logger.error(f"讀取健保卡失敗: {e}")
            raise
        except Exception as e:
            logger.error(f"讀取健保卡時發生未知錯誤: {e}")
            raise NHICardDLLError(f"讀取健保卡時發生未知錯誤: {e}")
        finally:
            # 釋放資源
            self.release()
    
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
        if self.initialized and hasattr(self.dll, 'NHI_Release'):
            try:
                self.dll.NHI_Release()
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
