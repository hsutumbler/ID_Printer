# -*- coding: utf-8 -*-
"""
中央健康保險署讀卡機 DLL 整合模組
此模組用於包裝健保卡讀卡機的 DLL 函式庫
"""

import os
import sys
import threading
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
    
    # 執行緒鎖定物件（參考程式的 SyncLock）
    _read_lock = threading.Lock()
    
    def __init__(self, dll_path=None):
        """
        初始化健保卡 DLL 包裝
        
        參數:
            dll_path: DLL 檔案路徑，如果為 None，則使用預設路徑
        """
        import threading
        self.dll = None
        self.com_object = None
        self.dll_path = dll_path or self._get_default_dll_path()
        self.initialized = False
        self.is_gnt_dll = False
        # 不再需要 COM 埠設定，DLL 會自動管理
        
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
        # 優先從 CardPortSet.xml 讀取（參考程式的方法）
        com_port = self._read_com_port_from_config()
        if com_port:
            return com_port
        
        # 備用：從 config.ini 讀取
        try:
            config = configparser.ConfigParser()
            config.read('config.ini', encoding='utf-8')
            com_port = config.getint('健保卡設定', 'com_port', fallback=1)
            logger.info(f"從 config.ini 讀取 COM 埠設定: COM{com_port}")
            return com_port
        except Exception as e:
            logger.warning(f"讀取 COM 埠設定失敗: {e}，使用預設值 COM1")
            return 1
    
    def _read_com_port_from_config(self):
        """從 CardPortSet.xml 讀取 COM 埠號（參考程式的方法）"""
        try:
            import xml.etree.ElementTree as ET
            
            # 參考程式讀取 CardPortSet.xml（從當前目錄）
            # 嘗試多個可能的路徑
            config_paths = [
                os.path.join(os.getcwd(), "CardPortSet.xml"),  # 當前工作目錄
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "CardPortSet.xml"),  # 專案根目錄
                r"C:\NHI\INI\CardPortSet.xml",  # 健保署標準路徑
            ]
            
            for config_path in config_paths:
                if os.path.exists(config_path):
                    tree = ET.parse(config_path)
                    root = tree.getroot()
                    
                    # 尋找 <num> 標籤
                    num_elem = root.find("num")
                    if num_elem is not None and num_elem.text:
                        port = int(num_elem.text.strip())
                        logger.info(f"從 CardPortSet.xml 讀取 COM 埠: {port} (檔案: {config_path})")
                        return port
                    
                    # 如果沒有 <num>，嘗試從 <port><num> 結構讀取
                    port_elem = root.find("port/num")
                    if port_elem is not None and port_elem.text:
                        port = int(port_elem.text.strip())
                        logger.info(f"從 CardPortSet.xml 讀取 COM 埠: {port} (檔案: {config_path})")
                        return port
        except Exception as e:
            logger.debug(f"讀取 CardPortSet.xml 失敗: {e}")
        
        return None
    
    def _get_default_dll_path(self):
        """取得預設的 DLL 路徑 - 只從專案下的 DLL 資料夾載入"""
        
        # 取得專案根目錄
        # 如果是打包後的執行檔，使用執行檔所在目錄
        # 如果是開發模式，使用腳本所在目錄
        if getattr(sys, 'frozen', False):
            # 打包後的執行檔模式
            project_dir = os.path.dirname(sys.executable)
        else:
            # 開發模式：使用模組所在目錄的上層（專案根目錄）
            project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        # 專案下的 DLL 資料夾路徑
        dll_folder = os.path.join(project_dir, "DLL")
        
        # 優先順序：CsHis.dll > CsHis50.dll
        dll_files = [
            "CsHis.dll",      # 優先使用 CsHis.dll
            "CsHis50.dll",    # 備用：CsHis50.dll
        ]
        
        # 搜尋 DLL 資料夾中的檔案
        for dll_file in dll_files:
            dll_path = os.path.join(dll_folder, dll_file)
            if os.path.exists(dll_path):
                logger.info(f"找到專案 DLL 資料夾中的健保卡 DLL: {dll_path}")
                return dll_path
        
        # 如果都找不到，返回錯誤訊息
        logger.error(f"在專案 DLL 資料夾中找不到健保卡 DLL: {dll_folder}")
        logger.error(f"請確認以下檔案是否存在：")
        for dll_file in dll_files:
            logger.error(f"  - {os.path.join(dll_folder, dll_file)}")
        
        # 返回第一個預期路徑（讓錯誤訊息更清楚）
        return os.path.join(dll_folder, dll_files[0])
    
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
                # 載入 DLL 前，設定 DLL 所在目錄為當前工作目錄
                # 這樣 DLL 可以找到相關的依賴檔案（如其他 DLL、設定檔等）
                dll_dir = os.path.dirname(self.dll_path)
                original_cwd = os.getcwd()
                try:
                    if os.path.exists(dll_dir):
                        os.chdir(dll_dir)
                        logger.debug(f"暫時切換工作目錄到 DLL 所在目錄: {dll_dir}")
                except:
                    pass
                
                # 載入 DLL（使用完整路徑）
                self.dll = ctypes.CDLL(self.dll_path)
                logger.info(f"使用標準健保署 DLL 函式簽名，載入: {self.dll_path}")
                
                # 恢復原始工作目錄
                try:
                    os.chdir(original_cwd)
                except:
                    pass
                
                # 設定標準健保署 DLL 函式簽名 (使用 cs 系列函式)
                
                # 注意：不再設定 csOpenCom 和 csCloseCom 的函式簽名
                # 因為我們直接呼叫 hisGetBasicData，讓 DLL 自動管理 COM 埠
                
                # 讀取健保卡基本資料 - 參考程式使用的關鍵函式
                if hasattr(self.dll, 'hisGetBasicData'):
                    # hisGetBasicData 函式簽名：
                    # int hisGetBasicData(byte[] pBuffer, ref int iBufferLen)
                    # 使用 POINTER(c_byte) 來表示 byte 陣列
                    self.dll.hisGetBasicData.argtypes = [
                        ctypes.POINTER(ctypes.c_byte),  # byte[] pBuffer
                        ctypes.POINTER(c_int)           # ref int iBufferLen
                    ]
                    self.dll.hisGetBasicData.restype = c_int
                    logger.info("已設定 hisGetBasicData 函式簽名（參考程式的方法）")
                
                # 讀取健保卡基本資料（備用方案）
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
        """初始化讀卡機（DLL 會自動管理 COM 埠，不需要手動初始化）"""
        if not self.initialized:
            raise NHICardDLLError("DLL 尚未載入")
        
        # DLL 會自動管理 COM 埠，不需要手動初始化
        logger.info("DLL 會自動管理 COM 埠，跳過手動初始化")
        return True
    
    def read_card(self):
        """讀取健保卡資料"""
        if not self.initialized:
            raise NHICardDLLError("DLL 尚未載入")
        
        try:
            # 優先使用 hisGetBasicData 函式（參考程式的正確方法）
            if hasattr(self.dll, 'hisGetBasicData'):
                logger.info("使用 hisGetBasicData 讀取健保卡（參考程式的方法）")
                return self._read_card_with_hisgetbasicdata()
            # 備用方案：使用 csReadCard 函式
            elif hasattr(self.dll, 'csReadCard'):
                logger.info("使用 csReadCard 讀取健保卡（備用方案）")
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
    
    def _read_card_with_hisgetbasicdata(self):
        """使用 hisGetBasicData 函式讀取健保卡（完全參考程式的邏輯）"""
        import time
        import locale
        
        # 儲存原始編碼設定
        original_locale = None
        
        # 使用執行緒鎖定（參考程式的 SyncLock PatientCardObj）
        with NHICardDLL._read_lock:
            try:
                # 設定系統編碼為 Big5（解決 CIE 視窗亂碼問題）
                try:
                    # 嘗試設定為 Big5 編碼
                    original_locale = locale.getlocale()
                    locale.setlocale(locale.LC_ALL, 'Chinese_Taiwan.950')  # Big5 編碼
                    logger.debug("已設定系統編碼為 Big5")
                except:
                    try:
                        # 備用方案：設定為 UTF-8
                        locale.setlocale(locale.LC_ALL, 'Chinese_Taiwan.65001')  # UTF-8
                        logger.debug("已設定系統編碼為 UTF-8")
                    except:
                        logger.warning("無法設定系統編碼，可能會有亂碼問題")
                
                # 設定環境變數（確保 DLL 視窗使用正確編碼）
                import os
                os.environ['LANG'] = 'zh_TW.BIG5'
                os.environ['LC_ALL'] = 'zh_TW.BIG5'
                
                logger.info("使用 hisGetBasicData 讀取健保卡（參考程式的完整邏輯）")
                
                # 參考程式的 CardCheck 屬性流程：
                # 1. _returnValue = Open() - Open() 直接返回 True，不需要手動開啟 COM 埠
                # 2. SysWait(10) - 等待 10 次，每次 40ms = 400ms
                # 3. If _returnValue Then _returnValue = GetPatientData()
                # 4. SysWait(15) - 等待 15 次，每次 40ms = 600ms
                
                # 模擬 Open() 返回 True（參考程式直接返回 True）
                # 參考程式：SysWait(10) - 等待 10 次，每次 40ms = 400ms
                time.sleep(0.4)
                
                # 檢查必要的函式是否存在
                if not hasattr(self.dll, 'hisGetBasicData'):
                    raise NHICardDLLError("DLL 中找不到 hisGetBasicData 函式")
                
                # 準備 72 bytes 的緩衝區（參考程式使用 72 bytes）
                # 注意：VB.NET 的 Dim pBuffer(iBufferLen) As Byte 會建立 iBufferLen+1 個元素（0 到 iBufferLen）
                # 所以實際上是 73 bytes（索引 0-72），但我們使用 72 bytes 應該也可以
                buffer_len = 72
                buffer = (ctypes.c_byte * buffer_len)()
                buffer_len_ref = ctypes.c_int(buffer_len)
                
                # 記錄 processID（如果可用）
                import os
                process_id = os.getpid()
                logger.info(f"Process ID: {process_id}")
                logger.info(f"API NO: hisGetBasicData")
                logger.info(f"緩衝區大小: {buffer_len} bytes")
                logger.info(f"緩衝區長度參考值: {buffer_len_ref.value}")
                
                # 參考程式的 GetPatientData() 方法
                # 直接呼叫 hisGetBasicData(pBuffer, iBufferLen)
                logger.info("呼叫 hisGetBasicData 函式（參考程式的方法）")
                result = self.dll.hisGetBasicData(buffer, ctypes.byref(buffer_len_ref))
                
                # 記錄回傳值
                logger.info(f"Return: {result}")
                logger.info(f"呼叫後 buffer_len_ref 值: {buffer_len_ref.value}")
                
                # 參考程式：SysWait(15) - 等待 15 次，每次 40ms = 600ms
                time.sleep(0.6)
                
                if result != 0:
                    # 記錄詳細的錯誤資訊以便除錯
                    logger.error(f"hisGetBasicData 執行失敗，錯誤碼: {result}")
                    logger.error(f"Process ID: {process_id}")
                    logger.error(f"API NO: hisGetBasicData")
                    logger.error(f"Return: {result}")
                    logger.error(f"緩衝區內容（前 20 bytes）: {bytes(buffer[:20])}")
                    logger.error(f"緩衝區長度: {len(buffer)}")
                    logger.error(f"buffer_len_ref 值: {buffer_len_ref.value}")
                    
                    # 根據錯誤碼提供更詳細的錯誤訊息
                    error_messages = {
                        4000: "健保卡未插入或讀取失敗\n\n請確認：\n1. 健保卡已正確插入讀卡機\n2. 讀卡機已正確連接至電腦\n3. 讀卡機電源已開啟\n4. 讀卡機驅動程式已正確安裝",
                        9200: "COM 埠開啟失敗\n\n請確認讀卡機連接正常",
                        9205: "SAM 認證失敗\n\n請確認讀卡機驅動程式已正確安裝",
                    }
                    
                    error_msg = error_messages.get(result, f"hisGetBasicData 執行失敗，錯誤碼: {result}")
                    raise NHICardDLLError(error_msg)
                
                # 4. 解析資料（使用 Big5 編碼，參考程式的方法）
                # 將 byte 陣列轉換為 bytes
                byte_array = bytes(buffer)
                
                # 使用 Big5 編碼解析（參考程式使用 Big5）
                big5_encoding = 'big5'
                card_whole_str = byte_array.decode(big5_encoding, errors='ignore')
                
                logger.info(f"hisGetBasicData 回傳原始資料長度: {len(byte_array)} bytes")
                logger.debug(f"解析後完整字串: {repr(card_whole_str)}")
                
                # 根據實際資料格式解析（範例：000073191033許欽豪    T1225224130670723M10409291077619438）
                # 格式規則：
                # 1. 前12碼為卡號
                # 2. 姓名(2-5字中文寬度，可能有空格填充)
                # 3. 身分證字號(英文字母+9碼數字，共10碼)
                # 4. 出生年月日(7碼數字=3(民國年)/2(月份)/2(日期))
                # 5. 性別(1碼：M/F/1/2)
                
                card_data = self._parse_card_data_by_format(card_whole_str)
                
                # 驗證必要資料
                if not card_data.get("ID_NUMBER"):
                    raise NHICardDLLError("無法從健保卡資料中提取身份證號碼")
                
                logger.info(f"成功讀取健保卡，身份證: {card_data.get('ID_NUMBER')}, 姓名: {card_data.get('FULL_NAME')}")
                return card_data
                
            except UnicodeDecodeError as e:
                raise NHICardDLLError(f"Big5 編碼解析失敗: {e}")
            except NHICardDLLError:
                raise
            except Exception as e:
                logger.error(f"hisGetBasicData 呼叫失敗: {e}")
                raise NHICardDLLError(f"hisGetBasicData 呼叫失敗: {e}")
            finally:
                # 恢復原始編碼設定
                try:
                    if original_locale:
                        locale.setlocale(locale.LC_ALL, original_locale)
                except:
                    pass
    
    def _parse_card_data_by_format(self, card_whole_str):
        """
        根據實際資料格式解析健保卡資料
        
        格式規則（範例：000073191033許欽豪    T1225224130670723M10409291077619438）：
        1. 前12碼為卡號
        2. 姓名(2-5字中文寬度，可能有空格填充)
        3. 身分證字號(英文字母+9碼數字，共10碼)
        4. 出生年月日(7碼數字=3(民國年)/2(月份)/2(日期))
        5. 性別(1碼：M/F/1/2)
        
        參數:
            card_whole_str: 完整的健保卡資料字串
        
        返回:
            dict: 包含解析後的健保卡資料
        """
        import re
        
        card_data = {
            "ID_NUMBER": "",
            "FULL_NAME": "",
            "BIRTH_DATE": "",
            "SEX": "",
            "CARD_NUMBER": "",
            "CARD_WHOLE_STR": card_whole_str
        }
        
        try:
            # 1. 提取卡號（前12碼）
            if len(card_whole_str) >= 12:
                card_data["CARD_NUMBER"] = card_whole_str[:12].strip()
                logger.debug(f"提取卡號: {card_data['CARD_NUMBER']}")
            
            # 2. 提取姓名（從位置12開始，尋找2-5個中文字）
            # 姓名後面可能有空格填充，然後是身分證字號
            name_pattern = r'[\u4e00-\u9fff]{2,5}'
            name_match = re.search(name_pattern, card_whole_str)
            name_end_pos = 12  # 預設從卡號後開始
            if name_match:
                card_data["FULL_NAME"] = name_match.group().strip()
                name_end_pos = name_match.end()
                logger.debug(f"提取姓名: {card_data['FULL_NAME']} (結束位置: {name_end_pos})")
            else:
                logger.warning("無法找到姓名（中文字）")
            
            # 3. 提取身分證字號（英文字母+9碼數字，共10碼）
            # 從姓名結束位置開始尋找，跳過可能的空格
            id_pattern = r'[A-Z][0-9]{9}'
            # 從姓名結束位置開始搜尋
            search_start = name_end_pos
            id_match = re.search(id_pattern, card_whole_str[search_start:])
            id_end_pos = search_start
            if id_match:
                card_data["ID_NUMBER"] = id_match.group().strip()
                id_end_pos = search_start + id_match.end()
                logger.debug(f"提取身分證字號: {card_data['ID_NUMBER']} (結束位置: {id_end_pos})")
            else:
                logger.warning("無法找到身分證字號（格式：英文字母+9碼數字）")
                # 備用方案：嘗試從位置32開始提取（參考程式的舊方法）
                if len(card_whole_str) >= 42:
                    card_data["ID_NUMBER"] = card_whole_str[32:42].strip()
                    id_end_pos = 42
            
            # 4. 提取出生年月日（7碼數字：3碼民國年+2碼月份+2碼日期）
            # 從身分證字號結束位置開始尋找
            birth_end_pos = id_end_pos
            if id_end_pos < len(card_whole_str):
                # 跳過可能的空格，尋找7碼數字
                birth_pattern = r'[0-9]{7}'
                birth_match = re.search(birth_pattern, card_whole_str[id_end_pos:])
                if birth_match:
                    birth_str = birth_match.group()
                    # 保持民國年格式（不轉換為西元年）
                    try:
                        roc_year = int(birth_str[:3])
                        month = int(birth_str[3:5])
                        day = int(birth_str[5:7])
                        # 儲存為民國年格式：YYYMMDD（7碼）
                        card_data["BIRTH_DATE"] = birth_str
                        logger.debug(f"提取出生日期: {birth_str} (民國{roc_year}年{month}月{day}日)")
                    except ValueError:
                        logger.warning(f"出生日期格式錯誤: {birth_str}")
                        card_data["BIRTH_DATE"] = birth_str
                    birth_end_pos = id_end_pos + birth_match.end()
                else:
                    logger.warning("無法找到出生年月日（7碼數字）")
            
            # 5. 提取性別（1碼：M/F/1/2）
            if birth_end_pos < len(card_whole_str):
                sex_char = card_whole_str[birth_end_pos:birth_end_pos+1].strip()
                if sex_char in ['M', '1']:
                    card_data["SEX"] = "男"
                elif sex_char in ['F', '2']:
                    card_data["SEX"] = "女"
                else:
                    card_data["SEX"] = sex_char if sex_char else ""
                logger.debug(f"提取性別: {sex_char} -> {card_data['SEX']}")
            
        except Exception as e:
            logger.error(f"解析健保卡資料格式失敗: {e}")
            # 如果解析失敗，嘗試使用舊的方法作為備用
            logger.info("嘗試使用備用解析方法")
            try:
                # 備用方法：使用舊的固定位置提取
                if len(card_whole_str) >= 12:
                    card_data["CARD_NUMBER"] = card_whole_str[:12].strip()
                if len(card_whole_str) >= 42:
                    card_data["ID_NUMBER"] = card_whole_str[32:42].strip()
                # 尋找姓名
                name_match = re.search(r'[\u4e00-\u9fff]+', card_whole_str)
                if name_match:
                    card_data["FULL_NAME"] = name_match.group().strip()
            except:
                pass
        
        return card_data
    
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
        """釋放資源（DLL 會自動管理 COM 埠，不需要手動關閉）"""
        if self.initialized:
            # DLL 會自動管理 COM 埠，不需要手動關閉
            logger.debug("DLL 會自動管理 COM 埠，跳過手動關閉")
            # 如果需要，可以呼叫其他釋放函式
            try:
                if hasattr(self.dll, 'NHI_Release'):
                    self.dll.NHI_Release()
                    logger.debug("資源釋放成功")
            except Exception as e:
                logger.debug(f"資源釋放: {e}")

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
