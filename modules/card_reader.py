# 健保卡讀取模組
import datetime
import time
import threading
import random
import os
import subprocess
import configparser
import serial
import serial.tools.list_ports
from .logger import logger
from .nhi_card_dll import NHICardDLL, NHICardDLLError

class CardReaderError(Exception):
    """健保卡讀取錯誤"""
    pass

class CardReader:
    def __init__(self, dll_path=None):
        # 初始化讀卡機 SDK
        self.is_reading = False
        self.dll_path = dll_path
        self.csfsim_path = None
        self.offline_mode = False
        self.offline_auto_print = False
        self.com_port = None
        self.auto_detect_com = True
        
        # 讀取設定檔
        try:
            config = configparser.ConfigParser()
            config.read('config.ini', encoding='utf-8')
            self.csfsim_path = config.get('健保卡設定', 'csfsim_path', fallback=r"C:\NHI\BIN\csfsim.exe")
            self.offline_mode = config.getboolean('健保卡設定', 'offline_mode', fallback=True)
            self.offline_auto_print = config.getboolean('健保卡設定', 'offline_auto_print', fallback=True)
            self.auto_detect_com = config.getboolean('健保卡設定', 'auto_detect_com', fallback=True)
            
            # 讀取COM埠設定
            if not self.auto_detect_com:
                self.com_port = config.getint('健保卡設定', 'com_port', fallback=3)
            else:
                # 自動偵測COM埠
                self.com_port = self._detect_com_port()
            
            logger.info(f"健保卡讀卡機控制軟體路徑: {self.csfsim_path}")
            logger.info(f"離線模式: {'啟用' if self.offline_mode else '停用'}")
            logger.info(f"COM埠自動偵測: {'啟用' if self.auto_detect_com else '停用'}")
            logger.info(f"使用COM埠: COM{self.com_port}")
        except Exception as e:
            logger.warning(f"讀取健保卡設定檔失敗: {e}")
            self.csfsim_path = r"C:\NHI\BIN\csfsim.exe"
            self.offline_mode = True
            self.offline_auto_print = True
            self.auto_detect_com = True
            self.com_port = self._detect_com_port()
        
        # 嘗試載入健保卡 DLL
        try:
            self.nhi_dll = NHICardDLL(dll_path)
            self.use_dll = True
            logger.info(f"健保卡讀取模組初始化完成 (使用 DLL: {self.nhi_dll.dll_path})")
        except NHICardDLLError as e:
            self.use_dll = False
            logger.warning(f"無法載入健保卡 DLL: {e}")
            if self.offline_mode:
                logger.info("健保卡讀取模組初始化完成 (離線模式)")
            else:
                logger.info("健保卡讀取模組初始化完成 (模擬模式)")
        except Exception as e:
            self.use_dll = False
            logger.error(f"健保卡讀取模組初始化時發生未知錯誤: {e}")
            logger.info("健保卡讀取模組初始化完成 (離線模式)" if self.offline_mode else "健保卡讀取模組初始化完成 (模擬模式)")

    def _detect_com_port(self):
        """自動偵測讀卡機COM埠"""
        logger.info("開始自動偵測讀卡機COM埠...")
        
        # 取得所有可用的COM埠
        available_ports = list(serial.tools.list_ports.comports())
        logger.info(f"發現 {len(available_ports)} 個COM埠: {[port.device for port in available_ports]}")
        
        # 常見的讀卡機COM埠 (按優先順序)
        preferred_ports = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        
        # 先檢查健保署設定檔案中的COM埠
        nhi_com_port = self._get_nhi_com_port()
        if nhi_com_port:
            logger.info(f"從健保署設定檔案找到COM埠: COM{nhi_com_port}")
            if self._test_com_port(nhi_com_port):
                logger.info(f"成功偵測到讀卡機: COM{nhi_com_port}")
                return nhi_com_port
        
        # 測試常見的COM埠
        for port_num in preferred_ports:
            if self._test_com_port(port_num):
                logger.info(f"成功偵測到讀卡機: COM{port_num}")
                return port_num
        
        # 如果都沒找到，使用預設值
        logger.warning("無法自動偵測讀卡機COM埠，使用預設值 COM3")
        return 3
    
    def _get_nhi_com_port(self):
        """從健保署設定檔案取得COM埠"""
        try:
            # 讀取 hisCom.xml
            import xml.etree.ElementTree as ET
            hiscom_path = r"C:\NHI\INI\hisCom.xml"
            if os.path.exists(hiscom_path):
                tree = ET.parse(hiscom_path)
                root = tree.getroot()
                com_port = root.text
                if com_port and com_port.isdigit():
                    return int(com_port)
        except Exception as e:
            logger.debug(f"無法讀取 hisCom.xml: {e}")
        
        try:
            # 讀取 csdll.ini
            config = configparser.ConfigParser()
            csdll_path = r"C:\NHI\INI\csdll.ini"
            if os.path.exists(csdll_path):
                config.read(csdll_path, encoding='utf-8')
                com_setting = config.get('CS', 'COM', fallback='')
                if 'COMX' in com_setting:
                    # 從 COMX1, COMX2 等格式中提取數字
                    port_num = com_setting.replace('COMX', '')
                    if port_num.isdigit():
                        return int(port_num)
        except Exception as e:
            logger.debug(f"無法讀取 csdll.ini: {e}")
        
        return None
    
    def _test_com_port(self, port_num):
        """測試COM埠是否有讀卡機連接"""
        port_name = f"COM{port_num}"
        logger.debug(f"測試COM埠: {port_name}")
        
        try:
            # 檢查COM埠是否存在
            available_ports = [port.device for port in serial.tools.list_ports.comports()]
            if port_name not in available_ports:
                logger.debug(f"COM埠 {port_name} 不存在")
                return False
            
            # 嘗試開啟COM埠
            ser = serial.Serial(
                port=port_name,
                baudrate=9600,  # 常見的讀卡機波特率
                timeout=1,
                write_timeout=1
            )
            
            # 嘗試讀取資料
            time.sleep(0.1)
            data = ser.read(1)
            ser.close()
            
            logger.debug(f"COM埠 {port_name} 測試成功")
            return True
            
        except Exception as e:
            logger.debug(f"COM埠 {port_name} 測試失敗: {e}")
            return False
    
    def set_com_port(self, port_num):
        """手動設定COM埠"""
        if self._test_com_port(port_num):
            self.com_port = port_num
            logger.info(f"手動設定COM埠成功: COM{port_num}")
            
            # 更新設定檔
            self._update_config_com_port(port_num)
            return True
        else:
            logger.warning(f"COM埠 {port_num} 測試失敗，無法設定")
            return False
    
    def _update_config_com_port(self, port_num):
        """更新設定檔中的COM埠"""
        try:
            config = configparser.ConfigParser()
            config.read('config.ini', encoding='utf-8')
            
            if not config.has_section('健保卡設定'):
                config.add_section('健保卡設定')
            
            config.set('健保卡設定', 'com_port', str(port_num))
            config.set('健保卡設定', 'auto_detect_com', 'false')
            
            with open('config.ini', 'w', encoding='utf-8') as f:
                config.write(f)
            
            logger.info(f"已更新設定檔COM埠為: COM{port_num}")
        except Exception as e:
            logger.warning(f"更新設定檔失敗: {e}")
    
    def get_available_com_ports(self):
        """取得所有可用的COM埠列表"""
        ports = []
        available_ports = serial.tools.list_ports.comports()
        
        for port in available_ports:
            port_info = {
                'device': port.device,
                'description': port.description,
                'manufacturer': port.manufacturer,
                'hwid': port.hwid
            }
            ports.append(port_info)
        
        return ports

    def launch_csfsim(self):
        """啟動健保署讀卡機控制軟體"""
        if not self.csfsim_path or not os.path.exists(self.csfsim_path):
            logger.warning(f"找不到健保署讀卡機控制軟體: {self.csfsim_path}")
            return False
            
        try:
            logger.info(f"嘗試啟動健保署讀卡機控制軟體: {self.csfsim_path}")
            
            # 檢查健保署環境設定
            self._check_nhi_environment()
            
            # 方法一：使用 subprocess.Popen 啟動
            try:
                process = subprocess.Popen([self.csfsim_path], 
                                         shell=True, 
                                         stdout=subprocess.DEVNULL, 
                                         stderr=subprocess.DEVNULL)
                logger.info("健保署讀卡機控制軟體啟動指令已發送 (方法一)")
                
                # 等待一段時間確認程式是否成功啟動
                time.sleep(3)
                
                # 檢查程序是否仍在運行
                if process.poll() is None:
                    logger.info("健保署讀卡機控制軟體已成功啟動")
                    return True
                else:
                    logger.warning("方法一失敗，嘗試方法二")
            except Exception as e:
                logger.warning(f"方法一啟動失敗: {e}")
            
            # 方法二：使用 os.startfile (Windows 專用)
            try:
                os.startfile(self.csfsim_path)
                logger.info("健保署讀卡機控制軟體啟動指令已發送 (方法二)")
                time.sleep(2)
                return True
            except Exception as e:
                logger.warning(f"方法二啟動失敗: {e}")
            
            # 方法三：使用 subprocess.run
            try:
                result = subprocess.run([self.csfsim_path], 
                                      shell=True, 
                                      timeout=5,
                                      stdout=subprocess.DEVNULL, 
                                      stderr=subprocess.DEVNULL)
                logger.info("健保署讀卡機控制軟體啟動指令已發送 (方法三)")
                return True
            except Exception as e:
                logger.warning(f"方法三啟動失敗: {e}")
            
            logger.error("所有啟動方法都失敗")
            return False
            
        except Exception as e:
            logger.error(f"啟動健保署讀卡機控制軟體失敗: {e}")
            return False
    
    def _check_nhi_environment(self):
        """檢查健保署環境設定"""
        logger.info("檢查健保署環境設定...")
        
        # 檢查健保署目錄結構
        nhi_base = r"C:\NHI"
        if not os.path.exists(nhi_base):
            logger.warning(f"健保署目錄不存在: {nhi_base}")
            return
        
        # 檢查關鍵檔案
        key_files = [
            r"C:\NHI\INI\csdll.ini",
            r"C:\NHI\INI\hisCom.xml",
            r"C:\NHI\INI\ReaderCFG0.ini"
        ]
        
        for file_path in key_files:
            if os.path.exists(file_path):
                logger.info(f"找到健保署設定檔案: {file_path}")
            else:
                logger.warning(f"健保署設定檔案不存在: {file_path}")
        
        # 檢查 SAM 檔案
        sam_dirs = [
            r"C:\NHI\SAM\COMX1",
            r"C:\NHI\SAM\COMX2", 
            r"C:\NHI\SAM\COMX3"
        ]
        
        for sam_dir in sam_dirs:
            if os.path.exists(sam_dir):
                sam_files = os.listdir(sam_dir)
                logger.info(f"找到 SAM 目錄 {sam_dir}，包含 {len(sam_files)} 個檔案")
            else:
                logger.warning(f"SAM 目錄不存在: {sam_dir}")
        
        # 檢查 COM 埠設定
        try:
            import configparser
            config = configparser.ConfigParser()
            config.read(r"C:\NHI\INI\csdll.ini", encoding='utf-8')
            com_setting = config.get('CS', 'COM', fallback='')
            logger.info(f"健保署 COM 埠設定: {com_setting}")
            
            # 讀取 hisCom.xml 中的 COM 埠設定
            try:
                import xml.etree.ElementTree as ET
                tree = ET.parse(r"C:\NHI\INI\hisCom.xml")
                root = tree.getroot()
                com_port = root.text
                logger.info(f"hisCom.xml COM 埠設定: COM{com_port}")
            except Exception as e:
                logger.warning(f"無法讀取 hisCom.xml: {e}")
                
        except Exception as e:
            logger.warning(f"無法讀取健保署設定檔案: {e}")
        
        logger.info("健保署環境檢查完成")
    
    def _simulate_read_card(self):
        """
        讀取健保卡
        優先使用 DLL 直接呼叫，失敗時使用 csfsim 控制軟體
        """
        logger.info("開始讀取健保卡...")
        
        # 優先嘗試使用 DLL 直接讀取
        if self.use_dll:
            try:
                logger.info("使用 DLL 直接讀取健保卡")
                card_data = self.nhi_dll.read_card()
                logger.info(f"DLL 讀取成功，病人: {card_data.get('FULL_NAME', 'N/A')}")
                return card_data
            except Exception as e:
                logger.warning(f"DLL 讀取失敗: {e}，嘗試使用 csReadCard")
                # DLL 讀取失敗，嘗試使用 csReadCard
                try:
                    card_data = self._read_card_with_csreadcard()
                    logger.info(f"csReadCard 讀取成功，病人: {card_data.get('FULL_NAME', 'N/A')}")
                    return card_data
                except Exception as e2:
                    logger.error(f"csReadCard 也失敗: {e2}")
                    # 兩種方法都失敗，拋出錯誤
                    raise CardReaderError(f"健保卡讀取失敗: {e2}")
        
        # 如果沒有 DLL，直接嘗試 csReadCard
        else:
            try:
                logger.info("使用 csReadCard 讀取健保卡")
                card_data = self._read_card_with_csreadcard()
                logger.info(f"csReadCard 讀取成功，病人: {card_data.get('FULL_NAME', 'N/A')}")
                return card_data
            except Exception as e:
                logger.error(f"csReadCard 讀取失敗: {e}")
                # csReadCard 失敗，嘗試使用 csfsim 作為備用方案
                return self._fallback_to_csfsim()
    
    def _read_card_with_csreadcard(self):
        """使用 csReadCard 函式讀取健保卡"""
        try:
            import ctypes
            from ctypes import c_char_p, c_int, byref, create_string_buffer
            
            # 載入健保署 DLL
            dll_path = r"C:\NHI\BIN\NHIDLL.dll"  # 預設路徑
            if not os.path.exists(dll_path):
                # 嘗試其他可能的路徑
                possible_paths = [
                    r"C:\NHI\BIN\csdll.dll",
                    r"C:\Program Files\NHI\BIN\NHIDLL.dll",
                    r"C:\Program Files (x86)\NHI\BIN\NHIDLL.dll"
                ]
                for path in possible_paths:
                    if os.path.exists(path):
                        dll_path = path
                        break
                else:
                    raise CardReaderError("找不到健保署 DLL 檔案")
            
            logger.info(f"載入健保署 DLL: {dll_path}")
            dll = ctypes.CDLL(dll_path)
            
            # 設定 csReadCard 函式簽名
            dll.csReadCard.restype = c_int
            dll.csReadCard.argtypes = [c_char_p]
            
            # 建立緩衝區接收資料
            buffer_size = 1024
            buffer = create_string_buffer(buffer_size)
            
            # 呼叫 csReadCard
            logger.info("呼叫 csReadCard 函式")
            result = dll.csReadCard(buffer)
            
            if result == 0:  # 假設 0 表示成功
                # 取得回傳的字串資料
                raw_data = buffer.value.decode('big5', errors='ignore')
                logger.info(f"csReadCard 回傳原始資料: {repr(raw_data)}")
                
                # 解析資料
                parsed_data = self._parse_csreadcard_data(raw_data)
                return parsed_data
            else:
                raise CardReaderError(f"csReadCard 執行失敗，錯誤碼: {result}")
                
        except Exception as e:
            logger.error(f"csReadCard 呼叫失敗: {e}")
            raise CardReaderError(f"csReadCard 呼叫失敗: {e}")
    
    def _parse_csreadcard_data(self, raw_data):
        """解析 csReadCard 回傳的資料"""
        try:
            logger.info(f"開始解析 csReadCard 資料: {repr(raw_data)}")
            
            # 移除空白字符
            data = raw_data.strip()
            
            if not data:
                raise CardReaderError("csReadCard 回傳空資料")
            
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
                raise CardReaderError("無法解析 csReadCard 回傳的資料格式")
            
            logger.info(f"解析結果: {parsed_data}")
            return parsed_data
            
        except Exception as e:
            logger.error(f"解析 csReadCard 資料失敗: {e}")
            raise CardReaderError(f"解析健保卡資料失敗: {e}")
    
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
    
    def _fallback_to_csfsim(self):
        """備用方案：使用 csfsim 控制軟體"""
        # 嘗試使用 csfsim.exe 讀取健保卡
        if os.path.exists(self.csfsim_path):
            logger.info("csReadCard 失敗，嘗試使用 csfsim 作為備用方案")
            raise CardReaderError(
                "請使用健保署讀卡機控制軟體讀取健保卡。\n\n"
                "操作步驟：\n"
                "1. 確認健保卡已正確插入讀卡機\n"
                "2. 開啟健保署讀卡機控制軟體\n"
                "3. 在軟體中完成讀卡操作\n"
                "4. 或切換到手工模式直接輸入資料"
            )
        else:
            if self.offline_mode:
                raise CardReaderError(
                    "健保卡讀取失敗。\n\n"
                    "請聯繫 IT 部門確認軟體安裝狀態。"
                )
            else:
                raise CardReaderError(
                    "健保卡讀取失敗：未找到健保署讀卡機控制軟體。\n\n"
                    "請聯繫 IT 部門確認：\n"
                    "1. 健保署讀卡機控制軟體已安裝\n"
                    "2. 讀卡機驅動程式已正確安裝\n\n"
                    "如需協助，請聯繫系統管理員。"
                )

    def _test_mode_read_card(self):
        """
        離線模式讀卡方法
        提示使用者手動輸入病人資料，避免使用固定測試資料
        """
        logger.info("離線模式：提示使用者手動輸入病人資料...")
        
        # 在離線模式下，不提供固定的測試資料
        # 而是提示使用者手動輸入真實的病人資料
        raise CardReaderError(
            "離線模式：請手動輸入病人資料\n\n"
            "由於無法讀取健保卡，請在程式中手動輸入：\n"
            "1. 病人身分證字號\n"
            "2. 病人姓名\n"
            "3. 出生日期\n"
            "4. 性別\n\n"
            "請確保輸入的資料正確無誤，避免醫療錯誤。"
        )

    def read_patient_info(self, callback=None, error_callback=None):
        """
        讀取病人資訊的公開方法。
        使用非同步方式避免阻塞 UI。
        """
        if self.is_reading:
            logger.warning("讀卡機正在讀取中，請稍候")
            if error_callback:
                error_callback(CardReaderError("讀卡機正在讀取中，請稍候"))
            return

        def _read_and_process():
            self.is_reading = True
            try:
                raw_data = self._simulate_read_card()
                if callback:
                    callback(raw_data)
            except Exception as e:
                logger.error(f"讀取健保卡失敗: {e}")
                if error_callback:
                    error_callback(e)
            finally:
                self.is_reading = False

        thread = threading.Thread(target=_read_and_process)
        thread.daemon = True  # 讓程式結束時自動終止執行緒
        thread.start()
