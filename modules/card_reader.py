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
        支援離線模式和 csfsim 控制軟體
        """
        logger.info("開始讀取健保卡...")
        
        # 嘗試使用 csfsim.exe 讀取健保卡
        if os.path.exists(self.csfsim_path):
            try:
                # 啟動健保署讀卡機控制軟體
                logger.info(f"使用健保署讀卡機控制軟體讀取健保卡: {self.csfsim_path}")
                success = self.launch_csfsim()
                
                if success:
                    # 等待使用者在 csfsim 中操作讀卡機
                    logger.info("健保署讀卡機控制軟體已啟動，請在其中讀取健保卡")
                    
                    # 由於我們無法直接從 csfsim 獲取資料，暫時拋出錯誤提示使用者
                    # 在實際環境中，應該實作與 csfsim 的資料交換機制
                    raise CardReaderError(
                        "請在健保署讀卡機控制軟體中完成讀卡操作。\n\n"
                        "操作步驟：\n"
                        "1. 確認健保卡已正確插入讀卡機\n"
                        "2. 在 csfsim 視窗中點擊讀取\n"
                        "3. 完成後關閉 csfsim 視窗\n"
                        "4. 回到本程式重新點擊讀取按鈕"
                    )
                else:
                    # 啟動失敗 - 可能是安全模組檔案問題
                    error_msg = (
                        "健保署讀卡機控制軟體啟動失敗。\n\n"
                        "常見問題解決方法：\n"
                        "1. 檢查健保卡是否正確插入讀卡機\n"
                        "2. 確認讀卡機電源已開啟\n"
                        "3. 重新安裝健保署讀卡機控制軟體\n"
                        "4. 以系統管理員身分執行本程式\n\n"
                        "如果出現「安全模組檔目錄錯誤」：\n"
                        "- 請聯繫 IT 部門重新設定健保署環境\n"
                        "- 可能需要重新下載安全模組檔案"
                    )
                    raise CardReaderError(error_msg)
            except Exception as e:
                logger.error(f"使用健保署讀卡機控制軟體讀取健保卡失敗: {e}")
                raise CardReaderError(f"讀取健保卡時發生錯誤: {e}")
        
        # 如果沒有找到 csfsim.exe
        else:
            if self.offline_mode:
                # 離線模式：簡化錯誤訊息
                time.sleep(1)  # 模擬讀卡時間
                logger.warning(f"離線模式：未找到健保署讀卡機控制軟體: {self.csfsim_path}")
                raise CardReaderError(
                    "未找到健保署讀卡機控制軟體。\n\n"
                    "請聯繫 IT 部門確認軟體安裝狀態。"
                )
            else:
                # 一般模式：提供完整安裝指引
                time.sleep(1)  # 模擬讀卡時間
                logger.warning(f"未找到健保署讀卡機控制軟體: {self.csfsim_path}")
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
