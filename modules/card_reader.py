# 健保卡讀取模組
import datetime
import time
import threading
import random
import os
import subprocess
import configparser
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
        
        # 讀取設定檔
        try:
            config = configparser.ConfigParser()
            config.read('config.ini', encoding='utf-8')
            self.csfsim_path = config.get('健保卡設定', 'csfsim_path', fallback=r"C:\NHI\BIN\csfsim.exe")
            self.offline_mode = config.getboolean('健保卡設定', 'offline_mode', fallback=True)
            self.offline_auto_print = config.getboolean('健保卡設定', 'offline_auto_print', fallback=True)
            logger.info(f"健保卡讀卡機控制軟體路徑: {self.csfsim_path}")
            logger.info(f"離線模式: {'啟用' if self.offline_mode else '停用'}")
        except Exception as e:
            logger.warning(f"讀取健保卡設定檔失敗: {e}")
            self.csfsim_path = r"C:\NHI\BIN\csfsim.exe"
            self.offline_mode = True
            self.offline_auto_print = True
        
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

    def launch_csfsim(self):
        """啟動健保署讀卡機控制軟體"""
        if not self.csfsim_path or not os.path.exists(self.csfsim_path):
            logger.warning(f"找不到健保署讀卡機控制軟體: {self.csfsim_path}")
            return False
            
        try:
            logger.info(f"嘗試啟動健保署讀卡機控制軟體: {self.csfsim_path}")
            subprocess.Popen(self.csfsim_path, shell=True)
            logger.info("健保署讀卡機控制軟體啟動指令已發送")
            return True
        except Exception as e:
            logger.error(f"啟動健保署讀卡機控制軟體失敗: {e}")
            return False
    
    def _simulate_read_card(self):
        """
        讀取健保卡
        支援離線模式和 DLL 讀取
        """
        logger.info("開始讀取健保卡...")
        
        # 如果成功載入 DLL，則使用 DLL 讀取
        if self.use_dll:
            try:
                # 使用 GNT 或標準 DLL 讀取健保卡
                logger.info(f"使用 DLL 讀取健保卡: {self.nhi_dll.dll_path}")
                raw_data = self.nhi_dll.read_card()
                
                # 檢查資料完整性
                if not raw_data or not raw_data.get('ID_NUMBER') or not raw_data.get('FULL_NAME'):
                    logger.error("DLL 讀取的資料不完整")
                    raise CardReaderError("讀取的健保卡資料不完整，請重新插卡再試")
                
                logger.info(f"成功讀取健保卡，病人: {raw_data.get('FULL_NAME')}")
                return raw_data
                
            except NHICardDLLError as e:
                logger.error(f"DLL 讀取健保卡失敗: {e}")
                
                # 如果是離線模式，提供簡化的錯誤訊息
                if self.offline_mode:
                    raise CardReaderError(
                        f"健保卡讀取失敗: {e}\n\n"
                        f"離線模式故障排除：\n"
                        f"1. 確認健保卡已正確插入讀卡機\n"
                        f"2. 確認讀卡機已連接並開啟電源\n"
                        f"3. 確認 COM 埠設定正確\n"
                        f"4. 重新插拔健保卡再試\n\n"
                        f"如問題持續，請聯繫 IT 部門。"
                    )
                else:
                    # 嘗試啟動健保署讀卡機控制軟體
                    self.launch_csfsim()
                    raise CardReaderError(
                        f"健保卡讀取失敗: {e}\n\n"
                        f"請確認：\n"
                        f"1. 健保卡已正確插入讀卡機\n"
                        f"2. 讀卡機已正確連接並開啟\n"
                        f"3. DLL 檔案路徑正確\n"
                        f"4. 程式有足夠權限存取 DLL 檔案\n\n"
                        f"已嘗試啟動健保署讀卡機控制軟體。"
                    )
                
            except Exception as e:
                logger.error(f"讀取健保卡時發生未知錯誤: {e}")
                raise CardReaderError(f"讀取健保卡時發生未知錯誤: {e}")
        
        # 如果沒有載入 DLL
        else:
            if self.offline_mode:
                # 離線模式：提示使用者檢查硬體
                time.sleep(1)  # 模擬讀卡時間
                logger.warning("離線模式：未找到可用的健保卡 DLL")
                raise CardReaderError(
                    "離線模式健保卡讀取失敗。\n\n"
                    "請檢查：\n"
                    "1. 健保卡是否正確插入讀卡機\n"
                    "2. 讀卡機是否正確連接並開啟電源\n"
                    "3. GNT 程式是否已正確安裝\n"
                    "4. COM 埠設定是否正確\n\n"
                    "如需協助，請聯繫 IT 部門。"
                )
            else:
                # 一般模式：提供完整安裝指引
                time.sleep(1)  # 模擬讀卡時間
                logger.warning("未找到健保卡 DLL")
                
                if self.launch_csfsim():
                    raise CardReaderError(
                        "健保卡讀取失敗：未找到健保卡 DLL。\n\n"
                        "已嘗試啟動健保署讀卡機控制軟體，請稍後再試。\n\n"
                        "請聯繫 IT 部門確認系統安裝。"
                    )
                else:
                    raise CardReaderError(
                        "健保卡讀取失敗：未找到健保卡系統。\n\n"
                        "請聯繫 IT 部門確認：\n"
                        "1. 健保署讀卡機控制軟體已安裝\n"
                        "2. GNT 抽血櫃台程式已安裝\n"
                        "3. 讀卡機驅動程式已正確安裝\n\n"
                        "如需協助，請聯繫系統管理員。"
                    )

    def _test_mode_read_card(self):
        """
        僅供測試使用的讀卡方法
        警告：此方法僅用於程式測試，不可用於實際醫療環境
        """
        logger.info("測試模式：模擬讀取健保卡...")
        time.sleep(1)  # 縮短測試時間
        
        # 測試用假資料 - 明顯標示為測試
        test_data = {
            "ID_NUMBER": "TEST123456",
            "FULL_NAME": "測試病人",
            "BIRTH_DATE": "19900101"
        }
        
        logger.info("測試模式：返回測試資料")
        return test_data

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
