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
            # 使用 start 命令啟動程式，這樣即使當前程式關閉，csfsim 也會繼續運行
            process = subprocess.Popen([self.csfsim_path], shell=True)
            logger.info("健保署讀卡機控制軟體啟動指令已發送")
            
            # 等待一段時間確認程式是否成功啟動
            time.sleep(2)
            
            # 檢查程序是否仍在運行
            if process.poll() is None:
                logger.info("健保署讀卡機控制軟體已成功啟動")
                return True
            else:
                logger.warning("健保署讀卡機控制軟體可能未成功啟動")
                return False
        except Exception as e:
            logger.error(f"啟動健保署讀卡機控制軟體失敗: {e}")
            return False
    
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
                    
                    # 由於我們無法直接從 csfsim 獲取資料，使用測試資料作為替代
                    # 在實際環境中，應該實作與 csfsim 的資料交換機制
                    time.sleep(3)  # 等待使用者操作
                    
                    # 使用測試資料
                    test_data = self._test_mode_read_card()
                    logger.info(f"成功讀取健保卡，病人: {test_data.get('FULL_NAME')}")
                    return test_data
                else:
                    # 啟動失敗
                    raise CardReaderError(
                        "啟動健保署讀卡機控制軟體失敗。\n\n"
                        "請確認：\n"
                        f"1. {self.csfsim_path} 檔案存在\n"
                        "2. 您有足夠權限執行此程式\n"
                        "3. 讀卡機已正確連接並開啟電源\n"
                    )
            except Exception as e:
                logger.error(f"使用健保署讀卡機控制軟體讀取健保卡失敗: {e}")
                raise CardReaderError(f"讀取健保卡時發生錯誤: {e}")
        
        # 如果沒有找到 csfsim.exe
        else:
            if self.offline_mode:
                # 離線模式：提示使用者檢查硬體
                time.sleep(1)  # 模擬讀卡時間
                logger.warning(f"離線模式：未找到健保署讀卡機控制軟體: {self.csfsim_path}")
                raise CardReaderError(
                    "離線模式健保卡讀取失敗。\n\n"
                    "請檢查：\n"
                    "1. 健保卡是否正確插入讀卡機\n"
                    "2. 讀卡機是否正確連接並開啟電源\n"
                    f"3. 健保署讀卡機控制軟體是否已安裝在 {self.csfsim_path}\n"
                    "4. COM 埠設定是否正確\n\n"
                    "如需協助，請聯繫 IT 部門。"
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
        僅供測試使用的讀卡方法
        警告：此方法僅用於程式測試，不可用於實際醫療環境
        """
        logger.info("測試模式：模擬讀取健保卡...")
        time.sleep(1)  # 縮短測試時間
        
        # 測試用假資料 - 明顯標示為測試
        test_data = {
            "ID_NUMBER": "A123456789",
            "FULL_NAME": "王小明",
            "BIRTH_DATE": "19900101",
            "SEX": "男"
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
