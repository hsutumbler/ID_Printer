# Error Handling & Logging 模組
import logging
import os
from datetime import datetime

class AppLogger:
    def __init__(self, log_file="logs/app_log.log"):
        # 確保日誌目錄存在
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        self.logger = logging.getLogger("MedicalCardApp")
        self.logger.setLevel(logging.INFO)

        # 避免重複添加處理器
        if not self.logger.handlers:
            # 檔案處理器
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setLevel(logging.INFO)
            
            # 控制台處理器
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)

            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            console_handler.setFormatter(formatter)

            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)

    def get_logger(self):
        return self.logger

# 初始化全域日誌器
logger = AppLogger().get_logger()
