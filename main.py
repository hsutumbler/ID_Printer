#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健保卡資料讀取與標籤列印系統
檢驗科抽血櫃台專用

作者: AI Assistant
版本: 1.0
日期: 2025/10/08
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox

# 將 modules 目錄加入 Python 路徑
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

try:
    from modules.ui import create_app
    from modules.logger import logger
except ImportError as e:
    print(f"模組匯入失敗: {e}")
    print("請確認所有必要的模組檔案都存在")
    sys.exit(1)

def check_dependencies():
    """檢查相依套件"""
    missing_packages = []
    
    # 檢查 tkinter (通常內建)
    try:
        import tkinter
    except ImportError:
        missing_packages.append("tkinter")
    
    # 檢查 reportlab (可選)
    try:
        import reportlab
        logger.info("ReportLab 已安裝，將使用 PDF 列印")
    except ImportError:
        logger.warning("ReportLab 未安裝，將使用文字檔列印")
    
    if missing_packages:
        error_msg = f"缺少必要套件: {', '.join(missing_packages)}"
        logger.error(error_msg)
        messagebox.showerror("相依套件錯誤", error_msg)
        return False
    
    return True

def get_dll_path():
    """從配置檔案讀取 DLL 路徑"""
    try:
        import configparser
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')
        
        dll_path = config.get('健保卡設定', 'dll_path', fallback='')
        if dll_path and os.path.exists(dll_path):
            logger.info(f"從配置檔案讀取 DLL 路徑: {dll_path}")
            return dll_path
        else:
            logger.info("使用預設 DLL 路徑")
            return None
    except Exception as e:
        logger.warning(f"讀取 DLL 路徑時發生錯誤: {e}")
        return None

def main():
    """主程式入口"""
    try:
        logger.info("=" * 50)
        logger.info("健保卡資料讀取與標籤列印系統啟動")
        logger.info("=" * 50)
        
        # 檢查相依套件
        if not check_dependencies():
            return 1
        
        # 讀取 DLL 路徑
        dll_path = get_dll_path()
        
        # 建立並啟動應用程式
        root, app = create_app(dll_path)
        
        logger.info("應用程式介面已建立，開始執行主迴圈")
        
        # 啟動 GUI 主迴圈
        root.mainloop()
        
        logger.info("應用程式正常結束")
        return 0
        
    except KeyboardInterrupt:
        logger.info("使用者中斷程式執行")
        return 0
        
    except Exception as e:
        error_msg = f"程式執行發生嚴重錯誤: {e}"
        logger.error(error_msg)
        
        # 嘗試顯示錯誤對話框
        try:
            root = tk.Tk()
            root.withdraw()  # 隱藏主視窗
            messagebox.showerror("系統錯誤", error_msg)
        except:
            print(error_msg)
        
        return 1

if __name__ == "__main__":
    # 設定工作目錄為程式所在目錄
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    # 執行主程式
    exit_code = main()
    sys.exit(exit_code)
