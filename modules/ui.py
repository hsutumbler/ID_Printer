# 使用者介面模組
import tkinter as tk
from tkinter import ttk, messagebox, font
import datetime
import os
from .logger import logger
from .card_reader import CardReader, CardReaderError
from .data_processor import DataProcessor, DataProcessingError
from .print_manager import PrintManager, PrintManagerError
from .record_manager import RecordManager, RecordManagerError

class MedicalCardApp:
    def __init__(self, root, dll_path=None):
        self.root = root
        self.root.title("健保卡資料讀取與標籤列印系統 v1.0")
        self.root.geometry("650x650")  # 調整視窗大小以適應頁籤
        self.root.resizable(True, True)  # 允許調整視窗大小
        self.root.minsize(600, 600)  # 設定最小視窗大小
        
        # 設定全局字體為微軟正黑體
        self.default_font = "微軟正黑體"
        self.setup_fonts()
        
        # 初始化各模組
        try:
            # 初始化健保卡讀取模組，傳入 DLL 路徑
            self.card_reader = CardReader(dll_path)
            
            # 初始化其他模組
            self.data_processor = DataProcessor()
            self.print_manager = PrintManager()
            self.record_manager = RecordManager()
            
            # 記錄 DLL 使用狀態和離線模式
            self.dll_path = dll_path
            self.dll_enabled = self.card_reader.use_dll
            self.offline_mode = self.card_reader.offline_mode
            self.offline_auto_print = self.card_reader.offline_auto_print
            
            logger.info("所有模組初始化完成")
        except Exception as e:
            logger.error(f"模組初始化失敗: {e}")
            messagebox.showerror("初始化錯誤", f"系統初始化失敗: {e}")
            return
        
        # 病人資料
        self.current_patient_data = None
        
        # 建立 UI
        self.create_widgets()
        
        # 顯示今日統計
        self.update_statistics()
        
    def setup_fonts(self):
        """設定全局字體為微軟正黑體"""
        # 獲取所有 Tkinter 字體名稱
        font_names = list(font.families())
        
        # 檢查微軟正黑體是否可用
        if "微軟正黑體" in font_names:
            self.default_font = "微軟正黑體"
        elif "Microsoft JhengHei" in font_names:
            self.default_font = "Microsoft JhengHei"
        elif "Microsoft JhengHei UI" in font_names:
            self.default_font = "Microsoft JhengHei UI"
        else:
            # 如果都不可用，使用系統預設字體
            self.default_font = "TkDefaultFont"
            logger.warning("找不到微軟正黑體字體，使用系統預設字體")
            
        # 設定 Tkinter 預設字體
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family=self.default_font)
        
        text_font = font.nametofont("TkTextFont")
        text_font.configure(family=self.default_font)
        
        fixed_font = font.nametofont("TkFixedFont")
        fixed_font.configure(family=self.default_font)
        
        # 設定 ttk 主題
        style = ttk.Style()
        style.configure(".", font=(self.default_font, 10))
        style.configure("TButton", font=(self.default_font, 10))
        style.configure("TLabel", font=(self.default_font, 10))
        style.configure("TEntry", font=(self.default_font, 10))
        
        # 創建一個主要按鈕樣式
        style.configure("primary.TButton", font=(self.default_font, 10, "bold"))
        
        logger.info(f"設定全局字體為: {self.default_font}")
        
        logger.info("使用者介面初始化完成")

    def create_widgets(self):
        """建立 UI 元件"""
        # 建立頁籤式設計
        self.tab_control = ttk.Notebook(self.root)
        self.tab_control.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 主畫面頁籤
        main_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(main_tab, text="主畫面")
        
        # 今日統計頁籤
        stats_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(stats_tab, text="今日統計")
        
        # 讀卡機狀態頁籤
        settings_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(settings_tab, text="系統狀態")
        
        # ===== 主畫面頁籤內容 =====
        # 主標題
        title_frame = ttk.Frame(main_tab)
        title_frame.pack(pady=10, fill='x')
        
        title_label = ttk.Label(title_frame, text="健保卡資料讀取與標籤列印系統", 
                               font=(self.default_font, 16, "bold"))
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, text="檢驗科抽血櫃台專用", 
                                  font=(self.default_font, 10), foreground="gray")
        subtitle_label.pack()
        
        # 狀態訊息區 - 使用固定高度避免推移佈局
        status_frame = ttk.LabelFrame(main_tab, text="系統狀態", padding=10)
        status_frame.pack(pady=5, padx=20, fill='x')
        
        # 創建一個固定高度的框架來容納狀態訊息
        status_inner_frame = ttk.Frame(status_frame, height=30)
        status_inner_frame.pack(fill='x')
        status_inner_frame.pack_propagate(False)  # 禁止自動調整大小
        
        self.status_text = tk.StringVar()
        # 根據離線模式設定初始狀態訊息
        if self.offline_mode:
            if self.dll_enabled:
                self.status_text.set("離線模式已啟用 - 請插入健保卡並點擊讀取")
            else:
                self.status_text.set("離線模式 - 未找到健保卡 DLL，請檢查健保署讀卡機控制軟體安裝")
        else:
            self.status_text.set("請插入健保卡並點擊讀取")
        
        # 使用可捲動的文字顯示，以便顯示較長的錯誤訊息
        self.status_label = ttk.Label(status_inner_frame, textvariable=self.status_text, 
                                     font=(self.default_font, 10), foreground="blue",
                                     wraplength=500)  # 設置文字換行寬度
        self.status_label.pack(anchor='w')
        
        # 病人資料顯示區
        patient_frame = ttk.LabelFrame(main_tab, text="病人基本資料", padding=15)
        patient_frame.pack(pady=10, padx=20, fill='x')
        
        # 建立資料顯示欄位
        self.patient_id_var = tk.StringVar()
        self.patient_name_var = tk.StringVar()
        self.patient_dob_var = tk.StringVar()
        self.patient_sex_var = tk.StringVar()
        self.read_time_var = tk.StringVar()
        
        # 身分證字號
        id_frame = ttk.Frame(patient_frame)
        id_frame.pack(fill='x', pady=2)
        ttk.Label(id_frame, text="身分證字號:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        id_display = ttk.Label(id_frame, textvariable=self.patient_id_var, 
                              font=(self.default_font, 13, "bold"), foreground="navy")
        id_display.pack(side='left', fill='x', expand=True)
        
        # 姓名
        name_frame = ttk.Frame(patient_frame)
        name_frame.pack(fill='x', pady=2)
        ttk.Label(name_frame, text="姓名:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        name_display = ttk.Label(name_frame, textvariable=self.patient_name_var, 
                                font=(self.default_font, 13, "bold"), foreground="navy")
        name_display.pack(side='left', fill='x', expand=True)
        
        # 出生年月日
        dob_frame = ttk.Frame(patient_frame)
        dob_frame.pack(fill='x', pady=2)
        ttk.Label(dob_frame, text="出生年月日:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        dob_display = ttk.Label(dob_frame, textvariable=self.patient_dob_var, 
                               font=(self.default_font, 13, "bold"), foreground="navy")
        dob_display.pack(side='left', fill='x', expand=True)
        
        # 性別 (如果有的話)
        sex_frame = ttk.Frame(patient_frame)
        sex_frame.pack(fill='x', pady=2)
        ttk.Label(sex_frame, text="性別:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        sex_display = ttk.Label(sex_frame, textvariable=self.patient_sex_var, 
                               font=(self.default_font, 13, "bold"), foreground="navy")
        sex_display.pack(side='left', fill='x', expand=True)
        
        # 讀取時間
        read_time_frame = ttk.Frame(patient_frame)
        read_time_frame.pack(fill='x', pady=2)
        ttk.Label(read_time_frame, text="讀取時間:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        read_time_display = ttk.Label(read_time_frame, textvariable=self.read_time_var, 
                                     font=(self.default_font, 11, "bold"), foreground="gray")
        read_time_display.pack(side='left', fill='x', expand=True)
        
        # 列印設定區
        print_frame = ttk.LabelFrame(main_tab, text="列印設定", padding=15)
        print_frame.pack(pady=10, padx=20, fill='x')
        
        count_frame = ttk.Frame(print_frame)
        count_frame.pack(fill='x')
        
        ttk.Label(count_frame, text="列印張數:", width=12, anchor='w').pack(side='left')
        self.print_count_var = tk.IntVar(value=1)
        count_spinbox = ttk.Spinbox(count_frame, from_=1, to=10, width=5, 
                                   textvariable=self.print_count_var)
        count_spinbox.pack(side='left', padx=5)
        ttk.Label(count_frame, text="張 (最多10張)").pack(side='left')
        
        # 主要按鈕區
        button_frame = ttk.Frame(main_tab)
        button_frame.pack(pady=20)
        
        # 讀取健保卡按鈕
        self.read_button = ttk.Button(button_frame, text="讀取健保卡", 
                                     command=self.start_read_card, width=15)
        self.read_button.pack(side='left', padx=10)
        
        # 列印標籤按鈕
        self.print_button = ttk.Button(button_frame, text="列印標籤", 
                                      command=self.print_labels, width=15, 
                                      state=tk.DISABLED)
        self.print_button.pack(side='left', padx=10)
        
        # 清除資料按鈕
        clear_button = ttk.Button(button_frame, text="清除資料", 
                                 command=self.clear_data, width=15)
        clear_button.pack(side='left', padx=10)
        
        # 統計資訊變數
        self.stats_text = tk.StringVar()
        self.mode_text = tk.StringVar()
        
        # ===== 今日統計頁籤內容 =====
        stats_main_frame = ttk.Frame(stats_tab, padding=20)
        stats_main_frame.pack(fill='both', expand=True)
        
        # 統計標題
        stats_title = ttk.Label(stats_main_frame, text="今日統計資訊", font=(self.default_font, 14, "bold"))
        stats_title.pack(pady=(10, 20))
        
        # 顯示簡易統計資訊
        stats_info_frame = ttk.LabelFrame(stats_main_frame, text="統計摘要", padding=15)
        stats_info_frame.pack(fill='x', pady=10)
        
        # 更新統計資訊
        self.update_statistics()
        
        # 顯示統計資訊
        stats_label = ttk.Label(stats_info_frame, textvariable=self.stats_text, 
                               font=(self.default_font, 12))
        stats_label.pack(pady=15)
        
        # 詳細統計資訊區
        details_frame = ttk.LabelFrame(stats_main_frame, text="詳細資訊", padding=15)
        details_frame.pack(fill='both', expand=True, pady=15)
        
        # 取得統計資料
        try:
            stats = self.record_manager.get_statistics()
            records = self.record_manager.get_today_records()
            
            # 建立表格式顯示
            details_text = tk.Text(details_frame, wrap='word', height=10, width=60, 
                                 font=(self.default_font, 10))
            details_text.insert('1.0', f"今日讀取次數: {stats['total_reads']} 次\n")
            details_text.insert('end', f"今日列印次數: {stats['total_prints']} 次\n")
            details_text.insert('end', f"今日列印標籤: {stats['total_labels']} 張\n")
            details_text.insert('end', f"今日總記錄數: {stats['total_records']} 筆\n\n")
            
            details_text.insert('end', "最近記錄:\n")
            recent_records = records[-5:] if len(records) > 5 else records
            for i, record in enumerate(reversed(recent_records), 1):
                details_text.insert('end', f"{i}. {record.get('時間戳記', '')} - {record.get('姓名', '')} - {record.get('操作類型', '')}\n")
                
            details_text.config(state='disabled')  # 設為唯讀
            
            scrollbar = ttk.Scrollbar(details_frame, orient='vertical', command=details_text.yview)
            details_text.configure(yscrollcommand=scrollbar.set)
            
            details_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
            
        except Exception as e:
            error_label = ttk.Label(details_frame, text=f"載入統計資料失敗: {e}", 
                                  foreground="red", font=(self.default_font, 10))
            error_label.pack(pady=20)
        
        # 操作按鈕區
        stats_button_frame = ttk.Frame(stats_main_frame)
        stats_button_frame.pack(pady=15, fill='x')
        
        # 查看詳細統計按鈕
        view_button = ttk.Button(stats_button_frame, text="查看完整統計", 
                               command=self.show_statistics, width=20)
        view_button.pack(side='left', padx=10)
        
        # 匯出按鈕
        def export_stats():
            # 直接調用 show_statistics 中的匯出功能
            self.show_statistics(auto_export=True)
            
        export_button = ttk.Button(stats_button_frame, text="匯出統計資料", 
                                  command=export_stats, width=20)
        export_button.pack(side='right', padx=10)
        
        # ===== 讀卡機設定頁籤內容 =====
        settings_main_frame = ttk.Frame(settings_tab, padding=20)
        settings_main_frame.pack(fill='both', expand=True)
        
        # 設定標題
        settings_title = ttk.Label(settings_main_frame, text="系統狀態與工具", 
                                 font=(self.default_font, 14, "bold"))
        settings_title.pack(pady=(10, 20))
        
        # 顯示讀卡機狀態
        status_frame = ttk.LabelFrame(settings_main_frame, text="讀卡機狀態", padding=15)
        status_frame.pack(fill='x', pady=10)
        
        # 讀卡機狀態
        self.dll_status_text = tk.StringVar()
        self.update_dll_status()
            
        status_label = ttk.Label(status_frame, textvariable=self.dll_status_text, 
                               font=(self.default_font, 11))
        status_label.pack(pady=15)
        
        # 重新檢查 DLL 狀態按鈕
        refresh_status_button = ttk.Button(status_frame, text="重新檢查 DLL 狀態", 
                                         command=self.refresh_dll_status, width=20)
        refresh_status_button.pack(pady=5)
        
        # COM埠設定區
        com_frame = ttk.LabelFrame(settings_main_frame, text="COM埠設定", padding=15)
        com_frame.pack(fill='x', pady=15)
        
        # COM埠狀態顯示
        self.com_status_text = tk.StringVar()
        self.update_com_status()
        
        com_status_label = ttk.Label(com_frame, textvariable=self.com_status_text, 
                                   font=(self.default_font, 11))
        com_status_label.pack(pady=10)
        
        # COM埠設定按鈕
        com_button_frame = ttk.Frame(com_frame)
        com_button_frame.pack(pady=10)
        
        # 自動偵測按鈕
        auto_detect_button = ttk.Button(com_button_frame, text="自動偵測COM埠", 
                                       command=self.auto_detect_com_port, width=20)
        auto_detect_button.pack(side='left', padx=5)
        
        # 手動設定按鈕
        manual_set_button = ttk.Button(com_button_frame, text="手動設定COM埠", 
                                      command=self.show_com_port_dialog, width=20)
        manual_set_button.pack(side='right', padx=5)
        
        # COM埠說明
        com_info_label = ttk.Label(com_frame, 
                                 text="如果讀卡機無法正常工作，請嘗試重新偵測或手動設定COM埠",
                                 foreground="gray", wraplength=500)
        com_info_label.pack(pady=5)
        
        # 印表機設定
        printer_frame = ttk.LabelFrame(settings_main_frame, text="印表機設定", padding=15)
        printer_frame.pack(fill='x', pady=15)
        
        # 顯示當前列印模式
        try:
            # 列印模式
            print_mode = self.print_manager.print_mode.upper()
            
            if print_mode == 'PDF':
                self.mode_text.set(f"列印模式: PDF")
            elif print_mode == 'TEXT':
                self.mode_text.set(f"列印模式: 文字檔")
            else:
                self.mode_text.set(f"列印模式: {print_mode}")
        except Exception as e:
            logger.warning(f"設定列印模式顯示失敗: {e}")
            self.mode_text.set("列印模式: 未知")
            
        mode_label = ttk.Label(printer_frame, textvariable=self.mode_text, 
                              font=(self.default_font, 11))
        mode_label.pack(pady=15)
        
        # 操作選項
        options_frame = ttk.LabelFrame(settings_main_frame, text="操作選項", padding=15)
        options_frame.pack(fill='x', pady=15)
        
        # 啟動健保署讀卡機控制軟體按鈕
        launch_button = ttk.Button(options_frame, text="啟動健保署讀卡機控制軟體", 
                                  command=self.launch_csfsim, width=30)
        launch_button.pack(pady=5)
        
        # 說明文字
        info_label = ttk.Label(options_frame, 
                               text="如果讀卡機無法正常工作，請點擊上方按鈕啟動健保署讀卡機控制軟體",
                               foreground="gray", wraplength=500)
        info_label.pack(pady=5)

    def start_read_card(self):
        """開始讀取健保卡"""
        self.status_text.set("健保卡讀取中，請稍候...")
        self.read_button.config(state=tk.DISABLED)
        self.print_button.config(state=tk.DISABLED)
        logger.info("使用者點擊讀取健保卡按鈕")

        # 在單獨的執行緒中執行讀卡，避免 UI 阻塞
        self.card_reader.read_patient_info(
            callback=self._on_read_success,
            error_callback=self._on_read_error
        )

    def _on_read_success(self, raw_data):
        """讀卡成功回調"""
        try:
            # 處理原始資料
            processed_data = self.data_processor.process_raw_data(raw_data)
            self.current_patient_data = processed_data

            # 更新 UI 顯示
            self.patient_id_var.set(processed_data["id"])
            self.patient_name_var.set(processed_data["name"])
            self.patient_dob_var.set(processed_data["dob"])
            self.patient_sex_var.set(processed_data.get("sex", ""))
            self.read_time_var.set(processed_data["read_time"])
            
            # 根據離線模式設定不同的狀態訊息
            if self.offline_mode:
                self.status_text.set("離線模式 - 讀取成功！請確認資料並設定列印張數")
            else:
                self.status_text.set("讀取成功！請確認資料並設定列印張數")
            
            self.print_button.config(state=tk.NORMAL)
            
            # 顯示確認對話框，提醒醫檢師核對病人身分
            confirm_msg = f"健保卡讀取成功！\n\n" \
                         f"病人: {processed_data['name']}\n" \
                         f"身分證: {processed_data['id']}\n" \
                         f"出生日期: {processed_data['dob']}\n\n" \
                         f"⚠️ 請仔細核對病人身分，確認無誤後再列印標籤！"
            
            messagebox.showinfo("讀取成功", confirm_msg)
            
            # 記錄讀取事件
            try:
                self.record_manager.log_operation(
                    processed_data, 
                    processed_data["read_time"], 
                    0, 
                    "讀取" + (" (離線模式)" if self.offline_mode else "")
                )
            except RecordManagerError as e:
                logger.warning(f"記錄讀取事件失敗: {e}")
            
            # 更新統計
            self.update_statistics()
            
            logger.info(f"讀取健保卡成功，病人: {processed_data['name']}")
            
        except DataProcessingError as e:
            self._on_read_error(e)
        except Exception as e:
            self._on_read_error(e)
        finally:
            self.read_button.config(state=tk.NORMAL)

    def _on_read_error(self, error):
        """讀卡失敗回調"""
        error_msg = str(error)
        
        # 截短錯誤訊息以避免狀態列過長
        short_error = error_msg
        if len(short_error) > 50:
            short_error = short_error[:47] + "..."
            
        # 在狀態列顯示簡短錯誤訊息
        self.status_text.set(f"讀取失敗: {short_error}")
        
        # 檢查是否為離線模式錯誤
        if "離線模式" in error_msg and "請手動輸入病人資料" in error_msg:
            # 在離線模式下，提供手動輸入選項
            if messagebox.askyesno("離線模式", 
                "無法讀取健保卡，是否要手動輸入病人資料？\n\n"
                "請確認：\n"
                "1. 輸入的資料必須正確無誤\n"
                "2. 避免醫療錯誤\n"
                "3. 仔細核對病人身分"):
                self.show_manual_input_dialog()
        else:
            # 在對話框中顯示完整錯誤訊息
            messagebox.showerror("讀取錯誤", f"讀取健保卡失敗:\n{error_msg}")
        
        logger.error(f"讀取健保卡失敗: {error_msg}")
        
        # 確保按鈕可用
        self.read_button.config(state=tk.NORMAL)
        self.print_button.config(state=tk.DISABLED)

    def print_labels(self):
        """列印標籤"""
        try:
            # 驗證列印張數
            print_count = self.print_count_var.get()
            if not isinstance(print_count, int) or print_count <= 0 or print_count > 10:
                messagebox.showwarning("列印錯誤", "列印張數必須為 1-10 之間的整數")
                return

            # 驗證病人資料
            if not self.current_patient_data:
                messagebox.showwarning("列印錯誤", "請先讀取健保卡資料")
                return

            # 直接列印，不顯示確認對話框

            # 開始列印
            self.status_text.set(f"標籤列印中... (共 {print_count} 張)")
            self.print_button.config(state=tk.DISABLED)
            self.read_button.config(state=tk.DISABLED)
            
            logger.info(f"使用者點擊列印標籤按鈕，列印 {print_count} 張")

            # 執行列印
            self.print_manager.print_labels(self.current_patient_data, print_count)
            
            # 記錄列印事件
            print_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            try:
                self.record_manager.log_operation(
                    self.current_patient_data, 
                    print_time, 
                    print_count, 
                    "列印"
                )
            except RecordManagerError as e:
                logger.warning(f"記錄列印事件失敗: {e}")
            
            # 更新狀態和統計
            self.status_text.set(f"列印完成！已列印 {print_count} 張標籤")
            self.update_statistics()
            
            logger.info(f"列印成功，病人: {self.current_patient_data['name']}，共 {print_count} 張")

        except PrintManagerError as e:
            error_msg = str(e)
            
            # 截短錯誤訊息以避免狀態列過長
            short_error = error_msg
            if len(short_error) > 50:
                short_error = short_error[:47] + "..."
                
            # 在狀態列顯示簡短錯誤訊息
            self.status_text.set(f"列印失敗: {short_error}")
            
            # 在對話框中顯示完整錯誤訊息
            messagebox.showerror("列印錯誤", f"列印標籤失敗:\n{error_msg}")
            logger.error(f"列印標籤失敗: {error_msg}")
            
        except Exception as e:
            error_msg = str(e)
            
            # 截短錯誤訊息以避免狀態列過長
            short_error = error_msg
            if len(short_error) > 50:
                short_error = short_error[:47] + "..."
                
            # 在狀態列顯示簡短錯誤訊息
            self.status_text.set(f"列印失敗: {short_error}")
            
            # 在對話框中顯示完整錯誤訊息
            messagebox.showerror("列印錯誤", f"列印標籤失敗:\n{error_msg}")
            logger.error(f"列印標籤失敗: {error_msg}")
            
        finally:
            self.print_button.config(state=tk.NORMAL)
            self.read_button.config(state=tk.NORMAL)
    
    def clear_data(self):
        """清除病人資料"""
        if messagebox.askyesno("確認清除", "確定要清除目前的病人資料嗎？"):
            self.current_patient_data = None
            self.patient_id_var.set("")
            self.patient_name_var.set("")
            self.patient_dob_var.set("")
            self.patient_sex_var.set("")
            self.read_time_var.set("")
            self.status_text.set("請插入健保卡並點擊讀取")
            self.print_button.config(state=tk.DISABLED)
            logger.info("使用者清除病人資料")

    def show_statistics(self, auto_export=False):
        """顯示今日統計，並提供匯出 CSV 功能
        
        參數:
            auto_export: 是否自動開啟匯出對話框
        """
        try:
            stats = self.record_manager.get_statistics()
            records = self.record_manager.get_today_records()
            
            stats_msg = f"""今日統計資訊:
            
讀取次數: {stats['total_reads']} 次
列印次數: {stats['total_prints']} 次  
列印標籤: {stats['total_labels']} 張
總記錄數: {stats['total_records']} 筆

最近 5 筆記錄:"""
            
            # 顯示最近 5 筆記錄
            recent_records = records[-5:] if len(records) > 5 else records
            for i, record in enumerate(reversed(recent_records), 1):
                stats_msg += f"\n{i}. {record.get('時間戳記', '')} - {record.get('姓名', '')} - {record.get('操作類型', '')}"
            
            def export_csv():
                try:
                    from tkinter import filedialog
                    import csv
                    
                    # 取得當前日期作為預設檔名
                    today_str = datetime.datetime.now().strftime("%Y%m%d")
                    default_filename = f"健保卡標籤統計_{today_str}.csv"
                    
                    # 詢問儲存位置
                    file_path = filedialog.asksaveasfilename(
                        initialfile=default_filename,
                        defaultextension=".csv",
                        filetypes=[("CSV 檔案", "*.csv"), ("所有檔案", "*.*")]
                    )
                    
                    if not file_path:
                        return False  # 使用者取消
                    
                    # 寫入 CSV 檔案
                    with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                        writer = csv.writer(f)
                        
                        # 寫入標題列
                        if records and len(records) > 0:
                            writer.writerow(records[0].keys())
                            
                            # 寫入資料列
                            for record in records:
                                writer.writerow(record.values())
                        
                        # 寫入統計資訊
                        writer.writerow([])
                        writer.writerow(["統計資訊"])
                        writer.writerow(["讀取次數", stats['total_reads']])
                        writer.writerow(["列印次數", stats['total_prints']])
                        writer.writerow(["列印標籤", stats['total_labels']])
                        writer.writerow(["總記錄數", stats['total_records']])
                    
                    messagebox.showinfo("匯出成功", f"已成功匯出統計資料至:\n{file_path}")
                    logger.info(f"使用者匯出統計資料: {file_path}")
                    return True
                    
                except Exception as e:
                    messagebox.showerror("匯出錯誤", f"匯出 CSV 檔案失敗:\n{e}")
                    logger.error(f"匯出 CSV 檔案失敗: {e}")
                    return False
            
            # 如果是自動匯出模式，直接開啟匯出對話框
            if auto_export:
                export_csv()
                return
            
            # 創建統計視窗
            stats_window = tk.Toplevel(self.root)
            stats_window.title("今日統計")
            stats_window.geometry("600x500")  # 增加視窗大小
            stats_window.resizable(True, True)
            stats_window.minsize(600, 500)    # 設定最小大小
            stats_window.transient(self.root)
            stats_window.grab_set()
            
            # 置中顯示
            stats_window.update_idletasks()
            x = (stats_window.winfo_screenwidth() // 2) - (600 // 2)
            y = (stats_window.winfo_screenheight() // 2) - (500 // 2)
            stats_window.geometry(f"600x500+{x}+{y}")
            
            # 主要框架
            main_frame = ttk.Frame(stats_window, padding=20)
            main_frame.pack(fill='both', expand=True)
            
            # 標題
            title_label = ttk.Label(main_frame, text="今日統計資訊", font=(self.default_font, 14, "bold"))
            title_label.pack(pady=(0, 20))
            
            # 統計資訊
            stats_frame = ttk.Frame(main_frame)
            stats_frame.pack(fill='both', expand=True)
            
            stats_text = tk.Text(stats_frame, wrap='word', height=15, width=50, font=(self.default_font, 10))
            stats_text.insert('1.0', stats_msg)
            stats_text.config(state='disabled')  # 設為唯讀
            
            scrollbar = ttk.Scrollbar(stats_frame, orient='vertical', command=stats_text.yview)
            stats_text.configure(yscrollcommand=scrollbar.set)
            
            stats_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
            
            # 按鈕框架 - 使用更好的布局
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill='x', pady=20)
            
            # 確保按鈕框架置中
            button_frame.columnconfigure(0, weight=1)
            button_frame.columnconfigure(1, weight=1)
            button_frame.columnconfigure(2, weight=1)
            
            def close_window():
                stats_window.destroy()
            
            # 建立按鈕容器 (置中顯示)
            inner_button_frame = ttk.Frame(button_frame)
            inner_button_frame.grid(row=0, column=1, sticky="nsew")
            
            # 匯出按鈕 - 使用更大的按鈕
            export_button = ttk.Button(
                inner_button_frame, 
                text="匯出 CSV", 
                command=export_csv, 
                width=20,  # 增加寬度
                padding=(10, 5)  # 增加內邊距
            )
            export_button.pack(side='left', padx=20, pady=10)
            
            # 關閉按鈕 - 使用更大的按鈕
            close_button = ttk.Button(
                inner_button_frame, 
                text="關閉", 
                command=close_window, 
                width=20,  # 增加寬度
                padding=(10, 5)  # 增加內邊距
            )
            close_button.pack(side='left', padx=20, pady=10)
            
        except Exception as e:
            messagebox.showerror("統計錯誤", f"取得統計資訊失敗:\n{e}")

    def update_statistics(self):
        """更新統計顯示"""
        try:
            stats = self.record_manager.get_statistics()
            stats_text = f"今日: 讀取 {stats['total_reads']} 次 | 列印 {stats['total_prints']} 次 | 共 {stats['total_labels']} 張標籤"
            self.stats_text.set(stats_text)
        except Exception as e:
            self.stats_text.set("統計資訊載入失敗")
            logger.warning(f"更新統計失敗: {e}")

    def update_dll_status(self):
        """更新 DLL 狀態顯示"""
        try:
            if self.dll_enabled:
                dll_path = self.dll_path or "預設路徑"
                # 檢查 DLL 檔案是否真的存在
                if self.dll_path and os.path.exists(self.dll_path):
                    status_text = f"✅ 已連接健保署讀卡機 DLL\n路徑: {dll_path}"
                else:
                    status_text = f"⚠️ DLL 已載入但檔案可能不存在\n路徑: {dll_path}"
            else:
                status_text = (
                    "❌ 未找到健保署讀卡機 DLL (使用模擬模式)\n\n"
                    "請確認：\n"
                    "1. 已安裝健保署讀卡機控制軟體\n"
                    "2. DLL 檔案路徑正確\n"
                    "3. 系統權限設定正確"
                )
            self.dll_status_text.set(status_text)
        except Exception as e:
            self.dll_status_text.set(f"檢查 DLL 狀態時發生錯誤: {e}")

    def refresh_dll_status(self):
        """重新檢查 DLL 狀態"""
        try:
            # 重新初始化讀卡機模組
            from .card_reader import CardReader
            from .nhi_card_dll import NHICardDLL, NHICardDLLError
            
            logger.info("重新檢查 DLL 狀態...")
            
            # 嘗試重新載入 DLL
            try:
                test_dll = NHICardDLL(self.dll_path)
                self.dll_enabled = True
                self.dll_path = test_dll.dll_path
                messagebox.showinfo("DLL 狀態", f"✅ 成功連接健保署讀卡機 DLL\n路徑: {test_dll.dll_path}")
                logger.info(f"DLL 重新載入成功: {test_dll.dll_path}")
            except NHICardDLLError as e:
                self.dll_enabled = False
                messagebox.showwarning("DLL 狀態", f"❌ 無法載入健保署讀卡機 DLL\n錯誤: {e}")
                logger.warning(f"DLL 重新載入失敗: {e}")
            
            # 更新狀態顯示
            self.update_dll_status()
            
        except Exception as e:
            error_msg = f"重新檢查 DLL 狀態時發生錯誤: {e}"
            messagebox.showerror("錯誤", error_msg)
            logger.error(error_msg)



    def show_card_reader_settings(self):
        """顯示健保卡讀卡機設定對話框"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("健保卡讀卡機設定")
        settings_window.geometry("650x650")  # 大幅增加視窗大小
        settings_window.resizable(True, True)  # 允許調整大小
        settings_window.minsize(650, 650)  # 設定最小大小
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        # 置中顯示
        settings_window.update_idletasks()
        x = (settings_window.winfo_screenwidth() // 2) - (650 // 2)
        y = (settings_window.winfo_screenheight() // 2) - (650 // 2)
        settings_window.geometry(f"650x650+{x}+{y}")
        
        # 設定視窗圖示 (如果有的話)
        try:
            settings_window.iconbitmap(self.root.iconbitmap())
        except:
            pass
        
        # 建立可捲動的畫布和框架
        canvas = tk.Canvas(settings_window)
        scrollbar = ttk.Scrollbar(settings_window, orient="vertical", command=canvas.yview)
        
        # 主要內容框架
        main_frame = ttk.Frame(canvas, padding=20)
        
        # 設定捲動區域
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar.pack(side="right", fill="y")
        
        # 將框架加入到畫布
        canvas_frame = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # 調整畫布大小
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_frame, width=event.width)
        
        canvas.bind("<Configure>", configure_canvas)
        main_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        # 標題
        title_label = ttk.Label(main_frame, text="健保卡讀卡機設定", font=(self.default_font, 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 讀取配置檔案
        import configparser
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')
        
        # 主要設定區域 - 使用 Grid 佈局以提高可讀性
        settings_notebook = ttk.Notebook(main_frame)
        settings_notebook.pack(fill='both', expand=True, pady=10)
        
        # === DLL 設定頁面 ===
        dll_page = ttk.Frame(settings_notebook, padding=15)
        settings_notebook.add(dll_page, text="DLL 設定")
        
        # DLL 路徑設定
        dll_frame = ttk.LabelFrame(dll_page, text="讀卡機 DLL 路徑", padding=15)
        dll_frame.pack(fill='x', pady=10, padx=5, ipady=10)
        
        dll_path_var = tk.StringVar(value=config.get('健保卡設定', 'dll_path', fallback=''))
        
        # 使用 Grid 佈局
        ttk.Label(dll_frame, text="DLL 檔案路徑:", font=(self.default_font, 10)).grid(row=0, column=0, sticky='w', pady=5)
        dll_path_entry = ttk.Entry(dll_frame, textvariable=dll_path_var, width=60)  # 增加寬度
        dll_path_entry.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        
        # 瀏覽按鈕
        def browse_dll():
            from tkinter import filedialog
            dll_file = filedialog.askopenfilename(
                title="選擇健保卡讀卡機 DLL 檔案",
                filetypes=[("DLL 檔案", "*.dll"), ("所有檔案", "*.*")]
            )
            if dll_file:
                dll_path_var.set(dll_file)
        
        browse_button = ttk.Button(dll_frame, text="瀏覽檔案...", command=browse_dll, width=15)
        browse_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 常見路徑提示
        path_hint_label = ttk.Label(
            dll_frame, 
            text="常見路徑: C:\\Program Files\\NHI\\NHICardReader.dll 或 C:\\Windows\\System32\\NHICardReader.dll",
            foreground="gray"
        )
        path_hint_label.grid(row=1, column=0, columnspan=3, sticky='w', pady=5)
        
        # 讀卡機類型設定
        type_frame = ttk.LabelFrame(dll_page, text="讀卡機控制軟體類型", padding=15)
        type_frame.pack(fill='x', pady=10, padx=5, ipady=5)
        
        reader_type_var = tk.StringVar(value=config.get('健保卡設定', 'reader_type', fallback='dll'))
        
        # 使用水平排列的單選按鈕
        type_inner_frame = ttk.Frame(type_frame)
        type_inner_frame.pack(fill='x', pady=5)
        
        ttk.Radiobutton(
            type_inner_frame, text="DLL 函式庫", value="dll", 
            variable=reader_type_var, padding=5
        ).pack(side='left', padx=20)
        
        ttk.Radiobutton(
            type_inner_frame, text="COM 元件", value="com", 
            variable=reader_type_var, padding=5
        ).pack(side='left', padx=20)
        
        ttk.Radiobutton(
            type_inner_frame, text="API 介面", value="api", 
            variable=reader_type_var, padding=5
        ).pack(side='left', padx=20)
        
        # 說明標籤
        ttk.Label(
            type_frame, 
            text="請選擇與健保署提供的讀卡機控制軟體相符的類型",
            foreground="gray"
        ).pack(pady=5)
        
        # === 進階設定頁面 ===
        advanced_page = ttk.Frame(settings_notebook, padding=15)
        settings_notebook.add(advanced_page, text="進階設定")
        
        # 其他設定
        other_frame = ttk.LabelFrame(advanced_page, text="讀卡參數設定", padding=15)
        other_frame.pack(fill='x', pady=10, padx=5)
        
        retry_var = tk.StringVar(value=config.get('健保卡設定', 'retry_count', fallback='3'))
        timeout_var = tk.StringVar(value=config.get('健保卡設定', 'read_timeout', fallback='30'))
        sound_var = tk.BooleanVar(value=config.get('健保卡設定', 'enable_sound', fallback='false').lower() == 'true')
        
        # 使用 Grid 佈局
        param_frame = ttk.Frame(other_frame)
        param_frame.pack(fill='x', pady=10)
        
        # 重試次數
        ttk.Label(param_frame, text="讀卡失敗重試次數:", width=20).grid(row=0, column=0, sticky='w', pady=8)
        retry_entry = ttk.Entry(param_frame, textvariable=retry_var, width=8)
        retry_entry.grid(row=0, column=1, sticky='w', pady=8, padx=5)
        ttk.Label(param_frame, text="次  (建議值: 1-5)", foreground="gray").grid(row=0, column=2, sticky='w', pady=8)
        
        # 逾時時間
        ttk.Label(param_frame, text="讀卡逾時時間:", width=20).grid(row=1, column=0, sticky='w', pady=8)
        timeout_entry = ttk.Entry(param_frame, textvariable=timeout_var, width=8)
        timeout_entry.grid(row=1, column=1, sticky='w', pady=8, padx=5)
        ttk.Label(param_frame, text="秒  (建議值: 10-60)", foreground="gray").grid(row=1, column=2, sticky='w', pady=8)
        
        # 音效設定
        sound_frame = ttk.Frame(other_frame)
        sound_frame.pack(fill='x', pady=5)
        
        sound_check = ttk.Checkbutton(
            sound_frame, text="啟用讀卡音效提示", 
            variable=sound_var, padding=5
        )
        sound_check.pack(anchor='w', pady=5)
        
        # 狀態顯示
        status_frame = ttk.LabelFrame(main_frame, text="讀卡機狀態", padding=15)
        status_frame.pack(fill='x', pady=15, padx=5)
        
        status_text = tk.StringVar()
        if self.dll_enabled:
            status_text.set(f"讀卡機狀態: 已連接 (使用 DLL: {self.dll_path or '預設路徑'})")
        else:
            status_text.set("讀卡機狀態: 未連接")
            
        status_label = ttk.Label(
            status_frame, 
            textvariable=status_text, 
            foreground="blue",
            font=(self.default_font, 10)
        )
        status_label.pack(pady=5)
        
        # 測試按鈕區域
        test_frame = ttk.LabelFrame(main_frame, text="DLL 連線測試", padding=15)
        test_frame.pack(fill='x', pady=10, padx=5)
        
        # 測試說明
        ttk.Label(
            test_frame, 
            text="點擊下方按鈕測試 DLL 連線是否正常。請先設定 DLL 路徑再進行測試。",
            wraplength=600
        ).pack(pady=5)
        
        def test_dll():
            dll_path = dll_path_var.get().strip()
            
            if not dll_path:
                messagebox.showinfo("測試", "請先指定 DLL 路徑")
                return
                
            if not os.path.exists(dll_path):
                messagebox.showerror("測試失敗", f"找不到 DLL 檔案: {dll_path}")
                return
                
            try:
                # 嘗試載入 DLL
                from .nhi_card_dll import NHICardDLL, NHICardDLLError
                try:
                    nhi_dll = NHICardDLL(dll_path)
                    status_text.set(f"讀卡機狀態: 已連接 (DLL 載入成功: {dll_path})")
                    messagebox.showinfo("測試成功", f"成功載入 DLL: {dll_path}")
                except NHICardDLLError as e:
                    status_text.set(f"讀卡機狀態: 未連接 (DLL 載入失敗)")
                    messagebox.showerror("測試失敗", f"DLL 載入失敗: {e}")
            except Exception as e:
                messagebox.showerror("測試錯誤", f"測試過程中發生錯誤: {e}")
        
        # 測試按鈕置中顯示
        test_button_frame = ttk.Frame(test_frame)
        test_button_frame.pack(pady=10)
        
        test_button = ttk.Button(
            test_button_frame, 
            text="測試 DLL 連線", 
            command=test_dll,
            width=20
        )
        test_button.pack(pady=5)
        
        # 分隔線
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.pack(fill='x', pady=15)
        
        # 操作按鈕區域 - 使用明顯的樣式和位置
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        # 提示標籤 - 使用明顯的顏色和字體
        hint_label = ttk.Label(
            button_frame, 
            text="完成設定後，請點擊「儲存設定」按鈕",
            foreground="blue",
            font=(self.default_font, 11, "bold")
        )
        hint_label.pack(pady=10)
        
        # 按鈕容器 - 確保按鈕居中顯示
        button_inner_frame = ttk.Frame(button_frame)
        button_inner_frame.pack(anchor='center', pady=5)
        
        def save_settings():
            try:
                # 驗證設定
                try:
                    retry_count = int(retry_var.get())
                    if retry_count < 0 or retry_count > 10:
                        messagebox.showerror("設定錯誤", "重試次數必須在 0-10 之間")
                        return
                except ValueError:
                    messagebox.showerror("設定錯誤", "重試次數必須是數字")
                    return
                    
                try:
                    timeout = int(timeout_var.get())
                    if timeout < 5 or timeout > 120:
                        messagebox.showerror("設定錯誤", "逾時時間必須在 5-120 秒之間")
                        return
                except ValueError:
                    messagebox.showerror("設定錯誤", "逾時時間必須是數字")
                    return
                
                # 更新配置檔案
                config.set('健保卡設定', 'dll_path', dll_path_var.get())
                config.set('健保卡設定', 'reader_type', reader_type_var.get())
                config.set('健保卡設定', 'retry_count', retry_var.get())
                config.set('健保卡設定', 'read_timeout', timeout_var.get())
                config.set('健保卡設定', 'enable_sound', str(sound_var.get()).lower())
                
                with open('config.ini', 'w', encoding='utf-8') as f:
                    config.write(f)
                
                messagebox.showinfo("設定儲存", "健保卡讀卡機設定已儲存！\n重新啟動程式後生效。")
                settings_window.destroy()
                
            except Exception as e:
                messagebox.showerror("儲存錯誤", f"儲存設定時發生錯誤: {e}")
        
        def cancel_settings():
            settings_window.destroy()
        
        # 使用更大、更明顯的按鈕
        button_style = {'font': (self.default_font, 11), 'width': 20, 'padding': 8}
        
        save_button = ttk.Button(
            button_inner_frame, 
            text="儲存設定", 
            command=save_settings,
            style='primary.TButton',  # 嘗試使用主題樣式
            width=20
        )
        save_button.pack(side='left', padx=30, pady=10)
        
        cancel_button = ttk.Button(
            button_inner_frame, 
            text="取消", 
            command=cancel_settings,
            width=20
        )
        cancel_button.pack(side='left', padx=30, pady=10)


    def launch_csfsim(self):
        """啟動健保署讀卡機控制軟體"""
        try:
            # 檢查檔案是否存在
            csfsim_path = self.card_reader.csfsim_path
            if not csfsim_path or not os.path.exists(csfsim_path):
                error_msg = f"找不到健保署讀卡機控制軟體:\n{csfsim_path}\n\n" \
                           f"請檢查配置檔案中的 csfsim_path 設定"
                self.status_text.set("找不到健保署讀卡機控制軟體")
                messagebox.showerror("檔案不存在", error_msg)
                return
            
            # 嘗試啟動
            result = self.card_reader.launch_csfsim()
            if result:
                self.status_text.set("已嘗試啟動健保署讀卡機控制軟體，請稍候...")
                messagebox.showinfo("啟動成功", 
                    f"已嘗試啟動健保署讀卡機控制軟體\n\n"
                    f"路徑: {csfsim_path}\n\n"
                    f"如果程式沒有出現，請檢查：\n"
                    f"1. 程式是否已在背景運行\n"
                    f"2. 防毒軟體是否阻擋\n"
                    f"3. 系統權限是否足夠")
            else:
                error_msg = f"無法啟動健保署讀卡機控制軟體\n\n" \
                           f"路徑: {csfsim_path}\n\n" \
                           f"可能的原因：\n" \
                           f"1. 程式檔案損壞\n" \
                           f"2. 系統權限不足\n" \
                           f"3. 防毒軟體阻擋\n" \
                           f"4. 程式已在運行\n\n" \
                           f"請查看日誌檔案獲取詳細錯誤資訊"
                self.status_text.set("無法啟動健保署讀卡機控制軟體")
                messagebox.showerror("啟動失敗", error_msg)
        except Exception as e:
            error_msg = f"啟動健保署讀卡機控制軟體時發生錯誤:\n{e}\n\n" \
                       f"請查看日誌檔案獲取詳細錯誤資訊"
            self.status_text.set(f"啟動健保署讀卡機控制軟體時發生錯誤: {e}")
            messagebox.showerror("啟動錯誤", error_msg)
            
    def show_manual_input_dialog(self):
        """顯示手動輸入病人資料對話框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("手動輸入病人資料")
        dialog.geometry("500x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 置中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (400 // 2)
        dialog.geometry(f"500x400+{x}+{y}")
        
        # 主要框架
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # 標題
        title_label = ttk.Label(main_frame, text="手動輸入病人資料", 
                               font=(self.default_font, 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 警告標籤
        warning_label = ttk.Label(main_frame, 
                                 text="⚠️ 請仔細核對病人身分，確保資料正確無誤！",
                                 foreground="red", font=(self.default_font, 10, "bold"))
        warning_label.pack(pady=(0, 15))
        
        # 輸入欄位
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill='x', pady=10)
        
        # 身分證字號
        ttk.Label(input_frame, text="身分證字號:", width=12, anchor='w').grid(row=0, column=0, sticky='w', pady=5)
        id_var = tk.StringVar()
        id_entry = ttk.Entry(input_frame, textvariable=id_var, width=20)
        id_entry.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        
        # 姓名
        ttk.Label(input_frame, text="姓名:", width=12, anchor='w').grid(row=1, column=0, sticky='w', pady=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(input_frame, textvariable=name_var, width=20)
        name_entry.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        
        # 出生日期
        ttk.Label(input_frame, text="出生日期:", width=12, anchor='w').grid(row=2, column=0, sticky='w', pady=5)
        dob_var = tk.StringVar()
        dob_entry = ttk.Entry(input_frame, textvariable=dob_var, width=20)
        dob_entry.grid(row=2, column=1, sticky='ew', pady=5, padx=5)
        ttk.Label(input_frame, text="(格式: YYYY/MM/DD)", foreground="gray").grid(row=2, column=2, sticky='w', pady=5, padx=5)
        
        # 設定欄位權重
        input_frame.columnconfigure(1, weight=1)
        
        # 按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=20)
        
        def confirm_input():
            # 驗證輸入
            if not id_var.get().strip():
                messagebox.showerror("輸入錯誤", "請輸入身分證字號")
                return
            if not name_var.get().strip():
                messagebox.showerror("輸入錯誤", "請輸入姓名")
                return
            if not dob_var.get().strip():
                messagebox.showerror("輸入錯誤", "請輸入出生日期")
                return
            
            # 建立病人資料
            from .data_processor import DataProcessor
            import datetime
            
            raw_data = {
                "ID_NUMBER": id_var.get().strip(),
                "FULL_NAME": name_var.get().strip(),
                "BIRTH_DATE": dob_var.get().strip(),
                "SEX": ""  # 移除性別欄位
            }
            
            try:
                # 處理資料
                processor = DataProcessor()
                processed_data = processor.process_raw_data(raw_data)
                self.current_patient_data = processed_data
                
                # 更新 UI 顯示
                self.patient_id_var.set(processed_data["id"])
                self.patient_name_var.set(processed_data["name"])
                self.patient_dob_var.set(processed_data["dob"])
                self.patient_sex_var.set(processed_data.get("sex", ""))
                self.read_time_var.set(processed_data["read_time"])
                
                # 更新狀態
                self.status_text.set("手動輸入完成！請確認資料並設定列印張數")
                self.print_button.config(state=tk.NORMAL)
                
                # 記錄操作
                try:
                    self.record_manager.log_operation(
                        processed_data, 
                        processed_data["read_time"], 
                        0, 
                        "手動輸入 (離線模式)"
                    )
                except Exception as e:
                    logger.warning(f"記錄手動輸入事件失敗: {e}")
                
                # 更新統計
                self.update_statistics()
                
                logger.info(f"手動輸入病人資料成功，病人: {processed_data['name']}")
                
                # 關閉對話框
                dialog.destroy()
                
                # 顯示確認訊息
                messagebox.showinfo("輸入完成", f"已成功輸入病人資料：\n\n姓名: {processed_data['name']}\n身分證: {processed_data['id']}")
                
            except Exception as e:
                messagebox.showerror("處理錯誤", f"處理輸入資料時發生錯誤:\n{e}")
                logger.error(f"處理手動輸入資料失敗: {e}")
        
        def cancel_input():
            dialog.destroy()
        
        # 按鈕
        confirm_button = ttk.Button(button_frame, text="確認輸入", command=confirm_input, width=15)
        confirm_button.pack(side='left', padx=10)
        
        cancel_button = ttk.Button(button_frame, text="取消", command=cancel_input, width=15)
        cancel_button.pack(side='left', padx=10)
        
        # 焦點設定
        id_entry.focus()
        
        # 綁定 Enter 鍵
        def on_enter(event):
            confirm_input()
        
        id_entry.bind('<Return>', on_enter)
        name_entry.bind('<Return>', on_enter)
        dob_entry.bind('<Return>', on_enter)

    def update_com_status(self):
        """更新COM埠狀態顯示"""
        try:
            com_port = self.card_reader.com_port
            auto_detect = self.card_reader.auto_detect_com
            
            if auto_detect:
                status_text = f"目前使用: COM{com_port} (自動偵測)"
            else:
                status_text = f"目前使用: COM{com_port} (手動設定)"
            
            # 測試COM埠是否可用
            if self.card_reader._test_com_port(com_port):
                status_text += " ✓ 可用"
            else:
                status_text += " ✗ 無法連接"
            
            self.com_status_text.set(status_text)
        except Exception as e:
            logger.error(f"更新COM埠狀態失敗: {e}")
            self.com_status_text.set("COM埠狀態: 未知")

    def auto_detect_com_port(self):
        """自動偵測COM埠"""
        try:
            self.status_text.set("正在自動偵測COM埠...")
            
            # 重新偵測COM埠
            new_port = self.card_reader._detect_com_port()
            
            if new_port != self.card_reader.com_port:
                self.card_reader.com_port = new_port
                logger.info(f"COM埠已更新為: COM{new_port}")
            
            # 更新狀態顯示
            self.update_com_status()
            
            # 顯示結果
            messagebox.showinfo("自動偵測完成", 
                               f"偵測完成！\n\n"
                               f"目前使用: COM{new_port}\n"
                               f"請測試讀卡機是否正常工作")
            
            self.status_text.set(f"COM埠已更新為: COM{new_port}")
            
        except Exception as e:
            logger.error(f"自動偵測COM埠失敗: {e}")
            messagebox.showerror("偵測失敗", f"自動偵測COM埠失敗:\n{e}")

    def show_com_port_dialog(self):
        """顯示COM埠設定對話框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("手動設定COM埠")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 置中顯示
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"400x300+{x}+{y}")
        
        # 主框架
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # 標題
        title_label = ttk.Label(main_frame, text="COM埠設定", 
                               font=(self.default_font, 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 說明文字
        info_label = ttk.Label(main_frame, 
                              text="請選擇要使用的COM埠：",
                              font=(self.default_font, 11))
        info_label.pack(pady=(0, 15))
        
        # 可用COM埠列表
        ports_frame = ttk.LabelFrame(main_frame, text="可用的COM埠", padding=10)
        ports_frame.pack(fill='x', pady=10)
        
        # 取得可用COM埠
        try:
            available_ports = self.card_reader.get_available_com_ports()
            
            if not available_ports:
                no_ports_label = ttk.Label(ports_frame, text="未發現任何COM埠", 
                                          foreground="red")
                no_ports_label.pack()
            else:
                # 建立COM埠選擇框架
                port_var = tk.StringVar()
                port_var.set(f"COM{self.card_reader.com_port}")  # 預設選擇目前使用的COM埠
                
                for i, port in enumerate(available_ports):
                    port_text = f"{port['device']} - {port['description']}"
                    if port['manufacturer']:
                        port_text += f" ({port['manufacturer']})"
                    
                    radio = ttk.Radiobutton(ports_frame, text=port_text, 
                                          variable=port_var, value=port['device'])
                    radio.pack(anchor='w', pady=2)
                    
        except Exception as e:
            error_label = ttk.Label(ports_frame, text=f"無法取得COM埠列表: {e}", 
                                   foreground="red")
            error_label.pack()
        
        # 手動輸入COM埠
        manual_frame = ttk.LabelFrame(main_frame, text="手動輸入", padding=10)
        manual_frame.pack(fill='x', pady=10)
        
        manual_label = ttk.Label(manual_frame, text="或手動輸入COM埠號碼:")
        manual_label.pack(anchor='w')
        
        manual_var = tk.StringVar()
        manual_entry = ttk.Entry(manual_frame, textvariable=manual_var, width=10)
        manual_entry.pack(anchor='w', pady=5)
        
        # 按鈕框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=20)
        
        def confirm_setting():
            try:
                # 取得選擇的COM埠
                selected_port = port_var.get()
                
                # 如果手動輸入了COM埠號碼
                if manual_var.get().strip():
                    try:
                        port_num = int(manual_var.get().strip())
                        selected_port = f"COM{port_num}"
                    except ValueError:
                        messagebox.showerror("輸入錯誤", "COM埠號碼必須是數字")
                        return
                
                # 提取COM埠號碼
                port_num = int(selected_port.replace('COM', ''))
                
                # 測試COM埠
                if self.card_reader.set_com_port(port_num):
                    # 更新狀態顯示
                    self.update_com_status()
                    
                    # 關閉對話框
                    dialog.destroy()
                    
                    # 顯示成功訊息
                    messagebox.showinfo("設定成功", 
                                       f"COM埠已設定為: COM{port_num}\n\n"
                                       f"請測試讀卡機是否正常工作")
                    
                    self.status_text.set(f"COM埠已設定為: COM{port_num}")
                else:
                    messagebox.showerror("設定失敗", 
                                        f"無法設定COM埠: COM{port_num}\n\n"
                                        f"請確認：\n"
                                        f"1. COM埠號碼正確\n"
                                        f"2. 讀卡機已連接\n"
                                        f"3. 沒有其他程式使用此COM埠")
                    
            except Exception as e:
                logger.error(f"設定COM埠失敗: {e}")
                messagebox.showerror("設定失敗", f"設定COM埠時發生錯誤:\n{e}")
        
        def cancel_setting():
            dialog.destroy()
        
        # 按鈕
        confirm_button = ttk.Button(button_frame, text="確認設定", 
                                   command=confirm_setting, width=15)
        confirm_button.pack(side='left', padx=10)
        
        cancel_button = ttk.Button(button_frame, text="取消", 
                                 command=cancel_setting, width=15)
        cancel_button.pack(side='left', padx=10)
        
        # 焦點設定
        manual_entry.focus()

    def on_closing(self):
        """程式關閉時的處理"""
        if messagebox.askokcancel("退出", "確定要退出程式嗎？"):
            logger.info("使用者關閉程式")
            self.root.destroy()

def create_app(dll_path=None):
    """建立並啟動應用程式"""
    root = tk.Tk()
    app = MedicalCardApp(root, dll_path)
    
    # 設定關閉事件
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # 設定圖示 (如果有的話)
    try:
        # root.iconbitmap("icon.ico")  # 如果有圖示檔案
        pass
    except:
        pass
    
    return root, app
