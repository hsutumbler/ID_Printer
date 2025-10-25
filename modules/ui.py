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
        self.root.geometry("650x700")  # 增加視窗高度以適應新增的備註欄位
        self.root.resizable(True, True)  # 允許調整視窗大小
        self.root.minsize(600, 650)  # 增加最小視窗高度
        
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
        
        # 初始化讀取時間變數
        self.read_time_var = tk.StringVar()
        
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
        
        # 綁定分頁切換事件
        self.tab_control.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # 主畫面頁籤
        main_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(main_tab, text="主畫面")
        
        # 今日統計頁籤
        self.stats_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.stats_tab, text="今日統計")
        
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

        # 模式選擇區
        mode_frame = ttk.LabelFrame(main_tab, text="操作模式", padding=10)
        mode_frame.pack(pady=5, padx=20, fill='x')

        # 建立模式選擇的單選按鈕
        self.mode_var = tk.StringVar(value="card")  # 預設為讀卡模式
        
        # 使用 Frame 來水平排列單選按鈕
        mode_button_frame = ttk.Frame(mode_frame)
        mode_button_frame.pack(fill='x', pady=5)
        
        # 讀卡模式單選按鈕
        card_radio = ttk.Radiobutton(mode_button_frame, text="讀卡模式", 
                                    variable=self.mode_var, value="card",
                                    command=self.on_mode_change)
        card_radio.pack(side='left', padx=20)
        
        # 手工模式單選按鈕
        manual_radio = ttk.Radiobutton(mode_button_frame, text="手工模式", 
                                      variable=self.mode_var, value="manual",
                                      command=self.on_mode_change)
        manual_radio.pack(side='left', padx=20)
        
        # 病人資料顯示區
        patient_frame = ttk.LabelFrame(main_tab, text="病人基本資料", padding=15)
        patient_frame.pack(pady=10, padx=20, fill='x')
        
        # 建立資料顯示/輸入欄位
        self.patient_id_var = tk.StringVar()
        self.patient_name_var = tk.StringVar()
        self.patient_dob_var = tk.StringVar()
        self.patient_note_var = tk.StringVar()
        self.card_no_var = tk.StringVar()
        
        # ID
        id_frame = ttk.Frame(patient_frame)
        id_frame.pack(fill='x', pady=2)
        ttk.Label(id_frame, text="ID:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        self.id_entry = ttk.Entry(id_frame, textvariable=self.patient_id_var, 
                                font=(self.default_font, 13), state='readonly',
                                width=20)
        self.id_entry.pack(side='left', padx=5)
        
        # 姓名
        name_frame = ttk.Frame(patient_frame)
        name_frame.pack(fill='x', pady=2)
        ttk.Label(name_frame, text="姓名:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        self.name_entry = ttk.Entry(name_frame, textvariable=self.patient_name_var, 
                                  font=(self.default_font, 13), state='readonly',
                                  width=20)
        self.name_entry.pack(side='left', padx=5)
        
        # 生日
        dob_frame = ttk.Frame(patient_frame)
        dob_frame.pack(fill='x', pady=2)
        ttk.Label(dob_frame, text="生日:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        self.dob_entry = ttk.Entry(dob_frame, textvariable=self.patient_dob_var, 
                                 font=(self.default_font, 13), state='readonly',
                                 width=20)
        self.dob_entry.pack(side='left', padx=5)
        ttk.Label(dob_frame, text="(格式: YYYY/MM/DD)", foreground="gray").pack(side='left', padx=5)
        
        # 備註
        note_frame = ttk.Frame(patient_frame)
        note_frame.pack(fill='x', pady=2)
        ttk.Label(note_frame, text="備註:", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        self.note_entry = ttk.Entry(note_frame, textvariable=self.patient_note_var, 
                                  font=(self.default_font, 13), width=20)
        self.note_entry.pack(side='left', padx=5)
        ttk.Label(note_frame, text="(限10字)", foreground="gray").pack(side='left', padx=5)
        
        # 綁定備註欄位的字數限制
        self.patient_note_var.trace_add('write', self._on_note_change)
        
        # 健保卡號(後四碼)
        card_no_frame = ttk.Frame(patient_frame)
        card_no_frame.pack(fill='x', pady=2)
        ttk.Label(card_no_frame, text="健保卡號(後四碼):", width=12, anchor='w', 
                 font=(self.default_font, 11, "bold")).pack(side='left')
        self.card_no_entry = ttk.Entry(card_no_frame, textvariable=self.card_no_var, 
                                     font=(self.default_font, 13), state='readonly',
                                     width=20)
        self.card_no_entry.pack(side='left', padx=5)
        ttk.Label(card_no_frame, text="(讀卡模式有效)", foreground="gray").pack(side='left', padx=5)
        
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

        # 綁定輸入欄位的變更事件
        self.patient_id_var.trace_add('write', self._on_data_change)
        self.patient_name_var.trace_add('write', self._on_data_change)
        self.patient_dob_var.trace_add('write', self._on_data_change)
        
        # 統計資訊變數
        self.stats_text = tk.StringVar()
        
        # ===== 今日統計頁籤內容 =====
        # 初始化統計頁面（將在第一次切換時建立內容）
        self.stats_content_created = False
        self.create_stats_content()
        
        # 版本號標籤 - 放在主視窗右下方
        version_label = ttk.Label(self.root, text="v1_20251025", 
                                 font=(self.default_font, 8, "underline"),
                                 foreground="gray")
        version_label.place(relx=1.0, rely=1.0, anchor='se', x=-10, y=-10)

    def on_tab_changed(self, event):
        """分頁切換事件處理器"""
        try:
            # 取得當前選中的分頁索引
            selected_tab = self.tab_control.select()
            tab_index = self.tab_control.index(selected_tab)
            
            # 如果切換到統計頁面（索引為1），重新載入統計資料
            if tab_index == 1:
                logger.info("切換到統計頁面，重新載入資料")
                self.refresh_stats_content()
                
        except Exception as e:
            logger.warning(f"分頁切換事件處理失敗: {e}")

    def create_stats_content(self):
        """建立統計頁面內容"""
        if self.stats_content_created:
            return
            
        # 主框架
        self.stats_main_frame = ttk.Frame(self.stats_tab, padding=20)
        self.stats_main_frame.pack(fill='both', expand=True)
        
        # 統計標題
        stats_title = ttk.Label(self.stats_main_frame, text="今日統計資訊", 
                               font=(self.default_font, 14, "bold"))
        stats_title.pack(pady=(10, 20))
        
        # 顯示簡易統計資訊
        self.stats_info_frame = ttk.LabelFrame(self.stats_main_frame, text="統計摘要", padding=15)
        self.stats_info_frame.pack(fill='x', pady=10)
        
        # 顯示統計資訊
        self.stats_label = ttk.Label(self.stats_info_frame, textvariable=self.stats_text, 
                                    font=(self.default_font, 12))
        self.stats_label.pack(pady=15)
        
        # 詳細統計資訊區
        self.details_frame = ttk.LabelFrame(self.stats_main_frame, text="詳細資訊", padding=15)
        self.details_frame.pack(fill='both', expand=True, pady=15)
        
        # 操作按鈕區
        stats_button_frame = ttk.Frame(self.stats_main_frame)
        stats_button_frame.pack(pady=15, fill='x')
        
        # 匯出完整資料按鈕
        def export_complete_data():
            self.export_today_complete_data()
            
        export_button = ttk.Button(stats_button_frame, text="匯出完整資料", 
                                  command=export_complete_data, width=20)
        export_button.pack(anchor='center', pady=10)
        
        self.stats_content_created = True
        
        # 初始載入資料
        self.refresh_stats_content()

    def refresh_stats_content(self):
        """重新載入統計頁面內容"""
        try:
            # 更新統計摘要
            self.update_statistics()
            
            # 清除詳細資訊區的舊內容
            for widget in self.details_frame.winfo_children():
                widget.destroy()
            
            # 取得最新統計資料
            stats = self.record_manager.get_statistics()
            records = self.record_manager.get_today_records()
            
            # 建立表格式顯示
            details_text = tk.Text(self.details_frame, wrap='word', height=15, width=80, 
                                 font=(self.default_font, 9))
            
            # 顯示統計摘要
            details_text.insert('1.0', "=== 今日統計摘要 ===\n")
            details_text.insert('end', f"讀取次數: {stats['total_reads']} 次\n")
            details_text.insert('end', f"列印次數: {stats['total_prints']} 次\n")
            details_text.insert('end', f"列印標籤: {stats['total_labels']} 張\n")
            details_text.insert('end', f"總記錄數: {stats['total_records']} 筆\n\n")
            
            # 顯示完整記錄
            details_text.insert('end', "=== 完整記錄 ===\n")
            if records:
                # 表頭 - 按照新的排序：序號、ID、姓名、生日、列印時間、備註
                details_text.insert('end', f"{'序號':<4} {'ID':<12} {'姓名':<8} {'生日':<12} {'列印時間':<18} {'備註':<10}\n")
                details_text.insert('end', "-" * 74 + "\n")
                
                # 記錄內容
                for i, record in enumerate(records, 1):
                    id_str = record.get('身分證字號', '')[:10]
                    name_str = record.get('姓名', '')[:6]
                    dob_str = record.get('出生年月日', '')[:10]
                    time_str = record.get('時間戳記', '')[:16]  # 顯示到分鐘
                    note_str = record.get('備註', '')[:8]
                    
                    # 按照新的排序顯示
                    details_text.insert('end', f"{i:<4} {id_str:<12} {name_str:<8} {dob_str:<12} {time_str:<18} {note_str:<10}\n")
            else:
                details_text.insert('end', "今日尚無記錄\n")
                
            details_text.config(state='disabled')  # 設為唯讀
            
            # 建立捲軸
            scrollbar = ttk.Scrollbar(self.details_frame, orient='vertical', command=details_text.yview)
            details_text.configure(yscrollcommand=scrollbar.set)
            
            details_text.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
            
            logger.info(f"統計頁面已更新，共 {len(records)} 筆記錄")
            
        except Exception as e:
            logger.error(f"重新載入統計內容失敗: {e}")
            error_label = ttk.Label(self.details_frame, text=f"載入統計資料失敗: {e}", 
                                  foreground="red", font=(self.default_font, 10))
            error_label.pack(pady=20)

    def start_read_card(self):
        """開始讀取健保卡"""
        # 只在讀卡模式下執行（手工模式下按鈕已禁用）
        if self.mode_var.get() == "card":
            # 讀卡模式：執行讀卡
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
            self.patient_note_var.set("")  # 讀卡後備註欄位清空，供使用者輸入
            self.card_no_var.set(processed_data.get("card_no", ""))
            
            # 根據離線模式設定不同的狀態訊息
            if self.offline_mode:
                self.status_text.set("離線模式 - 讀取成功！請確認資料並設定列印張數")
            else:
                self.status_text.set("讀取成功！請確認資料並設定列印張數")
            
            self.print_button.config(state=tk.NORMAL)
            
            # 顯示確認對話框，提醒醫檢師核對病人身分
            confirm_msg = f"健保卡讀取成功！\n\n" \
                         f"病人: {processed_data['name']}\n" \
                         f"ID: {processed_data['id']}\n" \
                         f"生日: {processed_data['dob']}\n\n" \
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
            # 只在讀卡模式下啟用讀取按鈕
            if self.mode_var.get() == "card":
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
        
        # 簡單顯示錯誤訊息，只有確認按鈕
        messagebox.showerror("無法讀取健保卡", "讀取健保卡失敗，請檢查健保卡是否正確插入。")
        
        logger.error(f"讀取健保卡失敗: {error_msg}")
        
        # 確保按鈕狀態正確（只在讀卡模式下啟用讀取按鈕）
        if self.mode_var.get() == "card":
            self.read_button.config(state=tk.NORMAL)
        self.print_button.config(state=tk.DISABLED)

    def _on_note_change(self, *args):
        """當備註欄位變更時，限制字數"""
        current_text = self.patient_note_var.get()
        if len(current_text) > 10:
            # 截斷到10字
            self.patient_note_var.set(current_text[:10])

    def _on_data_change(self, *args):
        """當手動輸入的資料變更時"""
        if self.mode_var.get() == "manual":
            # 檢查所有必填欄位是否都有值
            id_value = self.patient_id_var.get().strip()
            name_value = self.patient_name_var.get().strip()
            dob_value = self.patient_dob_var.get().strip()
            
            if id_value and name_value and dob_value:
                # 更新當前病人資料
                self.current_patient_data = {
                    "id": id_value,
                    "name": name_value,
                    "dob": dob_value,
                    "note": self.patient_note_var.get().strip(),
                    "read_time": datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                }
                # 啟用列印按鈕
                self.print_button.config(state=tk.NORMAL)
            else:
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

            # 準備列印資料，包含當前的備註內容
            print_data = self.current_patient_data.copy()
            print_data["note"] = self.patient_note_var.get().strip()

            # 開始列印
            self.status_text.set(f"標籤列印中... (共 {print_count} 張)")
            self.print_button.config(state=tk.DISABLED)
            self.read_button.config(state=tk.DISABLED)
            
            logger.info(f"使用者點擊列印標籤按鈕，列印 {print_count} 張")

            # 執行列印
            self.print_manager.print_labels(print_data, print_count)
            
            # 記錄列印事件
            print_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            try:
                self.record_manager.log_operation(
                    print_data,  # 使用包含備註的完整列印資料
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
            # 只在讀卡模式下啟用讀取按鈕
            if self.mode_var.get() == "card":
                self.read_button.config(state=tk.NORMAL)
    
    def on_mode_change(self):
        """處理模式切換"""
        mode = self.mode_var.get()
        if mode == "card":
            self.status_text.set("請插入健保卡並點擊讀取")
            self.read_button.configure(text="讀取健保卡", state=tk.NORMAL)
            # 轉換到讀卡模式：禁用輸入框
            self._set_entry_state('readonly')
        else:
            self.status_text.set("請直接在下方輸入病人資料")
            # 手工模式：按鈕保持原文字但變成灰色且無法按壓
            self.read_button.configure(text="讀取健保卡", state=tk.DISABLED)
            # 轉換到手工模式：啟用輸入框
            self._set_entry_state('normal')
        
        # 直接清除當前資料，不詢問
        self._clear_data_without_confirm()
        logger.info(f"切換至{'讀卡' if mode == 'card' else '手工'}模式")

    def _set_entry_state(self, state):
        """設定所有輸入框的狀態"""
        self.id_entry.configure(state=state)
        self.name_entry.configure(state=state)
        self.dob_entry.configure(state=state)
        # 備註欄位在兩種模式下都可以輸入
        self.note_entry.configure(state='normal')
        # 健保卡號欄位永遠保持唯讀狀態
        self.card_no_entry.configure(state='readonly')

    def _clear_data_without_confirm(self):
        """直接清除病人資料，不顯示確認對話框"""
        self.current_patient_data = None
        self.patient_id_var.set("")
        self.patient_name_var.set("")
        self.patient_dob_var.set("")
        self.patient_note_var.set("")
        self.card_no_var.set("")
        self.print_button.config(state=tk.DISABLED)
        logger.info("清除病人資料")

    def clear_data(self):
        """清除病人資料（帶確認對話框）"""
        if messagebox.askyesno("確認清除", "確定要清除目前的病人資料嗎？"):
            self._clear_data_without_confirm()
            # 如果是手工模式，清除後聚焦到身分證字號欄位
            if self.mode_var.get() == "manual":
                self.id_entry.focus()

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

    def export_today_complete_data(self):
        """匯出今日完整資料為 CSV 檔案"""
        try:
            from tkinter import filedialog
            import csv
            import datetime
            
            # 取得今日記錄
            records = self.record_manager.get_today_records()
            
            if not records:
                messagebox.showinfo("匯出資料", "今日尚無記錄可匯出")
                return
            
            # 取得當前日期作為預設檔名
            today_str = datetime.datetime.now().strftime("%Y%m%d")
            default_filename = f"健保卡標籤完整資料_{today_str}.csv"
            
            # 詢問儲存位置
            file_path = filedialog.asksaveasfilename(
                initialfile=default_filename,
                defaultextension=".csv",
                filetypes=[("CSV 檔案", "*.csv"), ("所有檔案", "*.*")]
            )
            
            if not file_path:
                return  # 使用者取消
            
            # 寫入 CSV 檔案
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                
                # 寫入標題列
                headers = ["ID", "姓名", "生日", "列印時間", "備註", "健保卡號", "操作類型", "列印張數"]
                writer.writerow(headers)
                
                # 寫入資料列
                for record in records:
                    row = [
                        record.get('身分證字號', ''),
                        record.get('姓名', ''),
                        record.get('出生年月日', ''),
                        record.get('時間戳記', ''),
                        record.get('備註', ''),
                        record.get('健保卡號', ''),
                        record.get('操作類型', ''),
                        record.get('列印張數', '')
                    ]
                    writer.writerow(row)
            
            messagebox.showinfo("匯出成功", f"已成功匯出 {len(records)} 筆記錄至:\n{file_path}")
            logger.info(f"使用者匯出完整資料: {file_path}, 共 {len(records)} 筆記錄")
            
        except Exception as e:
            messagebox.showerror("匯出錯誤", f"匯出 CSV 檔案失敗:\n{e}")
            logger.error(f"匯出完整資料失敗: {e}")





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
        
        # ID
        ttk.Label(input_frame, text="ID:", width=12, anchor='w').grid(row=0, column=0, sticky='w', pady=5)
        id_var = tk.StringVar()
        id_entry = ttk.Entry(input_frame, textvariable=id_var, width=20)
        id_entry.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        
        # 姓名
        ttk.Label(input_frame, text="姓名:", width=12, anchor='w').grid(row=1, column=0, sticky='w', pady=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(input_frame, textvariable=name_var, width=20)
        name_entry.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        
        # 生日
        ttk.Label(input_frame, text="生日:", width=12, anchor='w').grid(row=2, column=0, sticky='w', pady=5)
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
                messagebox.showerror("輸入錯誤", "請輸入ID")
                return
            if not name_var.get().strip():
                messagebox.showerror("輸入錯誤", "請輸入姓名")
                return
            if not dob_var.get().strip():
                messagebox.showerror("輸入錯誤", "請輸入生日")
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
                self.patient_note_var.set("")  # 手工模式備註欄位清空，供使用者輸入
                # 手工模式不顯示健保卡號
                self.card_no_var.set("")
                
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
                messagebox.showinfo("輸入完成", f"已成功輸入病人資料：\n\n姓名: {processed_data['name']}\nID: {processed_data['id']}")
                
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
