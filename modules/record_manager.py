# 記錄檔管理模組
import csv
import datetime
import os
from .logger import logger

class RecordManagerError(Exception):
    """記錄檔管理錯誤"""
    pass

class RecordManager:
    def __init__(self, record_dir="records"):
        self.record_dir = record_dir
        
        # 確保記錄目錄存在
        try:
            os.makedirs(record_dir, exist_ok=True)
            logger.info(f"記錄檔管理模組初始化完成，記錄目錄: {os.path.abspath(record_dir)}")
        except Exception as e:
            logger.error(f"建立記錄目錄失敗: {e}")
            raise RecordManagerError(f"建立記錄目錄失敗: {e}")

    def _get_record_filepath(self):
        """根據日期獲取記錄檔路徑"""
        today_str = datetime.datetime.now().strftime("%Y%m%d")
        return os.path.join(self.record_dir, f"record_{today_str}.csv")

    def log_operation(self, patient_data, operation_time, print_count=0, operation_type="讀取"):
        """
        記錄操作，包括病人資料、時間和列印張數。
        patient_data 應包含 'id', 'name', 'dob', 'note', 'card_no'。
        """
        filepath = self._get_record_filepath()
        header = ["時間戳記", "身分證字號", "姓名", "出生年月日", "列印張數", "操作類型", "備註", "健保卡號"]
        
        try:
            # 檢查檔案是否存在且有內容
            file_exists = os.path.exists(filepath)
            file_empty = not file_exists or os.path.getsize(filepath) == 0
            
            with open(filepath, 'a', newline='', encoding='utf-8-sig') as f:  # 使用 utf-8-sig 以支援 Excel
                writer = csv.writer(f)
                
                # 如果檔案不存在或為空，寫入標頭
                if file_empty:
                    writer.writerow(header)
                    logger.info(f"建立新記錄檔: {filepath}")
                
                # 寫入記錄
                row = [
                    operation_time,
                    patient_data.get("id", "N/A"),
                    patient_data.get("name", "N/A"),
                    patient_data.get("dob", "N/A"),
                    print_count,
                    operation_type,
                    patient_data.get("note", ""),  # 備註
                    patient_data.get("card_no", "")  # 健保卡號後四碼
                ]
                writer.writerow(row)
            
            logger.info(f"記錄成功: {operation_type} - {patient_data.get('name')} - {print_count} 張")
            
        except Exception as e:
            logger.error(f"寫入記錄檔失敗: {e}")
            raise RecordManagerError(f"寫入記錄檔失敗: {e}")

    def get_today_records(self):
        """取得今日的記錄"""
        filepath = self._get_record_filepath()
        
        if not os.path.exists(filepath):
            return []
        
        try:
            records = []
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    records.append(row)
            
            logger.info(f"讀取今日記錄: {len(records)} 筆")
            return records
            
        except Exception as e:
            logger.error(f"讀取記錄檔失敗: {e}")
            raise RecordManagerError(f"讀取記錄檔失敗: {e}")

    def get_statistics(self):
        """取得今日統計資訊"""
        try:
            records = self.get_today_records()
            
            total_reads = sum(1 for r in records if r.get('操作類型') == '讀取')
            total_prints = sum(1 for r in records if r.get('操作類型') == '列印')
            total_labels = sum(int(r.get('列印張數', 0)) for r in records if r.get('操作類型') == '列印')
            
            stats = {
                'total_reads': total_reads,
                'total_prints': total_prints,
                'total_labels': total_labels,
                'total_records': len(records)
            }
            
            logger.info(f"今日統計: 讀取 {total_reads} 次, 列印 {total_prints} 次, 共 {total_labels} 張標籤")
            return stats
            
        except Exception as e:
            logger.error(f"取得統計資訊失敗: {e}")
            return {
                'total_reads': 0,
                'total_prints': 0,
                'total_labels': 0,
                'total_records': 0
            }

    def backup_records(self, backup_dir="backup"):
        """備份記錄檔"""
        try:
            os.makedirs(backup_dir, exist_ok=True)
            
            # 取得所有記錄檔
            record_files = [f for f in os.listdir(self.record_dir) if f.startswith('record_') and f.endswith('.csv')]
            
            backup_count = 0
            for record_file in record_files:
                src_path = os.path.join(self.record_dir, record_file)
                dst_path = os.path.join(backup_dir, f"backup_{record_file}")
                
                # 複製檔案
                import shutil
                shutil.copy2(src_path, dst_path)
                backup_count += 1
            
            logger.info(f"備份完成: {backup_count} 個記錄檔")
            return True, f"成功備份 {backup_count} 個記錄檔"
            
        except Exception as e:
            logger.error(f"備份記錄檔失敗: {e}")
            return False, f"備份失敗: {e}"
