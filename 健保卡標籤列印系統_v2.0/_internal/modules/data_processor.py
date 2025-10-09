# 資料處理模組
import datetime
from .logger import logger

class DataProcessingError(Exception):
    """資料處理錯誤"""
    pass

class DataProcessor:
    def __init__(self):
        logger.info("資料處理模組初始化完成")

    def process_raw_data(self, raw_data):
        """
        處理從健保卡讀取到的原始資料。
        raw_data 是一個字典，包含從 SDK 取得的資訊。
        """
        try:
            logger.info("開始處理原始健保卡資料")
            
            patient_id = raw_data.get("ID_NUMBER", "").strip()
            patient_name = raw_data.get("FULL_NAME", "").strip()
            
            # 驗證身分證字號格式 (簡單驗證)
            if not patient_id or len(patient_id) != 10:
                raise DataProcessingError("身分證字號格式不正確")
            
            # 驗證姓名
            if not patient_name:
                raise DataProcessingError("姓名資料不完整")
            
            # 處理出生年月日
            raw_dob = raw_data.get("BIRTH_DATE", "").strip()
            if raw_dob and len(raw_dob) == 8:
                try:
                    dob_obj = datetime.datetime.strptime(raw_dob, "%Y%m%d")
                    patient_dob = dob_obj.strftime("%Y/%m/%d")
                except ValueError:
                    logger.warning(f"出生年月日格式錯誤: {raw_dob}")
                    patient_dob = "格式錯誤"
            else:
                logger.warning("出生年月日資料不完整")
                patient_dob = "資料不完整"

            processed_data = {
                "id": patient_id,
                "name": patient_name,
                "dob": patient_dob,
                "read_time": datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            }
            
            logger.info(f"資料處理完成，病人: {patient_name}")
            return processed_data
            
        except Exception as e:
            logger.error(f"處理原始資料失敗: {e}")
            raise DataProcessingError(f"處理原始資料失敗: {e}")

    def validate_patient_data(self, patient_data):
        """驗證病人資料的完整性"""
        required_fields = ["id", "name", "dob"]
        for field in required_fields:
            if not patient_data.get(field):
                return False, f"缺少必要欄位: {field}"
        return True, "資料驗證通過"
