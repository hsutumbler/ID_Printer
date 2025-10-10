# GNT 健保卡讀取系統參考文件

## 概述
此文件記錄了從 GNT 醫院抽血櫃台程式中提取的健保卡讀取相關資訊，供未來參考使用。

## GNT 系統架構

### 主要檔案結構
```
GNT/
├── HenCs/
│   ├── NhiCard.dll          # 主要健保卡讀取 DLL (.NET 組件)
│   ├── CardPortSet.xml      # COM 埠設定檔
│   ├── NhiCard.xml          # DLL 說明文件
│   └── Log/                 # 記錄檔目錄
├── HenCs1/                  # 備用系統
└── kanban/                  # 看板系統
```

### COM 埠設定 (CardPortSet.xml)
```xml
<?xml version="1.0" encoding="utf-8" ?>
<port>
  <!-- 設定讀卡機使用的port，只要輸入號碼就好 -->
  <num>1</num>
</port>
```

## NhiCard.dll COM 介面

### 主要類別和方法

#### NhiCard.Patient 類別
用於讀取病患健保卡資料的主要類別。

**屬性方法：**
- `GetPatientIdCard` - 取得病患身分證字號
- `GetPatientName` - 取得病患姓名  
- `GetPatientSex` - 取得病患性別
- `CardCheck` - 檢查是否有插入卡片
- `GetCardWholestr` - 取得完整卡片資料

**操作方法：**
- `Open()` - 開啟讀卡機連接埠
- `GetPatientData()` - 取得病患基本資料

#### NhiCard.CsHis 類別
健保局提供的標準函式庫。

**主要函式：**
- `hisGetBasicData()` - 1.1 讀取不需個人PIN碼資料
- `hisGetRegisterBasic()` - 1.2 掛號或報到時讀取基本資料
- `csOpenCom(int)` - 1.31 開啟讀卡機連結埠
- `csCloseCom()` - 1.32 關閉讀卡機連結埠
- `csGetCardNo()` - 1.35 讀取卡片號碼

### COM 物件使用方式
```python
import win32com.client

# 建立 COM 物件
patient = win32com.client.Dispatch("NhiCard.Patient")

# 開啟連接埠
result = patient.Open()

# 讀取病患資料
if patient.GetPatientData():
    id_card = patient.GetPatientIdCard
    name = patient.GetPatientName
    sex = patient.GetPatientSex
```

## 資料格式

### 病患資料結構
- **身分證字號：** 10 位英數字組合
- **姓名：** 中文姓名
- **性別：** "1"=男性, "2"=女性
- **出生日期：** YYYYMMDD 或 YYYMMDD (民國年) 格式

### 性別代碼對應
```
"1", "M", "男", "Male" → "男"
"2", "F", "女", "Female" → "女"
```

## 錯誤處理

### 常見錯誤代碼
- `-2147221005` - 無效的類別字串 (COM 物件未註冊)
- 讀卡機連接失敗
- 卡片讀取失敗

### 故障排除步驟
1. 檢查健保卡是否正確插入
2. 確認讀卡機電源和連接
3. 驗證 COM 埠設定 (預設 COM1)
4. 檢查 DLL 註冊狀態

## 記錄檔格式

### 記錄檔位置
- `GNT/HenCs/Log/NhiCardLOG_YYYYMMDD.txt`

### 記錄格式範例
```
::2025-10-09 07:29:53
*** 開始 開啟讀卡機 ****
連接埠 1 開啟測試...成功...
*** 結束 開啟讀卡機 ****

::2025-10-09 07:29:54
***開始取得病患資料***
***取得病患資料結束***
```

## 系統需求

### 軟體需求
- Windows 作業系統
- .NET Framework
- 健保署讀卡機驅動程式
- COM 介面支援 (win32com for Python)

### 硬體需求
- 健保卡讀卡機
- RS-232 或 USB 連接埠
- 健保卡

## 整合注意事項

### DLL 註冊
GNT 的 NhiCard.dll 是 .NET 組件，需要在系統中註冊才能透過 COM 介面使用。

### 相容性
- 不能直接使用 ctypes.CDLL 載入
- 必須透過 win32com.client.Dispatch 建立物件
- 需要醫院環境的特定設定

### 備援策略
當 GNT DLL 不可用時：
1. 自動切換至離線模式
2. 提供簡化的錯誤訊息
3. 維持基本的標籤列印功能

## 移除 GNT 後的因應措施

### 程式碼調整
1. 移除 GNT 路徑的搜尋邏輯
2. 直接使用標準健保署 DLL 路徑
3. 簡化 COM 介面相關程式碼

### 設定檔調整
```ini
[健保卡設定]
# 移除 GNT 相關路徑，使用標準路徑
dll_path = C:\NHI\LIB\csHis50.dll
com_port = 1
offline_mode = true
```

### 功能保留
- 離線模式運作
- 基本資料讀取
- 標籤列印功能
- 記錄檔管理

---

**備註：** 此文件基於 GNT 系統分析結果，移除 GNT 資料夾後仍可作為健保卡讀取功能的參考依據。
