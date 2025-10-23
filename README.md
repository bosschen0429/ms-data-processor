# 質譜數據處理工具 

智慧化質譜數據去重複與排序工具,支援自動欄位識別與多種檔案格式


## ✨ 主要功能

- 🔍 **智慧欄位識別** - 自動識別 RT, m/z, Intensity 欄位,無需手動設定
- 🧹 **訊號去重複** - 基於 m/z 和 RT 容差智慧去除重複訊號
- 📊 **強度排序** - 自動按訊號強度排序,快速找出最重要的峰
- 📁 **多格式支援** - 支援 Excel (.xlsx, .xls), CSV, TSV 檔案
- 💾 **完整資訊保留** - 保留所有其他欄位資訊 (ID, Tags, Adducts 等)
- 🎨 **圖形化介面** - 友善的 GUI,操作簡單直覺
- ⚙️ **彈性參數** - 可自訂 m/z 容差、RT 容差、輸出數量

## 💾 下載安裝

### 下載執行檔 (推薦)

前往 [**Releases 頁面**](../../releases/latest) 下載最新版本:

| 作業系統 | 下載連結 | 說明 |
|---------|---------|------|
| 🪟 **Windows** | [質譜數據處理工具-Windows.zip](../../releases/latest) | Windows 7+ |
| 🍎 **macOS** | [質譜數據處理工具-macOS.zip](../../releases/latest) | macOS 10.13+ |
| 🐧 **Linux** | [質譜數據處理工具-Linux.zip](../../releases/latest) | Ubuntu 18.04+ |

### 或使用 Python 原始碼

```bash
# 1. 克隆儲存庫
git clone https://github.com/YOUR_USERNAME/YOUR_REPO.git
cd YOUR_REPO

# 2. 安裝依賴
pip install pandas openpyxl numpy

# 3. 執行程式
python ms_processor.py
```

## 🚀 快速開始

### 1️⃣ 開啟程式
- **Windows**: 雙擊 `質譜數據處理工具.exe`
- **Mac**: 右鍵點擊 → 打開 (首次執行)
- **Linux**: `./質譜數據處理工具`

### 2️⃣ 選擇檔案
點擊「選擇檔案」按鈕,選擇你的數據檔案

### 3️⃣ 調整參數 (可選)
- **m/z 容差**: 預設 20 ppm
- **RT 容差**: 預設 1
- **輸出數量**: 預設前 10 個 (輸入 0 表示全部)

### 4️⃣ 開始處理
點擊「開始處理」按鈕,等待完成

### 5️⃣ 取得結果
處理完成後,會在原檔案目錄生成 `原檔名_processed.檔案格式`

## 📋 支援的欄位名稱

程式會自動識別以下欄位 (不區分大小寫):

| 欄位類型 | 支援的名稱 |
|---------|-----------|
| **RT** | RT, Retention Time, RT (min), retention_time, RetentionTime |
| **m/z** | m/z, mz, Precursor Ion m/z, Precursor m/z, mass |
| **Intensity** | Intensity, Int, Precursor Ion Intensity, Abundance, Height |

**注意**: 其他欄位 (如 ID, Tags, Adducts) 會完整保留

## 📊 使用範例

### 輸入檔案範例
```csv
ID,RT (min),Precursor Ion m/z,Precursor Ion Intensity,Tags matched
1,5.23,266.121324,88319,"4, 5, 6, 12"
2,5.25,266.123180,54138,"4, 5, 6, 12"
3,4.36,309.166030,938000,"8, 10"
```

### 處理結果
去除重複訊號後,按強度排序:
```csv
ID,RT (min),Precursor Ion m/z,Precursor Ion Intensity,Tags matched
3,4.36,309.166030,938000,"8, 10"
1,5.23,266.121324,88319,"4, 5, 6, 12"
```

## ⚙️ 技術細節

### 去重複演算法
1. 掃描所有數據點
2. 計算 m/z 相對差異: `|m/z₁ - m/z₂| / max(m/z₁, m/z₂)`
3. 如果 m/z 差異 ≤ 容差 **且** RT 差異 ≤ 容差,視為重複
4. 保留強度較高的訊號

### 系統需求
- **Python**: 3.8 或更新版本 (僅原始碼執行需要)
- **記憶體**: 建議至少 4GB RAM
- **儲存空間**: 約 100MB (執行檔版本)

### 依賴套件
```
pandas >= 2.0.0
openpyxl >= 3.0.0
numpy >= 1.24.0
```

## 🐛 常見問題

<details>
<summary><b>Mac 用戶無法開啟程式?</b></summary>

如果提示「無法打開,因為來自未識別的開發者」:

**方法 1 (推薦)**:
1. 右鍵點擊程式
2. 選擇「打開」
3. 再次點擊「打開」確認

**方法 2**:
```bash
xattr -cr 質譜數據處理工具
```

**方法 3**:
系統偏好設定 → 安全性與隱私 → 仍要打開
</details>

<details>
<summary><b>無法識別欄位?</b></summary>

確認你的檔案:
1. 第一列包含標頭
2. 標頭名稱包含 RT, m/z, Intensity 的關鍵字
3. 欄位中有有效數據 (非空白)

查看錯誤訊息中的「可用的欄位」列表
</details>

<details>
<summary><b>處理速度慢?</b></summary>

對於大型檔案 (>10000 筆):
1. 考慮先篩選數據
2. 使用較大的容差減少運算
3. 使用 Python 原始碼版本 (可能更快)
</details>

## 📝 版本紀錄

### v1.1.0 (2025-10-23)
- ✨ 初始版本發布
- 🔍 自動欄位識別功能
- 📁 支援 Excel, CSV, TSV 格式
- 🎨 圖形化使用者介面
- 💾 完整保留其他欄位資訊

查看完整的 [更新紀錄](CHANGELOG.md)

## 🤝 貢獻

歡迎提交 Issue 或 Pull Request!

1. Fork 這個專案
2. 創建你的功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request



## 👨‍💻 作者

**你的名字**
- GitHub: bosschen0429
- Email: j25057175@gmail.com

## 📧 聯絡

如有問題或建議,請:
- 提交 [Issue](../../issues)
- 發送 Email: j25057175@gmail.com

