# TWSE/TPEX 競拍資料庫

自動抓取台灣上市/上櫃/CB 競拍歷史資料（民國105年至今），每日自動更新並發布至 GitHub Pages。

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `auction_fetcher.py` | 主程式（資料抓取、HTML 產生、排程執行） |
| `auction_stocks.json` | 股票競拍資料庫 |
| `auction_cbs.json` | CB 可轉債競拍資料庫 |
| `price_cache.json` | 收盤價快取（避免重複呼叫 API） |
| `auction_viewer.html` | 網站頁面（自動產生，勿手動編輯） |
| `2025Q3TCRI.xlsx` | TCRI 信評對照表（依季更新） |
| `requirements.txt` | Python 套件清單 |

## 換電腦移機步驟

### 1. 複製資料夾
將整個 `Auction` 資料夾複製到新電腦（**含 `.git` 隱藏資料夾**，git 設定才會保留）。

### 2. 安裝 Python
至 https://www.python.org 下載安裝 Python 3.11 以上版本。

### 3. 安裝套件
```
pip install -r requirements.txt
```

### 4. 設定每日自動排程
以**系統管理員**身分開啟命令提示字元，執行：
```
schtasks /create /tn "TWSE_Auction_DailyUpdate" /tr "python \"d:\VS Code\Auction\auction_fetcher.py\" --update" /sc DAILY /st 08:00 /f
```
> 若資料夾路徑不同，請調整路徑。

### 5. 確認 GitHub 登入
第一次執行 `git push` 時會要求登入，使用 GitHub 帳號或 Personal Access Token 驗證即可。

---

## 手動操作指令

| 指令 | 說明 |
|------|------|
| `python auction_fetcher.py --update` | 增量更新（抓新資料＋補收盤價＋推送網站） |
| `python auction_fetcher.py --html` | 僅重新產生 HTML（不抓資料） |
| `python auction_fetcher.py` | 全量重建（重抓2016至今，通常不需要） |

## 資料來源

- 競拍清單：TWSE / TPEX OpenAPI
- 收盤價：FinMind API（付費帳號，6000 req/hr）
- TCRI 信評：依季更新，替換 `*TCRI*.xlsx` 檔案後執行 `--html` 即可套用
