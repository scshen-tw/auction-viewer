# CLAUDE.md — TWSE/TPEX 競拍資料庫

## 專案簡介
台灣上市／上櫃／CB 可轉債競拍歷史資料庫，自動抓取 2016 年至今的競拍資料，產生單一 HTML 頁面並推送至 GitHub。

---

## 架構說明

### 資料流
```
auction_fetcher.py --update
  ├─ 抓取 TWSE / TPEX 競拍清單
  ├─ 補收盤價（FinMind API）
  ├─ 補 CB 轉換價（TPEX 官方 API）
  ├─ 更新 auction_stocks.json / auction_cbs.json / price_cache.json
  ├─ 呼叫 generate_html() → 產生 auction_viewer.html
  └─ 呼叫 _git_push() → 自動 commit + push
```

### 關鍵檔案
| 檔案 | 說明 |
|------|------|
| `auction_fetcher.py` | 主程式：資料抓取、HTML 產生、git push |
| `auction_viewer.html` | **自動產生的輸出檔**，勿直接長期修改（見下方注意事項）|
| `auction_stocks.json` | 股票競拍資料庫 |
| `auction_cbs.json` | CB 可轉債競拍資料庫 |
| `price_cache.json` | 收盤價快取，避免重複呼叫 API |
| `2025Q3TCRI.xlsx` | TCRI 信評對照表（每季更新，替換後執行 `--html`）|
| `每日更新.bat` | Windows 排程呼叫的腳本 |

---

## ⚠️ 最重要的注意事項：CSS / HTML 修改方式

`auction_viewer.html` 是由 `auction_fetcher.py` 內的 `_HTML_TEMPLATE` 字串產生的。

**每次執行 `--update` 都會重新產生 auction_viewer.html，覆蓋掉對該檔案的直接修改。**

### 正確的修改流程
修改 CSS、欄位、JavaScript 時，必須同時修改兩處：
1. `auction_fetcher.py` 內的 `_HTML_TEMPLATE`（約第 724 行開始）← **永久生效**
2. `auction_viewer.html` ← 讓當下立即看到效果

修改完後執行：
```bash
python auction_fetcher.py --html   # 從模板重新產生 HTML（不抓資料）
```

---

## 執行模式

```bash
python auction_fetcher.py --update   # 增量更新（日常使用）
python auction_fetcher.py --html     # 僅重新產生 HTML
python auction_fetcher.py            # 全量重建（重抓 2016~現在，少用）
```

---

## 自動排程
Windows 工作排程器已設定：
- `競拍資料庫每日更新`：每天 **08:00**
- `競拍資料庫每日更新1010`：每天 **10:10**

兩者都執行 `每日更新.bat`，更新完有新資料會自動 git push。

---

## 資料來源
| 資料 | 來源 |
|------|------|
| 競拍清單 | TWSE / TPEX OpenAPI |
| 收盤價 | FinMind API（付費帳號，token 在 py 檔頂端）|
| CB 轉換價 | TPEX 官方 API |
| TCRI 信評 | 手動更新 Excel 檔 |

---

## 目前 UI 主題
**Bloomberg Terminal 風格**（2026-04 採用，使用者確認喜歡）
- 黑底 `#000000`、橘色標題 `#FF6600`
- 等寬字型 Courier New / Consolas
- 全無圓角（`border-radius: 0`）
- 漲綠 `#00CC44`、跌紅 `#FF3333`

歷史主題備份在 memory：
- 原始藍色系 → `memory/ui_original_colors.md`
- Hermès 橘棕系 → 已被 Bloomberg 取代

---

## 欄位結構（股票 / CB 共用）
- `投標結束日`、`市場`（上市/上櫃）、`代號`、`名稱`
- `競拍數量(張)`、`合格投標量(張)`
- `得標均價`、`最低得標價`、`收盤價`
- `掛牌日`（撥券日期 / 上市上櫃日期）
- CB 額外欄位：`轉換價`、`發行公司`、`到期日`、`TCRI 信評`

---

## 資料列狀態
| class | 意義 | 背景色 |
|-------|------|--------|
| `s-done` | 結標完成 | `#001400`（深綠）|
| `s-pending` | 待結果 | `#3E3E3E`（深灰）|
| `s-noresult` | 流標 | `#1A0800`（深橘紅）|
| `s-cancel` | 取消 | `#0D0D0D`（近黑）|

---

## Git
- Branch：`main`
- Remote：`https://github.com/scshen-tw/auction-viewer.git`
- `_git_push()` 在 `update_data()` 結尾自動執行
- 只有真正有資料變動才會 commit，不產生空 commit
