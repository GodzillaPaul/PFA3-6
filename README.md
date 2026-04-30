# 富邦人壽商品試算

線上即時試算富邦人壽外幣分紅、終身壽險等商品保單利益。引擎於前端 JavaScript 執行,計算結果與富邦原始 Excel 對齊。

## 線上瀏覽

部署到 GitHub Pages 後,訪問:
- 商品總覽: `https://<你的帳號>.github.io/<repo>/`
- PFA 試算: `https://<你的帳號>.github.io/<repo>/app.html?p=PFA`

## 目錄結構

```
.
├── index.html                  商品總覽頁(列出所有可選商品)
├── app.html                    試算頁本體(讀 ?p=XXX 切商品)
├── products/
│   └── PFA/
│       ├── config.json         商品設定(費率限制、折扣規則、模式、欄位、回饋條件)
│       └── tables.json         9 張精算對照表(從原始 Excel 抽取)
├── tools/
│   └── extract_pdata.py        從新版富邦 .xls 抽取 tables.json 的工具
├── .gitignore
└── README.md
```

## 新增商品流程

要把新商品(例如 PFW 美好優退)接進來:

1. **準備檔案**:富邦原始 .xls + 投保規則 PDF + 費率表 PDF
2. **抽精算表**:
   ```bash
   python tools/extract_pdata.py PFW path/to/PFW.xls
   ```
   會產生 `products/PFW/tables.json`。
3. **寫商品設定**:複製 `products/PFA/config.json` 為 `products/PFW/config.json`,依規則修改:
   - `code`、`name`、`fullName`、`keyPrefix`
   - `periodOptions`(年期選項)
   - `sumLimits`(年齡分段保額限制)
   - `discount`(折扣規則)
   - `modes`(計算模式定義)
   - `columns`(表格欄位)
   - `bonus`(馬上幸福 / 富邦錢包門檻)
4. **加入總覽頁卡片**:在 `index.html` 的「已上線商品」區加一張 `<a class="card" href="app.html?p=PFW">` 連結。
5. **檢查與微調**:`app.html?p=PFW` 開起來看版面、欄位、數字是否正確。
6. **commit & push** → GitHub Pages 自動部署。

## 計算引擎概要

引擎複現富邦 Excel 的「中分紅」情境(繳清保險金額累積),公式涵蓋:

| 欄位 | 公式 |
|---|---|
| A 身故保障 | `ROUNDUP(DIE × 萬)` |
| B 解約金 | `ROUNDUP(CV × 萬)` |
| K 年度紅利 | `ROUNDUP(萬 × BONU) + ROUNDUP(累積增額 × BONU / 10000)` |
| M 繳清保額(每年新增) | `ROUND(K × 10000 / 下一年 PVFB)` |
| Z 累計增加保險金額 | `Σ M[1..n]` |
| D 終期紅利-身故 | `ROUNDUP(BONUDIE × 萬)` |
| E 終期紅利-解約 | `ROUNDUP(BONUCV × 萬)` |
| X 身故增額 | `ROUNDUP(Z × _DIE2 / 10000)` (年屆滿 110 加 K) |
| Y 解約增額 | `ROUNDUP(Z × PVFB0 / 10000)` (年屆滿 110 加 K) |
| 身故總給付 | `A + D + X` |
| 解約總給付 | `B + E + Y` |

引擎以整數先乘後除避免 JavaScript 浮點誤差,在 PFA 已驗證 100% 對齊原 Excel 990 個儲存格。

## 不上傳的檔案(`.gitignore`)

- 原始 `.xls` / `.xlsx`(富邦版權)
- 備份 `.zip`
- 早期單檔 HTML
- Python 暫存

## 反推計算(模式)

每個商品可定義多種輸入模式,後端反推保額:

| 模式 | 輸入 | 反推 |
|---|---|---|
| `total` | 繳費期間總存(萬<幣別>) | 總存 ÷ 年期 → 年存 → 反推保額 |
| `annual` | 每年存(萬<幣別>) | 直接反推保額 |
| `target` | 目標金額(萬<幣別>) | 對 `targetField` 在 `targetYear` 的值,線性估算後微調至最接近 |

## 回饋金規則(共通)

- **馬上幸福**(分期繳 % 回饋):依年期決定門檻與回饋率,折扣後年繳達標可享 1% 或 2% × 折扣後年繳 × 匯率(回饋台幣)
- **富邦錢包**(NT$ 8,000):達門檻一次性 NT$8,000 刷卡金,每戶限一次

兩個門檻在 `config.json` 的 `bonus` 區設定。
