# Interdept Cup Draw Studio

足球系際盃抽籤與賽程產生工具。支援上傳同格式 Google 表單回覆 Excel，抽出分組，產生 JSON、Excel 與 PDF 隨機性說明。

目前有兩個入口：

- `streamlit_app.py`：給 Streamlit Community Cloud 部署使用，UI 是黑白精品風新版。
- `app.py`：Flask 版本，可本機或 Render/Railway 類平台使用。

## Features

- 上傳任意同格式 Excel，不綁死單一回覆表檔名。
- 隊名欄位預設讀取 `科系`。
- 支援非 12 隊抽籤，組數可調，最少 2 組、最多 26 組。
- 使用 `secrets.SystemRandom().shuffle()` 洗牌，亂數來自作業系統安全亂數來源。
- 可設定每組晉級名額、最佳名次補位數、四強或八強淘汰賽格式。
- 可選擇輸出 JSON、Excel、PDF。
- 12 隊 4 組時可沿用 `113系際盃賽程及裁判表_公開版.xlsx` 產完整賽程。
- 非 12 隊 4 組時會產分組結果 Excel，不硬套固定模板。

## Streamlit 部署

Streamlit Community Cloud 設定：

```text
Repository: an4131754-collab/football-draw-tool
Branch: master
Main file path: streamlit_app.py
Python version: 3.12
```

Streamlit 會從 repo root 讀取 `requirements.txt`，並執行：

```text
streamlit run streamlit_app.py
```

部署後請直接在網頁上傳報名表 Excel。repo 不包含 `114...回覆.xlsx` 與 `outputs/`，避免報名資料外流。

## 本機測試

Flask 版本：

```powershell
start_draw_site.bat
```

打開：

```text
http://127.0.0.1:8000
```

Streamlit 版本：

```powershell
streamlit run streamlit_app.py
```

如果你的終端機找不到 `streamlit`，可先安裝依賴：

```powershell
pip install -r requirements.txt
```

## 使用流程

1. 上傳 Google 表單回覆 Excel。
2. 選擇組數。
3. 設定每組前 N 名晉級。
4. 設定最佳名次補位數。
5. 選擇淘汰賽格式。
6. 選擇輸出 JSON、Excel、PDF。
7. 按下抽籤。
8. 下載需要的輸出檔。

## 12 隊 4 組賽程規則

當隊伍數為 12 且組數為 4 時，會套用公開版模板產完整賽程：

- DAY1 自動安排小組賽。
- DAY2 四強為 `A1 vs C1`、`B1 vs D1`。
- 四強預設 10:00 開打。
- 季軍賽與冠軍賽預設 14:00 開打。
- 11:00 與 15:00 時段保留空白。

## CLI

根據最新抽籤結果重新產生輸出：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json
```

只重新產生 PDF：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json --pdf-only
```

調整最晚結束時間：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json --day1-latest-end 18:45 --day2-latest-end 17:45
```

指定輸出格式：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json --outputs json,pdf,excel
```

## 設定檔

主要設定在 `config.json`：

- `team_column`：報名表隊名欄位，預設 `科系`。
- `default_group_count`：預設組數。
- `default_advance_per_group`：預設每組晉級名額。
- `default_wildcard_count`：預設最佳名次補位數。
- `fields`：場地名稱。
- `match_duration_minutes`：每場比賽分鐘數。
- `max_matches_per_team_per_day`：同隊每日最多比賽數。
- `day_slots`：DAY1 / DAY2 可用時段。
- `latest_end_options`：網站可選的最晚結束時間。
- `schedule_filename`：Excel 輸出檔名。
- `pdf_filename`：PDF 輸出檔名。
- `semifinal_labels`：12 隊 4 組模板的 DAY2 四強文字。
- `final_labels`：12 隊 4 組模板的 DAY2 決賽與季軍賽文字。

## Render / Railway

如果不用 Streamlit，而是部署 Flask 版本：

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
```

`Procfile` 已設定：

```text
web: gunicorn app:app
```

## 資料安全

`.gitignore` 已排除：

- `outputs/`
- `*回覆*.xlsx`
- `*最終版*.xlsx`
- `*最終版*.pdf`

也就是說，報名回覆表與產出的結果檔不會被推上 GitHub。
