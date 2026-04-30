# 足球系際盃抽籤工具

這個專案可以讀取同格式 Google 表單回覆 Excel，進行足球系際盃抽籤、分組與賽程安排。

## 本機啟動

雙擊：

```powershell
start_draw_site.bat
```

或在終端機執行：

```powershell
C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe app.py
```

啟動後打開：

```text
http://127.0.0.1:8000
```

## 網站流程

1. 上傳同格式 Google 表單回覆 Excel。
2. 選擇組數、每組晉級名額、最佳名次補位數。
3. 選擇 DAY1 / DAY2 最晚踢到幾點；預設沿用兩天兩場地。
4. 勾選本次要提供下載的項目：JSON、Excel、PDF。
5. 按「開始抽籤」。
6. 抽完只會顯示結果與下載按鈕，不會自動下載。

如果沒有上傳檔案，網站會嘗試使用資料夾內的本機 fallback 報名表。

## 賽程規則

- 小組賽採完整單循環。
- 2 隊小組可排；1 隊小組會顯示排程不可行，請調整組數。
- DAY1 預設排小組賽，時段為 `09:00`、`10:00`、`11:00`、`14:00`、`15:00`、`16:00`。
- DAY2 預設排淘汰賽，時段為 `10:00`、`11:00`、`14:00`、`15:00`。
- 每時段有甲、乙兩場地，每場 45 分鐘。
- 同一隊同一天最多 3 場，並盡量避免連續出賽。
- 排不下時不會硬塞不公平賽程，會提示延後最晚結束時間、調整組數或減少晉級隊數。

## 12 隊 4 組模板

當隊伍數是 12 且組數是 4，工具會沿用 `113系際盃賽程及裁判表_公開版.xlsx`：

- DAY1 小組賽自動填入模板指定位置。
- DAY2 四強固定為 `A1 vs C1`、`B1 vs D1`。
- 10:00 打兩場四強。
- 14:00 同時打季軍賽與冠軍賽。
- 11:00 與 15:00 時段保留空白。

非 12 隊 4 組時，工具會產生新版動態賽程 Excel，包含分組表、賽程表與排程說明。

## CLI

根據最新抽籤結果重新產生輸出：

```powershell
C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe scheduler.py --draw outputs/latest/draw_result.json
```

只重新產生 PDF：

```powershell
C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe scheduler.py --draw outputs/latest/draw_result.json --pdf-only
```

重新排程並指定加開時段：

```powershell
C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe scheduler.py --draw outputs/latest/draw_result.json --day1-latest-end 18:45 --day2-latest-end 17:45
```

指定公開下載項目：

```powershell
C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe scheduler.py --draw outputs/latest/draw_result.json --outputs json,pdf,excel
```

## 設定檔

可修改 `config.json`：

- `team_column`：報名表隊名欄位，預設 `科系`
- `default_group_count`：網站預設組數
- `default_advance_per_group`：預設每組晉級名額
- `default_wildcard_count`：預設最佳名次補位數
- `fields`：場地名稱，預設 `甲`、`乙`
- `match_duration_minutes`：每場分鐘數，預設 45
- `max_matches_per_team_per_day`：同隊每日最多場次，預設 3
- `day_slots`：DAY1 / DAY2 預設開始時間
- `latest_end_options`：網站可選的加開最晚結束時間
- `schedule_filename`：Excel 輸出檔名
- `pdf_filename`：PDF 輸出檔名
- `semifinal_labels`：12 隊 4 組模板的 DAY2 四強文字
- `final_labels`：12 隊 4 組模板的 DAY2 季軍賽與冠軍賽文字

## 部署準備

部署到 Render/Railway 類平台時，安裝：

```powershell
pip install -r requirements.txt
```

平台啟動指令可用：

```text
gunicorn app:app
```

`Procfile` 已設定：

```text
web: gunicorn app:app
```
