# Interdept Cup Draw Studio

足球系際盃抽籤、賽程與裁判表產生工具。可在 Streamlit 線上部署，也可用本機 Flask 版本執行。

## Features

- 上傳同格式 Google 表單回覆 Excel，預設讀取 `科系` 欄位作為隊名。
- 支援非 12 隊抽籤，可自選組數、每組晉級名額、外卡名額。
- 使用 `secrets.SystemRandom().shuffle()` 進行抽籤，並輸出 JSON / Excel / PDF。
- 支援直接四強或八強賽制；12 隊 4 組時可沿用 113 公開版模板產完整賽程。
- 抽籤完成後可安排裁判，每場 3 位，支援所屬隊伍避開與不可排場次。
- Excel 會在同一份賽程檔新增或更新「裁判」工作表。
- 支援「既有賽程只排裁判」模式：上傳已排好的賽程 Excel，不重新抽籤、不改賽程，只新增裁判表。

## Streamlit 部署

Streamlit Community Cloud 設定：

```text
Repository: an4131754-collab/football-draw-tool
Branch: master
Main file path: streamlit_app.py
Python version: 3.12
```

本機測試 Streamlit：

```powershell
streamlit run streamlit_app.py
```

如果本機沒有 `streamlit`：

```powershell
pip install -r requirements.txt
```

## 本機 Flask

雙擊或在終端執行：

```powershell
start_draw_site.bat
```

開啟：

```text
http://127.0.0.1:8000
```

## 使用流程

1. 上傳 Google 表單回覆 Excel。
2. 選擇組數、晉級規則、淘汰賽格式與輸出格式。
3. 按下 `開始抽籤` 完成抽籤與賽程。
4. 在 `Referee Setup` 輸入裁判姓名，一行一位。
5. 每位裁判可選所屬隊伍，並勾選不能擔任裁判的場次。
6. 按下 `產生裁判表`。
7. 下載 Excel，檔案會包含「裁判」工作表。

## 既有賽程只排裁判

如果你已經自己排好賽程：

1. 在左側找到 `既有賽程只排裁判`。
2. 上傳已排好的賽程 Excel。
3. 輸入裁判名單，一行一位。
4. 按下 `讀取賽程並建立裁判設定`。
5. 右側會顯示裁判設定表，可補上所屬隊伍與不可排場次。
6. 按下 `產生裁判表`，再下載 Excel。

目前支援兩種賽程格式：

- 類似 `113系際盃賽程及裁判表_公開版.xlsx` 的固定賽程表格式。
- 本工具自己產出的動態賽程 Excel。

## 裁判排班規則

- 每場固定安排 3 位裁判。
- 同一位裁判不能在同一時間被排到甲、乙兩場。
- 裁判可以連續場次執法，但系統會優先平均分配。
- 若裁判有填所屬隊伍，小組賽會避開該隊實際比賽。
- 淘汰賽尚未輸入勝負，因此會保守避開該裁判所屬隊伍可能出現的場次。
- 若人手不足，Excel 會留白並在頁面與 JSON 顯示警告。

## CLI

根據最新抽籤結果重建輸出：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json
```

只重建 PDF：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json --pdf-only
```

調整 DAY1 / DAY2 最晚結束時間：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json --day1-latest-end 18:45 --day2-latest-end 17:45
```

指定輸出格式：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json --outputs json,pdf,excel
```

用 JSON 補裁判表：

```powershell
python scheduler.py --draw outputs/latest/draw_result.json --referees referees.json
```

`referees.json` 範例：

```json
{
  "referees": [
    {
      "name": "王小明",
      "affiliated_team": "電機",
      "unavailable_match_nos": [1, 2, 13]
    },
    {
      "name": "陳小華",
      "affiliated_team": "",
      "unavailable_match_nos": []
    }
  ]
}
```

## Config

常用設定在 `config.json`：

- `team_column`：隊名欄位，預設 `科系`。
- `default_group_count`：預設組數。
- `default_advance_per_group`：每組晉級名額。
- `default_wildcard_count`：外卡名額。
- `default_knockout_format`：`semifinal` 或 `quarterfinal`。
- `fields`：場地名稱。
- `referees_per_match`：每場裁判人數，預設 3。
- `referee_sheet_name`：裁判工作表名稱，預設 `裁判`。
- `match_duration_minutes`：每場比賽分鐘數。
- `day_slots`：DAY1 / DAY2 可用開賽時間。
- `latest_end_options`：網站上可選的最晚結束時間。

## Data Safety

`.gitignore` 已排除：

- `outputs/`
- `*回覆*.xlsx`
- `*最終版*.xlsx`
- `*最終版*.pdf`

請不要把未公開的報名表或個資檔案推上 GitHub。
