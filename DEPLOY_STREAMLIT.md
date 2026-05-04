# Streamlit Community Cloud 部署

如果你已經選 Streamlit Community Cloud，請用這些設定：

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

本機測試也可以在專案根目錄跑：

```powershell
streamlit run streamlit_app.py
```

注意事項：

- 線上部署後請在網站上傳報名表 Excel。
- repo 沒有推 `114...回覆.xlsx` 與 `outputs/`，避免報名資料外流。
- Streamlit Cloud 的檔案系統是暫時性的，抽籤後下載檔請當次下載保存。
