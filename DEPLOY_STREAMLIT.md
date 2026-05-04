# Streamlit Community Cloud 部署

Cloud Run 如果因為 billing 行不通，可以改用 Streamlit Community Cloud。官方文件說 Community Cloud 可以從 GitHub repo 免費部署 Streamlit app，部署後會得到 `streamlit.app` 公開網址。

## 部署步驟

1. 到 https://share.streamlit.io/
2. 用 GitHub 登入。
3. 點 Create app。
4. 選 Yup, I have an app。
5. Repository 選：

```text
an4131754-collab/football-draw-tool
```

6. Branch 選：

```text
master
```

7. Main file path 填：

```text
streamlit_app.py
```

8. Python version 選 3.12。
9. 點 Deploy。

部署完成後，網址會長得像：

```text
https://football-draw-tool.streamlit.app
```

或 Streamlit 自動分配的子網域。

## 注意

- 線上版請直接上傳報名表 Excel，不要依賴本機 fallback 檔案。
- GitHub repo 已排除 `114...回覆.xlsx`、`outputs/` 和最終版輸出檔。
- Streamlit Cloud 適合這種抽籤工具，但如果很多人同時用，免費資源可能會比較慢。

官方文件：

- https://docs.streamlit.io/deploy/streamlit-community-cloud
- https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy
