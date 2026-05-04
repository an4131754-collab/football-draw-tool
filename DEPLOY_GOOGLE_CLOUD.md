# Google Cloud Run 部署

建議使用 Google Cloud Run，因為它適合部署這種 Flask 小網站，部署完成會得到真正公開網址。

## 推薦設定

- Service name: `football-draw-tool`
- Region: `asia-east1`，也就是 Taiwan
- Build type: `Dockerfile`
- Dockerfile source location: `/Dockerfile`
- Authentication: Allow unauthenticated invocations

## 從 GitHub 部署

1. 到 Google Cloud Console。
2. 開啟 Cloud Run。
3. 選 Create service。
4. 選 Continuously deploy from a repository。
5. 連接 GitHub repository：`an4131754-collab/football-draw-tool`。
6. Branch 選 `master`。
7. Build type 選 Dockerfile。
8. Source location 填 `/Dockerfile`。
9. 允許 public access。
10. 建立服務後，Cloud Run 會給一個公開網址。

公開網址會長得像：

```text
https://football-draw-tool-xxxxx-de.a.run.app
```

## 用 gcloud 從本機部署

如果你有安裝 Google Cloud CLI，也可以在專案資料夾執行：

```powershell
gcloud init
gcloud config set project YOUR_PROJECT_ID
gcloud run deploy football-draw-tool --source . --region asia-east1 --allow-unauthenticated
```

Google Cloud 官方文件：

- https://cloud.google.com/run/docs/quickstarts/deploy-continuously
- https://cloud.google.com/run/docs/deploying-source-code
