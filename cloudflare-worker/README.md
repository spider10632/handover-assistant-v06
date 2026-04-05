# Cloudflare D1 雲端資料庫

此資料夾提供 `工作交接助手` 的雲端 API（Workers + D1）。

## 1) 建立 D1

```bash
cd cloudflare-worker
npm install
npm run db:create
```

建立後會拿到 `database_id`，請填入 `wrangler.toml`：

```toml
[[d1_databases]]
binding = "DB"
database_name = "handover-db"
database_id = "你的 database_id"
```

## 2) 建表

```bash
npm run db:migrate
```

## 3) 部署 Worker

```bash
npm run deploy
```

部署成功後會得到網址，例如：

`https://handover-cloud-api.<your-subdomain>.workers.dev`

## 4) 前端設定 API

在 `index.html` 的 `app.js` 前加入（或修改）以下設定：

```html
<script>
  window.HANDOVER_CLOUD_API_BASE = "https://你的-workers-url.workers.dev";
</script>
```

前端會呼叫：

`GET/PUT {CLOUD_API_BASE}/v1/state/caesarmetro`

## 5) kvdb 舊資料搬移

前端已內建一次性搬移：
- 當 D1 目前沒有資料時
- 會自動從舊的 kvdb 讀取 `caesarmetro` 資料
- 再寫入 D1
- 之後會標記完成，不重複搬移

## API

- `GET /health`
- `GET /v1/state/:serverId`
- `PUT /v1/state/:serverId`

`wrangler.toml` 的 `ALLOWED_USERS` 可設定允許的 serverId（逗號分隔）。
