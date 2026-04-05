# 工作交接助手（Web）

網頁版交接工具，支援：
- 待辦新增 / 修改 / 完成 / 置頂
- 當日事項與倒數提醒
- 日期查詢與狀態篩選
- 一鍵輸出 Word（Arr / Dep / Occ + Daily Briefing）
- 雲端同步（Cloudflare Workers + D1）

## 本地開啟

直接開啟 `index.html` 即可使用。

## 雲端同步設定（Cloudflare D1）

1. 進入 `cloudflare-worker` 並部署 Worker（見 `cloudflare-worker/README.md`）。
2. 取得 Worker 網址後，在 `index.html` 的底部設定：

```html
<script>
  window.HANDOVER_CLOUD_API_BASE = "https://你的-workers-url.workers.dev";
</script>
```

3. 重新整理網頁，即可讓同一個使用者（目前 `caesarmetro`）共用同一份雲端資料。

## kvdb 舊資料搬移

前端已內建一次性搬移流程：
- 當 D1 還沒有資料時，會先讀舊 kvdb
- 匯入後寫入 D1
- 後續不重複搬移

## 專案檔案

- `index.html`：頁面結構
- `styles.css`：介面樣式
- `app.js`：待辦邏輯、提醒、Word 匯出、雲端同步
- `cloudflare-worker/`：D1 API（Worker）
