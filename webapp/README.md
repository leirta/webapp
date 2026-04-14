# 年表記帳本 Web App

這是一個純前端的個人記帳工具，資料存在瀏覽器 `localStorage`。

## 已完成

- 左側三頁切換，並同步網址 `hash`
- 年度檢視可直接編輯收入、帳戶餘額、固定支出
- 每日支出可新增、編輯、刪除，並顯示每日小計
- 支出分類可新增與編輯
- 月明細有收入區、固定支出區、支出分類區與圓餅圖

## 本機開啟

直接打開 `index.html` 即可。

## 同 Wi-Fi 給手機開

在 `webapp` 資料夾執行：

```powershell
node server.js
```

或：

```powershell
npm.cmd start
```

啟動後終端會顯示：

- `http://localhost:4173`
- `http://你的區網IP:4173`

讓 iPhone 和這台電腦連同一個 Wi-Fi，再用 Safari 打開那個區網網址即可。
