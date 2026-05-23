# Claude Desktop 自動化指示

## 啟動時自動執行

每次開啟 Claude Desktop 時，請自動依序執行以下步驟：

### 步驟一：確認本週是否有新週報

從 GitHub repo `sqmochen/TDCC_WeeklyStockNumInfo` 的 `output/summary/` 資料夾，
讀取最新的 `notebooklm_summary_YYYYMMDD.txt` 檔案名稱，
確認是否有比 NotebookLM 中現有資料更新的週報。

### 步驟二：上傳至 NotebookLM 指定資料夾

若有新週報，透過 NotebookLM MCP Server 執行以下動作：

- **目標筆記本**：`TDCC 每週股東資訊`
  - 若筆記本不存在，自動建立同名筆記本
  - 不可上傳至其他筆記本
- **上傳內容**：最新的 `notebooklm_summary_YYYYMMDD.txt` 全文
- **重複防護**：若該週日期的資料已存在於筆記本中，跳過上傳，不重複新增

### 步驟三：回報執行結果

完成後主動告知使用者：

```
✅ TDCC 週報已更新
─────────────────────────
本週基準日：YYYY-MM-DD（週六）
NotebookLM 筆記本：TDCC 每週股東資訊
符合條件股票數：XX 檔（連續3週大戶持股比上漲）
─────────────────────────
```

若本週資料尚未產出（GitHub Actions 尚未執行），則告知：

```
⏳ 本週 TDCC 週報尚未產出
   GitHub Actions 將於每週日 09:00 自動執行
   上次更新：YYYY-MM-DD
```

---

## 重要規則

- NotebookLM 上傳目標**只能**是 `TDCC 每週股東資訊` 筆記本
- 不可刪除筆記本中的舊週資料
- 每週只上傳一次，不重複
