# 圖書館尋寶・老師體驗版

此版本位於 `demo/teachers` 分支；`main` 保留學生版。

- 仍需以一張借書證註冊兩人小隊，借書證格式沿用原版。
- 隊名沒有字數上限；隊員填姓名或編號，不限定數字、長度，也允許同名。
- 隊名及兩位隊員資料仍為必填。同一張借書證重新註冊會更新小隊。
- 註冊、刷書及排行榜都不使用班級參數。
- 指定書單最多 100 本，小隊找到 10 本不同的指定書即完成。

## 網址

將下列參數加在部署網址後：

| 頁面 | 參數 |
| --- | --- |
| 小隊註冊 | `?page=register` |
| 刷書 | 不需參數 |
| 排行榜 | `?page=leaderboard` |
| 書單管理 | `?page=admin` |

## 資料與同步

老師小隊存入 `DemoTeams`，尋寶紀錄存入 `DemoBorrowRecords`，兩張工作表會自動建立。指定書單沿用 `TreasureBooks`；學生版的 `Teams`、`BorrowRecords` 不會被讀寫。

推送分支：`git push -u origin demo/teachers`。在 Apps Script 的 GitHub Assistant 選擇 `demo/teachers` 後拉取，再建立或更新老師版部署。若要保留原學生網站，請另建老師版部署，讓學生部署繼續使用原版本。

## 本機驗證

執行 `node tests/teacher-demo.cjs`。測試使用記憶體試算表與網頁環境替身，不會寫入 Google 帳號；實際掃描器與 Apps Script 授權仍須部署後確認。
