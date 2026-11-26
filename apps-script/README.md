# AngelNote Gmail 自動提醒設定

## 功能

- 在 AngelNote 的待辦編輯視窗勾選「Gmail 自動傳送提醒」，設定提前天數（0–365）。
- 排程每天約上午 9 點（台北時間）檢查尚未完成的任務，並把當日符合條件的任務合併成一封提醒信，寄到執行 Apps Script 的 Google 帳號。
- 只寄送待辦任務與截止日期；已完成任務、行程與未勾選 Gmail 提醒的項目不會寄送。
- 每筆待辦同一截止日只寄一次。完成待辦或更改截止日會更新後續判斷。

## 建立排程

1. 用與 AngelNote 相同的 Google 帳號開啟 <https://script.google.com/home/start>，建立「新專案」。
2. 將 `Code.gs` 全部貼入 Apps Script 編輯器並儲存。
3. 在專案設定啟用「在編輯器中顯示 `appsscript.json` 資訊清單檔案」，打開資訊清單並貼上本資料夾的 `appsscript.json`。
4. 在函式下拉選單選擇 `setupDailyGmailReminder`，按「執行」。首次執行時依 Google 提示授權寄信、讀取 Firestore 與外部連線；需使用 AngelNote 登入相同的 Google 帳號。
5. 成功後 Apps Script 會建立每日台北時間上午 9 點的觸發器，可在左側「觸發條件」確認。

如需停止自動寄送，在 Apps Script 執行 `removeDailyGmailReminder`，或刪除 `sendDueGmailReminders` 觸發器。
