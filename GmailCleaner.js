/*
 *本程式的限制：Google Apps Script有固定的執行時長限制，所以如果信件太多，可能會刪不完，要自己再次執行<-後續有提供一個auto_click_gmail_cleaner.py做自動點擊的工作
 * /

 // --- 設定區：請在此修改您要執行的操作 ---
const CONFIG = {
  //TODO 改成您要搜尋的郵件關鍵字，語法和 Gmail 搜尋框完全相同
  SEARCH_QUERY: "unsubscribe",

  // 您要執行的操作：'DELETE' (移至垃圾桶) 或 'ARCHIVE' (封存)
  OPERATION: 'DELETE', 

  // 每一批次處理的會話群組數量。Gmail API 建議一次最多處理 500 個。
  BATCH_SIZE: 100 
};
// --- 設定區結束 ---


/**
 * 主執行函式 - 帶有迴圈，可以處理所有符合條件的郵件
 */
function cleanUpAllEmails() {
  Logger.log(`--- 開始執行郵件清理任務 ---`);
  Logger.log(`搜尋條件: "${CONFIG.SEARCH_QUERY}"`);
  Logger.log(`執行操作: ${CONFIG.OPERATION}`);

  let threads;         // 用來存放每一批的會話群組
  let totalProcessed = 0; // 紀錄總共處理了多少會話群組

  // 使用 do...while 迴圈來確保至少執行一次，並在還有郵件時繼續執行
  do {
    try {
      // 獲取下一批符合條件的會話群組
      // 注意：GmailApp.search 會自動處理分頁，我們只需要重複呼叫它
      threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, CONFIG.BATCH_SIZE);

      if (threads.length > 0) {
        Logger.log(`找到新的一批 ${threads.length} 個會話群組，準備處理...`);
        
        // 根據設定執行操作
        if (CONFIG.OPERATION.toUpperCase() === 'DELETE') {
          GmailApp.moveThreadsToTrash(threads);
        } else if (CONFIG.OPERATION.toUpperCase() === 'ARCHIVE') {
          GmailApp.moveThreadsToArchive(threads);
        } else {
          Logger.log(`錯誤：無效的操作 "${CONFIG.OPERATION}"。`);
          return; // 設定錯誤，直接終止
        }

        totalProcessed += threads.length;
        Logger.log(`本批處理完畢。目前已累計處理 ${totalProcessed} 個會話群組。`);

        // 短暫暫停一下，避免觸發 Google 的 API 頻率限制
        Utilities.sleep(1000); // 暫停 1 秒

      } else {
        // 如果 threads.length 是 0，表示已經沒有符合條件的郵件了
        Logger.log("所有符合條件的郵件都已處理完畢。");
      }

    } catch (e) {
      Logger.log(`處理過程中發生錯誤: ${e.toString()}`);
      Logger.log(`已處理 ${totalProcessed} 個會話群組後中斷。`);
      return; // 發生錯誤，中斷執行
    }
  // 迴圈條件：只要上一批次處理的數量等於設定的 BATCH_SIZE，就代表可能還有更多郵件，繼續迴圈
  } while (threads.length === CONFIG.BATCH_SIZE); 

  Logger.log(`--- 郵件清理任務執行完畢 ---`);
  Logger.log(`總共處理了 ${totalProcessed} 個會話群組。`);
}