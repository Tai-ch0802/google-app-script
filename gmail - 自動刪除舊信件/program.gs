/**
 * 自動刪除 Gmail 中多個特定標籤且超過指定天數的郵件 (自訂設定版本)。
 */
function deleteOldMailsWithLabel() {
  // **程式碼區 -  請勿修改以下程式碼，除非您了解程式碼運作原理 **

  // 1. 取得標籤設定 (從 private function _getLabelConfigurations 取得)
  const labelConfigurations = _getLabelConfigurations();

  // 檢查是否有標籤設定，若無則結束程式
  if (!labelConfigurations || labelConfigurations.length === 0) {
    Logger.log("錯誤：未設定任何標籤和天數。請檢查 _getLabelConfigurations() 函式。");
    return; // 結束程式
  }

  Logger.log("開始處理郵件...");

  // 2. 迴圈處理每個標籤設定
  for (let i = 0; i < labelConfigurations.length; i++) {
    const config = labelConfigurations[i];
    const labelName = config.labelName;
    const daysToDelete = config.daysToDelete;

    // 取得指定的 Gmail 標籤
    const label = GmailApp.getUserLabelByName(labelName);

    // 檢查標籤是否存在，若不存在則記錄並繼續處理下一個標籤
    if (!label) {
      Logger.log("標籤 '" + labelName + "' 不存在，略過此標籤。");
      continue; // 略過不存在的標籤，繼續處理下一個
    }

    // 取得標籤下的所有郵件串 (Threads)
    const threads = label.getThreads();
    const now = new Date(); // 取得目前時間

    Logger.log("處理標籤 '" + labelName + "'，郵件超過 " + daysToDelete + " 天的郵件串，共有 " + threads.length + " 個郵件串。");

    // 3. 迴圈處理每個郵件串
    for (let j = 0; j < threads.length; j++) { // 修改迴圈變數名稱，避免與外層迴圈衝突
      const thread = threads[j];
      const messages = thread.getMessages(); // 取得郵件串中的所有郵件

      // 4. 取得郵件串中第一封郵件的日期 (以第一封郵件的日期作為整個郵件串的日期判斷)
      const firstMessage = messages[0]; // 假設第一封郵件為郵件串的代表日期
      const sentDate = firstMessage.getDate(); // 取得第一封郵件的寄送日期

      // 5. 計算郵件寄送日期至今的天數差
      const timeDiff = now.getTime() - sentDate.getTime(); // 時間差 (毫秒)
      const daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24)); // 毫秒轉天數 (無條件進位)

      // 6. 判斷郵件是否超過指定天數，若超過則刪除 (移動到垃圾桶)
      if (daysDiff > daysToDelete) {
        thread.moveToTrash(); // 將郵件串移動到垃圾桶 (刪除)
        Logger.log("  - 郵件串主旨: '" + thread.getFirstMessageSubject() + "'，寄送日期: " + sentDate + "，已超過 " + daysToDelete + " 天，已移動到垃圾桶。");
      }
    }
    Logger.log("標籤 '" + labelName + "' 郵件處理完成。");
  }

  Logger.log("所有標籤郵件處理完成。");
}

/**
 * **Private Function (底線開頭表示私有函式)**
 * 定義要處理的 Gmail 標籤及其對應的郵件刪除天數。
 * 您可以在這個函式中新增或修改標籤設定。
 *
 *  回傳格式為一個陣列，每個元素都是一個物件，包含 labelName (標籤名稱) 和 daysToDelete (刪除天數)。
 *  例如:
 *  [
 *    { labelName: "舊郵件", daysToDelete: 7 },
 *    { labelName: "待辦事項", daysToDelete: 3 },
 *    { labelName: "促銷郵件", daysToDelete: 30 }
 *  ]
 */
function _getLabelConfigurations() {
  //  ** 在這裡修改您的標籤和天數設定 **
  return [
    { labelName: "財經/seeking-alpha", daysToDelete: 4 },
    { labelName: "財經/財經 M 平方", daysToDelete: 3 },
    { labelName: "財經/個人提醒通知", daysToDelete: 2 },
  ];
}
