/**
 * 每日獲取市場指數資料 (從 Google 試算表), CNN 恐懼與貪婪指數，並寄送郵件報告 (表格化版本).
 */
function sendDailyMarketIndexReport() {
  // ** 設定區 **
  const recipientEmail = "";  //  請將 "您的 Gmail 信箱地址"  替換成您的 Gmail 電子郵件地址
  const emailSubject = "每日市場指數報告 (表格化)"; // 修改郵件主旨
  const timeZone = "Asia/Taipei"; // 時區設定為台北，您可以根據您的所在地調整時區

  // ** 程式碼區 - 請勿修改以下程式碼，除非您了解程式碼運作原理 **

  try {
    // 1. 取得市場指數資料 (從 Google 試算表)
    const indexData = getIndex();

    //  檢查 getIndex() 是否回傳錯誤訊息字串
    if (typeof indexData === 'string' && indexData.startsWith("N/A -")) {
      throw new Error(indexData); //  如果是錯誤訊息，直接拋出錯誤，讓 catch 區塊處理錯誤通知郵件
    }


    // 2. 取得 CNN 恐懼與貪婪指數
    const fearGreedIndexData = getFearGreedIndex();
    const fearGreedIndexValue = fearGreedIndexData.indexValue;
    const fearGreedIndexCategory = fearGreedIndexData.indexCategory;


    // 3. 建立郵件內容 (HTML 表格化)
    let emailBody = `<!DOCTYPE html><html><head><style>body { font-family: Arial, sans-serif; } .index-value { font-size: 18px; font-weight: bold; color: #3367D6; } .index-category { font-style: italic; color: #666; } table { border-collapse: collapse; width: 100%; margin-top: 20px; } th, td { border: 1px solid #ccc; padding: 8px; text-align: left; } th { background-color: #f0f0f0; } .positive-change { background-color: #e0f7fa; /* 淺綠色 */ } .negative-change { background-color: #ffdddd; /* 淺紅色 */ }</style></head><body>`; // 加入 positive-change 和 negative-change CSS class
    emailBody += `<h1>每日市場指數報告</h1>`;
    emailBody += `<p>日期：${Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss")}</p>`; // 顯示報告產生時間 (台北時區)

    //  建立市場指數資料表格
    emailBody += `<h2>市場指數</h2>`;
    emailBody += `<table>`;
    emailBody += `<thead><tr><th>名稱</th><th>價格/指數</th><th>和前一日變動百分比</th><th>前一天的收盤價</th><th>52週最高價</th><th>52週最低價</th></tr></thead><tbody>`;

    if (Array.isArray(indexData)) { // 檢查 indexData 是否為陣列
      indexData.forEach(item => {
        const changePercent = parseFloat(item.changePercent); // 先將變動百分比轉換為數字

        let changePercentClass = ''; // 預設為空 class
        if (!isNaN(changePercent)) { // 檢查是否為有效數字
          if (changePercent > 0) {
            changePercentClass = 'positive-change'; // 正值套用 positive-change class
          } else if (changePercent < 0) {
            changePercentClass = 'negative-change'; // 負值套用 negative-change class
          }
        }

        emailBody += `<tr>`;
        emailBody += `<td>${item.name}</td>`;
        emailBody += `<td class="index-value">${item.price}</td>`;
        emailBody += `<td class="${changePercentClass}">${item.changePercent}</td>`; // 動態加入 class
        emailBody += `<td>${item.previousClose}</td>`;
        emailBody += `<td>${item.fiftyTwoWeekHigh}</td>`;
        emailBody += `<td>${item.fiftyTwoWeekLow}</td>`;
        emailBody += `</tr>`;
      });
    } else {
      emailBody += `<tr><td colspan="6">無法取得市場指數資料</td></tr>`; // 如果 indexData 不是陣列，顯示錯誤訊息
    }

    emailBody += `</tbody></table>`;


    //  加入 CNN 恐懼與貪婪指數
    emailBody += `<h2>CNN 恐懼與貪婪指數 (Fear & Greed Index)</h2>`;
    emailBody += `<p class="index-value">${fearGreedIndexValue} <span class="index-category">(${fearGreedIndexCategory})</span></p>`;
    emailBody += `<p>資料來源：Google 試算表 (市場指數), CNN Money (Fear & Greed Index)</p>`; // 更新資料來源說明
    emailBody += `</body></html>`;


    // 4. 寄送電子郵件
    MailApp.sendEmail({
      to: recipientEmail,
      subject: emailSubject,
      htmlBody: emailBody
    });

    Logger.log("每日市場指數報告郵件 (表格化) 已成功寄出至 " + recipientEmail);


  } catch (error) {
    let errorBody = `<!DOCTYPE html><html><head><style>body { font-family: Arial, sans-serif; color: red; }</style></head><body>`;
    errorBody += `<h1>每日市場指數報告 - 錯誤通知 (表格化)</h1>`; // 修改錯誤通知郵件標題
    errorBody += `<p>程式執行時發生錯誤，無法取得市場指數資料或寄送郵件。</p>`;
    errorBody += `<p>錯誤訊息：</p><pre style="background-color:#f0f0f0; padding:10px; border: 1px solid #ccc;">${error}</pre>`; // 將錯誤訊息包含在郵件中
    errorBody += `<p>請檢查 Apps Script 執行記錄以獲取更詳細的錯誤資訊。</p>`;
    errorBody += `</body></html>`;


    MailApp.sendEmail({
      to: recipientEmail,
      subject: "每日市場指數報告 - 錯誤通知 (表格化)", // 修改錯誤通知郵件主旨
      htmlBody: errorBody
    });

    Logger.log("錯誤報告郵件 (表格化) 已寄出至 " + recipientEmail + ". 原始錯誤訊息: " + error);
    Logger.log("完整錯誤堆疊追蹤: " + error.stack); // 記錄更詳細的錯誤堆疊追蹤資訊
  }
}


/**
 * 從 Google 試算表讀取市場指數資料 (通用版本).
 * @returns {Array<object>|string}  包含市場指數資料的物件陣列.
 *                                  每個物件代表試算表中的一行資料 (一個市場指數),
 *                                  包含 'name', 'price', 'changePercent', 'previousClose', 'fiftyTwoWeekHigh', 'fiftyTwoWeekLow' 屬性.
 *                                  如果讀取失敗，則回傳錯誤訊息字串 "N/A - 試算表讀取錯誤" 或 "N/A - 試算表 ID 未設定".
 */
function getIndex() {
  try {
    // **請將 "您的試算表ID" 替換成您的 Google 試算表 ID**
    const spreadsheetId = "1j7wtyywUQ2S76r7ryQNcucF9PgPiwYn-iGpnEBKeu9A";         
    
    const dataRange = "A2:F";   //  資料範圍 (假設資料從 A2 開始到 F 欄，包含多行)

    // ** 檢查是否已正確設定試算表 ID，若未設定則直接回傳錯誤 **
    if (spreadsheetId === "您的試算表ID") {
      Logger.log("錯誤：請在 getIndex() 函式中設定您的 Google 試算表 ID。");
      return "N/A - 試算表 ID 未設定"; //  回傳特殊錯誤訊息，方便辨識問題
    }

    //  使用 Sheets API 取得指定範圍的資料 (包含標題列)
    const values = Sheets.Spreadsheets.Values.get(spreadsheetId, dataRange).values;

    if (!values || values.length === 0) {
      Logger.log("錯誤：在 Google 試算表中找不到任何市場指數資料，範圍: " + dataRange + ", 試算表 ID: " + spreadsheetId + ". API 回應: " + JSON.stringify(values));
      return "N/A - 試算表讀取錯誤"; // 如果在試算表中找不到資料，回傳錯誤訊息
    }

    //  假設試算表欄位順序為: 名稱, 價格/指數, 和前一日變動百分比, 前一天的收盤價, 52週期間的最高價, 52週期間的最低價
    const indexData = values.map(row => {
      return {
        name: row[0] || "N/A",             // 名稱 (A 欄)
        price: row[1] !== undefined ? parseFloat(row[1]).toFixed(2) : "N/A",     // 價格/指數 (B 欄), 轉換為數字並取小數點後兩位
        changePercent: row[2] || "N/A",    // 和前一日變動百分比 (C 欄)
        previousClose: row[3] || "N/A",    // 前一天的收盤價 (D 欄)
        fiftyTwoWeekHigh: row[4] || "N/A", // 52週期間的最高價 (E 欄)
        fiftyTwoWeekLow: row[5] || "N/A"   // 52週期間的最低價 (F 欄)
      };
    });

    return indexData; //  回傳包含市場指數資料的物件陣列

  } catch (e) {
    Logger.log("Error getting market index data from Google Sheet: " + e);
    return "N/A - 試算表讀取錯誤"; // 發生錯誤時回傳錯誤訊息
  }
}

/**
 * 取得 CNN 恐懼與貪婪指數 (從 CNN 資料 API 取得，模擬 curl 命令).
 * @returns {object} 包含 indexValue (指數值) 和 indexCategory (指數分類) 的物件. 如果獲取失敗，則回傳包含 "N/A" 值的物件.
 */
function getFearGreedIndex() {
  try {
    const url = "https://production.dataviz.cnn.io/index/fearandgreed/graphdata"; // CNN 恐懼與貪婪指數資料 API 端點

    //  模擬 curl 命令的 header 設定
    const headers = {
      'accept': '*/*',
      'accept-language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
      'dnt': '1',
      'if-none-match': 'W/-5407781859672843994',
      'origin': 'https://edition.cnn.com',
      'priority': 'u=1, i',
      'referer': 'https://edition.cnn.com/',
      'sec-ch-ua': '"Not;A=Brand";v="24", "Chromium";v="128"',
      'sec-ch-ua-mobile': '?0',
      'sec-ch-ua-platform': '"macOS"',
      'sec-fetch-dest': 'empty',
      'sec-fetch-mode': 'cors',
      'sec-fetch-site': 'cross-site',
      'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    };

    const options = {
      'headers': headers,
      'muteHttpExceptions': true //  設定為 true 以避免 HTTP 錯誤導致程式碼停止，方便錯誤處理
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode(); // 取得 HTTP 狀態碼

    if (responseCode !== 200) { // 檢查 HTTP 狀態碼是否為 200 (成功)
      Logger.log("HTTP Error: 取得 CNN 恐懼與貪婪指數 API 失敗，狀態碼: " + responseCode);
      return { indexValue: "N/A", indexCategory: "N/A" }; //  如果 API 回應非 200，回傳 "N/A"
    }


    const json = JSON.parse(response.getContentText());

    //  從 JSON 回應中取得 fear_and_greed.score 和 fear_and_greed.rating
    const fearGreedData = json.fear_and_greed;
    if (fearGreedData && fearGreedData.score !== undefined && fearGreedData.rating) { // 檢查 fearGreedData, score 和 rating 是否存在
      const indexValue = parseFloat(fearGreedData.score).toFixed(2); // 取得 score 並格式化到小數點後兩位
      const indexCategory = fearGreedData.rating; // 取得 rating (指數分類)

      return {
        indexValue: indexValue,
        indexCategory: indexCategory
      };
    } else {
      Logger.log("Error: Fear & Greed index data (score or rating) not found in CNN API response: " + JSON.stringify(json)); // 記錄詳細的 API 回應內容，方便debug
      return { indexValue: "N/A", indexCategory: "N/A" }; //  如果 JSON 中找不到 score 或 rating，回傳 "N/A"
    }

  } catch (e) {
    Logger.log("Error getting Fear & Greed index from CNN API: " + e);
    return {
      indexValue: "N/A",
      indexCategory: "N/A"
    }; // 發生錯誤時回傳 "N/A"
  }
}
