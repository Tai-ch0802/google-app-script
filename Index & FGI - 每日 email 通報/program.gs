/**
 * 每日獲取市場指數資料 (從 Google 試算表), CNN 恐懼與貪婪指數，並寄送郵件報告 (GEEK 版).
 */
function sendDailyMarketIndexReport() {
  // ** 設定區 **
  const recipientEmail = "";  //  請將 "您的 Gmail 信箱地址"  替換成您的 Gmail 電子郵件地址
  const emailSubject = "每日市場指數報告 (表格化)"; // 修改郵件主旨
  const timeZone = "Asia/Taipei"; // 時區設定為台北，您可以根據您的所在地調整時區
  const ALERT_THRESHOLD = 2.0; // 設定警報門檻 (百分比)

  try {
    const indexData = getIndex(); 
    if (typeof indexData === 'string' && indexData.startsWith("N/A -")) throw new Error(indexData);
    
    const fearGreedData = getFearGreedIndex(); 
    const fgiValue = parseFloat(fearGreedData.indexValue);

    let isAlert = false;
    let alertMessages = [];

    // 1. 偵測 FGI 警報
    if (fgiValue <= 25) {
      alertMessages.push("EXTREME FEAR");
      isAlert = true;
    } else if (fgiValue >= 75) {
      alertMessages.push("MARKET OVERHEAT");
      isAlert = true;
    }

    // 2. 處理市場數據：整合計算水位與波動監控
    if (Array.isArray(indexData)) {
      indexData.forEach(item => {
        // --- A. 計算水位百分比 ---
        const current = parseFloat(item.price);
        const high = parseFloat(item.fiftyTwoWeekHigh);
        const low = parseFloat(item.fiftyTwoWeekLow);
        
        if (!isNaN(current) && !isNaN(high) && !isNaN(low) && high !== low) {
          let position = ((current - low) / (high - low) * 100);
          item.rangePosition = Math.max(0, Math.min(100, position)).toFixed(0);
        } else {
          item.rangePosition = null; 
        }

        // --- B. 偵測波動警報 (修復處) ---
        // 確保排除百分比符號後轉換為數字進行比較
        const changeVal = parseFloat(item.changePercent.toString().replace('%', ''));
        if (!isNaN(changeVal) && Math.abs(changeVal) >= ALERT_THRESHOLD) {
          isAlert = true;
          item.isHighVolatility = true; // 標記此行需要警示顏色
          alertMessages.push(`${item.name} ${item.changePercent}`);
        }
      });
    }

    // 3. 動態建構主旨
    let alertPrefix = isAlert ? `[ALERT: ${alertMessages.join(' | ')}]` : "[INFO]";
    const emailSubject = `${alertPrefix} 每日市場指數報告`;

    // 4. 建立 HTML 模板
    const template = HtmlService.createTemplateFromFile('emailTemplate');
    template.indexData = indexData;
    template.fearGreedValue = fearGreedData.indexValue;
    template.fearGreedCategory = fearGreedData.indexCategory;
    template.isAlert = isAlert; 
    template.timestamp = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss");

    const htmlBody = template.evaluate().getContent();

    // 5. 寄送郵件
    GmailApp.sendEmail(recipientEmail, emailSubject, '', {
      from: 'tai@taiwan-no1.net',
      htmlBody: htmlBody,
      name: 'Terminal-Market-Bot'
    });
    Logger.log("多重警報監控郵件已成功寄出！");

  } catch (error) {
    sendErrorEmail(recipientEmail, error); // 封裝錯誤郵件邏輯
  }
}

/** 錯誤通知函式 */
function sendErrorEmail(recipient, error) {
  let errorBody = `<!DOCTYPE html><html><head><style>body { font-family: monospace; color: #f85149; background: #0d1117; padding: 20px; }</style></head><body>`;
  errorBody += `<h2>[SYSTEM ERROR] Market Report Failure</h2>`;
  errorBody += `<p>執行發生異常，請檢查日誌：</p><pre style="background:#161b22; padding:15px; border:1px solid #30363d;">${error}</pre></body></html>`;
  
  MailApp.sendEmail({
    to: recipient,
    subject: "[CRITICAL] Market Report Error",
    htmlBody: errorBody
  });
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
