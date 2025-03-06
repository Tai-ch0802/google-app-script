function summarizeDailyEmailPDF() {

  const subject = '華爾街日報重點摘要 - ' + getTodayDate();

  // **---  修改：直接從 URL 取得 PDF 檔案內容  ---**
  var todayPaperPdfUrl = 'https://customercenter.wsj.com/todaysPaper/'; //  直接使用 PDF URL
  var fetchOptions = {
    'headers': { // 模擬 curl 命令中的 headers (與之前相同)
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:135.0) Gecko/20100101 Firefox/135.0',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.5',
      'Accept-Encoding': 'gzip, deflate, br, zstd',
      'Connection': 'keep-alive',
      'Upgrade-Insecure-Requests': '1',
      'Sec-Fetch-Dest': 'document',
      'Sec-Fetch-Mode': 'navigate',
      'Sec-Fetch-Site': 'none',
      'Sec-Fetch-User': '?1',
      'Priority': 'u=0, i',
      'TE': 'trailers'
    },
    'muteHttpExceptions': true //  設定為 true，即使 HTTP 錯誤也不會拋出例外，方便處理錯誤情況
  };

  try {
    var pdfResponse = UrlFetchApp.fetch(todayPaperPdfUrl, fetchOptions); //  直接使用 PDF URL 下載
    if (pdfResponse.getResponseCode() !== 200) {
      Logger.log('下載 PDF 檔案失敗，狀態碼: ' + pdfResponse.getResponseCode());
      return; //  下載 PDF 失敗，結束腳本
    }
    var pdfBlob = pdfResponse.getBlob();
    Logger.log('成功下載 PDF 檔案');

    // **---  新增：指定 pdfBlob 的檔案名稱  ---**
    var pdfData = pdfBlob.getBytes(); // 取得原始 pdfBlob 的二進制資料
    var pdfContentType = pdfBlob.getContentType(); // 取得原始 pdfBlob 的 Content-Type
    var filename = 'WSJ_Daily_Paper_' + getTodayDate() + '.pdf'; 
    var namedPdfBlob = Utilities.newBlob(pdfData, pdfContentType, filename); // 使用 Utilities.newBlob() 建立帶有檔名的 Blob


    var pdfBase64 = Utilities.base64Encode(pdfData); // Base64 编码 PDF 内容
  
    // **调用 Gemini API 文件解读 API (需要替换为真实的文件解读 API 端点)**
    var geminiApiKey = '';
    var geminiFileApiEndpoint = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + geminiApiKey; //  **<--  需要确认正确的文件解读 API 端点**
    var payload = {
      contents: [{
        parts: [{
          inlineData: {
            mimeType: "application/pdf", // 指定 MIME 类型为 PDF
            data: pdfBase64 // Base64 编码的 PDF 数据
          },
        },
        {
          text: "請分析這份華爾街日報 PDF，並總結其要點，以 html 的格式做呈現並回傳，用 \"```html\" 和 \"```\" 前後包住 html 內容。報告文字請用繁體中文，並且回傳給我的文案只能出現一次\"```html\" 和 \"```\"。"
        }]
      }]
    };
    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    };
    
      var response = UrlFetchApp.fetch(geminiFileApiEndpoint, options);
      var jsonResponse = JSON.parse(response.getContentText());
      var analysisResult = jsonResponse.candidates[0].content.parts[0].text;

      Logger.log('郵件主題: ' + subject);
      Logger.log('PDF 分析結果: ' + analysisResult);

      // **---  新增：擷取 HTML 內容  ---**
      var htmlStartIndex = analysisResult.indexOf("```html") + 7; // 找到 ```html 的起始位置，並往後移動 7 個字元 (跳過 ```html)
      var htmlEndIndex = analysisResult.indexOf("```", htmlStartIndex); // 找到 ``` 的結束位置，從 htmlStartIndex 後開始找
      var htmlContent = analysisResult.substring(htmlStartIndex, htmlEndIndex).trim(); // 擷取 ```html 和 ``` 之間的內容，並去除前後空白

      //Logger.log('擷取後的 HTML 內容: ' + htmlContent); // 記錄擷取後的 HTML 內容

      // **---  新增：发送邮件  ---**
      var userEmail = Session.getActiveUser().getEmail(); // 获取当前用户邮箱
      var emailSubject = subject; // 邮件主题，可以自定义
      var emailBody = '您好，\n\n這是您每日郵件 "' + subject + '"  的 PDF 附件分析結果，以 HTML 格式呈現：\n\n (如果郵件顯示異常，請嘗試以網頁瀏覽器開啟附件 html 檔案)\n\n--- 原始郵件資訊 ---\n發件人: tai@taiwan-no1.net \n日期: ' + getTodayDate(); //  修改郵件正文提示
      var emailHtmlBody = htmlContent; //  設定 emailHtmlBody 為擷取出的 HTML 內容


      GmailApp.sendEmail(userEmail, emailSubject, emailBody, {
        htmlBody: emailHtmlBody, //  設定 htmlBody 參數為 HTML 內容
        attachments: [namedPdfBlob], // 夹带原始 PDF 附件
        name: '華爾街日報 - 分析機器人' //  自定义发件人名称
      });

      Logger.log('郵件已發送至: ' + userEmail);

    } catch (error) {
      Logger.log('Gemini API 文件解读 API 调用失败: ' + error);
    }
}

function getTodayDate() {
  var today = new Date();
  return Utilities.formatDate(today, Session.getTimeZone(), 'yyyy-MM-dd');
}
