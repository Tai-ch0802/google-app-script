function summarizeDailyEmailPDF() {

  const geminiApiKey = getScriptProperties('GEMINI_API_KEY');
  const subject = '華爾街日報重點摘要 - ' + getTodayDate();

  // **---  新增：发送邮件  ---**
  var userEmail = Session.getActiveUser().getEmail(); // 获取当前用户邮箱
  var bccRecipients = [];

  // **---  修改：直接從 URL 取得 PDF 檔案內容  ---**
  var emailsource = getLatestEmailSourceFromGmail();
  var todayPaperPdfUrl = getTodaysEditionURL(emailsource);
  // var todayPaperPdfUrl = 'https://customercenter.wsj.com/todaysPaper/'; //  直接使用 PDF URL
  var fetchOptions = {
    'headers': {
      'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
      'accept-language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
      'cache-control': 'max-age=0',
      'dnt': '1',
      // 'if-modified-since': 'Mon, 10 Mar 2025 06:03:25 GMT',
      // 'if-none-match': 'W/"f88e7b2853b463bb47d40c47a05f9e03"',
      'priority': 'u=0, i',
      'sec-ch-ua': '"Not;A=Brand";v="24", "Chromium";v="128"',
      'sec-ch-ua-mobile': '?0',
      'sec-ch-ua-platform': '"macOS"',
      'sec-fetch-dest': 'document',
      'sec-fetch-mode': 'navigate',
      'sec-fetch-site': 'none',
      'sec-fetch-user': '?1',
      'upgrade-insecure-requests': '1',
      'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
    },
    'muteHttpExceptions': true
  };

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

  try {
    // **---  新增：Upload PDF to File API  ---**
    var geminiFileUploadEndpoint = 'https://generativelanguage.googleapis.com/upload/v1beta/files?key=' + geminiApiKey; // **File API Upload Endpoint**

    var numBytes = pdfData.length; //  取得 PDF 檔案大小
    var displayName = filename; //  使用檔案名稱作為 Display Name

    var resumableUploadStartOptions = {
      'method': 'post',
      'headers': {
        "X-Goog-Upload-Protocol": "resumable",
        "X-Goog-Upload-Command": "start",
        "X-Goog-Upload-Header-Content-Length": String(numBytes), //  檔案大小，需要轉換為字串
        "X-Goog-Upload-Header-Content-Type": pdfContentType,
        "Content-Type": "application/json"
      },
      'payload': JSON.stringify({'file': {'display_name': displayName}}), //  Metadata: display_name
      'muteHttpExceptions': true
    };

    var resumableUploadStartResponse = UrlFetchApp.fetch(geminiFileUploadEndpoint, resumableUploadStartOptions);
    if (resumableUploadStartResponse.getResponseCode() !== 200) {
      Logger.log('Resumable Upload Start 失敗，狀態碼: ' + resumableUploadStartResponse.getResponseCode() + ', Response: ' + resumableUploadStartResponse.getContentText());
      return;
    }
    var uploadUrl = resumableUploadStartResponse.getHeaders()['x-goog-upload-url']; // **從 Header 中取得 upload_url**
    Logger.log('Resumable Upload Start 成功，Upload URL: ' + uploadUrl);


    // **---  步驟 2: Upload File Bytes (Upload and Finalize) ---**
    var resumableUploadBytesEndpoint = uploadUrl; //  使用步驟 1 取得的 uploadUrl

    var resumableUploadBytesOptions = {
      'method': 'post', //  仍然是 POST 方法
      'headers': {
        "X-Goog-Upload-Header-Content-Length": String(numBytes), // 檔案大小，需要轉換為字串
        "X-Goog-Upload-Offset": "0",
        "X-Goog-Upload-Command": "upload, finalize"
      },
      'payload': pdfBlob.getBytes(), //  直接傳送 PDF 二進制資料
      'muteHttpExceptions': true
    };

    var resumableUploadBytesResponse = UrlFetchApp.fetch(resumableUploadBytesEndpoint, resumableUploadBytesOptions);
    if (resumableUploadBytesResponse.getResponseCode() !== 200) {
      Logger.log('Resumable Upload Bytes 失敗，狀態碼: ' + resumableUploadBytesResponse.getResponseCode() + ', Response: ' + resumableUploadBytesResponse.getContentText());
      return;
    }
    var fileInfoJsonResponse = JSON.parse(resumableUploadBytesResponse.getContentText());
    var fileResource = fileInfoJsonResponse.file.uri; // **從 Response Body 中取得 file_uri**
    Logger.log('Resumable Upload Bytes 成功，File Resource: ' + fileResource);


    // **步驟 3: 调用 Gemini API 文件解读 API (需要替换为真实的文件解读 API 端点)**
    var geminiFileApiEndpoint = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + geminiApiKey; //  **<--  需要确认正确的文件解读 API 端点**
    var payload = {
      systemInstruction: {
        parts: [{
          text: "您是一位在華爾街擁有三十年經驗的資深投信經理人。您的任務是解讀華爾街日報的摘要內容，並分析其對整體股市可能造成的影響。您的分析應保持完全中立的立場，不帶有個人情緒或偏見。請專注於新聞事件的客觀事實，並基於您多年的市場經驗，評估這些事件可能如何影響投資人情緒、資金流向以及整體市場的風險偏好。您的目標是提供一份簡潔扼要、觀點清晰的分析報告，幫助使用者快速掌握新聞摘要的重點，並了解其潛在的市場意涵。請避免使用過於專業或晦澀的金融術語，力求分析結果通俗易懂，讓即使是不具備深厚金融背景的讀者也能理解。在分析過程中，請特別留意以下幾個方面：\n\n1.  **新聞事件的核心內容：** 準確把握新聞摘要所傳達的最主要訊息，例如：政策變動、經濟數據發布、企業財報、地緣政治風險、重大技術突破等。\n2.  **市場情緒的潛在影響：**  評估新聞事件可能如何影響投資者的情緒，例如：是會增強市場的樂觀情緒，還是引發避險情緒？是會刺激投資人積極承擔風險，還是促使他們轉向更安全的資產？\n3.  **資金流動的可能變化：**  分析新聞事件是否可能導致資金在不同資產類別之間流動，例如：是否會促使資金從債市流向股市，或從成長股轉向價值股？是否會引發資金從新興市場流向成熟市場，或反之？\n4.  **主要產業 sector 的連動效應：**  判斷新聞事件可能對哪些產業 sector 產生直接或間接的影響？例如：能源價格上漲可能利好能源產業，但對航空業和運輸業可能不利；科技創新可能推動科技股上漲，但對傳統產業可能構成競爭壓力。\n5.  **總體經濟的潛在風險：**  評估新聞事件是否可能對總體經濟環境構成風險，例如：通膨壓力上升、經濟成長放緩、利率變動、貨幣政策調整等。\n\n請將您的分析結果整理成結構清晰、重點突出的報告，並在報告的開頭明確指出新聞摘要的來源（華爾街日報），以及新聞事件發生的日期。在報告的結尾，您可以簡要總結您的觀點，並提供一些投資建議（請注意，這些建議僅供參考，不構成任何形式的投資保證）。最重要的是，請始終保持專業、中立和客觀的態度，以一位資深投信經理人的視角，為使用者解讀華爾街日報的新聞摘要，並洞察其背後的市場意涵。"
        }]
      },
      contents: [{
        // parts: [{
        //   inlineData: {
        //     mimeType: "application/pdf", // 指定 MIME 类型为 PDF
        //     data: pdfBase64 // Base64 编码的 PDF 数据
        //   },
        // },
        parts: [{
          fileData: { // **Use fileData instead of inlineData**
            mimeType: pdfContentType, // Specify MIME type
            fileUri: fileResource // **Reference the file resource identifier**
          },
        },
        {
          text: "請分析這份華爾街日報 PDF，並總結其要點，以 html 的格式做呈現並回傳，style 標籤內容請用這組 body { font-family: Arial, sans-serif; } h1 { color: navy; } h2 { color: darkgreen; } h3 { color: darkred; } ul { list-style-type: square; } .section { margin-bottom: 20px; } .subsection { margin-left: 20px; }。用 \"```html\" 和 \"```\" 前後包住 html 內容。報告文字請用繁體中文，並且回傳給我的文案只能出現一次\"```html\" 和 \"```\"。"
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

    

    var emailSubject = subject; // 邮件主题，可以自定义
    var emailBody = '您好，\n\n這是您每日郵件 "' + subject + '"  的 PDF 附件分析結果，以 HTML 格式呈現：\n\n (如果郵件顯示異常，請嘗試以網頁瀏覽器開啟附件 html 檔案)\n\n--- 原始郵件資訊 ---\n發件人: tai@taiwan-no1.net \n日期: ' + getTodayDate(); //  修改郵件正文提示
    var emailHtmlBody = htmlContent; //  設定 emailHtmlBody 為擷取出的 HTML 內容


    GmailApp.sendEmail(userEmail, emailSubject, emailBody, {
      bcc: bccRecipients.join(', '), //  **新增 bcc 參數，設定為收件人 email 地址字串 (逗號分隔)**
      htmlBody: emailHtmlBody, //  設定 htmlBody 參數為 HTML 內容
      attachments: [namedPdfBlob], // 夹带原始 PDF 附件
      name: '華爾街日報 - 分析機器人' //  自定义发件人名称
    });

    Logger.log('HTML 格式郵件已 BCC 發送至: ' + bccRecipients.join(', ') + ', 附件檔名: ' + filename); //  記錄發送郵件和收件人列表

  } catch (error) {
    Logger.log('Gemini API 文件解读 API 调用失败: ' + error);
    var emailSubject = 'gemini 解析失敗，只有附件內容（待修正中） - '　+ subject; // 邮件主题，可以自定义
    var emailBody = '您好，\n\n　gemini 解析失敗，此問題待修正當中\n\n--- 原始郵件資訊 ---\n發件人: tai@taiwan-no1.net \n日期: ' + getTodayDate();

    GmailApp.sendEmail(userEmail, emailSubject, emailBody, {
      bcc: bccRecipients.join(', '), //  **新增 bcc 參數，設定為收件人 email 地址字串 (逗號分隔)**
      attachments: [namedPdfBlob], // 夹带原始 PDF 附件
      name: '華爾街日報 - 分析機器人' //  自定义发件人名称
    });
    Logger.log('HTML 格式郵件已 BCC 發送至: ' + bccRecipients.join(', ') + ', 附件檔名: ' + filename); //  記錄發送郵件和收件人列表
  }
}

/**
 * 從郵件原始碼中提取 "READ TODAY'S EDITION" 按鈕的連結網址。
 *
 * @param {string} emailSource 郵件的原始碼 (字串格式)。
 * @return {string} 按鈕連結網址，如果找不到則返回空字串。
 */
function getTodaysEditionURL(emailSource) {
  // 尋找 "Read Today's Edition" 按鈕文字在郵件原始碼中的位置
  var buttonText = "Read Today's Edition [";
  var buttonIndex = emailSource.indexOf(buttonText);

  // 如果找不到按鈕文字，則返回空字串
  if (buttonIndex === -1) {
    return "";
  }

  // 從按鈕文字位置開始，尋找連結網址的起始和結束位置
  var urlStartIndex = buttonIndex + buttonText.length;
  var urlEndIndex = emailSource.indexOf("]", urlStartIndex);

  // 如果找不到網址結束符號，則返回空字串
  if (urlEndIndex === -1) {
    return "";
  }

  // 提取連結網址
  var url = emailSource.substring(urlStartIndex, urlEndIndex);

  // 移除網址中可能存在的換行符號 (包括 \n, \r, \r\n)
  // 移除網址結尾可能存在的 '=' 符號 (處理 Quoted-printable 編碼)
  url = url.replace(/[\r\n=]+/g, ''); // 使用正則表達式移除所有換行符號

  // 移除網址前後可能存在的空白字元
  url = url.trim();

  // 移除網址結尾可能存在的 '=' 符號 (處理 Quoted-printable 編碼)
  // url = url.replace(/=+$/, ''); // 使用正則表達式移除結尾的所有 '=' 符號

  Logger.log("下載網址：" + url);

  // 返回提取到的網址 (前後 trim() 處理)
  return url;
}

/**
 * 從 Gmail 取得最新一封標籤為 '財經/華爾街日報/每日報導' 的信件原始碼。
 *
 * @return {string} 最新郵件的原始碼 (字串格式)，如果找不到郵件則返回空字串。
 */
function getLatestEmailSourceFromGmail() {
  var labelName = '財經/華爾街日報/每日報導';
  var searchQuery = 'label:' + labelName;
  var threads = GmailApp.search(searchQuery, 0, 1); // 搜尋最新一封符合標籤的信件 (最多 1 封)

  if (threads.length > 0) {
    var latestThread = threads[0];
    var messages = latestThread.getMessages();
    if (messages.length > 0) {
      var latestMessage = messages[messages.length - 1]; // 假設最新郵件是最後一封
      var emailSource = latestMessage.getRawContent();
      return emailSource;
    } else {
      Logger.log("在標籤 '" + labelName + "' 的郵件串中找不到郵件。");
      return ""; // 郵件串中沒有郵件
    }
  } else {
    Logger.log("找不到標籤為 '" + labelName + "' 的郵件串。");
    return ""; // 找不到符合標籤的郵件串
  }
}

function getScriptProperties(key) {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(key);
}

function getTodayDate() {
  var today = new Date();
  return Utilities.formatDate(today, Session.getTimeZone(), 'yyyy-MM-dd');
}
