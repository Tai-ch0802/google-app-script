function summarizeDailyEmailPDF() {
  // 搜索今天的邮件 (假设邮件主题包含 "Daily Report")
  var threads = GmailApp.search('subject:"華爾街日報" after:' + getYesterdayDate());
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var attachments = message.getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        if (attachment.getContentType() == 'application/pdf') {
          // **关键修改: 准备 PDF 文件 для Gemini API 文件解读 API**
          var pdfBlob = attachment.copyBlob();
          var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes()); // Base64 编码 PDF 内容

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

          try {
            var response = UrlFetchApp.fetch(geminiFileApiEndpoint, options);
            var jsonResponse = JSON.parse(response.getContentText());
            var analysisResult = jsonResponse.candidates[0].content.parts[0].text;

            Logger.log('郵件主題: ' + message.getSubject());
            Logger.log('PDF 分析結果: ' + analysisResult);

            // **---  新增：擷取 HTML 內容  ---**
            var htmlStartIndex = analysisResult.indexOf("```html") + 7; // 找到 ```html 的起始位置，並往後移動 7 個字元 (跳過 ```html)
            var htmlEndIndex = analysisResult.indexOf("```", htmlStartIndex); // 找到 ``` 的結束位置，從 htmlStartIndex 後開始找
            var htmlContent = analysisResult.substring(htmlStartIndex, htmlEndIndex).trim(); // 擷取 ```html 和 ``` 之間的內容，並去除前後空白

            Logger.log('擷取後的 HTML 內容: ' + htmlContent); // 記錄擷取後的 HTML 內容

            // **---  新增：发送邮件  ---**
            var userEmail = Session.getActiveUser().getEmail(); // 获取当前用户邮箱
            var emailSubject = '華爾街日報 - 分析结果: ' + message.getSubject(); // 邮件主题，可以自定义
            var emailBody = '您好，\n\n這是您每日郵件 "' + message.getSubject() + '"  的 PDF 附件分析結果，以 HTML 格式呈現：\n\n (如果郵件顯示異常，請嘗試以網頁瀏覽器開啟附件 html 檔案)\n\n--- 原始郵件資訊 ---\n發件人: ' + message.getFrom() + '\n日期: ' + message.getDate(); //  修改郵件正文提示
            var emailHtmlBody = htmlContent; //  設定 emailHtmlBody 為擷取出的 HTML 內容


            GmailApp.sendEmail(userEmail, emailSubject, emailBody, {
              htmlBody: emailHtmlBody, //  設定 htmlBody 參數為 HTML 內容
              attachments: [attachment.copyBlob()], // 夹带原始 PDF 附件
              name: '華爾街日報 - 分析機器人' //  自定义发件人名称
            });

            Logger.log('郵件已發送至: ' + userEmail);

          } catch (error) {
            Logger.log('Gemini API 文件解读 API 调用失败: ' + error);
          }
        }
      }
    }
  }
}

function getYesterdayDate() {
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  return Utilities.formatDate(yesterday, Session.getTimeZone(), 'yyyy/MM/dd');
}
