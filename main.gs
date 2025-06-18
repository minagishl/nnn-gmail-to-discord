function saveGmailToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Retrieve existing Message-IDs (assuming they are recorded in column 3)
  var existingIds = sheet.getRange(1, 3, lastRow).getValues().flat();

  var threads = GmailApp.getInboxThreads(0, 10);
  var newEntries = [];

  for (var i = 0; i < threads.length; i++) {
    var message = threads[i].getMessages()[0];
    var date = message.getDate();
    var subject = message.getSubject();
    var from = message.getFrom();
    var plainBody = message.getPlainBody();
    var messageId = message.getId();

    // Check forwarding conditions
    if (shouldForwardMessage(from, plainBody)) {
      if (!existingIds.includes(messageId)) {
        newEntries.push([date, subject, messageId]); // Add messageId
        sendEmailToDiscord(message);
      }
    }
  }

  if (newEntries.length > 0) {
    sheet.getRange(lastRow + 1, 1, newEntries.length, 3).setValues(newEntries); // Save to 3 columns
  } else {
    Logger.log("No new emails matching the conditions.");
  }
}

function shouldForwardMessage(from, body) {
  const domainPattern = /@.*\.(nnn\.ed\.jp|nnn\.ac\.jp)/i;
  const keyword = "学校法人角川ドワンゴ学園";

  return domainPattern.test(from) || body.includes(keyword);
}

function sendEmailToDiscord(message) {
  // Retrieve Webhook URL from script properties
  var discordWebhookUrl = PropertiesService.getScriptProperties().getProperty("DISCORD_WEBHOOK_URL");
  if (!discordWebhookUrl) {
    Logger.log("Webhook URL is not set.");
    return;
  }

  const from = message.getFrom();
  const subject = message.getSubject();
  const plainBody = message.getPlainBody();

  const payload = {
    content: subject,
    embeds: [{
      title: subject,
      author: {
        name: from,
      },
      description: plainBody.substr(0, 2048),
    }],
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(discordWebhookUrl, options);
}
