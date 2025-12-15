// ======================================================
// 🧩 CONFIG
// ======================================================
const SHEET_ID = "1Rj9Hr76N85_wP7f5PL6AWYzRqskiLc1iH6u3gVUJzCQ";
const SHEET_NAME = "Transactions";
const FOLDER_ID = "1kECv-ukZ6B6ywD59RapTZyfuGUVRlT_3";

const LINE_CHANNEL_ACCESS_TOKEN = "FjHk7yf4+S99f0s7dMc0ule5mXe2jAZpYDShYmpUc83YumxcQr6QD6ljyJHCdYucYRAHWGshdZ3s+4lSdwKClgBy7qP2LIUkvOAYlYJz72qAIwSBLxAMrzQtTxkbV+SzsxDfBn57Zq7vThiDh9TrNgdB04t89/1O/w1cDnyilFU=";
const LINE_GROUP_ID = "Cecaac2607f7a8b9f51a53dfcad4a5040";

// ======================================================
// 📤 ส่ง Flex Message
// ======================================================
function sendFlexToGroup(name, amount, desc, type, imageUrl, balance) {
  const color = (type === "เงินเข้า") ? "#16a34a" : "#dc2626";
  const symbol = (type === "เงินเข้า") ? "💰" : "💸";

  const flex = {
    type: "flex",
    altText: `${symbol} ${type} ${amount.toLocaleString()} บาท`,
    contents: {
      type: "bubble",
      hero: imageUrl ? {
        type: "image",
        url: imageUrl,
        size: "full",
        aspectRatio: "20:13",
        aspectMode: "cover"
      } : undefined,
      body: {
        type: "box",
        layout: "vertical",
        contents: [
          { type: "text", text: `${symbol} รายการ${type}`, weight: "bold", size: "lg", color: color },
          { type: "separator", margin: "sm" },
          { type: "text", text: `👤 ${name}`, size: "sm" },
          { type: "text", text: `💵 ${amount.toLocaleString()} บาท`, weight: "bold", color: color },
          { type: "text", text: `📝 ${desc}`, size: "sm", wrap: true },
          { type: "separator", margin: "sm" },
          { type: "text", text: `💹 ยอดคงเหลือ: ${balance.toLocaleString()} บาท`, color: "#2563eb", weight: "bold" }
        ]
      },
      footer: imageUrl ? {
        type: "box",
        layout: "vertical",
        contents: [{
          type: "button",
          style: "link",
          color: "#4f46e5",
          action: { type: "uri", label: "📸 ดูภาพหลักฐาน", uri: imageUrl }
        }]
      } : undefined
    }
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({ to: LINE_GROUP_ID, messages: [flex] })
  });
}


// ======================================================
// ✅ POST: เพิ่มข้อมูลเงินเข้า/ออก แล้วส่ง Flex เข้า LINE
// ======================================================
function doPost(e) {
  try {
    if (e.postData && e.postData.type === "application/json") {
      return handleLineWebhook(e); // สำหรับ groupid?
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const action = e.parameter.action;
    const name = e.parameter.personName;
    const amount = parseFloat(e.parameter.amount || 0);
    const desc = e.parameter.description || "-";
    const type = (action === "addIncome") ? "เงินเข้า" : "เงินออก";
    const now = new Date();
    let imageUrl = "";

    // ✅ อัปโหลดรูปถ้ามี
    if (e.files && e.files.image) {
      const blob = e.files.image;
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = file.getUrl();
    }

    // ✅ เพิ่มข้อมูลในชีต
    sheet.appendRow([
      Utilities.formatDate(now, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss"),
      name,
      type,
      amount,
      desc,
      imageUrl
    ]);

    // ✅ คำนวณยอดคงเหลือปัจจุบัน
    const data = sheet.getDataRange().getValues();
    data.shift();
    let income = 0, expense = 0;
    data.forEach(r => {
      if (r[2] === "เงินเข้า") income += Number(r[3]);
      if (r[2] === "เงินออก") expense += Number(r[3]);
    });
    const balance = income - expense;

    // ✅ ส่ง Flex Message
    sendFlexToGroup("รายการใหม่", name, amount, desc, type, imageUrl, balance);

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ======================================================
// 📥 GET: ดึงข้อมูลสรุป & ประวัติ
// ======================================================
function doGet(e) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  rows.shift();
  const action = e.parameter.action;

  if (action === "getSummary") {
    let income = 0, expense = 0, personTotals = {};
    rows.forEach(r => {
      const name = r[1];
      const type = r[2];
      const amount = Number(r[3]);
      if (!personTotals[name]) personTotals[name] = 0;
      if (type === "เงินเข้า") { income += amount; personTotals[name] += amount; }
      if (type === "เงินออก") { expense += amount; personTotals[name] -= amount; }
    });
    return ContentService.createTextOutput(JSON.stringify({ income, expense, persons: personTotals }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getHistory") {
    const data = rows.reverse().map(r => ({
      date: r[0],
      name: r[1],
      type: r[2],
      amount: r[3],
      desc: r[4],
      image: r[5]
    }));
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ message: "No action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ======================================================
// 🤖 LINE Webhook: ดึง groupId อัตโนมัติ
// ======================================================
function handleLineWebhook(e) {
  const json = JSON.parse(e.postData.contents);
  const event = json.events[0];
  const replyToken = event.replyToken;
  const source = event.source;
  const userMessage = event.message?.text || "";
  const replyUrl = "https://api.line.me/v2/bot/message/reply";

  const sourceId = source.groupId || source.roomId || source.userId || "unknown";

  if (userMessage.trim().toLowerCase() === "groupid?") {
    const replyText = `🆔 groupId ของที่นี่คือ:\n${sourceId}`;
    const payload = {
      replyToken: replyToken,
      messages: [{ type: "text", text: replyText }],
    };
    const options = {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN,
      },
      payload: JSON.stringify(payload),
    };
    UrlFetchApp.fetch(replyUrl, options);
  }

  return ContentService.createTextOutput("OK");
}
