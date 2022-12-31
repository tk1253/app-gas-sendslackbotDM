const APIToken = "xoxb-xxxxxxxxxxxx-xxxxxxxxxxxxx-xxxxxxxxxxxxxxxxxxxxxxxx"; // 使うSlackAppの「Bot User OAuth Token」を記載(裏に格納して隠すのがベター)
const ss = SpreadsheetApp.getActiveSpreadsheet(); 
const sheetName = ss.getSheetByName('シート名'); // 配信内容を記載したシート名
const infoCol = {
  "SlackID": "C", // DMを送る相手のIDが記載された列
  "content1": "D", // 送る内容が記載された列
  "content2": "E",
  "content3": "F",
  "content4": "G",
  "content5": "H",
  "content6": "I",
  "content7": "J",
  "content8": "K",
}

function main() {
  const rowBegin = 2; // シートに合わせて変更
  const rowEnd = 33; // シートに合わせて変更
  const APIMethodUrl = "https://slack.com/api/chat.postMessage";
　// 実行ボタンの確認メッセージ
  const MsgBox = Browser.msgBox("GASを実行し、ユーザーにメッセージを送信します", Browser.Buttons.OK_CANCEL);
  if (MsgBox == "cancel") {
    Browser.msgBox("GASの実行を中止しました");
  } else {
    // 「postmessage」を実行！
  postMessage(rowBegin, rowEnd, APIMethodUrl);
    Browser.msgBox("GASを実行しました");
  }
}

function getUserID(row) {
  let userID = sheetName.getRange(infoCol.SlackID + row).getValue();

  return userID;

}

function createMessage(row) {
  let msgContent1 = sheetName.getRange(infoCol.content1 + row).getValue();
  let msgContent2 = sheetName.getRange(infoCol.content2 + row).getValue();
  let msgContent3 = sheetName.getRange(infoCol.content3 + row).getValue();
  let msgContent4 = sheetName.getRange(infoCol.content4 + row).getValue();
  let msgContent5 = sheetName.getRange(infoCol.content5 + row).getValue();
  let msgContent6 = sheetName.getRange(infoCol.content6 + row).getValue();
  let msgContent7 = sheetName.getRange(infoCol.content7 + row).getValue();
  let msgContent8 = sheetName.getRange(infoCol.content8 + row).getValue();
// D列からのメッセージコンテンツの件名
  let msgSubject1 = sheetName.getRange(1,4).getValue();
  let msgSubject2 = sheetName.getRange(1,5).getValue();
  let msgSubject3 = sheetName.getRange(1,6).getValue();
  let msgSubject4 = sheetName.getRange(1,7).getValue();
  let msgSubject5 = sheetName.getRange(1,8).getValue();
  let msgSubject6 = sheetName.getRange(1,9).getValue();
  let msgSubject7 = sheetName.getRange(1,10).getValue();
  let msgSubject8 = sheetName.getRange(1,11).getValue();
  

  let message = ("\n\n"+ msgSubject1 +"\n"+ msgContent1+ "\n\n"+ msgSubject2 +"\n"+ msgContent2+ "\n\n"+ msgSubject3 +"\n"+ msgContent3+ "\n\n"+ msgSubject4 +"\n"+ msgContent4+ "\n\n"+ msgSubject5 +"\n"+ msgContent5+ "\n\n"+ msgSubject6 +"\n"+ msgContent6+ "\n\n"+ msgSubject7 +"\n"+ msgContent7+ "\n\n"+ msgSubject8 +"\n"+ msgContent8 ); 

  return message;

}

function postMessage(rowBegin, rowEnd, APIMethodUrl) {
  for (let row = rowBegin; row < rowEnd; row++) {  //for(初期化式; 条件式; 増減式)
    let userID = getUserID(row);

    if (userID == "")
      continue;

    let payload = {
      "token": APIToken,
      "channel": userID,
      "text": createMessage(row)
    };

    let params = {
      "method" : "post",
      "payload" : payload
    };

    UrlFetchApp.fetch(APIMethodUrl, params);
  }  //for文ここまで
}
