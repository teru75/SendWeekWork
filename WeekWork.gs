// 週報のブックをIDで取得
let spreadSheet = SpreadsheetApp.openById('1xUNYqS9l91BIDbI4oDO8ZOPe4bOl3RB_d-CuC2ulwxE');

function UpdateWeekWork(){
  // 実行する日の日付を取得
  var date = new Date();

  // 月曜日用の日付取得
  var monday = new Date();

  // コピー元のシートを指定
  var originalSheet = spreadSheet.getSheetByName('20210125');

  //  コピーを作成
  var copySheet = originalSheet.copyTo(spreadSheet);

　// コピー先のシート名を変更
  var sheetName = copySheet.getSheetName(); 
　sheetName = Utilities.formatDate(date,'JST', "yyyyMMdd");
  copySheet.setName(sheetName);

  // 月、金曜日の日付を取得、セルに入力
  monday.setDate(monday.getDate() - 6);
  var mondayDate = Utilities.formatDate(monday,'JST', 'yyyy/MM/dd');
  var sundayDate = Utilities.formatDate(date,'JST', 'yyyy/MM/dd');
  copySheet.getRange("B5").setValue(mondayDate);
  copySheet.getRange("D5").setValue(sundayDate);
}

function SendWeekWork(){
  //スプレッドシートオブジェクトからIDを取り出す
  var fileId   = spreadSheet.getId();

  //Excelファイルの拡張子を変更
  var xlsxName = spreadSheet.getName() + ".xlsx";

  //エクスポート用のURLはこちら
  var fetchUrl = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + fileId + "&amp;exportFormat=xlsx";

  //OAuth2対応が必要
  var fetchOpt = {
    "headers" : { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    "muteHttpExceptions" : true
  };

  //URLをダウンロード
  var xlsxFile = UrlFetchApp.fetch(fetchUrl, fetchOpt).getBlob().setName(xlsxName)

  // 日曜日の夜9時に自動送信
  MailApp.sendEmail('shiki@citycom.co.jp', '週報について','お疲れ様です。仲井です。\n週報を提出させていただきます。よろしくお願い致します。', {attachments:[xlsxFile]});
}
