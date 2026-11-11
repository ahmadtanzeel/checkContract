var file_name = "学習者情報入力フォーム";

// フォームの回答に応じ、スプレッドシートに生徒情報を記録する関数。
// フォームと同期したスプレッドシートも同じ機能を果たしうるが、手作業での修正との相性が悪いため、別シートに転記している。
function generateStudentInfo(e) {
  var ss = SpreadsheetApp.openById(FUF.CustomerInfoSsId[env]); //学習者・決済者情報
  var sheet = FUF.ReadSheetById(ss, FUF.StudentInfoSheetId[env]);

  // フォームの回答を取得
  var itemResponse = e.response.getItemResponses();
  var formResponses = FormResponses2Array(itemResponse); //回答を配列に変換する。

  // スプレッドシートに記録
  sheet.appendRow(formResponses);
}


function onStudentInfoFormSubmit(e) {
  // myFunctionの引数eを追加↑
  // Googleフォームの投稿内容をitemResponsesへ入れる
  var itemResponses = e.response.getItemResponses();

  // Googleフォームの質問項目名と回答を順番に変数textへ入れる
  var text = "学習者情報フォームに新規提出がありました。";
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    var question = itemResponse.getItem().getTitle();
    var answer = FormResponse2String(itemResponse);
    text += "\n" + question + " : " + answer;
  }

  // Googleチャットのスペースで取得した着信WebhookのURLを設定。コピペすればOK
  var post_url = FUF.StudentInfoWebhook[env];

  // スペースに投稿する本文(JSON形式) → 変数textの中身を入れているだけ
  var payload = {
    "text": text
  }

  // スペースに投稿する際のパラメータ情報(JSON形式)
  var options = {
    "method": "POST",
    "headers": { "Content-Type": "application/json; charset=UTF-8" },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  }

  // スペースの着信Webhook URLへ投稿
  var result = UrlFetchApp.fetch(post_url, options);
}