// refactored on 2024/11/26
// 振込期限がxxx日以上前である場合に通知する
function paymentReminderTemp() {
  const xxx = 7 // 閾値を設定

  /// 支払情報の振分
  var remindList = tempPaymentFilter(xxx);
  console.log(remindList) //振分結果

  if (remindList.length != 0) {
    temporaryRemind(remindList);
  }
}

// refactored on 2024/11/26
// 支払期限日の翌営業日が xxx 日前に該当する支払を抽出する関数
function tempPaymentFilter(xxx) {

  /// 参照情報（本命★）の読込
  var contractDbSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]); // 契約手続DBスプレッドシート
  var paymentDbSheet = FUF.ReadSheetById(contractDbSs, FUF.paymentDbGId[env]) //支払情報シート
  var paymentDbArray = paymentDbSheet.getDataRange().getValues(); //★支払情報行列

  /// アクティブな支払情報のみ抽出（未入金残金2円以上、かつ、取消・解約がない）
  var availablePayment = paymentDbArray.filter(function (paymentRow) {
    return paymentRow[11] >= 2 && paymentRow[12] == "";
  });

  /// 重要な日付を準備。.getTime() にするのは数式化するため。
  var currentDate = new Date();
  var xxxDaysBefore = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - xxx).getTime(); //切迫通知予定

  /// リストの準備
  var rowToRemind = [];

  /// 支払情報の各行を、各リストに振り分ける
  for (var i = 1; i < availablePayment.length; i++) {
    var rowDate = new Date(availablePayment[i][8]); // 日付欄を読込
    var rowDateFormatted = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate()).getTime();
    var nextBusinessDay = FUF.getNextBusinessDay(rowDateFormatted).getTime();

    if (nextBusinessDay < xxxDaysBefore) { // 支払期限日の翌営業日が xxx 日前
      rowToRemind.push(availablePayment[i]);
    }
  }
  return rowToRemind
}

// refactored on 2024/11/26
// 入力された支払内容に応じて、期限超過メールを作成する関数
function temporaryRemind(remindList) {

  /// ユニーク学習者IDリストを生成
  var studentIds = FUF.extractColumn(remindList, 3, 1); // 生徒ID一覧の列を抽出 //to FUF
  var uniQueStudentIds = FUF.getUniqueValues(studentIds); // ユニーク値のみ取得


  /// 学習者毎に支払情報を抽出しメールを送信する
  for (studentId of uniQueStudentIds) {

    //// 学習者IDで該当生徒の支払い情報を抽出 
    var targetPayments = remindList.filter(function (row) {
      return row[3] == studentId;
    });
    var paymentBody = FUF.extractColumn(targetPayments, [8, 7, 9, 2], 1) //必要な列だけ抽出 //to FUF

    //// 学習者情報の取出
    var customerInfo = FUF.studentId2Data(studentId, env); //学習者IDから学習者情報を取得
    var contractantMail = customerInfo["contractantMail"]; //決済者メールアドレス
    var studentName = customerInfo["studentName"]; //学習者氏名
    var contractantName = customerInfo["contractantName"]; //決済者氏名
    var senderName = studentId + customerInfo["contractantSei"] + customerInfo["contractantMei"] //振込依頼人名義

    var subject = "【SSS Education】お振込依頼内容の通知"

    var body = greeting(contractantName); //挨拶
    body +=
      `日頃より、当塾へのご理解とご協力を賜り、誠にありがとうございます。<br>`
      + `この度は${studentName}様の授業料につきまして、振込期限日が超過しておりますが，お振込みが確認できませんでしたので、ご案内申し上げます。<br>`
      + `<br>`;

    body +=
      `お振込みが確認できていませんサービス内容は，以下の通りです。<br>`
      + paymentOverviewTable(paymentBody) + `<br>` //支払概要表
      + `振込先情報と振込依頼人名義は、以下の通りでお願いいたします。<br>`
      + transferInfoTable(senderName) + `<br>`; //振込先情報と振込依頼人名義

    body += `お手数をおかけいたしますが、可能な限り早急にご確認の上、お振込いただけますようお願い申し上げます。もし、すでにお振込みが完了している場合は、お手数をおかけし大変申し訳ございませんが <a href="https://lin.ee/JyviL3u">公式LINE</a>までお問い合わせください。<br>`
      + `その他、何かご不明な点やご質問等がございましたら、どうぞお気軽にお申し付けください。<br>`
      + `<br>`;

    body += regard(studentName);


    GmailApp.createDraft(contractantMail, subject, "", { htmlBody: body });

    /// 参照情報（本命★）の読込
    var contractDbSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]); // 契約手続DB
    var paymentDbSheet = FUF.ReadSheetById(contractDbSs, FUF.paymentDbGId[env]) // ★支払情報シート
    var paymentDbArray = paymentDbSheet.getDataRange().getValues(); //★支払情報行列

    /// 支払いIDをもとに通知状況欄に追記
    var paymentIds = FUF.extractColumn(targetPayments, 1, 1)
    //// 各支払ID毎に、支払情報スプレッドシートを検索し、支払IDが合致する行の通知状況欄に追記
    for (paymentId of paymentIds) {
      for (var i = 0; i < paymentDbArray.length; i++) {
        if (paymentDbArray[i][1] === paymentId && paymentDbArray[i][1] != "") {
          var cell = paymentDbSheet.getRange(i + 1, 14, 1, 1) // 通知状況欄
          var cellBody = cell.getValue(); // 通知状況欄の記載
          var kaigyo = "\n"
          if (cellBody == "") { kaigyo = "" }; // 通知状況欄が空白なら改行は不要
          paymentDbSheet.getRange(i + 1, 14, 1, 1).setValue(
            cellBody + kaigyo
            + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd') + "臨時期限超過通知メール作成済"
          );// 「yyyyMMdd臨時期限超過通知メール作成済」という記載を追加
        }
      }
    }
  }
}