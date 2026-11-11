//　【用語】期限切迫：期限日がもうすぐであること

// 支払いスケジュールを調べ、切迫通知予定Chat・切迫通知メール下書・超過通知予定Chat・超過通知メール下書を実施する
function paymentReminder() {

  /// 支払情報の振分
  var remindList = paymentFilter();

  /// 支払い情報の振分結果
  console.log(remindList.rowsDueToPrePreRemind) //切迫通知予定Chat
  console.log(remindList.rowsDueToPreRemind) //切迫通知メール下書
  console.log(remindList.rowsDueToPreOverdueAlert) //超過通知予定Chat
  console.log(remindList.rowsDueToOverdueAlert) //超過通知メール下書
  // console.log(remindList.rowsPerfectlyOverdueAlert) //ブラックリスト

  if (remindList.rowsDueToPrePreRemind.length != 0) {
    preRemindToOffice(FUF.extractColumn(remindList.rowsDueToPrePreRemind, paymentChatReminderColumns, 1));
  }
  if (remindList.rowsDueToPreRemind.length != 0) {
    preRemindToCustomer(remindList.rowsDueToPreRemind);
  }
  if (remindList.rowsDueToPreOverdueAlert.length != 0) {
    dueRemind(FUF.extractColumn(remindList.rowsDueToPreOverdueAlert, paymentChatReminderColumns, 1));
  }
  if (remindList.rowsDueToOverdueAlert.length != 0) {
    overdueRemind(remindList.rowsDueToOverdueAlert);
  }

  //１週間以上たっても振込がない場合に，ブラックリストに記載する．
  // blackListPost(remindList.rowsPerfectlyOverdueAlert)
}


// refactored on 2024/11/18
// 支払情報をリストに振り分ける
function paymentFilter() {

  /// paymentReminderSpan = X とすると、支払期限が X 日後に迫っていると通知メールを送る
  /// salesPersonMargin = Y とすると、支払期限が X+Y 日後に迫っていると Google Chat に通知を送る
  /// paymentOverdueSpan = Z とすると、支払期限の翌営業日が Z 日前でも支払いが確認できなければ支払依頼メールを送る (Z-Y 日前に Google Chat に通知)
  /// paymentPerfectlyOverdueSpan = A とすると、支払期限が A 日前ならブラックリストに入る

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
  var prePreRemindDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() + paymentReminderSpan + salesPersonMargin).getTime(); //切迫通知予定
  var preRemindDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() + paymentReminderSpan).getTime(); //切迫通知対象
  var preOverdueDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - paymentOverdueSpan + salesPersonMargin).getTime(); //超過通知予定
  var overdueDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - paymentOverdueSpan).getTime(); //超過通知対象
  var perfectlyOverdueDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - paymentPerfectlyOverdueSpan).getTime(); //ブラックリスト対象

  /// リストの準備
  var rowsDueToPrePreRemind = []; //切迫通知予定
  var rowsDueToPreRemind = [];  //切迫通知対象
  var rowsDueToPreOverdueAlert = []; //超過通知予定
  var rowsDueToOverdueAlert = []; //超過通知対象
  var rowsPerfectlyOverdueAlert = []; //ブラックリスト対象

  /// 支払情報の各行を、各リストに振り分ける
  for (var i = 1; i < availablePayment.length; i++) {
    var rowDate = new Date(availablePayment[i][8]); // 日付欄を読込
    var rowDateFormatted = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate()).getTime();
    var nextBusinessDay = FUF.getNextBusinessDay(rowDateFormatted).getTime();

    if (rowDateFormatted === prePreRemindDate) { //支払期限日がX+Y日後
      rowsDueToPrePreRemind.push(availablePayment[i]);
    } else if (rowDateFormatted === preRemindDate) { //支払期限日がX日後
      rowsDueToPreRemind.push(availablePayment[i]);
    } else if (nextBusinessDay === preOverdueDate) { // 支払期限日の翌営業日がZ-Y日前
      rowsDueToPreOverdueAlert.push(availablePayment[i]);
    } else if (nextBusinessDay === overdueDate) { //支払期限日の翌営業日がZ日前
      rowsDueToOverdueAlert.push(availablePayment[i]);
    } else if (rowDateFormatted < perfectlyOverdueDate) { //支払期限日がA日前
      rowsPerfectlyOverdueAlert.push(availablePayment[i]);
    }
  }

  return {
    rowsDueToPrePreRemind: rowsDueToPrePreRemind,
    rowsDueToPreRemind: rowsDueToPreRemind,
    rowsDueToPreOverdueAlert: rowsDueToPreOverdueAlert,
    rowsDueToOverdueAlert: rowsDueToOverdueAlert,
    rowsPerfectlyOverdueAlert: rowsPerfectlyOverdueAlert
  };
}

// refactored on 2024/11/18
// 翌日のメール通知予定リストを、Google Chat に通知を送信する．
function preRemindToOffice(preRemindList) {

  /// 参照情報（本命★）の読込
  var contractDbSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]); // 契約手続DB
  var paymentDbSheet = FUF.ReadSheetById(contractDbSs, FUF.paymentDbGId[env]) // 支払情報シート
  var paymentDbArray = paymentDbSheet.getDataRange().getValues(); // ★支払情報行列
  var notificationListSs = SpreadsheetApp.openById(FUF.prePreRemindSsId[env]); //★通知予定リストのスプレッドシート

  /// 新規シート名生成：yyyyMMdd作成, 一つ既存の同一名シートがあれば yyyyMMdd作成(1)
  var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var baseSheetName = formattedDate + "作成";
  var sheetName = baseSheetName;
  var sheetIndex = 1;
  //// 同じ名称のシートが存在する場合に番号を付加
  while (notificationListSs.getSheetByName(sheetName) !== null) {
    sheetName = baseSheetName + ' (' + sheetIndex + ')';
    sheetIndex++;
  }

  /// ヘッダー行生成
  var headRow = FUF.extractColumn([paymentDbArray[0]], paymentChatReminderColumns, 1)[0]

  /// 新規シート作成
  var newNotificationListSheet = notificationListSs.insertSheet(sheetName);
  newNotificationListSheet.appendRow(headRow); // ヘッダー行
  preRemindList.forEach(function (row) {
    newNotificationListSheet.appendRow(row);
  }); //データ行

  /// 新しいシートを一番左に移動
  notificationListSs.setActiveSheet(newNotificationListSheet);
  notificationListSs.moveActiveSheet(1);
  SpreadsheetApp.flush();

  /// 新規シートへの URL 生成
  var sheetUrl = notificationListSs.getUrl() + '#gid=' + newNotificationListSheet.getSheetId();

  /// Google Chatに通知を送信
  var message = {
    "text": `【振込期限日４日前リスト】\n${salesPerson["nickname"]} （<users/${salesPerson["chatId"]}>）様\n以下の振込依頼を明日通知予定です。すれ違いにならないよう、本日中に、振込確認漏れがないかを確認してください。\n<${sheetUrl}|振込依頼通知予定リスト - ${sheetName}>`
  };
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(message)
  };
  UrlFetchApp.fetch(FUF.PaymentReminderWebhook[env], options);

}


// refactored on 2024/11/18
// 振込通知依頼メールを送信する．
function preRemindToCustomer(preRemindList) {

  /// ユニーク学習者IDリストを生成
  var studentIds = FUF.extractColumn(preRemindList, 3, 1); // 生徒ID一覧の列を抽出 //to FUF
  var uniQueStudentIds = FUF.getUniqueValues(studentIds); // ユニーク値のみ取得

  /// 学習者毎に支払情報を抽出しメールを送信する
  for (studentId of uniQueStudentIds) {

    //// 学習者IDで該当生徒の支払い情報を抽出 
    var targetPayments = preRemindList.filter(function (row) {
      return row[3] == studentId;
    });
    var paymentBody = FUF.extractColumn(targetPayments, [8, 7, 9, 2], 1) //必要な列だけ抽出 //to FUF

    //// 学習者情報の取出
    var customerInfo = FUF.studentId2Data(studentId, env); //学習者IDから学習者情報を取得
    var contractantMail = customerInfo["contractantMail"]; //決済者メールアドレス
    var studentName = customerInfo["studentName"]; //学習者氏名
    var contractantName = customerInfo["contractantName"]; //決済者氏名
    var senderName = studentId + customerInfo["contractantSei"] + customerInfo["contractantMei"] //振込依頼人名義

    //// メールの文面生成
    var subject = "【SSS Education】振込期限３日前をお知らせいたします。" //タイトル
    var body = greeting(contractantName); //挨拶

    body +=
      `日頃より、弊塾へのご理解とご協力を賜り、誠にありがとうございます。<br>`
      + `この度は${studentName}様の授業料につきまして、振込期限が近づいてまいりましたので、ご案内申し上げます。<br>`
      + `<br>`;

    body +=
      `以下の通り、振込期限が3日後となっております。<br>`
      + paymentOverviewTable(paymentBody) + `<br>` //支払概要表
      + `振込先情報と振込依頼人名義は、以下の通りでお願いいたします。<br>`
      + transferInfoTable(senderName) + `<br>`; //振込先情報と振込依頼人名義

    body +=
      `お手数をおかけいたしますが、期日までにお手続きいただけますようお願い申し上げます。もし、すでにお振込みが完了している場合は、本書面はご放念くださいませ。何かご不明な点やご質問等がございましたら、どうぞお気軽にお申し付けください。<br>`
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
          if(cellBody == ""){kaigyo = ""}; // 通知状況欄が空白なら改行は不要
          paymentDbSheet.getRange(i + 1, 14, 1, 1).setValue(
            cellBody + kaigyo 
            + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd') + "期限切迫通知送付済"
            ); // 「yyyyMMdd期限切迫通知送付済」という記載を追加
        }
      }
    }
  }

  /// Google Chatへの通知メッセージを作成
  var message = {
    "text": `【振込期限通知メールの下書きを作成しました！】\n ${salesPerson["nickname"]} （<users/${salesPerson["chatId"]}>）様\n admin@sss-edu.org の Gmail に振込期限通知メールの下書きを作成しました。問題なければ送信してください。`
  };
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(message)
  };
  UrlFetchApp.fetch(FUF.PaymentReminderWebhook[env], options);
}


// refactored on 2024/11/18
// 動作未確認
// 振込期限超過リストを Google Chat に通知を送信する．
function dueRemind(dueRemindList) {

  /// 参照情報（本命★）の読込
  var contractDbSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]); // 契約手続DB
  var paymentDbSheet = FUF.ReadSheetById(contractDbSs, FUF.paymentDbGId[env]) // 支払情報シート
  var paymentDbArray = paymentDbSheet.getDataRange().getValues(); // ★支払情報行列
  var overdueListSs = SpreadsheetApp.openById(FUF.preOverdueSsId[env]); //★通知予定リストのスプレッドシート

  /// 新規シート名生成：yyyyMMdd作成, 一つ既存の同一名シートがあれば yyyyMMdd作成(1)
  var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var baseSheetName = formattedDate + "作成";
  var sheetName = baseSheetName;
  var sheetIndex = 1;
  //// 同一名称のシートが存在する場合に番号を付加
  while (overdueListSs.getSheetByName(sheetName) !== null) {
    sheetName = baseSheetName + ' (' + sheetIndex + ')';
    sheetIndex++;
  }

  /// ヘッダー行生成
  var headRow = FUF.extractColumn([paymentDbArray[0]], paymentChatReminderColumns, 1)[0]

  /// 新規シート作成
  var newOverdueListSheet = overdueListSs.insertSheet(sheetName);
  newOverdueListSheet.appendRow(headRow);// ヘッダー行
  dueRemindList.forEach(function (row) {
    newOverdueListSheet.appendRow(row);
  });// データ行

  /// 新規シートを一番左に移動
  overdueListSs.setActiveSheet(newOverdueListSheet);
  overdueListSs.moveActiveSheet(1);
  SpreadsheetApp.flush();

  /// 新規シートへの URL 生成
  var sheetUrl = overdueListSs.getUrl() + '#gid=' + newOverdueListSheet.getSheetId();

  /// Google Chatへの通知メッセージを作成
  var message = {
    "text": `【振込期限日超過リスト】\n${salesPerson["nickname"]} （<users/${salesPerson["chatId"]}>）様\n以下の振込期限超過通知を明日通知予定です。すれ違いにならないよう、本日中に、振込確認漏れがないかを確認してください。\n<${sheetUrl}|振込期限超過通知予定リスト - ${sheetName}>`
  };
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(message)
  };
  UrlFetchApp.fetch(FUF.PaymentReminderWebhook[env], options);
}


// refactored on 2024/11/18
// 振込期限超過メールを送信する
function overdueRemind(postRemindList) {

  /// ユニーク学習者IDリストを生成
  var studentIds = FUF.extractColumn(postRemindList, 3, 1);
  var uniQueStudentIds = FUF.getUniqueValues(studentIds);

  /// 学習者毎に支払情報を抽出しメールを送信する
  for (studentId of uniQueStudentIds) {

    //// 学習者IDで該当生徒の支払い情報を抽出 
    var targetPayments = postRemindList.filter(function (row) {
      return row[3] == studentId;
    });
    var paymentBody = FUF.extractColumn(targetPayments, [8, 7, 9, 2], 1) //必要な列だけ抽出 //to FUF

    //// 学習者情報の取出
    var customerInfo = FUF.studentId2Data(studentId, env); //学習者IDから学習者情報を取得
    var contractantMail = customerInfo["contractantMail"]; //決済者メールアドレス
    var studentName = customerInfo["studentName"]; //学習者氏名
    var contractantName = customerInfo["contractantName"]; //決済者氏名
    var senderName = studentId + customerInfo["contractantSei"] + customerInfo["contractantMei"] //振込依頼人名義

    var subject = "【SSS Education】振込期限超過をお知らせいたします。"

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
  }
  /// Google Chatへの通知メッセージを作成
  var message = {
    "text": `【振込期限超過通知メールの下書きを作成しました！】\n ${salesPerson["nickname"]} （<users/${salesPerson["chatId"]}>）様\n admin@sss-edu.org の Gmail に振込期限超過通知メールの下書きを作成しました。問題なければ送信してください。`
  };
  var options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(message)
  };
  UrlFetchApp.fetch(FUF.PaymentReminderWebhook[env], options);

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
          if(cellBody == ""){kaigyo = ""}; // 通知状況欄が空白なら改行は不要
          paymentDbSheet.getRange(i + 1, 14, 1, 1).setValue(
            cellBody + kaigyo 
            + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd') + "期限超過通知送付済"
            );// 「yyyyMMdd期限超過通知送付済」という記載を追加
        }
      }
    }
}


// 
function blackListPost(blackList){
  
  /// 参照情報（本命★）の読込
  var contractDbSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]); // 契約手続DB
  var paymentDbSheet = FUF.ReadSheetById(contractDbSs, FUF.paymentDbGId[env]) // 支払情報シート
  var paymentDbArray = paymentDbSheet.getDataRange().getValues(); // ★支払情報行列
  var blackListSs = SpreadsheetApp.openById(FUF.prePreRemindSsId[env]); //★通知予定リストのスプレッドシート
  var blackListSheet = FUF.ReadSheetById(blackListSs, FUF.blackListSheetId[env]) // 支払情報シート
  

}

// refactored on 2024/11/18
// （メール文面用）振込先・振込依頼人名義の表
function transferInfoTable(senderName) {
  var body
    = `<table border='1' style='border-collapse:collapse;'>`
    + `<tr><td style='padding: 8px;'>銀行名</td><td style='padding: 8px;'>GMOあおぞら銀行(金融機関コード：0310)</td></tr>`
    + `<tr><td style='padding: 8px;'>支店名</td><td style='padding: 8px;'>法人第二営業部(支店コード：102)</td></tr>`
    + `<tr><td style='padding: 8px;'>種類</td><td style='padding: 8px;'>普通預金</td></tr>`
    + `<tr><td style='padding: 8px;'>口座番号</td><td style='padding: 8px;'>1307672</td></tr>`
    + `<tr><td style='padding: 8px;'>口座名義名</td><td style='padding: 8px;'>株式会社 SSS Education</td></tr>`
    + `<tr><td style='padding: 8px;'>振込依頼人名義</td><td style='padding: 8px;'<strong><font color="red">${senderName}</font></strong></td></tr>`
    + "</table>";

  return body
}

// refactored on 2024/11/18
// （メール文面用）支払概要の表
function paymentOverviewTable(targetPayments) {

  /// 参照情報（本命★）の読込
  var contractDbSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]);  // 契約手続DBスプレッドシート
  var contractDbSheet = FUF.ReadSheetById(contractDbSs, FUF.contractDBGId[env]); // 契約手続シート
  var contractArray = contractDbSheet.getDataRange().getValues(); // ★契約手続行列

  /// 表の準備
  var body
    = `<table border='1' style='border-collapse:collapse;'>`
    + `<tr>`
    + `<th style='padding: 8px;'>振込期限日</th>`
    + `<th style='padding: 8px;'>項目</th>`
    + `<th style='padding: 8px;'>金額</th>`
    + `<th style='padding: 8px;'>契約書ID</th>`
    + `</tr>`; // 表
  var total = 0; // 合計額

  /// 支払の各行を表形式で追記していく
  targetPayments.forEach(function (row) {
    var paymentDueDate = Utilities.formatDate(row[0], "JST", "yyyy/MM/dd")
    var item = row[1]; //項目（サービス名）
    var value = Math.round(row[2]); // 金額
    var formattedValue = "￥" + value.toLocaleString(); // ￥123,456形式
    var contractID = row[3] // 契約書ID
    for (var i = 0; i < contractArray.length; i++) {
      if (contractArray[i][2] === contractID) {
        var link = contractArray[i][13]; // 契約書URL
        break
      }
    }

    var paymentRow
      = `<tr>`
      + `<td style='padding: 8px;'>${paymentDueDate}</td>`
      + `<td style='padding: 8px;'>${item}</td>`
      + `<td style='padding: 8px;'>${formattedValue}</td>`
      + `<td style='padding: 8px;'><a href="${link}">${contractID}</a></td>`
      + `</tr>`;

    body += paymentRow; // 表に追記
    total += value; // 合計額に加算
  });

  ///合計金額の行を追記
  var formattedTotal = "￥" + total.toLocaleString(); // ￥123,456形式
  var hooter
    = `<tr>`
    + `<td></td>`
    + `<td style='padding: 8px;'><strong>合計</strong></td>`
    + `<td style='padding: 8px;'><strong>${formattedTotal}</strong></td>`
    + `<td></td>`
    + `</tr>`;
  body += hooter;

  /// 表の完成
  body += "</table>";

  return body
}

// （メール文面用）挨拶
function greeting(surName) {
  var body =
    `${surName}様<br>`
    + `<br>`
    + `平素よりお世話になっております。SSS Education 事務局です。<br>`
    + `<br>`;

  return body;
}

// （メール文面用）謝辞
function regard(studentName) {
  var body =
    `引き続き、${studentName}様の学習をサポートさせていただきます。<br>`
    + `今後とも SSS Education をどうぞよろしくお願い申し上げます。<br>`
    + `<br>`
    + `SSS Education 事務局`;

  return body;
}