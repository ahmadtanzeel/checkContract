//to structuralize は具体的な数字を変数に置き換えたい部分。優先度は低いので現状そのままにしている

// refactored on 2024/11/25
// 領収書を自動生成し送信する関数
function receiptAutoCreate() {

  /// 必要な情報の取得（本命：★）
  var contractDbSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]); // 契約手続DBスプレッドシート
  var paymentDbSheet = FUF.ReadSheetById(contractDbSs, FUF.paymentDbGId[env]) //★支払情報シート
  var paymentArray = paymentDbSheet.getDataRange().getValues(); //★支払情報行列

  /// 領収書送付対象となる領収書を取得
  /// 条件は、支払い残金が１以下で、「解約/取消」および「領収書URL」が空欄となっていること。
  var targetPayments = paymentArray.filter(function (row) {
    return Math.abs(row[11]) <= 1 && row[11] !== "" && row[12] === '' && row[14] === ''; //to structuralize
  });
  /// 領収書送付対象となる生徒IDを取得
  var uniqueStudentID = FUF.getUniqueValues(targetPayments.map(function (row) { return row[3]; })); //to structuralize
  Logger.log(targetPayments);

  /// 学習者毎に領収書情報を検索し（学習者IDで識別）、領収書を作成し送付する
  for (id of uniqueStudentID) {
    var individualPayments = targetPayments.filter(function (row) {return row[3] === id}); //該当生徒の支払情報

    var studentName = individualPayments[0][4]; //学習者氏名 //to structuralize
    var contractantName = individualPayments[0][5]; // 決済者氏名 //to structuralize
    var paymentIds = individualPayments.map(function (row) {return row[1]}); // 該当生徒の支払いID一覧 //to structuralize

    //// 領収書を作成し送信する
    var receiptData = createAndSendReceipt(id, contractantName, studentName, paymentIds);

    //// 領収書を作成した支払情報に領収書 pdf の URL を転記する
    for (paymentId of paymentIds) {
      for (var i = 0; i < paymentArray.length; i++) {
        if (paymentArray[i][1] === paymentId && paymentArray[i][1] != "") {
          paymentDbSheet.getRange(i + 1, 15, 1, 4).setValues(receiptData); //to structuralize
        }
      }
    }
  }
}

// refactored on 2024/11/25
// 領収書を自動生成し送信する関数
function createAndSendReceipt(id, contractantName, studentName, paymentIds) {

  /// 必要な情報の取得
  var receiptSs = SpreadsheetApp.openById(FUF.receiptSsId[env]); //領収書スプレッドシート
  var receiptSheet = FUF.ReadSheetById(receiptSs, FUF.receiptGid[env]); //領収書テンプレシート
  var receiptId = receiptSheet.getRange(receiptIdCoordinate).getValue();  //今回作成する領収書Id

  /// テンプレシートに値を入力する
  /// サービス名称、契約書IDなどは、スプレッドシート内在の関数で出力される
  receiptSheet.getRange(receiptContractantNameCoordinate).setValue(contractantName); //決済者氏名を入力
  receiptSheet.getRange(receiptStudentNameCoordinate).setValue(studentName); //生徒氏名を入力
  receiptSheet.getRange(receiptStudentIdCoordinate).setValue(id); // 学習者IDを入力
  var receiptPaymentIdRange = receiptSheet.getRange(receiptPaymentIdsCoordinate);
  FUF.setValuesWithPadding(receiptPaymentIdRange, FUF.transpose(paymentIds)); // 行数や列数が合わなくても空白を補って setValues する
  SpreadsheetApp.flush();

  /// 領収書シートをコピー
  receiptSs.duplicateActiveSheet(); //領収書シートをコピー
  var sheet = receiptSs.getActiveSheet().setName(receiptId);  //シート名を領収書IDに変更。
  
  /// 領収書 pdf を生成し生徒の領収書フォルダに格納する
  /// 契約書IDにハイパーリンクを適用するため、値のみに変換する前に pdf 生成が必要
  var receiptPdfName = `SSS Education 領収書 - ` + id + ` - ` + studentName + `様 - ` + receiptId + '.pdf';
  var pdf = FUF.generatePDF(FUF.receiptSsId[env],sheet.getSheetId().toString(),receiptCustomPages, receiptPdfName);
  var receiptFolder = FUF.studentContractFolderConfig(id, studentName, env).receiptFolder;
  var receiptPdf = receiptFolder.createFile(pdf); //領収書PDF を領収書フォルダ内に作成

  /// 生成したシートを値のみに変換する
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).copyTo(receiptSs.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); 

  /// エラーを自動検知
  var data = sheet.getDataRange().getValues();
  var hasError = false;
  var errorCells = ""
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] == '#N/A' || data[i][j] == '#REF!') {
        hasError = true;
        Logger.log("Error found at row " + (i + 1) + " column " + (j + 1) + "in receipt:" + receiptId);
        var errorColumn = FUF.columnNumberToLetter(j)
        var errorRow = (i + 1).toString()
        errorCells = errorCells + errorColumn + errorRow + ",";
      }
    }
  }
  if (!hasError) {
    var autoValidation = "異常なし．";;
  } else {
    var autoValidation = "異常あり：" + errorCells;
  }

  /// 領収書シートを初期化
  receiptSheet.getRange(receiptIdCoordinate).setValue(FUF.addOneToStringNumber(receiptId)); //領収書IDは+1して書換 
  receiptSheet.getRange(receiptContractantNameCoordinate).setValue(""); //決済者氏名を空欄に
  receiptSheet.getRange(receiptStudentNameCoordinate).setValue(""); //学習者氏名を空欄に
  receiptSheet.getRange(receiptStudentIdCoordinate).setValue(""); //学習者ID欄を空欄に
  FUF.setValuesWithPadding(receiptSheet.getRange(receiptPaymentIdsCoordinate), [[""]]) //支払いID欄を空欄に
  receiptSheet.getRange(receiptBikouCoordinate).setValue(""); //備考欄を空欄に
  SpreadsheetApp.flush();

  // 支払情報シートへの転記用の情報を取得
  var receiptUrl = receiptPdf.getUrl();
  var customerInfo = FUF.studentId2Data(id, env);
  var contractantMail = customerInfo["contractantMail"];

  // 自動検知で問題なければ、領収書を Gmail で送付する
  if (autoValidation == "異常なし．") {
    var receiptFileId = FUF.extractFileId(receiptUrl);
    if (receiptFileId) {
      var receiptFile = DriveApp.getFileById(receiptFileId);
      receiptFile.addViewer(contractantMail);
      var receiptSent = sendReceiptMail(contractantMail, studentName, contractantName, receiptUrl)
    } else {
      Logger.log("領収書URLが不正です．")
    }
  }

  return [[receiptUrl, receiptId, autoValidation, receiptSent]] //契約書ファイルIDを返す
}


// refactored on 2024/11/25
// 領収書をメールで送付する
function sendReceiptMail(contractantMail, studentName, contractantName, receiptUrl) {
  var subject = '【SSS Education】領収証を送付いたします。'
  var body = `${contractantName}様 \n \n`
    + `平素よりお世話になっております。株式会社 SSS Education 事務局です。\n \n`
    + `この度は、${studentName}様の授業料をお支払いいただき、誠にありがとうございます。お支払いいただいた授業料を確かに受領いたしましたことを、下記領収書URLを共有する形でご報告申し上げます。\n \n`
    + `${receiptUrl}\n \n`
    + `引き続き，スタッフ一同で${studentName}様の学習を誠心誠意サポートさせていただきますので、今後ともどうぞよろしくお願いいたします。\n何かご不明な点やご質問等がございましたら、どうぞお気軽にご連絡ください。\n \n`
    + `引き続き、SSS Education をよろしくお願い申し上げます。\n \n`
    + `SSS Education 事務局`;
  try {
    GmailApp.sendEmail(contractantMail, subject, body);
    return true
  } catch (e) {
    console.log(`エラー内容:` + e.message);
    return false
  }

}