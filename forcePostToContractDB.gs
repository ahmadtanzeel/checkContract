function forcePostToContractDB() {
  // ポップアップで ID を入力させる
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('ID入力', '7桁のIDを入力してください:', ui.ButtonSet.OK_CANCEL);

  // ユーザーが OK を押したかを確認
  if (response.getSelectedButton() == ui.Button.OK) {
    var id = response.getResponseText();

    // 入力された ID が 7 桁かどうかをチェック
    if (/^\d{7}$/.test(id)) {
      // スプレッドシートのIDを指定
      var spreadsheetId = FUF.contractDBSsId[env];
      // スプレッドシートを取得
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      // 「契約手続」シートを取得
      var sheet = spreadsheet.getSheetByName('契約手続');

      // 「契約手続」シートのC列のデータを取得
      var columnCValues = sheet.getRange('C:C').getValues();

      // 7桁のIDがC列に存在するかをチェック
      var idExists = columnCValues.some(function (row) {
        return row[0] == id;
      });

      // 同一IDが無い場合、特定の関数を実行
      if (!idExists) {
        var contract = SpreadsheetApp.openById(FUF.contractSsID[env]);//契約書スプシ
        var contractSheet = contract.getSheetByName(id); //該当する契約書IDのシートを取得
        if (contractSheet === null) {
          //ダイアログを表示
          var html = HtmlService.createHtmlOutput('<a href="https://docs.google.com/spreadsheets/d/' + FUF.contractSsID[env] + '" target="_blank" rel="noopener noreferrer">契約書一覧を確認する。</a>')
            .setHeight(100)
            .setWidth(200)
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          SpreadsheetApp.getUi().showModalDialog(html, "指定IDの契約書が存在しません。");
          return
        } else {
          var studentId = contractSheet.getRange(studentIdCoordinate).getValue();
          var studentName = contractSheet.getRange(studentNameCoordinate).getValue(); //本人氏名
          var contractPdfName = `SSS Education 契約書 - ${studentId} - ${studentName}様 - ${id}.pdf`
          var contractPdfs = DriveApp.getFilesByName(contractPdfName);
          if (contractPdfs.hasNext()) {
            var pdf = contractPdfs.next();
            var pdfUrl = pdf.getUrl();
            if (contractPdfs.hasNext()) {
              ui.alert(`「SSS Education 契約書 - ${studentId} - ${studentName}様 - ${id}.pdf」という名称の契約書pdfが複数あります．整理してやり直してください．`)
              return
            }
          } else {
            ui.alert(`「SSS Education 契約書 - ${studentId} - ${studentName}様 - ${id}.pdf」という名称の契約書pdfが見つかりません。`);
            return
          }
          postContractToDB(contractSheet, pdfUrl)
          ui.alert("関数を実行しました。")
          return
        }

      } else {
        ui.alert('IDが既に存在します。');
      }
    } else {
      ui.alert('無効なIDです。7桁の数字を入力してください。');
    }
  }
}

function postContractToDB(contractSheet, pdfUrl) {

  var contractID = contractSheet.getRange(contractIdCoordinate).getValue();
  var studentId = contractSheet.getRange(studentIdCoordinate).getValue();
  var studentName = contractSheet.getRange(studentNameCoordinate).getValue(); //本人氏名
  var contractantName = contractSheet.getRange(contractantNameCoordinate).getValue(); //決済者氏名 
  var paymentMethod = contractSheet.getRange(paymentMethodCoordinate).getValue(); //支払方法
  var conferenceSchedule = contractSheet.getRange(mtgDateCoordinate).getValue(); //面談予定日
  var paymentReccurence = contractSheet.getRange(paymentReccurenceCoordinate).getValue() //支払回数
  var services = contractSheet.getRange(contractServiceInfoRange[0], contractServiceInfoRange[1], contractServiceInfoRange[2], contractServiceInfoRange[3]).getValues()
  var validService = removeInvalidServices(services)

  var approvalList = [];
  for (_service of validService) {
    _service = FUF.removeBlankElements(_service)
    var service = [];
    service[0] = new Date();
    service[1] = "成約";
    service[2] = contractID;
    service[3] = studentId;
    service[4] = studentName;

    service[5] = _service[0];
    service[6] = _service[2];
    service[7] = _service[3];
    service[8] = _service[4] * 1.10;
    service[9] = _service[5];
    service[10] = _service[6] * 1.10;

    service[11] = paymentMethod; //支払方法
    service[12] = paymentReccurence; //支払回数
    service[13] = pdfUrl; //契約書 URL
    approvalList.push(service)
  }

  var contractdBSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]);
  var contractDbSheet = FUF.ReadSheetById(contractdBSs, FUF.contractDBGId[env])
  var lastRow = contractDbSheet.getLastRow();
  contractDbSheet.getRange(lastRow + 1, 1, approvalList.length, 14).setValues(approvalList);
  postPaymentData(studentId, studentName, contractID, contractantName, contractSheet);

  Logger.log("特定の関数が実行されました。");
}
