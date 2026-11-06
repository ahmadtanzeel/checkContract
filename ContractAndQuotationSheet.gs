function newContractOrQuotation() {
  var spreadsheet = SpreadsheetApp.openById(FUF.contractSsID[env]);
  var sheet = spreadsheet.getActiveSheet();
  var distinguishId = sheet.getRange(distinguishCell).getValue();
  var corq;

  if (distinguishId.toString().includes('契約書')) {
    corq = 'contract';
  } else {
    corq = 'quotation';
  }

  createPDF(corq, sheet);
}

function createPDF(corq, sheet) {
  var id = sheet.getRange(idCoordinate).getValue();
  if (corq == 'contract') {
    var ui = SpreadsheetApp.getUi();
    var sheetName = sheet.getSheetName()
    checkValue = sheet.getRange(mtgDateCoordinate).getValue();
    if (checkValue == []) {
      var response = ui.alert(sheetName + '面談予定日が入力されていません', '実行を継続しますか？', ui.ButtonSet.OK_CANCEL);
      if (response == 'CANCEL') {
        Browser.msgBox('契約書の発行を中止しました。');
        return;
      }
    }
    var spreadsheet = SpreadsheetApp.openById(FUF.contractSsID[env]);
    spreadsheet.duplicateActiveSheet();//契約書シートをコピー
    var sheet = spreadsheet.getActiveSheet().setName(id);  //名前を契約書IDに設定する。
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); //値のみに変換する。
    createContractPDF(sheet);
  }
  
}

function newContract(sheet) {
  var ui = SpreadsheetApp.getUi();
  var sheetName = sheet.getSheetName()
  checkValue = sheet.getRange(mtgDateCoordinate).getValue();
  if (checkValue == []) {
    var response = ui.alert(sheetName + '面談予定日が入力されていません', '実行を継続しますか？', ui.ButtonSet.OK_CANCEL);
    if (response == 'OK') createContractPDF(contract);
    else if (response == 'CANCEL') Browser.msgBox('契約書の発行を中止しました。');
  } else {
  }
}

function newQuotation() {
  createQuotationPDF();
}

function createContractPDF(sheet) {
  var spreadsheet = SpreadsheetApp.openById(FUF.contractSsID[env]);
  var contractID = spreadsheet.getRange(contractIdCoordinate).getValue(); //契約書ID。
  var studentId = sheet.getRange(studentIdCoordinate).getValue(); //学習者 ID。
  var studentName = sheet.getRange(studentNameCoordinate).getValue(); //生徒名。

  contractPdfUrl = contractPdfSave(sheet.getSheetId().toString(), studentName, studentId, contractID); //契約書PDFを作成+格納。
  var js = '<a href="' + contractPdfUrl + '" target="_blank" rel="noopener noreferrer">契約書PDFを確認する。</a>';

  var contractFormatSheet = spreadsheet.getSheetByName(corqGenerateSheetName["contract"]);
  var contractUneditableSheet = spreadsheet.getSheetByName(contractUneditableSheetName);
  var previousContractIdRange = contractFormatSheet.getRange(contractIdCoordinate);
  var newContractId = previousContractIdRange.getValue() + 1;

  //ダイアログを表示
  var html = HtmlService.createHtmlOutput(js)
    .setHeight(100)
    .setWidth(200)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  SpreadsheetApp.getUi().showModalDialog(html, "契約書PDFを作成しました。");

  for (let range of unprotect) {
    var format = contractUneditableSheet.getRange(range);
    format.copyTo(contractFormatSheet.getRange(format.getA1Notation()), { contentsOnly: false });
  }
  previousContractIdRange.setValue(newContractId);
};

function createQuotationPDF() {
  var spreadsheet = SpreadsheetApp.openById(FUF.contractSsID[env]).getSheetByName(quotationFormatSheetName);
  var quotationId = spreadsheet.getRange(quotationIdCoordinate).getValue(); //見積書ID
  var studentId = spreadsheet.getRange(studentIdCoordinate).getValue(); //学習者 ID。
  var studentName = spreadsheet.getRange(studentNameCoordinate).getValue(); //生徒名。

  quotationPdfUrl = quotationPdfSave(spreadsheet.getSheetId().toString(), studentName, studentId, quotationId);
  var js = '<a href="' + quotationPdfUrl + '" target="_blank" rel="noopener noreferrer">見積書PDFを確認する。</a>';


  //ダイアログを表示
  var html = HtmlService.createHtmlOutput(js)
    .setHeight(100)
    .setWidth(200)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  SpreadsheetApp.getUi().showModalDialog(html, "見積書PDFを作成しました。");

  var newQuotationId = quotationId + 1;
  // var quotationFormatSheet = spreadsheet.getSheetByName(quotationFormatSheetName);
  // var quotationUneditableSheet = spreadsheet.getSheetByName(quotationUneditableSheetName);

  // var ranges = ['B12:J27', 'B30:G30','C33:F38', 'N10:Z10', 'N12:S17'];

  // for (let range of ranges) {
  //   var format = quotationUneditableSheet.getRange(range);
  //   format.copyTo(quotationFormatSheet.getRange(format.getA1Notation()), { contentsOnly: false });
  // }
  spreadsheet.getRange(quotationIdCoordinate).setValue(newQuotationId);
};

//参照サイト１：https://stackoverflow.com/questions/56853670/google-sheets-export-to-pdf-script-why-is-it-exporting-hidden-rows/56856418
//参照サイト２：https://kandiral.ru/googlescript/eksport_tablic_google_sheets_v_pdf_fajl.html
//契約書格納フォルダ直下の指定した生徒名のフォルダに、契約書を格納する。
function contractPdfSave(source, studentName, studentId, contractID) {
  var corq = "contract"
  //PDF発行用URL
  var pdfURL = "https://docs.google.com/spreadsheets/d/" + FUF.contractSsID[env] + "/pdf?esid=" + (Math.round(Math.random() * 10000000)) + "&id=" + FUF.contractSsID[env];

  //PDF出力の際の設定
  var dt = new Date();
  var d = encodeDate(dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds());
  var pc = customPages(source, d, corq); //ここがカスタム改ページの指定場所

  var options = {
    'method': 'post',
    'payload': "a=true&pc=" + JSON.stringify(pc) + "&gf=[]",
    'headers': { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, //アクセストークンの取得。
    'muteHttpExceptions': true
  };

  //pdfを生成
  var pdf = UrlFetchApp.fetch(pdfURL, options).getBlob().setName(`SSS Education 契約書 - ` + studentId + ` - ` + studentName + `様 - ` + contractID + '.pdf');

  const contractFolders = DriveApp.getFolderById(FUF.contractFolderRootId[env]); //契約書格納フォルダ
  var studentFolderTitle = `${studentId} - ${studentName}様 - SSS Education 契約情報`;
  let studentFolderIterator = contractFolders.getFoldersByName(studentFolderTitle); //生徒名フォルダ一覧

  //契約書格納フォルダに生徒名フォルダがなかったら作成する。
  let studentFolder;
  let contractFolder; //契約書
  let quotationFolder; //見積書
  let receiptFolder; //領収書
  if (studentFolderIterator.hasNext()) {
    studentFolder = studentFolderIterator.next();
    contractFolder = studentFolder.getFoldersByName(`${studentId} - ${studentName}様 - 契約書`).next();
    quotationFolder = studentFolder.getFoldersByName(`${studentId} - ${studentName}様 - 見積書`).next();
    receiptFolder = studentFolder.getFoldersByName(`${studentId} - ${studentName}様 - 領収書`).next();
  } else {
    studentFolder = contractFolders.createFolder(studentFolderTitle);
    contractFolder = studentFolder.createFolder(`${studentId} - ${studentName}様 - 契約書`);
    quotationFolder = studentFolder.createFolder(`${studentId} - ${studentName}様 - 見積書`);
    receiptFolder = studentFolder.createFolder(`${studentId} - ${studentName}様 - 領収書`);
  }

  //契約書PDFを契約書フォルダ内に作成
  var contractPdf = contractFolder.createFile(pdf);
  return contractPdf.getUrl()//契約書ファイルIDを返す
}

function quotationPdfSave(source, studentName, studentId, quotationId) {
  var corq = "quotation"
  //PDF発行用URL
  var pdfURL = "https://docs.google.com/spreadsheets/d/" + FUF.contractSsID[env] + "/pdf?esid=" + (Math.round(Math.random() * 10000000)) + "&id=" + FUF.contractSsID[env];

  //PDF出力の際の設定
  var dt = new Date();
  var d = encodeDate(dt.getFullYear(), dt.getMonth(), dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds());
  var pc = customPages(source, d, corq); //ここがカスタム改ページの指定場所

  var options = {
    'method': 'post',
    'payload': "a=true&pc=" + JSON.stringify(pc) + "&gf=[]",
    'headers': { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, //アクセストークンの取得。
    'muteHttpExceptions': true
  };

  //pdfを生成
  var pdf = UrlFetchApp.fetch(pdfURL, options).getBlob().setName(`SSS Education 見積書 - ` + studentId + ` - ` + studentName + `様 - ` + quotationId + '.pdf');

  const contractFolders = DriveApp.getFolderById(FUF.contractFolderRootId[env]); //契約書格納フォルダ
  var studentFolderTitle = `${studentId} - ${studentName}様 - SSS Education 契約情報`;
  let studentFolderIterator = contractFolders.getFoldersByName(studentFolderTitle); //生徒名フォルダ一覧

  //契約書格納フォルダに生徒名フォルダがなかったら作成する。
  let studentFolder;
  let contractFolder; //契約書
  let quotationFolder; //見積書
  let receiptFolder; //領収書
  if (studentFolderIterator.hasNext()) {
    studentFolder = studentFolderIterator.next();
    contractFolder = studentFolder.getFoldersByName(`${studentId} - ${studentName}様 - 契約書`).next();
    quotationFolder = studentFolder.getFoldersByName(`${studentId} - ${studentName}様 - 見積書`).next();
    receiptFolder = studentFolder.getFoldersByName(`${studentId} - ${studentName}様 - 領収書`).next();
  } else {
    studentFolder = contractFolders.createFolder(studentFolderTitle);
    contractFolder = studentFolder.createFolder(`${studentId} - ${studentName}様 - 契約書`);
    quotationFolder = studentFolder.createFolder(`${studentId} - ${studentName}様 - 見積書`);
    receiptFolder = studentFolder.createFolder(`${studentId} - ${studentName}様 - 領収書`);
  }

  //契約書PDFを契約書フォルダ内に作成
  var quotationPdf = quotationFolder.createFile(pdf);
  return quotationPdf.getUrl()//契約書ファイルIDを返す
}

//日付をエンコードして生成する
function encodeDate(yy, mm, dd, hh, ii, ss) {
  var days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (((yy % 4) == 0) && (((yy % 100) != 0) || ((yy % 400) == 0))) days[1] = 29;
  for (var i = 0; i < mm; i++)dd += days[i];
  yy--;
  return ((((yy * 365 + ((yy - (yy % 4)) / 4) - ((yy - (yy % 100)) / 100) + ((yy - (yy % 400)) / 400) + dd - 693594) * 24 + hh) * 60 + ii) * 60 + ss) / 86400.0;
}
