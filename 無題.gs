function tempPosts(){
  var list = [2000159];
  var contractDBSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]);
  for(num of list){ 
    tempPost(num)
    SpreadsheetApp.flush();
  }
}

function tempPost(num) {
  var id = String(num);
  console.log(id)
  // 入力された ID が 7 桁かどうかをチェック
  if (/^\d{7}$/.test(id)) {
    var contract = SpreadsheetApp.openById(FUF.contractSsID[env]);//契約書スプシ
    var contractSheet = contract.getSheetByName(id); //該当する契約書IDのシートを取得
    if (contractSheet === null) {
      console.log(`ID ${id} の契約書が存在しません。`);
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
          console.log(`「SSS Education 契約書 - ${studentId} - ${studentName}様 - ${id}.pdf」という名称の契約書pdfが複数あります．整理してやり直してください．`)
          return
        }
      } else {
        console.log(`「SSS Education 契約書 - ${studentId} - ${studentName}様 - ${id}.pdf」という名称の契約書pdfが見つかりません。`);
        return
      }
      postContractToDB(contractSheet, pdfUrl)
    }
  } else {
    console.log('無効なIDです。7桁の数字を入力してください。');
  }
}