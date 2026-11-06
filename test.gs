function slideToPdf(presentationId, folderId) {

  // プレゼンテーションIDからプレゼンテーションを取得
  var presentation = DriveApp.getFileById("1b5rMvbe4ToPW46Rg1Z8lgjFO9nc9cH07rccp4NnMCvo");

  // PDFに変換
  var blob = presentation.getAs('application/pdf');

  // フォルダIDからフォルダを取得
  var folder = DriveApp.getFolderById("1lQ3hDXAJBnj-xqMvGU8jOBdTP0wsdaqD");

  // PDFをフォルダに保存
  folder.createFile(blob).setName(presentation.getName() + '.pdf');
}

function dummy()
{//契約書格納フォルダに生徒名フォルダがなかったら作成する。
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
}}


//各人のオンライン資料を取得する
function dummy2(){
let onlinematerial; //オンライン資料
if (studentFolderIterator.hasNext()) {
  studentFolder = studentFolderIterator.next();
  onlinematerial = studentFolder.getFilesByName(`${studentId} - ${studentName}様 - オンライン資料`).next();
} else {
  const message = "該当するファイルがありません";
  console.log(message);
}}

//オンライン資料をpdfに変換する





function convertPresentationToPdf() {
  // プレゼンテーションのファイルIDを格納
  const presentationId = onlinematerial.getId(); // onlinematerial は Drive ファイルオブジェクト
  const presentationName = onlinematerial.getName(); // プレゼンテーション名

  // PDFエクスポート用のURLを作成
  const pdfExportUrl = `https://docs.google.com/presentation/d/${presentationId}/export/pdf`;

  // リクエスト設定
  const options = {
    method: "get",
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    }
  };

  // PDFデータを取得
  const response = UrlFetchApp.fetch(pdfExportUrl, options);

  // PDFをGoogleドライブに保存
  const pdfBlob = response.getBlob();
  pdfBlob.setName(`${presentationName}.pdf`);
  const pdfFile = DriveApp.createFile(pdfBlob);

  Logger.log(`PDF created: ${pdfFile.getName()} (ID: ${pdfFile.getId()})`);
}
