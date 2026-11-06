// フォームの回答に応じ、Google Chat 用の通知を作成し、カレンダーの予定を作成する関数
function studentIdPwGenReception(e) {
  var answerObject = formAnswer2Object(e) //フォームの回答をオブジェクト形式に整える。
  console.log(answerObject);
  var [studentId, studentName, parentName] = studentIdPwGenerator(e, answerObject); //学習者ID, Pass の生成。学習者 Id のみを出力する。
  console.log(studentId)
  var [studentFolderUrl, studentDocUrl] = generatePersonalFiles(studentId, studentName, parentName); //生徒個人用のフォルダを作成。指導情報の Id と受験戦略ドキュメントのURL を出力する。
  chatNotification(e, studentDocUrl); //Google Chat に通知を送る。
};

function studentIdPwGenerator(e) {
  var ss = SpreadsheetApp.openById(FUF.CustomerInfoSsId[env]); //学習者・決済者情報
  var sheet = FUF.ReadSheetById(ss, FUF.ContactAnsSheetTd[env]);
  var numOfColumns = sheet.getMaxColumns()

  // 学習者 ID を生成する。
  // スプレッドシートから最大値を取得し、+1する
  var lastRow = sheet.getLastRow();
  var maxValue = 0;

  var twoDigitsYear = parseFloat(new Date().getFullYear() % 100); //西暦年の下２桁の数字を取り出す。
  var start = twoDigitsYear * 10000; //2023年なら 230000
  var end = twoDigitsYear * 10000 + 10000; //2023年なら 240000

  if (lastRow > 1) { //最大値の抽出
    var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    values.forEach(function (row) {
      if (row[0] > start && row[0] < end) {
        maxValue = Math.max(maxValue, row[0]);
      }
    });
  }
  var studentId = extractLastFourDigits(maxValue) + start + 1; //最大値（4桁）に230001を加える

  // パスワードを生成
  var password = generatePassword();


  var columnTitles = sheet.getRange(1, 3, 1, numOfColumns - 2).getValues()[0];
  var newRow = [studentId, password];
  var answerObject = formAnswer2Object(e)

  for (i = 0; i < columnTitles.length; i++) {
    var newElement = answerObject[columnTitles[i]]
    if (newElement) {
      newRow.push(newElement)
    }
    else {
      newRow.push("")
    }
  }

  // スプレッドシートに記録
  sheet.appendRow(newRow);

  var itemResponses = e.response.getItemResponses();
  var formResponses = FormResponses2Array(itemResponses); //回答を配列に変換する。
  var studentEmail = extractEmailsFromArray(formResponses);
  var studentName = formResponses[0] //１番目の回答を学習者氏名とする。
  var parentName = formResponses[1] //２番目の回答を決済者氏名とする。

  // 生徒に学習者IDとPassをメールで連絡する。
  mailToStudent(studentName, parentName, studentEmail, studentId, password)
  return [studentId, studentName, parentName];
}