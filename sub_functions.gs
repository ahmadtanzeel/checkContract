function extractLastFourDigits(number) {
  if (number === 0) { return 0 }
  else {
    // 数値を文字列に変換
    var numberString = number.toString();

    // 文字列の長さが4未満の場合は、そのまま返す
    if (numberString.length < 4) {
      return numberString;
    }

    // 文字列の末尾4文字を抽出
    return parseFloat(numberString.substring(numberString.length - 4));
  }
}


// 配列の中で、メールアドレス形式の要素のみ抽出する関数。
function extractEmailsFromArray(array) {
  // メールアドレス形式をチェックするための正規表現
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  // filterメソッドを使用して、メールアドレス形式に適合する要素のみを抽出
  var emails = array.filter(function (item) {
    return emailRegex.test(item);
  });

  return emails[0];
}

function getMessageBasedOnTime() {
  var now = new Date();
  var hours = now.getHours();

  if (hours < 11) {
    return "おはようございます！";
  } else if (hours < 18) {
    return "こんにちは！";
  } else {
    return "こんばんは！";
  }
}

// フォームの回答に対し，，未回答でないものをオブジェクト形式で整理する関数。
function validFormResponses(e) {
  var itemResponses = e.getItemResponses();
  var responseObj = {};

  // 各質問項目に対する回答を処理
  itemResponses.forEach(function (itemResponse) {
    var question = itemResponse.getItem().getTitle();
    var answer = itemResponse.getResponse();

    if (answer && answer != "") {
      // Checkbox, Grid, CheckboxGrid の回答をCSV形式に変換
      if (Array.isArray(answer)) {
        if (answer.every(Array.isArray)) {
          // GridItemの場合、各サブリストをCSVに変換
          answer = answer.map(subAnswer => subAnswer.join(',')).join('\n');
        } else {
          // CheckboxItem の場合
          answer = answer.join(',');
        }
      }

      // ファイルアップロードの場合、IDとURLを結合
      if (itemResponse.getItem().getType() === FormApp.ItemType.FILE_UPLOAD) {

        var fileUrls = answer.map(function (fileId) {
          var file = DriveApp.getFileById(fileId);
          return fileId + ',' + file.getUrl();
        });
        answer = fileUrls.join('\n');
      }
      responseObj[question] = answer;
    }
  });

  return responseObj;
}

//フォームの回答を配列形式に変換する。ファイルのアップロードは URL を csv で列挙する。
function FormResponses2Array(itemResponses) {
  var responsesArray = [];

  // 各質問項目に対する回答を処理
  itemResponses.forEach(function (itemResponse) {
    var answer = itemResponse.getResponse();

    if (Array.isArray(answer)) {
      if (answer.every(Array.isArray)) {
        // GridItemの場合、各サブリストをCSVに変換
        answer = answer.map(subAnswer => subAnswer.join(',')).join('\n');
      } else {
        // CheckboxItem の場合
        answer = answer.join(',');
      }
    }

    // ファイルアップロードの場合、IDとURLを結合
    if (itemResponse.getItem().getType() === FormApp.ItemType.FILE_UPLOAD) {
      var fileIds = itemResponse.getResponse(); // ファイルのIDを取得
      var idAndUrls = ""
      // 各ファイルIDに対して処理
      for (var k = 0; k < fileIds.length; k++) {
        if (k > 0) { idAndUrls += "\n" }
        var file = DriveApp.getFileById(fileIds[k]);
        idAndUrls += fileIds[k] + ", " + file.getUrl();
      }
      answer = idAndUrls;
    }

    // var fileUrls = answer.map(function (fileId) {
    //   var file = DriveApp.getFileById(fileId);
    //   return fileId + ',' + file.getUrl();
    // });
    // answer = fileUrls.join('\n');

    responsesArray = responsesArray.concat(answer);
  });

  return responsesArray;
}

function FormResponse2String(itemResponse) {
  var answer = itemResponse.getResponse();
  if (Array.isArray(answer)) {
    if (answer.every(Array.isArray)) {
      // GridItemの場合、各サブリストをCSVに変換
      answer = answer.map(subAnswer => subAnswer.join(',')).join('\n');
    } else {
      // CheckboxItem の場合
      answer = answer.join(',');
    }
  }
  // ファイルアップロードの場合、IDとURLを結合
  if (itemResponse.getItem().getType() === FormApp.ItemType.FILE_UPLOAD) {
    var fileIds = itemResponse.getResponse(); // ファイルのIDを取得
    var idAndUrls = ""
    // 各ファイルIDに対して処理
    for (var k = 0; k < fileIds.length; k++) {
      if (k > 0) { idAndUrls += "\n" }
      var file = DriveApp.getFileById(fileIds[k]);
      idAndUrls += fileIds[k] + ", " + file.getUrl();
    }
    answer = idAndUrls;
  }
  return answer;
}


function formAnswer2Object(e) {
  // 引数 e はフォームから送信される情報が格納されている。
  // Googleフォームの投稿内容をitemResponsesへ入れる
  var itemResponses = e.response.getItemResponses();
  var answerObject = {};

  for (var i = 0; i < itemResponses.length; i++) {
    var key = itemResponses[i].getItem().getTitle();
    var value = FormResponse2String(itemResponses[i])
    answerObject[key] = value;
  }

  return answerObject;
}

function calculateMonthlyAmounts(totalAmount, startDate, endDate) {
  var start = new Date(startDate);
  var end = new Date(endDate);
  var datesList = [];
  var amountsList = [];
  var accumulatedAmount = 0;

  // 総利用日数を計算（ミリ秒での差を日数に変換）
  var totalDays = (end - start) / (1000 * 60 * 60 * 24) + 1;

  // 各月の利用日数を計算し、金額を分配
  var currentDate = new Date(start);
  while (currentDate <= end) {
    var year = currentDate.getFullYear();
    var month = currentDate.getMonth() + 1; // 月は0から始まるため、+1
    var nextMonth = new Date(year, month, 0); // 翌月の0日（前月の最終日）
    var lastDayOfMonth = nextMonth.getDate();

    // 月末またはサービス利用終了日のどちらか早い方
    var endDay = (year === end.getFullYear() && month === end.getMonth() + 1) ? end.getDate() : lastDayOfMonth;
    var daysInMonth = endDay - currentDate.getDate() + 1;

    // 月ごとの利用金額を計算し、整数に丸める
    var amountForMonth = Math.round((totalAmount / totalDays) * daysInMonth);
    accumulatedAmount += amountForMonth;

    // 月ごとの第1日をDate型でリストに追加
    datesList.push(new Date(year, month - 1, 1));
    amountsList.push(amountForMonth);

    // 次の月へ
    currentDate = new Date(year, month, 1);
  }

  // 合計金額との調整
  amountsList[amountsList.length - 1] += totalAmount - accumulatedAmount;

  return [datesList, amountsList];
}


// たまに、サービス欄の１行が「オンライン」だけのような、ミスのある契約書がある。
// 価格の欄が空白の場合は、その行を無視する関数をつくる。
function removeInvalidServices(data) {
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var service = data[i][contractServiceInfoRange[3]-1];
    if (service != "") {
      result.push(data[i]);
    }
  }
  return result;
}
