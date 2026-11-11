var file_name = "お問合せフォーム";

// フォームの回答に応じ、Google Chat 用の通知を作成し、カレンダーの予定を作成する関数
function contactFormAnswerReception(e) {

  // フォームの回答をオブジェクト形式に整える
  var answerObject = formAnswer2Object(e)
  console.log(answerObject);

  // 学習者ID, PW の生成し、学習者・決裁者情報DBに転記
  var [studentName, parentName, studentEmail, studentId, password] = generateContactInfo(e, answerObject);

  // 生徒個人用のフォルダを作成
  // 指導情報フォルダと受験戦略ドキュメントのURL を出力する
  var [contractInfoFolderUrl, studentFolderUrl, studentDocUrl, onlinematerialUrl] = generatePersonalFiles(studentId, studentName, parentName, studentEmail);

  if(answerObject['ご希望に関する情報'] == '資料請求がしたい'){
    sendOnlinematerialToStudent(studentName, parentName, studentEmail, studentId, password, onlinematerialUrl)
  }else{
   // 学習者ID, PW をメールで通知
  mailToStudent(studentName, parentName, studentEmail, studentId, password, onlinematerialUrl)}

  // 面談カレンダーのイベントを作成し招待する。 
  addCalendarEvents(studentId, studentFolderUrl, answerObject);

  // Google Chat に通知を送信
  chatNotification(e, contractInfoFolderUrl, studentFolderUrl, studentDocUrl);

};


// refactored on 2024/11/26
// フォームの回答に応じ、学習者ID, PW を作成し、スプレッドシートに生徒情報を記録する関数
function generateContactInfo(e, answerObject) {

  // 必要な情報の取得
  var contactDbSs = SpreadsheetApp.openById(FUF.CustomerInfoSsId[env]); //学習者・決済者情報DBスプレッドシート
  var contactFormAnsSheet = FUF.ReadSheetById(contactDbSs, FUF.ContactAnsSheetTd[env]); //「お問合わせ」シート
  var lastRow = contactFormAnsSheet.getLastRow(); // 行数
  var numOfColumns = contactFormAnsSheet.getMaxColumns(); // 列数

  // 学習者 ID を生成
  // 既存の学習者IDの最大値を取得し、+1する
  var twoDigitsYear = parseFloat(new Date().getFullYear() % 100); //西暦年の下２桁の数字を取り出す。
  var start = twoDigitsYear * 10000; //2023年なら 230000
  var end = (twoDigitsYear + 1) * 10000; //2023年なら 240000

  if (lastRow >= 1) {
    var studentIds = FUF.transpose(contactFormAnsSheet.getRange(2, 1, lastRow - 1, 1).getValues())[0]; //学習者ID
    var validStudentIds = studentIds.filter(function (id) { return id >= start && id < end }); //有効な学習者ID
    if (validStudentIds.length >= 1) {
      var studentId = extractLastFourDigits(Math.max(...validStudentIds)) + start + 1; //最大値（4桁）に230001を加える
    }
    else {
      var studentId = start;
    }
  } else {
    var studentId = start;
  }

  // 学習者PWを生成
  var password = generatePassword();

  // 新たに挿入する行（お問合わせ情報）を作成
  var newRow = [studentId, password];
  var columnTitles = contactFormAnsSheet.getRange(1, 3, 1, numOfColumns - 2).getValues()[0];
  for (i = 0; i < columnTitles.length; i++) {
    var newElement = answerObject[columnTitles[i]]
    if (newElement) {
      newRow.push(newElement)
    }
    else {
      newRow.push("")
    }
  }

  // シートにお問合わせ情報を挿入
  contactFormAnsSheet.appendRow(newRow);

  // 生徒に学習者IDとPassをメールで連絡する。
  var studentEmail = answerObject[mailQuestionTitle];
  var studentName = answerObject[studentNameQuestionTitle];
  var parentName = answerObject[parentNameQuestionTitle];

  return [studentName, parentName, studentEmail, studentId, password];
}


// refactored on 2024/11/26
// ランダムなパスワードを生成する関数
function generatePassword() {
  var charset = "acdefghjkmnpqrtwxyzABCDEFGHJKLMNPRTWXYZ234789";
  //oO0, 1liI, uvUV, sS5, b6 は間違えやすいので除外
  var password = "";
  for (var i = 0, n = charset.length; i < passwordLength; ++i) {
    password += charset.charAt(Math.floor(Math.random() * n));
  }
  return password;
}


// refactored on 2024/11/26
// 申込者に学習者IDとパスワードをメールで送信する
function mailToStudent(studentName, parentName, studentEmail, studentId, password, onlinematerialUrl) {

  // 件名
  var subject = `【SSS Education】学習者IDとパスワード(PW)のお知らせ（送信専用）`;

  // 宛名を生成
  var atena = `${studentName}様　`;
  if (studentName != parentName) { atena += `${parentName}様` }
  atena += `<br>`

  var body = `（※当メールアドレスは送信専用です。）<br>`
    + atena
    + `初めまして。難関大・医学部対策予備校 SSS Education です。<br>`
    + `<br>`
    + `この度はお問合せいただきましてありがとうございます。<br>`
    + `後ほど、担当の者から改めてご連絡いたしますので少々お待ち下さいませ。<br>`
    + `下記に、もし入塾される際に必要となる ID, パスワードを共有いたします。（請求情報などではありませんので、ご安心ください。）<br>`
    + `<br>`
    + `${studentName}様の学習者 ID , パスワードは以下の通りです。<br>`
    + `<br>`
    + `学習者ID：<strong>${studentId}</strong><br>`
    + `学習者PW：<strong>${password}</strong><br>`
    + `<br>`
    + `<strong>ID, PW はもし入塾される際に必要となりますので、お手元に保管していただくようにお願いいたします。</strong><br>`
    + `<br>`
    + `では、担当者からの追っての連絡を今しばらくお待ちくださいませ。<br>`
    + `なお、当メールアドレスは送信専用となります。返信等は公式ＬＩＮＥ( https://lin.ee/CbmNbRP )までお送りください。<br>`
    + `引き続きよろしくお願いいたします。<br>`
    + `<br>`
    + `Exceed Your Potential. <br>`
    + `株式会社 SSS Education`

  var options = {
    from: "info@sss-edu.org", // 送信元のアドレス
    htmlBody: body
  };

  try {
    GmailApp.sendEmail(studentEmail, subject, body, options);
  } catch (e) {
    console.log(`エラー内容:` + e.message);
  }
}

function sendOnlinematerialToStudent(studentName, parentName, studentEmail, studentId, password, onlinematerialUrl) {

  // 件名
  var subject = `【SSS Education】資料と学習者ID・パスワード(PW)のお知らせ（送信専用）`;

  // 宛名を生成
  var atena = `${studentName}様　`;
  if (studentName != parentName) { atena += `${parentName}様` }
  atena += `<br>`

  var body = `（※当メールアドレスは送信専用です。）<br>`
    + atena
    + `初めまして。難関大・医学部対策予備校 SSS Education です。<br>`
    + `<br>`
    + `この度はお問合せいただきましてありがとうございます。<br>`
    + `<br>`
    + `当塾のオンライン資料は以下の URL からご覧いただけます。<br>`
    + `${onlinematerialUrl}`
    + `<br>`
    + `その他の具体的なご相談は、公式ライン( https://lin.ee/CbmNbRP )にメッセージをお送りください。<br>`
    + `<br>`
    + `下記に、もし入塾される際に必要となる ID, パスワードを共有いたします。（請求情報などではありませんので、ご安心ください。）<br>`
    + `<br>`
    + `${studentName}様の学習者 ID , パスワードは以下の通りです。<br>`
    + `<br>`
    + `学習者ID：<strong>${studentId}</strong><br>`
    + `学習者PW：<strong>${password}</strong><br>`
    + `<br>`
    + `<strong>ID, PW はもし入塾される際に必要となりますので、お手元に保管していただくようにお願いいたします。</strong><br>`
    + `<br>`
    + `では、担当者からの追っての連絡を今しばらくお待ちくださいませ。<br>`
    + `引き続きよろしくお願いいたします。<br>`
    + `<br>`
    + `Exceed Your Potential. <br>`
    + `株式会社 SSS Education`

  var options = {
    from: "info@sss-edu.org", // 送信元のアドレス
    htmlBody: body
  };

  try {
    GmailApp.sendEmail(studentEmail, subject, body, options);
  } catch (e) {
    console.log(`エラー内容:` + e.message);
  }
}

// refactored on 2024/11/27
// 生徒用の指導情報フォルダ + 契約情報フォルダを作成する
function generatePersonalFiles(studentId, studentName, parentName, studentEmail) {

  // 契約情報フォルダを作成する
  FUF.checkAndSyncTemplateFolderDynamic(tempOfContractFolder.id, FUF.contractFolderRootId[env], tempFolderTags.slice(0, 4), [studentId, studentName, parentName, studentEmail])

  // オンライン資料のpdfを作成
  // オンライン資料のpdfを各々のフォルダに保存

  // 指導情報フォルダを作成する
  FUF.checkAndSyncTemplateFolderDynamic(tempOfGuidanceFolder.id, FUF.studentGuidanceFolderRootId[env], tempFolderTags.slice(0, 4), [studentId, studentName, parentName, studentEmail])

  // 受験戦略フォルダにサービス取得状況のリンクを添付する。
  var serviceFileUrl = DriveApp.getFilesByName(`${studentId} - ${studentName}様 - サービス取得状況`).next().getUrl();
  var strategyFile = DriveApp.getFilesByName(`受験戦略 - ${studentId} - ${studentName}様`).next();
  var strategyDoc = DocumentApp.openById(strategyFile.getId());
  var urlElement = strategyDoc.getBody().findText(`<指導情報URL>`);
  if (urlElement) {
    var startOffset = urlElement.getStartOffset();
    var endOffset = urlElement.getEndOffsetInclusive();
    var element = urlElement.getElement();
    element.asText().setLinkUrl(startOffset, endOffset, serviceFileUrl);
  } else {
    console.error('指定されたURLが文書内に見つかりませんでした。');
  }
  FUF.replaceTagsInGoogleDocs(strategyFile, [`<指導情報URL>`], [serviceFileUrl]);

  var contractInfoFolder = DriveApp.getFoldersByName(`${studentId} - ${studentName}様 - SSS Education 契約情報`).next()
  var onlinematerialSlide = DriveApp.getFilesByName(`${studentId} - ${studentName}様 - オンライン資料`).next();
  var blob = onlinematerialSlide.getAs('application/pdf');
  // PDFをフォルダに保存
  contractInfoFolder.createFile(blob).setName(onlinematerialSlide.getName() + '.pdf');
  console.log(`${studentId} - ${studentName}様 - オンライン資料.pdf を作成しました。`)
  var guidanceInfoFolderUrl = DriveApp.getFoldersByName(`${studentId} - ${studentName}様 - SSS Education 指導情報`).next().getUrl();
  var onlinematerial = DriveApp.getFilesByName(`${studentId} - ${studentName}様 - オンライン資料.pdf`).next();
  onlinematerial.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  //指導情報フォルダのIDと受験戦略ファイルのurlを返す
  return [contractInfoFolder.getUrl(), guidanceInfoFolderUrl, strategyFile.getUrl(), onlinematerial.getUrl()];
}

// refactored on 2024/11/27
// 面談カレンダーのイベントを作成し招待する
function addCalendarEvents(studentId, studentFolderUrl, answerObject) {

  // フォームの回答から必要な情報を読み出す
  var mail = answerObject[mailQuestionTitle]; //メール
  var studentName = answerObject[studentNameQuestionTitle]; //学習者氏名
  var meetingFormat = answerObject[mtgFormatQuestionTitle]; //面談形式

  // オンライン・対面の場合で場合分け
  if (meetingFormat == "オンライン") {
    for (var i = 0; i < onlineReceptionQuestionTitles.length; i++) {
      var mtgTime = answerObject[onlineReceptionQuestionTitles[i]];
      createCalendarEvent(studentId, studentName, mail, mtgTime, meetingFormat, studentFolderUrl);
    }
  } else if (meetingFormat == "対面") {
    for (var i = 0; i < onsiteReceptionQuestionTitles.length; i++) {
      var mtgTime = answerObject[onsiteReceptionQuestionTitles[i]];
      createCalendarEvent(studentId, studentName, mail, mtgTime, meetingFormat, studentFolderUrl);
    }
  } else { ; }

}

// refactored on 2024/11/27
// カレンダーを作成し招待する。
function createCalendarEvent(studentId, studentName, mail, dateTime, meetingFormat, urls) {

  let adminCalendar = CalendarApp.getCalendarById('admin@sss-edu.org');
  if (env == "test") { var testalert = "テスト・" } else { var testalert = "" }
  const title = `【仮・${testalert}${meetingFormat}】${studentName}様ご面談予約(学習者ID:${studentId})`;  //予定のタイトル
  const startTime = new Date(dateTime); //開始時間
  const endTime = new Date(startTime.getTime() + (salesInterviewerHours * 60 * 60 * 1000)); //終了時間

  var attendants = mail // 参加者一覧

  if (meetingFormat == "対面") {
    for (var i = 0; i < interviewRooms.length; i++) {
      if (FUF.isAvailable(startTime, endTime, interviewRooms[i])) {
        if (attendants != ``) { attendants += `, ` }
        attendants += `${interviewRooms[i]["mail"]}`;
      }
    }
    //　対応可能な契約担当者を列挙する。
    for (var i = 0; i < salesInterviewers.length; i++) {
      if (FUF.isAvailable(startTime, endTime, salesInterviewers[i])) {
        if (attendants != ``) { attendants += `, ` }
        attendants += `${salesInterviewers[i]["mail"]}`;
      }
    }
    //　対応可能な教務担当者を列挙する。
    for (var i = 0; i < onsiteAcademicInterviwers.length; i++) {
      if (FUF.isAvailable(startTime, endTime, onsiteAcademicInterviwers[i])) {
        if (attendants != ``) { attendants += `, ` }
        attendants += `${onsiteAcademicInterviwers[i]["mail"]}`;
      }
    }
  }


  if (meetingFormat == "オンライン") {
    for (var i = 0; i < interviewRooms.length; i++) {
      if (FUF.isAvailable(startTime, endTime, interviewRooms[i])) {
        if (attendants != ``) { attendants += `, ` }
        attendants += `${interviewRooms[i]["mail"]}`;
      }
    }
    //　対応可能な契約担当者を列挙する。
    for (var i = 0; i < salesInterviewers.length; i++) {
      if (FUF.isAvailable(startTime, endTime, salesInterviewers[i])) {
        if (attendants != ``) { attendants += `, ` }
        attendants += `${salesInterviewers[i]["mail"]}`;
      }
    }
    //　対応可能な教務担当者を列挙する。
    for (var i = 0; i < onlineAcademicInterviewers.length; i++) {
      if (FUF.isAvailable(startTime, endTime, onlineAcademicInterviewers[i])) {
        if (attendants != ``) { attendants += `, ` }
        attendants += `${onlineAcademicInterviewers[i]["mail"]}`;
      }
    }
  }

  //オプションを付けて送信
  const options = {
    guests: attendants,
    sendInvites: true,
    description: "指導情報フォルダ URL：" + urls
  }

  //Googleカレンダーに予定を登録
  var calendarEvent = adminCalendar.createEvent(title, startTime, endTime, options);
  calendarEvent.setGuestsCanModify(true)
}

// refactored on 2024/11/27
// Google Chat にお問合せ内容と関連フォルダURLを送信する
function chatNotification(e, contractInfoFolderUrl, studentFolderUrl, studentDocUrl) {

  // 引数 e はフォームから送信される情報が格納されている。
  // Googleフォームの投稿内容をitemResponsesへ入れる
  var itemResponses = e.response.getItemResponses();

  // Googleフォームの質問項目名と回答を順番に変数textへ入れる
  var text = `<users/${salesPerson["chatId"]}>, <users/${salesPerson1["chatId"]}> \n ${salesPerson["nickname"]}様， ${salesPerson1["nickname"]}様,新規お問合せがありましたよ！！`;
  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    var question = itemResponse.getItem().getTitle();
    var answer = FormResponse2String(itemResponse);
    text += `\n${question} : ${answer}`;
  }
  text += `\n<${contractInfoFolderUrl}|契約情報>`
  text += `\n<${studentFolderUrl}|指導情報>`
  text += `\n<${studentDocUrl}|受験戦略>`

  // Googleチャットのスペースで取得した着信WebhookのURLを設定。
  var post_url = FUF.ContactFormWebhook[env];

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
  UrlFetchApp.fetch(post_url, options);
}

// Google フォームに営業担当者と教務担当者が１名ずつ面談可能な時間をプルダウン形式で出力する関数
function putAvailableTimesToForm() {
  var today = new Date();
  var onlineAvailableTimes = [];
  var onsiteAvailableTimes = [];

  //面談可能日時を抽出する。
  for (let i = 2; i <= receptionDays; i++) {
    var start = new Date();
    start.setDate(today.getDate() + i);
    var interviewRoomFlags = [];
    var salesInterviewerFlags = [];
    var onsiteAcademicInterviewerFlags = [];
    var onlineAcademicInterviewerFlags = [];

    for (let j = 0; j < interviewReceptionTimes.length; j++) {
      start.setHours(interviewReceptionTimes[j][0], interviewReceptionTimes[j][1], 01);
      var end = new Date(start.getTime() + (salesInterviewerHours * 60 * 60 * 1000) - 2000);

      // 面談室の予定おさえる
      var flags = [];
      for (let i = 0; i < interviewRooms.length; i++) {
        flags = flags.concat(FUF.isAvailable(start, end, interviewRooms[i]));
      }
      interviewRoomFlags = interviewRoomFlags.concat(flags.some(function (element) { return element === true }))

      // 営業担当者の予定おさえる。
      var flags = [];
      for (let i = 0; i < salesInterviewers.length; i++) {
        flags = flags.concat(FUF.isAvailable(start, end, salesInterviewers[i]));
      }
      salesInterviewerFlags = salesInterviewerFlags.concat(flags.some(function (element) { return element === true }))

      // オンサイトでの面談担当者の予定おさえる。
      var flags = [];
      for (let i = 0; i < onsiteAcademicInterviwers.length; i++) {
        flags = flags.concat(FUF.isAvailable(start, end, onsiteAcademicInterviwers[i]))
      }
      onsiteAcademicInterviewerFlags = onsiteAcademicInterviewerFlags.concat(flags.some(function (element) { return element === true }));

      // オンラインでの面談担当者の予定押さえる。
      var flags = [];
      for (let i = 0; i < onlineAcademicInterviewers.length; i++) {
        flags = flags.concat(FUF.isAvailable(start, end, onlineAcademicInterviewers[i]))
      }
      onlineAcademicInterviewerFlags = onlineAcademicInterviewerFlags.concat(flags.some(function (element) { return element === true }));


    }

    // onlineAvailableTimes の可否を、契約・教務担当が両方可能な日におさえる日を true, どちらかが不可な日を false で示す
    onlineAvailableFlags = FUF.andByIndex(salesInterviewerFlags, onlineAcademicInterviewerFlags);
    onsiteAvailableFlags = FUF.andByIndex(interviewRoomFlags, salesInterviewerFlags, onsiteAcademicInterviewerFlags);

    // 可能な日時を配列形式で出力する
    onlineAvailableTimes = onlineAvailableTimes.concat(getAvailableTime(onlineAvailableFlags, start));
    onsiteAvailableTimes = onsiteAvailableTimes.concat(getAvailableTime(onsiteAvailableFlags, start))
  }

  //フォームにプルダウンを作成する。
  refreshFormPulldowns(FUF.contactFormId[env], onlineReceptionQuestionTitles, onlineAvailableTimes);
  refreshFormPulldowns(FUF.contactFormId[env], onsiteReceptionQuestionTitles, onsiteAvailableTimes);

}


// １日の面談受付時間のうち、配列のtrueの位置に対応する時間を配列形式で抽出する。
function getAvailableTime(availableFlags, date) {
  var availableTimes = [];

  // flags の true の時間を対応する日時に変換。
  for (let j = 0; j < interviewReceptionTimes.length; j++) {
    if (availableFlags[j] == true) {
      availableTimes = availableTimes.concat(new Date(date.setHours(interviewReceptionTimes[j][0], interviewReceptionTimes[j][1], 00)));
    }
  }
  // 日付形式を "yyyy/mm/dd HH:mm" 文字列に変換。
  for (let k = 0; k < availableTimes.length; k++) {
    availableTimes[k] = Utilities.formatDate(availableTimes[k], "JST", "yyyy/MM/dd HH:mm");
  }
  return availableTimes;
}

// 指定のフォームID, 質問項目に対し、プルダウンの選択肢を設定する関数。
function refreshFormPulldowns(formId, itemNamesArray, choicesArray) {
  const form = FormApp.openById(formId);
  const items = form.getItems();

  for (let i = 0; i < items.length; i++) {
    const item = items[i]
    const itemName = item.getTitle()

    if (itemNamesArray.some(function (element) { return element === itemName })) {
      items[i].asListItem().setChoiceValues(choicesArray)
    }
  }
}
