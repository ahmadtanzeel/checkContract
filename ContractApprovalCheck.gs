// 契約書の承認に関するメールを検索し、スレッドを抽出する。
// メールの検索条件
// "subject:「SSS Education 契約書 .pdf」に対する承認リクエスト, from:approvals-noreply@google.com" これらに加え、時間帯で絞り込む。
// たとえ社員から転送されたメールでも、from:approvals-noreply@google.com で検索可能。

function mailSearch() {
  const interval = 5; //5分おきにトリガーを設定する

  // 現在時刻を取得
  var date = new Date();
  var formatnow = Utilities.formatDate(date, "JST", "yyyy-MM-dd HH:mm");
  var now = Date.parse(formatnow) / 1000; //ISO 8601 ミリ秒形式

  // １分前の時刻を取得
  var pastdate = new Date(date);
  pastdate.setMinutes(date.getMinutes() - interval);
  var formatpast = Utilities.formatDate(pastdate, "JST", "yyyy-MM-dd HH:mm");
  var past = Date.parse(formatpast) / 1000;

  // Gmail 受信トレイを取得
  var threads = GmailApp.search(`subject:「SSS Education 契約書 .pdf」に対する承認リクエスト from:approvals-noreply@google.com after: ${past} before:${now}`);

  // 契約書の承認に関するメールが届いていたら、contractData を実行する。
  if (threads != []) {
    contractData(threads);
  }

}


// スレッドの最新のメッセージに基づき、承認されたなら[[契約書Id, pdfUrl],...]を返す。
// 拒否・期限切れなら Google Chat に通知を送る。
function contractData(threads) {

  var _approvalList = [];
  
  for (i=0;i<threads.length; i++) {
    var thread = threads[i];
    var msgs = thread.getMessages();

    // 同スレッド内の契約書については、最新のものだけ用いる。
    var msg = msgs[msgs.length - 1];
    var contractID = msg.getSubject().match(/\d{7}/)[0]; //契約書 ID
    var studentId = msg.getSubject().match(/\d{6}/)[0]; //学習者 ID
    var body = msg.getPlainBody();
    var pdfUrl = body.match(/https:\/\/drive\.google\.com\/file\/d\/\S*/)[0];// string
    var salesStaffMail = msg.getTo();
    var laborerSs = SpreadsheetApp.openById(FUF.LaborerInfoSsId[env])
    var laborerSheet = FUF.ReadSheetById(laborerSs,FUF.LaborerInfoGId[env])
    var laborerLastRow = laborerSheet.getLastRow();
    var salesStaffName;

    for(let i=2; i <= laborerLastRow; i++){
      if(laborerSheet.getRange(i, 5).getValue() == salesStaffMail){
        salesStaffName = laborerSheet.getRange(i, 3).getValue();
        break;
      }
    }

    var salesStaffMention = `<users/${salesStaffName}>`

    //承認完了と、拒否または期限切れで処理を分岐する。
    //承認が得られた場合は[契約ID, 契約書pdfURL]を_approvalListに追加する。
    if (msg.getPlainBody().includes(`承認が完了しました`)) {

      _approvalList.push([contractID, pdfUrl, studentId]);
    }

    //承認が得られなかった場合。承認が下りなかった通知を Google Chat にへ送信する。
    else {
      var contract = SpreadsheetApp.openById(FUF.contractSsID[env]);//契約書スプシ
      var contractSheet = contract.getSheetByName(contractID); //該当する契約書IDのシートを取得
      var name = contractSheet.getRange(studentNameCoordinate).getValue(); //本人氏名
      var rejectionBody = `${salesStaffMention}様\n\n承認申請していた${name}様の契約書の承認がおりませんでした。\n再リクエストおよび顧客への承認依頼を行いましょう。\n送信状況は添付の pdf から確認してください。\n\n${pdfUrl}`

      var text = {
        'text': `${rejectionBody}`
      };

      var json = JSON.stringify(text);

      var options = {
        'method': 'POST',
        'headers': {
          'Content-type': 'application/json; charset=utf-8'
        },
        'payload': json,
      };

      UrlFetchApp.fetch(FUF.ApprovalCheckWebhook[env], options);
    }

  }

  postContractData(_approvalList, salesStaffMention);
}


function postContractData(_approvalList, salesStaffMention) {
  //_approvalList にデータが貯まっていたら、契約情報に情報を転記する。
  if (_approvalList != []) {
    for (element of _approvalList) {
      var contract = SpreadsheetApp.openById(FUF.contractSsID[env]);//契約書スプシ
      var contractSheet = contract.getSheetByName(element[0]); //該当する契約書IDのシートを取得
      var contractID = element[0];
      var studentId = element[2];
      var studentName = contractSheet.getRange(studentNameCoordinate).getValue(); //本人氏名
      var contractantName  = contractSheet.getRange(contractantNameCoordinate).getValue(); //決済者氏名 
      var paymentMethod = contractSheet.getRange(paymentMethodCoordinate).getValue(); //支払方法
      var conferenceSchedule = contractSheet.getRange(mtgDateCoordinate).getValue(); //面談予定日
      var paymentReccurence = contractSheet.getRange(paymentReccurenceCoordinate).getValue() //支払回数

      var services = contractSheet.getRange(contractServiceInfoRange[0], contractServiceInfoRange[1],contractServiceInfoRange[2],contractServiceInfoRange[3]).getValues()
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
        service[8] = _service[4]*1.10;
        service[9] = _service[5];
        service[10] = _service[6]*1.10;

        service[11] = paymentMethod; //支払方法
        service[12] = paymentReccurence; //支払回数
        service[13] = element[1]; //契約書 URL

        approvalList.push(service);
      }

      var ss = SpreadsheetApp.openById(FUF.contractDBSsId[env]);
      var contractDbSheet = FUF.ReadSheetById(ss,FUF.contractDBGId[env])
      var lastRow = contractDbSheet.getLastRow();

      contractDbSheet.getRange(lastRow + 1, 1, approvalList.length, 14).setValues(approvalList);

      postToChat(studentName, contractantName, element, salesStaffMention);
      postPaymentData(studentId, studentName, contractID, contractantName, contractSheet);
    }
  }
}

function postToChat(studentName, contractantName, element, salesStaffMention) {
  var approvalBody = `${salesStaffMention}様\n\n承認申請していた ${studentName} 様の契約書の承認が完了しました！\n\n承認が完了した契約の情報を契約情報に転記しました。\n承認が完了したpdf は以下の通りです。\n\n${studentName}様（決済者：${contractantName}様, 契約書ID: ${element[0]}, 学習者ID:${element[3]}）\n${element[1]}\n\n契約情報一覧：${FUF.contractDBSsURL[env]}\n 個別指導・英作文パーソナルトレーニング・集団授業（映像・LIVE）: <users/117258935252382614384> \n 自習室：<users/111552729343943889615> \n オンライン学習サポート・その他：<users/106909858067625788719>`

  var text = {
    'text': `${approvalBody}`
  };

  var json = JSON.stringify(text);

  var options = {
    'method': 'POST',
    'headers': {
      'Content-type': 'application/json; charset=utf-8'
    },
    'payload': json,
  };

  UrlFetchApp.fetch(FUF.ApprovalCheckWebhook[env], options);
}


function postPaymentData(studentId, studentName, contractID, contractantName, contractSheet){
  var paymentDataSs = SpreadsheetApp.openById(FUF.contractDBSsId[env]);
  var paymentDataSheet = FUF.ReadSheetById(paymentDataSs,FUF.paymentDbGId[env])
  var dates = contractSheet.getRange(paymentScheduleDatesCoordinate["contract"]).getValues().flat();
  var servicelist = FUF.removeBlankElements(contractSheet.getRange(paymentScheduleServiceListCoordinate).getValues().flat());
  var payments = FUF.removeEmptyRows(contractSheet.getRange(paymentScheduleAmountFullCoordinate).getValues());
  var method = contractSheet.getRange(paymentMethodCoordinate).getValue();

  for(let i = 0; i < payments.length; i++){
    for(let j = 0; j < payments[i].length; j++)
    if(payments[i][j] != 0){
      var lastRow = paymentDataSheet.getLastRow();
      var columnA = paymentDataSheet.getRange(1 , 1, lastRow, 1).getValues().flat();
      var dataInColumnA = columnA.filter((row) => {
        return row != ''
      });
      var lastRowInColumnA = dataInColumnA.length;
      var data = [new Date(), contractID+ i.toString()+j.toString(),contractID, studentId, studentName, contractantName, method, servicelist[i], dates[j], payments[i][j]];
      paymentDataSheet.getRange(lastRowInColumnA + 1, 1, 1, 10).setValues([data]);
    }
  }

}
