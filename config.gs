// 注意すべき規則
// 契約IDは6桁、生徒Idは７桁になっている。

const digitsOfStudentId = 7;
const digitsOfContractId = 6;
const passwordLength = 8 //8文字

const indexOfDriveFileQuestion = 15 //添付資料の質問を 15 番目としている。（１から数えている）

// #お問合せフォーム関連
// 面談受付時間
const interviewReceptionTimes = [
  [8, 00], [9, 00], [10, 00],
  [11, 00], [12, 00], [13, 00],
  [14, 00], [15, 00], [16, 00],
  [17, 00], [18, 00], [19, 00],
  [20, 00]
];
const receptionDays = 14 //面談候補日時を提示する日数

const salesInterviewerHours = 1.5;  //契約担当者が確保すべき時間
const academicInterViewerHours = 1; //教務担当者が確保すべき時間

// 面談必須出席者
const interviewRooms = [FUF.sssJimu]; //面談室
const salesInterviewers = [FUF.tanzeelAhmad]; //契約担当者
const onsiteAcademicInterviwers = [FUF.tanzeelAhmad]; //教務担当者
const onlineAcademicInterviewers = [FUF.tanzeelAhmad] // オンラインのみで対応可能な教務担当者。

// 面談希望日時の質問名
const onlineReceptionQuestionTitles = ["第１希望日時(OL)", "第２希望日時(OL)", "第３希望日時(OL)"];
const onsiteReceptionQuestionTitles = ["第１希望日時(対面)", "第２希望日時(対面)", "第３希望日時(対面)"];
const mailQuestionTitle = "メール";
const studentNameQuestionTitle = "学習者（生徒）氏名";
const parentNameQuestionTitle = "保護者様（決済者様）氏名";
const mtgFormatQuestionTitle = "ご希望の面談形式";

//　契約担当者
const salesPerson = FUF.tanzeelAhmad;
const salesPerson1 = FUF.nomiReina;

// #契約書スプレッドシート関連
// 契約書の情報欄
const corqGenerateSheetName = {
  "contract": "契約書",
  "quotation": "見積書"
}
const contractUneditableSheetName = '契約書（フォーマット用）' //フォーマット用の契約書シートの名称
const quotationFormatSheetName = '見積書' //見積書生成シートの名称
const quotationUneditableSheetName = '見積書（フォーマット用）'
const contractDateCoordinate = `I5` //契約日の座標
const contractIdCoordinate = `I4` //契約書IDの座標
const studentIdCoordinate = 'C4' //学習者 ID の座標
const studentNameCoordinate = 'C5' //本人氏名の座標
const contractantNameCoordinate = 'C11'//決済者氏名の座標
const contractantAddressCoordinate = 'C13'//決済者メールアドレスの座標
const contractServiceInfoCoordinate = 'B19:J34' //契約内容の座標
const contractServiceInfoRange = [19, 2, 16, 9] //契約内容のレンジ
const paymentScheduleDatesCoordinate = {
  "contract": 'R17:AF17',
  "quotation": "N10:Z10" //支払いスケジュール表の日付の座標\
}
const paymentScheduleAmountCoordinate = {
  "contract": 'R19:AF34',
  "quotation": "N12:Z27"
} //支払いスケジュール表の金額の座標
const paymentScheduleServiceListCoordinate = 'Q19:Q35' //支払スケジュール表の項目一覧の座標
const paymentScheduleAmountFullCoordinate = 'R19:AF35' //入塾金支払いスケジュール表の金額の座標
const paymentReccurenceCoordinate = 'I9' //支払回数の座標
const paymentAmountCoordinate = 'R18:AF18' //支払合計額の座標
const paymentMethodCoordinate = 'C40' //支払方法の座標
const contractDueDateIndex = 0; //支払期限インデックス（転置後）
const contractPayAmountIndex = 1; //支払合計額インデックス（転置後）
const unprotect = ['C4:D5', 'I8:J9', 'B19:J34', 'C39:E39', 'R19:AD34', 'R17:AD17']; //契約書の保護解除セル
const mtgDateCoordinate = 'I8'

const quotationIdCoordinate = `I4` //見積書IDの座標
//カスタム改ページの指定（//参照サイト２：https://kandiral.ru/googlescript/eksport_tablic_google_sheets_v_pdf_fajl.html)
function customPages(source, d, corq) {
  if (corq == "contract") {
    return [null, null, null, null, null, null, null, null, null, 0,
      [[source]],
      10000000, null, null, null, null, null, null, null, null, null, null, null, null, null, null,
      d,
      null, null, [0, null, 1, 0, 0, 0, 0, 0, 1, 1, 2, 1, null, null, 2, 1],
      ["A4", 1, 6, 1, [0.75, 0.75, 0.25, 0.25]], null, 0,
      [[source, null, [[0, 12], [12, 16], [16, 30]]]]];
  }
  if (corq == "quotation") {
    console.log([null, null, null, null, null, null, null, null, null, 0,
      [[source]],
      10000000, null, null, null, null, null, null, null, null, null, null, null, null, null, null,
      d, null, null, [0, null, 1, 0, 0, 0, 0, 0, 1, 1, 2, 1, null, null, 2, 1],
      ["A4", 1, 6, 1, [0.75, 0.75, 0.25, 0.25]], null, 0,
      [[source, null, [[0, 12], [12, 16], [16, 30]]]]])
    return [null, null, null, null, null, null, null, null, null, 0,
      [[source]],
      10000000, null, null, null, null, null, null, null, null, null, null, null, null, null, null,
      d, null, null, [0, null, 1, 0, 0, 0, 0, 0, 1, 1, 2, 1, null, null, 2, 1],
      ["A4", 1, 6, 1, [0.75, 0.75, 0.25, 0.25]], null, 0,
      [[source, null, [[0, 11], [11, 26]]]]];

  }

}

const serviceInfoCoordinate = {
  "contract": "B19:J34",
  "quotation": "B12:J27"
}

const serviceInfoRange = {
  "contract": [19, 2, 16, 9],
  "quotation": [12, 2, 16, 9]
}

const scheduleDatesCoordinate = {
  "contract": 'R17:AD17',
  "quotation": 'N10:Z10'
}

const scheduleAmountCoordinate = {
  "contract": 'R19:AD35',
  "quotation": 'N12:Z28'
}

const paymentOptionRange = {
  "contract": "I9",
  "quotation": "I8"
}

const amountCoordinate = {
  "contract": 'R18:AD18',
  "quotation": 'N11:Z11'
}

const methodCoordinate = {
  "contract": 'C40',
  "quotation": 'C33'
}

const unprotectCoordinates = {
  "contract": ['C4:D5', 'I8:J9', 'B19:J34', 'C39:E39', 'R19:AD35', 'R17:AD17'],
  "quotation": ['C4:D5', 'I8:J8', 'B12:J27', 'C32:E32', 'N12:Z28', 'N10:Z10']
}

const serviceStartDateCoordinate = {
  "contract": "F19:F34",
  "quotation": "F12:F27"
}


const serviceEndDateCoordinate = {
  "contract": "G19:G34",
  "quotation": "G12:G27"
}

const idCoordinate = "I4:J4";

const distinguishCell = 'H4';

const monthlyPaymentDueDate = 27;

const tempFolderTags = [`<学習者ID>`, `<学習者姓名>`, `<保護者姓名>`,`<連絡先Mail>`,`<学習者Mail>`, `<保護者Mail>`];

const tempOfContractFolder = {
    name: '<学習者ID> - <学習者姓名>様 - SSS Education 契約情報',
    id: `1lQ3hDXAJBnj-xqMvGU8jOBdTP0wsdaqD`
}

const tempOfGuidanceFolder = {
    name: '<学習者ID> - <学習者姓名>様 - SSS Education 指導情報',
    id: `1VZ9PTy3510mu3XXjImbPlcPf7OH8-YSx`
}

const tempFolders = {
  contract: tempOfContractFolder,
  guidance: tempOfGuidanceFolder
};

const contractInfoTempFile = [
  '<学習者ID> - <学習者姓名>様 - 契約状況',
  '1UHnoPMJ3x2c4hBi90-1GoQhukOHy5cycvT0_w8FYJ6E'
]

const contractInfoTempFolders = [
  '<学習者ID> - <学習者姓名>様 - 領収書',
  '<学習者ID> - <学習者姓名>様 - 契約書', '<学習者ID> - <学習者姓名>様 - 見積書'
];

const guidanceInfoTempFiles = [
  ['受験戦略_<学習者ID> - <学習者姓名>様', '1GypcuMCnRqnxcQs6YBKbCTBpF5BhqIjEz8r1JudWHPg'],
  ['<学習者ID> - <学習者姓名>様 - サービス取得状況', '1AJ-PGE4ofD60LPSKAknlqWhdyREXCqoo7sn0omthGPE']
]

const guidanceInfoTempFolders = [
  '<学習者ID> - <学習者姓名>様 - 成績表',
  '<学習者ID> - <学習者姓名>様 - 指導記録'
];

const guidanceRecordTempFolders = [
  '<学習者ID> - <学習者姓名>様  - その他',
  '<学習者ID> - <学習者姓名>様  - 添削指導',
  '<学習者ID> - <学習者姓名>様 - 面談記録',
  '<学習者ID> - <学習者姓名>様  - 個別指導記録'
];

const studentMailColumn = 5;
const parentMailColumn = 9;
const contactMailColumn = 7;

const receiptContractantNameCoordinate = "C11";
const receiptStudentNameCoordinate = "C12";
const receiptStudentIdCoordinate = "C13";
const receiptPaymentIdsCoordinate = "B18:B39";
const receiptPaymentsCoordinate = "B18:F39";
const receiptContractUrlsCoordinate = "B41:D44";
const receiptBikouCoordinate = "B47";
const receiptIdCoordinate = "F11";
const receiptUneditableSheetName = "領収書-フォーマット用";

function receiptCustomPages(source, d) {
  return [null, null, null, null, null, null, null, null, null, 0, [[source]], 10000000, null, null, null, null, null, null, null, null, null, null, null, null, null, null, d, null, null, [1, null, 1, 0, 0, 0, 0, 0, 1, 1, 2, 1, null, null, 2, 1], ["A4", 1, 6, 1, [0.75, 0.75, 0.25, 0.25]], null, 0, [[source, [[0, 47]], [[0, 7]]]]];
}

const paymentReminderSpan = 3;
const salesPersonMargin = 1;
const paymentOverdueSpan = 2;
const paymentPerfectlyOverdueSpan = 7;
const paymentChatReminderColumns = [8, 1, 3, 4, 5, 7, 9, 10, 11];