/**
 * 自習室案内メール送信
 *
 * 契約書シート（シート名＝契約書ID）から生徒情報と申込内容を読み取り、
 * 下の SR_MAIL_TEMPLATE の文面で、info@sss-edu.org から生徒本人へ送る。
 * （文面はプログラム内に固定。変えるときは SR_MAIL_TEMPLATE を直す）
 *
 * - 実行はスプレッドシートのメニュー「SSSメール」から（info@sss-edu.org でログインして実行）
 * - 送信前に必ず確認画面（宛先・件名・本文）を出す
 * - 同じ契約書IDへの二重送信を防ぐ（スクリプト プロパティに送信記録）
 * - 本人Gmailが保護者Gmailと同じ／空欄なら、生徒本人のアドレスを入力してもらう
 *
 * 既存プロジェクトと名前がぶつからないよう、グローバル名はすべて SR_MAIL / srMail_ で始める。
 */

const SR_MAIL_CONFIG = {
  SENDER: 'info@sss-edu.org',
  CELLS: {
    studentId: { label: 'B4', value: 'C4', expected: '学習者ID' },
    studentName: { label: 'B5', value: 'C5', expected: '本人氏名' },
    studentMail: { label: 'B7', value: 'C7', expected: '本人Gmail' },
    parentName: { label: 'B11', value: 'C11', expected: '保護者氏名' },
    parentMail: { label: 'B13', value: 'C13', expected: '本人Gmail' },
    contractId: { label: 'H4', value: 'I4', expected: '契約書ID' },
  },
  SERVICES_RANGE: 'B19:B34', // 申込内容の項目列
  ROOM_TYPES: { OL: '自習室（OL）', FACE: '自習室（対面）' },
  MARKERS: { OL: '（自習室（OL）の場合）', FACE: '（自習室（対面）の場合）' },
  PLACEHOLDERS: { name: '{生徒氏名}', mail: '{生徒Gmail}' },
  PROP_PREFIX: 'SR_MAIL_SENT_',
};

/**
 * メール文面（2026-10-29 にユーザー承認済みの文面）
 * - {生徒氏名} {生徒Gmail} は自動で置き換える
 * - 「（自習室（OL）の場合）」「（自習室（対面）の場合）」の見出しの下は、契約している種類だけ残す
 *   （見出しの行自体はメールに入らない。ブロックの終わりは空行）
 */
const SR_MAIL_TEMPLATE = {
  subject: '自習室ご利用開始に向けた事前準備のお願い',
  lines: [
    '{生徒氏名}様',
    '',
    'この度はSSS Educationの自習室にお申し込みいただきまして、誠にありがとうございます。',
    '自習室のご利用開始に向けて、以下の事前準備をお願いいたします。',
    '',
    '【１. アプリのダウンロード】',
    '自習室のご案内やご連絡には、Googleのアプリを使用いたします。',
    'お使いのスマートフォンまたはパソコンに、以下の3つをダウンロードしてください。',
    '',
    '・Google Chat',
    '・Google カレンダー',
    '・Google ドライブ',
    '',
    'ログインには、ご契約時にご登録いただいたGmailアカウント（{生徒Gmail}）をお使いください。',
    '',
    '【２. 初回のご利用日について】',
    '（自習室（OL）の場合）',
    'オンライン自習室を初めてご利用になる日はいつがよろしいでしょうか。ご希望の日にちをお知らせください。',
    '',
    '（自習室（対面）の場合）',
    '自習室へ初めてご来室される日はいつがよろしいでしょうか。ご希望の日にちと時間をお知らせください。',
    '',
    'お手数ですが、本メールへのご返信をお願いいたします。',
    '何かご不明な点がございましたら、お気軽にご連絡ください。',
    '今後ともどうぞよろしくお願い申し上げます。',
  ],
};

/** メニューを追加する（スプレッドシートを開いたときに自動で実行） */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('SSSメール')
      .addItem('自習室案内メールを送信（開いている契約書シート）', 'srMail_sendForActiveSheet')
      .addItem('自習室案内メールをテスト送信（自分宛て）', 'srMail_testForActiveSheet')
      .addSeparator()
      .addItem('自習室案内メールの送信履歴', 'srMail_showHistory')
      .addToUi();
  } catch (e) {
    // エディタの「実行」など、スプレッドシートの画面がない状態で呼ばれたときは何もしない
    console.log('onOpen: メニューは追加しませんでした（' + e.message + '）');
  }
}

/**
 * 初回の権限許可用。エディタで関数「srMail_authorize」を選んで「実行」する。
 * メールは送らない。必要なサービスに触れるだけ。
 */
function srMail_authorize() {
  const user = srMail_currentUser_();
  GmailApp.getAliases();
  PropertiesService.getScriptProperties().getKeys();
  LockService.getScriptLock();
  console.log('権限OK：実行アカウント=' + user);
}

/**
 * 実行しているアカウントのメールアドレス。
 * このプロジェクトは権限（スコープ）が固定されていて Session.getActiveUser() が使えないため、
 * 既に許可済みの Drive 権限でマイドライブの所有者から取得する。
 */
function srMail_currentUser_() {
  try {
    return String(DriveApp.getRootFolder().getOwner().getEmail() || '').toLowerCase();
  } catch (e) {
    return '';
  }
}

function srMail_sendForActiveSheet() {
  srMail_run_(false);
}

function srMail_testForActiveSheet() {
  srMail_run_(true);
}

function srMail_run_(isTest) {
  const ui = SpreadsheetApp.getUi();
  const cfg = SR_MAIL_CONFIG;
  try {
    const user = srMail_currentUser_();
    if (!isTest && user !== cfg.SENDER) {
      throw new Error(
        '送信元を ' + cfg.SENDER + ' にするため、' + cfg.SENDER + ' のアカウントで実行してください。\n' +
        '現在のアカウント：' + (user || '不明')
      );
    }

    const sheet = SpreadsheetApp.getActiveSheet();
    const info = srMail_readContract_(sheet);
    if (info.roomTypes.length === 0) {
      throw new Error(
        'この契約（契約書ID ' + info.contractId + '）の申込内容に「' + cfg.ROOM_TYPES.OL + '」も「' +
        cfg.ROOM_TYPES.FACE + '」もありません。自習室案内メールは送りません。'
      );
    }

    // 二重送信チェック（本番のみ）
    const props = PropertiesService.getScriptProperties();
    const propKey = cfg.PROP_PREFIX + info.contractId;
    if (!isTest) {
      const prev = props.getProperty(propKey);
      if (prev) {
        const p = JSON.parse(prev);
        const again = ui.alert(
          '送信済みです',
          '契約書ID ' + info.contractId + ' には ' + p.sentAt + ' に ' + p.to + ' へ送信済みです。\nもう一度送りますか？',
          ui.ButtonSet.YES_NO
        );
        if (again !== ui.Button.YES) {
          ui.alert('送信を中止しました。');
          return;
        }
      }
    }

    // 生徒本人のアドレス（保護者と同じ・空欄・形式不正なら入力してもらう）
    let studentMail = info.studentMail;
    let mailNote = '';
    if (srMail_needsStudentMail_(studentMail, info.parentMail)) {
      const res = ui.prompt(
        '生徒本人のメールアドレスを入力してください',
        '契約書の本人Gmail（' + (studentMail || '空欄') + '）が、保護者Gmailと同じか空欄、または形式が正しくありません。\n' +
          info.studentName + '様 本人のメールアドレスを入力してください。',
        ui.ButtonSet.OK_CANCEL
      );
      if (res.getSelectedButton() !== ui.Button.OK) {
        ui.alert('送信を中止しました。');
        return;
      }
      studentMail = res.getResponseText().trim();
      if (srMail_needsStudentMail_(studentMail, info.parentMail)) {
        throw new Error('入力されたアドレス（' + studentMail + '）は使えません。形式が正しく、保護者のアドレスと違うものを入力してください。');
      }
      mailNote = '（契約書の本人Gmailではなく、入力されたアドレスを使います）';
    }

    const mail = srMail_buildMail_(info, studentMail);
    const to = isTest ? user : studentMail;
    if (!srMail_isEmail_(to)) throw new Error('宛先のメールアドレスが正しくありません：' + to);
    const subject = (isTest ? '[テスト] ' : '') + mail.subject;

    const preview =
      (isTest ? '【テスト送信：自分宛てに送ります】\n' : '') +
      '送信元：' + user + '\n' +
      '宛先：' + to + (isTest ? '' : ' ' + mailNote) + '\n' +
      '件名：' + subject + '\n' +
      '契約書ID：' + info.contractId + '　学習者ID：' + info.studentId + '\n' +
      '自習室：' + info.roomTypes.map(function (t) { return cfg.ROOM_TYPES[t]; }).join('、') + '\n' +
      '\n──────── 本文 ────────\n' + mail.body;

    const ok = ui.alert('この内容で送信しますか？', preview, ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) {
      ui.alert('送信を中止しました。');
      return;
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) throw new Error('ほかの送信処理が実行中です。少し待ってからやり直してください。');
    try {
      GmailApp.sendEmail(to, subject, mail.body);
      if (!isTest) {
        props.setProperty(propKey, JSON.stringify({
          to: to,
          studentName: info.studentName,
          studentId: info.studentId,
          roomTypes: info.roomTypes,
          sentAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
          by: user,
        }));
      }
    } finally {
      lock.releaseLock();
    }

    ui.alert('送信しました', to + ' に「' + subject + '」を送信しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('自習室案内メール：送信していません', String(e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

/** 契約書シートから必要な値を読む。形式が想定と違えばエラーで止める。 */
function srMail_readContract_(sheet) {
  const cfg = SR_MAIL_CONFIG;
  const name = sheet.getName();
  if (!/^\d+$/.test(name)) {
    throw new Error('契約書IDのシート（例：2000472）を開いてから実行してください。\n現在のシート：' + name);
  }
  const get = function (a1) { return String(sheet.getRange(a1).getDisplayValue()).trim(); };

  const values = {};
  Object.keys(cfg.CELLS).forEach(function (key) {
    const c = cfg.CELLS[key];
    const label = srMail_normLabel_(get(c.label));
    if (label !== c.expected) {
      throw new Error('シートの形式が想定と違います。' + c.label + ' が「' + c.expected + '」ではなく「' + get(c.label) + '」です。');
    }
    values[key] = get(c.value);
  });

  if (values.contractId && values.contractId !== name) {
    throw new Error('シート名（' + name + '）と契約書ID（' + values.contractId + '）が一致しません。');
  }
  if (!values.studentName) throw new Error('本人氏名（' + cfg.CELLS.studentName.value + '）が空欄です。');

  const services = sheet.getRange(cfg.SERVICES_RANGE).getDisplayValues()
    .map(function (r) { return srMail_normParen_(String(r[0]).trim()); })
    .filter(function (s) { return s; });

  return {
    contractId: name,
    studentId: values.studentId,
    studentName: values.studentName,
    studentMail: values.studentMail,
    parentName: values.parentName,
    parentMail: values.parentMail,
    services: services,
    roomTypes: srMail_roomTypes_(services),
  };
}

function srMail_roomTypes_(services) {
  const t = SR_MAIL_CONFIG.ROOM_TYPES;
  const types = [];
  if (services.some(function (s) { return s.indexOf(t.OL) !== -1; })) types.push('OL');
  if (services.some(function (s) { return s.indexOf(t.FACE) !== -1; })) types.push('FACE');
  return types;
}

/** SR_MAIL_TEMPLATE から件名と本文を作る */
function srMail_buildMail_(info, studentMail) {
  return srMail_composeFromTemplate_(SR_MAIL_TEMPLATE, info.studentName, studentMail, info.roomTypes);
}

/** 文面テンプレート {subject, lines} から件名・本文を組み立てる */
function srMail_composeFromTemplate_(template, studentName, studentMail, roomTypes) {
  const cfg = SR_MAIL_CONFIG;
  const subject = template.subject;
  const lines = template.lines;

  const want = {};
  roomTypes.forEach(function (t) { want[t] = true; });

  const out = [];
  let mode = null;
  lines.forEach(function (raw) {
    const line = raw.replace(/\s+$/, '');
    const trimmed = line.trim();
    if (trimmed === cfg.MARKERS.OL) { mode = 'OL'; return; }
    if (trimmed === cfg.MARKERS.FACE) { mode = 'FACE'; return; }
    if (mode) {
      if (trimmed === '') {
        if (want[mode]) out.push(''); // 使わなかった種類の空行は残さない
        mode = null;
        return;
      }
      if (want[mode]) out.push(line);
      return;
    }
    out.push(line);
  });

  // 連続する空行を1つにまとめ、前後の空行を削る
  const collapsed = [];
  out.forEach(function (l) {
    const blank = l.trim() === '';
    if (blank && (collapsed.length === 0 || collapsed[collapsed.length - 1].trim() === '')) return;
    collapsed.push(blank ? '' : l);
  });
  while (collapsed.length && collapsed[collapsed.length - 1] === '') collapsed.pop();

  let body = collapsed.join('\n')
    .split(cfg.PLACEHOLDERS.name).join(studentName)
    .split(cfg.PLACEHOLDERS.mail).join(studentMail);

  const leftover = body.match(/\{[^{}\n]{1,20}\}/);
  if (leftover) throw new Error('本文に置き換えられていない項目があります：' + leftover[0]);
  if (body.indexOf(cfg.MARKERS.OL) !== -1 || body.indexOf(cfg.MARKERS.FACE) !== -1) {
    throw new Error('本文に自習室の種類の見出しが残っています。テンプレートの形式を確認してください。');
  }
  if (body.indexOf(studentName + '様') !== 0) {
    throw new Error('本文の1行目が「' + studentName + '様」になっていません。テンプレートを確認してください。');
  }
  return { subject: subject, body: body };
}

function srMail_needsStudentMail_(studentMail, parentMail) {
  if (!srMail_isEmail_(studentMail)) return true;
  return String(studentMail).trim().toLowerCase() === String(parentMail || '').trim().toLowerCase();
}

function srMail_isEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim());
}

function srMail_normLabel_(s) {
  return String(s).replace(/[\s　]/g, '');
}

function srMail_normParen_(s) {
  return String(s).replace(/\(/g, '（').replace(/\)/g, '）');
}

/** 送信履歴（新しい順に最大30件）を表示 */
function srMail_showHistory() {
  const ui = SpreadsheetApp.getUi();
  const all = PropertiesService.getScriptProperties().getProperties();
  const rows = Object.keys(all)
    .filter(function (k) { return k.indexOf(SR_MAIL_CONFIG.PROP_PREFIX) === 0; })
    .map(function (k) {
      const p = JSON.parse(all[k]);
      p.contractId = k.slice(SR_MAIL_CONFIG.PROP_PREFIX.length);
      return p;
    })
    .sort(function (a, b) { return a.sentAt < b.sentAt ? 1 : -1; })
    .slice(0, 30)
    .map(function (p) {
      return p.sentAt + '　' + p.contractId + '　' + p.studentName + '様　→ ' + p.to;
    });
  ui.alert('自習室案内メールの送信履歴', rows.length ? rows.join('\n') : 'まだ送信記録はありません。', ui.ButtonSet.OK);
}
