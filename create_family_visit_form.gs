/**
 * Googleフォーム + 回答スプレッドシートを自動作成するGAS。
 *
 * 使い方:
 * 1) STAFF_OPTIONS を実際の職員名に更新
 * 2) Apps Script で setupSurveySystem() を1回実行
 * 3) ログに出る Form/Sheet URL を共有
 */

const CONFIG = {
  formTitle: '面会時 ご家族要望・苦情・申し出 記録フォーム',
  formDescription:
    '【趣旨】\n' +
    'ご家族の面会時に出る要望・苦情・申し出を、個人メモで終わらせず職員間で共有するための記録です。\n\n' +
    '【目的】\n' +
    '1. 状況把握: いつ、誰のご家族から、どんな内容が出たかを一覧で把握\n' +
    '2. 情報共有: 主担当・対応職員・ステータスを明確化\n' +
    '3. 改善: 対策内容と効果確認を記録し、再発防止へ\n' +
    '4. 業務効率化: フォーム入力→スプレッドシート自動反映（案件ID付与）',
  spreadsheetName: '面会時_要望苦情申し出_管理台帳',
  caseIdPrefix: 'MV',
};

const STAFF_OPTIONS = [
  '職員A',
  '職員B',
  '職員C',
  '職員D',
];

/**
 * 初期構築（フォーム作成、スプレッドシート連携、トリガー設定）
 */
function setupSurveySystem() {
  const form = FormApp.create(CONFIG.formTitle)
    .setDescription(CONFIG.formDescription)
    .setCollectEmail(false)
    .setAllowResponseEdits(true)
    .setProgressBar(true)
    .setConfirmationMessage('ご入力ありがとうございました。内容を受け付けました。');

  buildFormItems_(form);

  const spreadsheet = SpreadsheetApp.create(CONFIG.spreadsheetName);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  const ss = SpreadsheetApp.openById(spreadsheet.getId());
  const sheet = ss.getSheets()[0];
  sheet.setName('フォーム回答');

  prepareResponseSheet_(sheet);
  recreateSubmitTrigger_(form.getId());

  Logger.log('=== 作成完了 ===');
  Logger.log('Form URL: %s', form.getEditUrl());
  Logger.log('Form Public URL: %s', form.getPublishedUrl());
  Logger.log('Spreadsheet URL: %s', spreadsheet.getUrl());
}

/**
 * フォーム項目作成
 */
function buildFormItems_(form) {
  form.addDateItem()
    .setTitle('面会日')
    .setRequired(true);

  form.addTextItem()
    .setTitle('利用者名')
    .setRequired(true);

  form.addTextItem()
    .setTitle('ご家族名')
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('種別（複数可）')
    .setChoiceValues(['要望', '苦情', '申し出'])
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('聞き取り内容')
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('重要度')
    .setChoiceValues(['高', '中', '低'])
    .setRequired(true);

  form.addListItem()
    .setTitle('主担当職員')
    .setChoiceValues(STAFF_OPTIONS)
    .setRequired(true);
}

/**
 * 回答シートの先頭に案件ID列を追加
 */
function prepareResponseSheet_(sheet) {
  sheet.insertColumnBefore(1);
  sheet.getRange(1, 1).setValue('案件ID');
  sheet.setFrozenRows(1);
}

/**
 * フォーム送信トリガーを作り直し
 */
function recreateSubmitTrigger_(formId) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forForm(formId)
    .onFormSubmit()
    .create();
}

/**
 * フォーム送信時に案件IDを自動採番
 */
function onFormSubmit(e) {
  if (!e || !e.range) {
    return;
  }

  const row = e.range.getRow();
  const sheet = e.range.getSheet();
  const caseId = generateCaseId_(sheet, row);
  sheet.getRange(row, 1).setValue(caseId);
}

/**
 * 案件IDを採番（例: MV-20260423-0001）
 */
function generateCaseId_(sheet, row) {
  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const datePart = Utilities.formatDate(new Date(), timezone, 'yyyyMMdd');

  let serial = row - 1;
  if (serial < 1) {
    serial = 1;
  }

  return `${CONFIG.caseIdPrefix}-${datePart}-${String(serial).padStart(4, '0')}`;
}
