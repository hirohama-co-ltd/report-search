// ========================================
// 製品PDF検索 — スプレッドシート初期化
// ========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('製品PDF検索')
    .addItem('シートを初期化', 'initializePdfSearchSpreadsheet')
    .addSeparator()
    .addItem('社外フォルダ接続テスト', 'menuTestExternalFolder_')
    .addItem('社内フォルダ接続テスト', 'menuTestInternalFolder_')
    .addItem('社外で検索テスト（先頭200件）', 'runPdfSearchFromMenu_')
    .addToUi();
}

function initializePdfSearchSpreadsheet() {
  var ss = getPdfSpreadsheet_();

  var config = ss.getSheetByName(PDF_CONFIG_SHEET);
  if (!config) config = ss.insertSheet(PDF_CONFIG_SHEET);
  config.clear();
  config.getRange(1, 1, 1, 3).setValues([['項目', '値', '説明']]);
  config.getRange(2, 1, 5, 3).setValues([
    ['社外ルートフォルダID', PDF_ROOT_FOLDER_IDS.external, '加工済PDF（_originalなし）'],
    ['社内ルートフォルダID', PDF_ROOT_FOLDER_IDS.internal, '原本PDF（_originalあり）'],
    ['検索上限件数', PDF_SEARCH_MAX_RESULTS, '1回の最大表示件数'],
    ['WebアプリデプロイID', PDF_WEBAPP_DEPLOY_ID || '（未設定）', '設定.js の PDF_WEBAPP_DEPLOY_ID']
  ]);
  config.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#dbeafe');
  config.setFrozenRows(1);

  var guide = ss.getSheetByName(PDF_SETUP_SHEET);
  if (!guide) guide = ss.insertSheet(PDF_SETUP_SHEET);
  guide.clear();
  guide.getRange(1, 1).setValue('製品PDF検索 — セットアップ手順');
  guide.getRange(3, 1, 10, 1).setValues([
    ['1. clasp push でコードを反映'],
    ['2. デプロイ → 新しいデプロイ → Webアプリ'],
    ['3. 実行者: 自分 / アクセス: 組織内'],
    ['4. デプロイIDを 設定.js の PDF_WEBAPP_DEPLOY_ID に記載（任意）'],
    ['5. デプロイ者が社内・社外フォルダの閲覧権限を持つこと'],
    [''],
    ['フォルダ構造: ルート → 製品コード → 西暦 → PDF'],
    ['ファイル名: 製品コード-ロットNo-製造日.pdf'],
    ['社内のみ: …_original.pdf']
  ]);

  return '製品PDF検索シートを初期化しました。';
}

function menuTestExternalFolder_() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = testPdfFolderConnection_(PDF_SCOPE_EXTERNAL);
    ui.alert(
      '社外フォルダ接続OK',
      'フォルダ: ' + r.rootName + '\n'
        + '製品コードフォルダ数: ' + r.productFolderCount + '\n'
        + '例: ' + (r.sampleProductCodes.join(', ') || '（なし）'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('接続エラー', e.message, ui.ButtonSet.OK);
  }
}

function menuTestInternalFolder_() {
  var ui = SpreadsheetApp.getUi();
  try {
    var r = testPdfFolderConnection_(PDF_SCOPE_INTERNAL);
    ui.alert(
      '社内フォルダ接続OK',
      'フォルダ: ' + r.rootName + '\n'
        + '製品コードフォルダ数: ' + r.productFolderCount + '\n'
        + '例: ' + (r.sampleProductCodes.join(', ') || '（なし）'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('接続エラー', e.message, ui.ButtonSet.OK);
  }
}
