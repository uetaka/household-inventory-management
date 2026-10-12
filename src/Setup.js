/**
 * 初期設定とスプレッドシートのメニュー。
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在庫管理')
    .addItem('初期設定（シート作成）', 'setupSpreadsheet')
    .addItem('Claude API 接続テスト', 'testClaudeConnectionFromMenu')
    .addItem('買い物リストを更新', 'refreshShoppingListFromMenu')
    .addItem('買いたい物: AI への指示文を表示', 'showWishInstructionsFromMenu')
    .addItem('買い物リストの自動更新（1時間ごと）を設定', 'installShoppingTriggerFromMenu')
    .addToUi();
}

/**
 * 必要なシートをすべて作成し、ヘッダーとサンプルを入れる。
 * 既存シートは壊さない（ヘッダーだけ確認する）。
 */
function setupSpreadsheet() {
  const ss = getSpreadsheet_();
  const S = CONFIG.SHEETS;
  const H = CONFIG.HEADERS;

  const drawers = ensureSheet_(ss, S.DRAWERS, H.DRAWERS);
  if (drawers.getLastRow() === 1) {
    drawers.getRange(2, 1, SAMPLE_DRAWERS.length, H.DRAWERS.length).setValues(SAMPLE_DRAWERS);
  }

  const items = ensureSheet_(ss, S.ITEMS, H.ITEMS);
  if (items.getLastRow() === 1) {
    items.getRange(2, 1, SAMPLE_ITEMS.length, H.ITEMS.length).setValues(SAMPLE_ITEMS);
  }

  const inventory = ensureSheet_(ss, S.INVENTORY, H.INVENTORY);
  inventory.getRange('I:I').setNumberFormat('yyyy/mm/dd hh:mm');

  const history = ensureSheet_(ss, S.HISTORY, H.HISTORY);
  history.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  history.setColumnWidth(5, 300);
  history.setColumnWidth(6, 300);

  const shopping = ensureSheet_(ss, S.SHOPPING, H.SHOPPING);
  formatShoppingSheet_(shopping);

  formatWishSheet_(ensureSheet_(ss, S.WISHLIST, H.WISHLIST));

  // AI が「買いたい物」に書き込んでも、外部からの編集では編集時トリガーが動かないので、定期的に作り直す。
  try {
    installShoppingTrigger_();
  } catch (e) {
    Logger.log('買い物リストの自動更新を設定できませんでした: ' + e.message);
  }

  // 空の「シート1」が残っていれば消す。
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }

  try {
    SpreadsheetApp.getUi().alert('初期設定が完了しました。\n次に「プロジェクトの設定 > スクリプト プロパティ」に ANTHROPIC_API_KEY を登録してください。');
  } catch (e) {
    Logger.log('初期設定が完了しました。');
  }
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold').setBackground('#e8eaed');
  sheet.setFrozenRows(1);
  return sheet;
}

function testClaudeConnectionFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = testClaudeConnection();
    ui.alert('接続OK\nモデル: ' + result.model + ' (' + result.displayName + ')');
  } catch (e) {
    ui.alert('接続に失敗しました\n' + e.message);
  }
}

function refreshShoppingListFromMenu() {
  const list = refreshShoppingList();
  SpreadsheetApp.getUi().alert(
    (list.shortage.length ? list.shortage.length + ' 品目が最低在庫を下回っています。' : '不足している品目はありません。') +
    (list.wishes.length ? '\n買いたい物（単発）: ' + list.wishes.length + ' 件' : ''));
}

function showWishInstructionsFromMenu() {
  SpreadsheetApp.getUi().alert('AI への指示文', wishInstructions_().sheet, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ---------- 買い物リストの自動更新 ----------

const SHOPPING_TRIGGER_HANDLER = 'refreshShoppingListByTrigger';

/** 時間主導トリガーから呼ばれる。AI や人が「買いたい物」に直接書いた行を買い物リストに反映する。 */
function refreshShoppingListByTrigger() {
  refreshShoppingList();
}

/**
 * 1時間ごとに買い物リストを作り直すトリガーを、無ければ作る。
 * トリガーは作った人のものになるので、アプリをデプロイした本人として実行される場所から呼ぶ。
 * @return {boolean} 新しく作ったら true
 */
function installShoppingTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === SHOPPING_TRIGGER_HANDLER;
  });
  if (exists) return false;
  ScriptApp.newTrigger(SHOPPING_TRIGGER_HANDLER).timeBased().everyHours(1).create();
  return true;
}

function installShoppingTriggerFromMenu() {
  const created = installShoppingTrigger_();
  SpreadsheetApp.getUi().alert(created ? '買い物リストを1時間ごとに自動更新するよう設定しました。' : '自動更新はすでに設定されています。');
}
