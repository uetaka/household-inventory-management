/**
 * スプレッドシートの読み書き。
 * スクリプトがシートに紐付いている場合は getActive()、
 * スタンドアロンの場合はスクリプトプロパティ SPREADSHEET_ID を使う。
 */

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActive();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('スプレッドシートが見つかりません。スクリプトをシートに紐付けるか、スクリプトプロパティ SPREADSHEET_ID を設定してください。');
  }
  return SpreadsheetApp.openById(id);
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」がありません。メニューの「在庫管理 > 初期設定」を実行してください。');
  }
  return sheet;
}

/** ヘッダー行をキーにしたオブジェクト配列として読む。空行は除く。 */
function readTable_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(function (v) { return v === '' || v === null; })) continue;
    const obj = { _row: i + 1 };
    headers.forEach(function (h, idx) { obj[h] = row[idx]; });
    rows.push(obj);
  }
  return rows;
}

function toStr_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function toNum_(v, fallback) {
  const n = Number(v);
  return isNaN(n) || v === '' || v === null ? (fallback === undefined ? 0 : fallback) : n;
}

// ---------- 引き出し ----------

/**
 * 引き出しの一覧。アプリの選択肢はこの順で並ぶ。
 * 「表示順」列の昇順、未設定の行はその後ろにシートの行順で並べる。
 */
function getDrawers() {
  return readTable_(getSheet_(CONFIG.SHEETS.DRAWERS))
    .map(function (r) {
      return {
        id: toStr_(r['引き出しID']),
        name: toStr_(r['名前']),
        location: toStr_(r['場所']),
        note: toStr_(r['備考']),
        order: toOrder_(r['表示順']),
        _row: r._row,
      };
    })
    .filter(function (d) { return d.id && d.name; })
    .sort(function (a, b) {
      const ao = a.order === null ? Infinity : a.order;
      const bo = b.order === null ? Infinity : b.order;
      return ao === bo ? a._row - b._row : ao - bo;
    })
    .map(function (d) { delete d._row; return d; });
}

/** 表示順のセル値を数値に。空欄や数値でない物は null。 */
function toOrder_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function getDrawer_(drawerId) {
  const d = getDrawers().filter(function (x) { return x.id === drawerId; })[0];
  if (!d) throw new Error('引き出しID「' + drawerId + '」が見つかりません。');
  return d;
}

/**
 * 引き出しシートの列位置をヘッダー名から引く（1始まり）。
 * 古いシートに「表示順」列が無い場合は createOrder が true のときだけ追加する。
 */
function drawerColumns_(sheet, createOrder) {
  const H = CONFIG.HEADERS.DRAWERS;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const find = function (name) {
    const i = headers.indexOf(name);
    if (i === -1) throw new Error('引き出しシートに「' + name + '」列がありません。メニューの「在庫管理 > 初期設定」を実行してください。');
    return i + 1;
  };
  const cols = { id: find(H[0]), name: find(H[1]), location: find(H[2]), note: find(H[3]) };
  let orderCol = headers.indexOf(H[4]) + 1;
  if (!orderCol && createOrder) {
    orderCol = Math.max(lastCol + 1, H.length);
    sheet.getRange(1, orderCol).setValue(H[4]).setFontWeight('bold').setBackground('#e8eaed');
  }
  cols.order = orderCol;
  return cols;
}

/**
 * 引き出しを1件追加する。ID は D01, D02… の続き番号を自動で振る。
 * 表示順は既存の最大値 + 1（誰も表示順を使っていなければ空欄のまま末尾に並ぶ）。
 * @return {string} 追加した引き出しID
 */
function addDrawer(input) {
  const name = toStr_(input && input.name);
  if (!name) throw new Error('引き出しの名前を入力してください。');
  const location = toStr_(input && input.location);
  const note = toStr_(input && input.note);

  const sheet = getSheet_(CONFIG.SHEETS.DRAWERS);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rows = readTable_(sheet);
    const used = {};
    let maxNum = 0, maxOrder = 0, hasOrder = false;
    rows.forEach(function (r) {
      const id = toStr_(r['引き出しID']);
      if (id) used[id] = true;
      const m = /^D(\d+)$/i.exec(id);
      if (m) maxNum = Math.max(maxNum, Number(m[1]));
      const o = toOrder_(r['表示順']);
      if (o !== null) { hasOrder = true; maxOrder = Math.max(maxOrder, o); }
      if (toStr_(r['名前']) === name && toStr_(r['場所']) === location) {
        throw new Error('同じ名前・場所の引き出し「' + name + '」が既にあります。');
      }
    });
    let id;
    do {
      maxNum++;
      id = 'D' + (maxNum < 10 ? '0' + maxNum : String(maxNum));
    } while (used[id]);

    const cols = drawerColumns_(sheet, true);
    const row = [];
    for (let i = 0; i < Math.max(cols.id, cols.name, cols.location, cols.note, cols.order); i++) row.push('');
    row[cols.id - 1] = id;
    row[cols.name - 1] = name;
    row[cols.location - 1] = location;
    row[cols.note - 1] = note;
    row[cols.order - 1] = hasOrder ? maxOrder + 1 : '';
    sheet.appendRow(row);
    return id;
  } finally {
    lock.releaseLock();
  }
}

/**
 * アプリで並べ替えた順序を「表示順」列に 1, 2, 3… として書き込む。
 * 一覧に無い行（名前が空など）は既存の値をそのまま残す。
 * @param {string[]} ids 表示したい順に並べた引き出しID
 */
function saveDrawerOrder(ids) {
  ids = (ids || []).map(toStr_).filter(Boolean);
  const pos = {};
  ids.forEach(function (id, i) { if (!pos[id]) pos[id] = i + 1; });

  const sheet = getSheet_(CONFIG.SHEETS.DRAWERS);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cols = drawerColumns_(sheet, true);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const idValues = sheet.getRange(2, cols.id, lastRow - 1, 1).getValues();
    const orderRange = sheet.getRange(2, cols.order, lastRow - 1, 1);
    const current = orderRange.getValues();
    const next = idValues.map(function (r, i) {
      const id = toStr_(r[0]);
      return [pos[id] ? pos[id] : current[i][0]];
    });
    orderRange.setValues(next);
  } finally {
    lock.releaseLock();
  }
}

// ---------- 品目マスタ ----------

function getItemMaster() {
  return readTable_(getSheet_(CONFIG.SHEETS.ITEMS))
    .map(function (r) {
      return {
        name: toStr_(r['品目名']),
        unit: toStr_(r['単位']),
        aliases: toStr_(r['別名']).split(/[,、，]/).map(function (s) { return s.trim(); }).filter(Boolean),
        minStock: toNum_(r['最低在庫'], 0),
        category: toStr_(r['カテゴリ']),
        product: toStr_(r['定番商品']),
        store: toStr_(r['購入先']),
      };
    })
    .filter(function (m) { return m.name; });
}

/** 認識結果に含まれる未登録品目をマスタへ追加する。 */
function addNewItemsToMaster_(items) {
  const master = getItemMaster();
  const known = {};
  master.forEach(function (m) {
    known[m.name] = true;
    m.aliases.forEach(function (a) { known[a] = true; });
  });
  const additions = [];
  items.forEach(function (it) {
    const name = toStr_(it.name);
    if (!name || known[name]) return;
    known[name] = true;
    // 定番商品は人が決めるものなので空のまま。写真から読めた商品名だけ参考に入れておく。
    additions.push([name, toStr_(it.unit) || '個', '', '', '', '', '']);
  });
  if (additions.length) {
    const sheet = getSheet_(CONFIG.SHEETS.ITEMS);
    sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, additions[0].length).setValues(additions);
  }
  return additions.map(function (a) { return a[0]; });
}

/**
 * 商品登録モードからのマスタ更新。
 * 同じ品目名があれば「定番商品」「単位」「カテゴリ」の空欄を埋めて別名をマージし、無ければ行を追加する。
 */
function upsertItemMaster_(products) {
  const sheet = getSheet_(CONFIG.SHEETS.ITEMS);
  const rows = readTable_(sheet);
  const byName = {};
  rows.forEach(function (r) { byName[toStr_(r['品目名'])] = r; });
  const added = [], updated = [];

  products.forEach(function (pr) {
    const name = toStr_(pr.name);
    if (!name) return;
    const aliases = toStr_(pr.aliases).split(/[,、，]/).map(function (a) { return a.trim(); }).filter(function (a) { return a && a !== name; });
    const existing = byName[name];
    if (existing) {
      const row = existing._row;
      const current = {
        unit: toStr_(existing['単位']),
        aliases: toStr_(existing['別名']).split(/[,、，]/).map(function (a) { return a.trim(); }).filter(Boolean),
        category: toStr_(existing['カテゴリ']),
        product: toStr_(existing['定番商品']),
      };
      const mergedAliases = current.aliases.slice();
      aliases.forEach(function (a) { if (mergedAliases.indexOf(a) === -1) mergedAliases.push(a); });
      // 定番商品は画面で入力された値を優先（人が確認済み）。空なら既存を維持。
      const product = toStr_(pr.product) || current.product;
      sheet.getRange(row, 2).setValue(current.unit || toStr_(pr.unit) || '個');
      sheet.getRange(row, 3).setValue(mergedAliases.join(','));
      sheet.getRange(row, 5).setValue(current.category || toStr_(pr.category));
      sheet.getRange(row, 6).setValue(product);
      updated.push(name);
    } else {
      sheet.appendRow([name, toStr_(pr.unit) || '個', aliases.join(','), toNum_(pr.minStock, ''), toStr_(pr.category), toStr_(pr.product), '']);
      byName[name] = { _row: sheet.getLastRow() };
      added.push(name);
    }
  });
  return { added: added, updated: updated };
}

// ---------- 在庫 ----------

function getInventoryForDrawer(drawerId) {
  return readTable_(getSheet_(CONFIG.SHEETS.INVENTORY))
    .filter(function (r) { return toStr_(r['引き出しID']) === drawerId; })
    .map(inventoryRowToObject_);
}

function getAllInventory_() {
  return readTable_(getSheet_(CONFIG.SHEETS.INVENTORY)).map(inventoryRowToObject_);
}

function inventoryRowToObject_(r) {
  return {
    drawerId: toStr_(r['引き出しID']),
    drawerName: toStr_(r['引き出し名']),
    name: toStr_(r['品目名']),
    quantity: toNum_(r['数量'], 0),
    unit: toStr_(r['単位']),
    remaining: toStr_(r['残量目安']),
    confidence: toStr_(r['確信度']),
    needsReview: r['要確認'] === true || toStr_(r['要確認']) === 'TRUE',
    updatedAt: r['最終更新'] instanceof Date ? r['最終更新'].toISOString() : toStr_(r['最終更新']),
    note: toStr_(r['メモ']),
    product: toStr_(r['見えた商品名']),
  };
}

/**
 * 指定した引き出しの在庫行を丸ごと入れ替える（スナップショット方式）。
 * 他の引き出しの行はそのまま残す。
 */
function replaceDrawerInventory_(drawer, items, now) {
  const sheet = getSheet_(CONFIG.SHEETS.INVENTORY);
  const headers = CONFIG.HEADERS.INVENTORY;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const all = sheet.getDataRange().getValues();
    const kept = all.slice(1).filter(function (row) {
      return toStr_(row[0]) !== drawer.id && !row.every(function (v) { return v === ''; });
    });
    const fresh = items.map(function (it) {
      return [
        drawer.id,
        drawer.name,
        toStr_(it.name),
        toNum_(it.quantity, 0),
        toStr_(it.unit),
        remainingLabel_(it.remaining),
        CONFIG.CONFIDENCE_LABELS[it.confidence] || toStr_(it.confidence) || '',
        it.needsReview === true,
        now,
        toStr_(it.note),
        toStr_(it.product),
      ];
    });
    const body = kept.concat(fresh);
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
    }
    if (body.length) {
      sheet.getRange(2, 1, body.length, headers.length).setValues(body);
      sheet.getRange(2, 8, body.length, 1).insertCheckboxes();
    }
  } finally {
    lock.releaseLock();
  }
}

/** 残量目安のキー（full など）でも表示名（未開封/満 など）でも表示名に揃える。 */
function remainingLabel_(v) {
  const s = toStr_(v);
  if (CONFIG.REMAINING_LABELS[s]) return CONFIG.REMAINING_LABELS[s];
  return s;
}

// ---------- 履歴 ----------

function appendHistory_(record) {
  const sheet = getSheet_(CONFIG.SHEETS.HISTORY);
  sheet.appendRow([
    record.at,
    record.drawerId,
    record.drawerName,
    (record.photoUrls || []).join('\n'),
    truncateForCell_(JSON.stringify(record.recognized)),
    truncateForCell_(JSON.stringify(record.committed)),
    record.user || '',
    record.inputTokens || '',
    record.outputTokens || '',
    record.model || '',
  ]);
}

function truncateForCell_(s) {
  // セルの上限は 50,000 文字。
  return s.length > 49000 ? s.slice(0, 49000) + '…(省略)' : s;
}

// ---------- 買いたい物（単発） ----------

// 「買いたい物」シートの列位置（0始まり）。CONFIG.HEADERS.WISHLIST の順番と揃える。
const WISH_COL = { NAME: 0, MEMO: 1, STORE: 2, ADDED_AT: 3, ADDED_BY: 4, DONE: 5, DONE_AT: 6 };

/**
 * 「買いたい物」シート。以前の版で作った台帳には無いので、無ければその場で作る。
 * 人も AI（Google ドライブ連携）も直接書き込むので、品名以外は空欄でもよい作りにする。
 */
function getWishSheet_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.WISHLIST);
  return sheet || formatWishSheet_(ensureSheet_(ss, CONFIG.SHEETS.WISHLIST, CONFIG.HEADERS.WISHLIST));
}

/** 日付の書式、購入済み欄のチェックボックス、AI 向けの説明（品名セルのメモ）を付ける。 */
function formatWishSheet_(sheet) {
  sheet.getRange('D:D').setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.getRange('G:G').setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 200);
  // insertCheckboxes() は既存の値を FALSE で上書きするので、入力規則だけ付けて値は残す。
  // 空のセルに規則を付けても最終行は伸びないので、AI の「最終行の下に追記」もずれない。
  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, WISH_COL.DONE + 1, sheet.getMaxRows() - 1, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
  sheet.getRange(1, 1).setNote(wishInstructions_().sheet);
  return sheet;
}

/** 購入済み欄の値を真偽に。チェックのほか「済」「✓」や日付も購入済みとみなし、空・FALSE・その他は未購入。 */
function isPurchased_(v) {
  if (v === true) return true;
  if (v instanceof Date) return true;
  if (typeof v === 'number') return v === 1;
  return CONFIG.PURCHASED_MARKS.indexOf(toStr_(v).toUpperCase()) !== -1;
}

/** 同じ物かどうかの判定用。全角半角・大文字小文字・空白の違いは同じ物とみなす。 */
function wishKey_(name) {
  return toStr_(name).normalize('NFKC').replace(/\s+/g, '').toLowerCase();
}

/**
 * AI や人が書き込んだ行を整える。
 * 追加日が空なら今の日時を入れ、購入済み欄を TRUE/FALSE に揃える（日付が書かれていれば購入日に移す）。
 * 値が変わる列だけ書き戻すので、品名・メモ・購入先には触らない。
 */
function normalizeWishSheet_(sheet, now) {
  const width = CONFIG.HEADERS.WISHLIST.length;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const added = values.map(function (r) { return [r[WISH_COL.ADDED_AT]]; });
  const done = values.map(function (r) { return [r[WISH_COL.DONE], r[WISH_COL.DONE_AT]]; });
  let addedChanged = false, doneChanged = false;
  values.forEach(function (r, i) {
    if (!toStr_(r[WISH_COL.NAME])) return;
    if (r[WISH_COL.ADDED_AT] === '' || r[WISH_COL.ADDED_AT] === null) {
      added[i][0] = now;
      addedChanged = true;
    }
    const mark = r[WISH_COL.DONE];
    const purchased = isPurchased_(mark);
    if (mark !== purchased) {
      done[i][0] = purchased;
      if (mark instanceof Date && toStr_(r[WISH_COL.DONE_AT]) === '') done[i][1] = mark;
      doneChanged = true;
    }
  });
  if (addedChanged) sheet.getRange(2, WISH_COL.ADDED_AT + 1, values.length, 1).setValues(added);
  if (doneChanged) sheet.getRange(2, WISH_COL.DONE + 1, values.length, 2).setValues(done);
  // シートの行数を超えて追記された行にもチェックボックスを付ける。
  sheet.getRange(2, WISH_COL.DONE + 1, lastRow - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
}

/**
 * まだ買っていない物を、同じ品名ごとに1件へまとめて返す。
 * 品目マスタの品目名か別名に一致すれば、定番商品と購入先（空欄のとき）をマスタから補う。
 */
function getOpenWishes_(master) {
  const byKey = {};
  master.forEach(function (m) {
    [m.name].concat(m.aliases).forEach(function (n) {
      const k = wishKey_(n);
      if (k && !byKey[k]) byKey[k] = m;
    });
  });
  const groups = [];
  const index = {};
  readTable_(getWishSheet_()).forEach(function (r) {
    const name = toStr_(r['品名']);
    if (!name || isPurchased_(r['購入済み'])) return;
    const key = wishKey_(name);
    let g = index[key];
    if (!g) {
      const m = byKey[key];
      g = index[key] = {
        key: key, name: name, memos: [], store: '', product: m ? m.product : '', masterName: m ? m.name : '',
        masterStore: m ? m.store : '', addedBy: [], addedAt: null,
      };
      groups.push(g);
    }
    const memo = toStr_(r['メモ']);
    if (memo && g.memos.indexOf(memo) === -1) g.memos.push(memo);
    if (!g.store) g.store = toStr_(r['購入先']);
    const by = toStr_(r['追加者']);
    if (by && g.addedBy.indexOf(by) === -1) g.addedBy.push(by);
    const at = r['追加日'];
    if (at instanceof Date && (!g.addedAt || at < g.addedAt)) g.addedAt = at;
  });
  return groups.map(function (g) {
    return {
      key: g.key, name: g.name, memo: g.memos.join(' / '), store: g.store || g.masterStore, product: g.product,
      masterName: g.masterName, addedBy: g.addedBy.join('、'), addedAt: g.addedAt,
    };
  });
}

/**
 * 買いたい物を追加する。同じ品名がまだ買っていない状態で残っていても行は足す
 * （買い物リストでは1件にまとまり、「買った」で全部まとめて購入済みになる）。
 * @param {Array<{name: string, memo: string, store: string}>} items
 * @param {string} addedBy 追加者として記録する名前
 * @return {number} 追加した件数
 */
function addWishes(items, addedBy) {
  const now = new Date();
  const rows = (items || []).map(function (it) {
    return [toStr_(it && it.name), toStr_(it && it.memo), toStr_(it && it.store), now, toStr_(addedBy), false, ''];
  }).filter(function (r) { return r[0]; });
  if (!rows.length) return 0;
  const sheet = getWishSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(start, WISH_COL.DONE + 1, rows.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  } finally {
    lock.releaseLock();
  }
  return rows.length;
}

/**
 * 指定した品名（wishKey_ で比べる）のまだ買っていない行を、すべて購入済みにする。
 * 行番号ではなく品名で探すので、その間に AI が行を足していてもずれない。
 * @return {number} 購入済みにした行数
 */
function markWishesPurchased(keys) {
  const want = {};
  (keys || []).forEach(function (k) { const key = wishKey_(k); if (key) want[key] = true; });
  const sheet = getWishSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    const range = sheet.getRange(2, WISH_COL.DONE + 1, lastRow - 1, 2);
    const names = sheet.getRange(2, WISH_COL.NAME + 1, lastRow - 1, 1).getValues();
    const done = range.getValues();
    const now = new Date();
    let count = 0;
    names.forEach(function (r, i) {
      if (!want[wishKey_(r[0])] || isPurchased_(done[i][0])) return;
      done[i] = [true, now];
      count++;
    });
    if (count) range.setValues(done);
    return count;
  } finally {
    lock.releaseLock();
  }
}

/**
 * AI に渡す指示文。sheet はスプレッドシートに直接書き込んでもらうとき、
 * paste はアプリに貼り付ける形で出力してもらうとき用。
 */
function wishInstructions_() {
  const ss = getSpreadsheet_();
  const sheetText = [
    'Google スプレッドシート「' + ss.getName() + '」（' + ss.getUrl() + '）の「' + CONFIG.SHEETS.WISHLIST + '」シートに、決めた買い物を追記してください。',
    '- 1行目は見出し。既存の最後の行のすぐ下に、1品1行で追加する',
    '- A列「品名」は必須。B列「メモ」にサイズ・色・数量・用途など、C列「購入先」に URL か店名（分からなければ空欄）',
    '- D列「追加日」は空欄でよい（自動で入る）。E列「追加者」にはあなたの名前（例: ChatGPT）',
    '- F列「購入済み」と G列「購入日」は空欄のままにする',
    '- ほかのシートと、既存の行は変更・削除しない',
  ].join('\n');
  const pasteText = [
    '決めた買い物を、家の在庫管理アプリに貼り付けられる形で出力してください。',
    '- 1行1品で「品名 / メモ / 購入先」の形式。区切りは前後に空白を入れたスラッシュ（ / ）',
    '- メモにはサイズ・色・数量・用途など、購入先には URL か店名。分からなければ省略してよい',
    '- 見出しや説明文、箇条書きの記号は付けず、その行だけをコードブロックで出力する',
  ].join('\n');
  return { sheet: sheetText, paste: pasteText };
}

// ---------- 買い物リスト ----------

/**
 * 数量の列（在庫合計・最低在庫・不足数）を数値に、更新日時の列を日時にする。
 * 以前の版は不足数の列に日時の書式を付けていたので、書き出すたびに付け直して既存シートも直す。
 */
function formatShoppingSheet_(sheet) {
  sheet.getRange('D:F').setNumberFormat('General');
  sheet.getRange('H:H').setNumberFormat('yyyy/mm/dd hh:mm');
}

/**
 * 「買い物リスト」シートを作り直す。
 * 先に最低在庫を下回る品目（種別「在庫不足」）、続けて「買いたい物」のまだ買っていない物（種別「単発」）を並べる。
 * 単発の物が在庫不足の品目と同じなら、行は分けずに在庫不足の行のメモに載せる。
 * 単発の物は購入済みにされるまで残る。在庫不足の物は在庫を数え直して最低在庫以上になると消える。
 * @return {{shortage: Array, wishes: Array}}
 */
function refreshShoppingList() {
  const master = getItemMaster();
  const inventory = getAllInventory_();
  const totals = {};
  const places = {};
  inventory.forEach(function (r) {
    totals[r.name] = (totals[r.name] || 0) + r.quantity;
    if (!places[r.name]) places[r.name] = [];
    if (places[r.name].indexOf(r.drawerName) === -1) places[r.name].push(r.drawerName);
  });
  const now = new Date();

  const wishSheet = getWishSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    normalizeWishSheet_(wishSheet, now);
  } finally {
    lock.releaseLock();
  }
  const allWishes = getOpenWishes_(master);

  const shortage = master
    .filter(function (m) { return m.minStock > 0 && (totals[m.name] || 0) < m.minStock; })
    .map(function (m) {
      const have = totals[m.name] || 0;
      const wish = allWishes.filter(function (w) { return w.masterName === m.name; })[0];
      return {
        name: m.name, product: m.product, store: m.store, have: have, min: m.minStock, shortage: m.minStock - have,
        places: (places[m.name] || []).join('、'), memo: wish ? wish.memo : '', wishKey: wish ? wish.key : '',
      };
    });
  const merged = {};
  shortage.forEach(function (s) { if (s.wishKey) merged[s.wishKey] = true; });
  const wishes = allWishes.filter(function (w) { return !merged[w.key]; });

  const rows = shortage.map(function (s) {
    return [s.name, s.product, s.store, s.have, s.min, s.shortage, s.places, now, '在庫不足', s.memo];
  }).concat(wishes.map(function (w) {
    return [w.name, w.product, w.store, '', '', '', '', now, '単発', w.memo];
  }));

  const sheet = getSheet_(CONFIG.SHEETS.SHOPPING);
  const headers = CONFIG.HEADERS.SHOPPING;
  // 以前の版のシートは「種別」「メモ」列が無いので、見出しも毎回書き直す。
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  formatShoppingSheet_(sheet);
  return {
    // wishKey があれば、その在庫不足の行は「買いたい物」にも載っている。「買った」でそちらを購入済みにできる。
    shortage: shortage,
    wishes: wishes.map(function (w) {
      return { key: w.key, name: w.name, memo: w.memo, store: w.store, product: w.product, addedBy: w.addedBy,
        addedAt: w.addedAt ? w.addedAt.toISOString() : '' };
    }),
  };
}
