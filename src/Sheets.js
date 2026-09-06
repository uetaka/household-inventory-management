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

function getDrawers() {
  return readTable_(getSheet_(CONFIG.SHEETS.DRAWERS))
    .map(function (r) {
      return {
        id: toStr_(r['引き出しID']),
        name: toStr_(r['名前']),
        location: toStr_(r['場所']),
        note: toStr_(r['備考']),
      };
    })
    .filter(function (d) { return d.id && d.name; });
}

function getDrawer_(drawerId) {
  const d = getDrawers().filter(function (x) { return x.id === drawerId; })[0];
  if (!d) throw new Error('引き出しID「' + drawerId + '」が見つかりません。');
  return d;
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

// ---------- 買い物リスト ----------

/** 最低在庫を下回る品目を「買い物リスト」シートに書き出す。 */
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
  const rows = master
    .filter(function (m) { return m.minStock > 0 && (totals[m.name] || 0) < m.minStock; })
    .map(function (m) {
      const have = totals[m.name] || 0;
      return [m.name, m.product, m.store, have, m.minStock, m.minStock - have, (places[m.name] || []).join('、'), now];
    });

  const sheet = getSheet_(CONFIG.SHEETS.SHOPPING);
  const headers = CONFIG.HEADERS.SHOPPING;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  return rows.map(function (r) { return { name: r[0], product: r[1], store: r[2], have: r[3], min: r[4], shortage: r[5] }; });
}
