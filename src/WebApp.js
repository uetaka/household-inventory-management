/**
 * スマホ向け Web アプリのエントリポイントと、画面から呼ばれる API。
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('在庫スキャン')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** 画面の初期データ。 */
function apiGetInitialData() {
  return {
    drawers: getDrawers(),
    master: getItemMaster().map(function (m) { return { name: m.name, unit: m.unit }; }),
    config: {
      maxPhotos: CONFIG.MAX_PHOTOS,
      model: CONFIG.MODEL,
      remainingLabels: CONFIG.REMAINING_LABELS,
      confidenceLabels: CONFIG.CONFIDENCE_LABELS,
    },
  };
}

/** 引き出しを選んだときの現在在庫。 */
function apiGetDrawerInventory(drawerId) {
  const inv = getInventoryForDrawer(drawerId);
  return {
    items: inv,
    lastUpdated: inv.length ? inv[0].updatedAt : null,
  };
}

/**
 * 写真を認識する。まだシートには書き込まない。
 * @param {string} drawerId
 * @param {Array<{data: string, mediaType: string}>} photos
 */
function apiRecognize(drawerId, photos) {
  const drawer = getDrawer_(drawerId);
  const master = getItemMaster();
  const previous = getInventoryForDrawer(drawerId);

  const result = recognizeInventory(photos, drawer, master, previous);

  let photoUrls = [];
  if (CONFIG.SAVE_PHOTOS) {
    try {
      photoUrls = savePhotos_(drawer, photos);
    } catch (e) {
      Logger.log('写真の保存に失敗: ' + e.message);
    }
  }

  return {
    items: result.items,
    notes: result.notes,
    usage: {
      input: result.usage.input_tokens || 0,
      output: result.usage.output_tokens || 0,
    },
    model: result.model,
    photoUrls: photoUrls,
  };
}

/**
 * 確認済みの内容でシートを更新する。
 * @param {string} drawerId
 * @param {Array} items 画面で修正済みの品目
 * @param {Object} meta { recognized, photoUrls, usage, model }
 */
function apiCommit(drawerId, items, meta) {
  const drawer = getDrawer_(drawerId);
  const now = new Date();
  meta = meta || {};

  const cleaned = (items || []).map(function (it) {
    return {
      name: String(it.name || '').trim(),
      quantity: Math.max(0, Math.round(Number(it.quantity) || 0)),
      unit: String(it.unit || '').trim(),
      remaining: it.remaining || 'unknown',
      confidence: it.confidence || 'high',
      note: String(it.note || '').trim(),
      needsReview: it.needsReview === true,
    };
  }).filter(function (it) { return it.name; });

  const newNames = addNewItemsToMaster_(cleaned);
  replaceDrawerInventory_(drawer, cleaned, now);

  let user = '';
  try { user = Session.getActiveUser().getEmail(); } catch (e) { /* 取得できない設定もある */ }

  appendHistory_({
    at: now,
    drawerId: drawer.id,
    drawerName: drawer.name,
    photoUrls: meta.photoUrls || [],
    recognized: meta.recognized || null,
    committed: cleaned,
    user: user,
    inputTokens: meta.usage ? meta.usage.input : '',
    outputTokens: meta.usage ? meta.usage.output : '',
    model: meta.model || CONFIG.MODEL,
  });

  const shortage = refreshShoppingList();

  return {
    ok: true,
    count: cleaned.length,
    newItems: newNames,
    shortage: shortage,
    updatedAt: now.toISOString(),
  };
}

// ---------- 写真の保存 ----------

function getPhotoFolder_() {
  const ss = getSpreadsheet_();
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const existing = parent.getFoldersByName(CONFIG.PHOTO_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(CONFIG.PHOTO_FOLDER_NAME);
}

function savePhotos_(drawer, photos) {
  const folder = getPhotoFolder_();
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  return photos.map(function (p, i) {
    const ext = (p.mediaType || 'image/jpeg').split('/')[1] || 'jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mediaType || 'image/jpeg',
      drawer.id + '_' + stamp + '_' + (i + 1) + '.' + ext);
    return folder.createFile(blob).getUrl();
  });
}
