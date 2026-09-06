/**
 * Claude API 呼び出し（Apps Script には公式 SDK がないため UrlFetchApp で直接叩く）。
 */

const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODELS_URL = 'https://api.anthropic.com/v1/models/';

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) {
    throw new Error('ANTHROPIC_API_KEY が未設定です。Web アプリの初回セットアップ画面か、Apps Script の「プロジェクトの設定 > スクリプト プロパティ」で登録してください。');
  }
  return key;
}

/**
 * Claude API 共通ヘッダー。
 * キーを特定ワークスペースに限定して作っていれば ANTHROPIC_WORKSPACE_ID は不要。
 * 複数ワークスペース対応のキーの場合はスクリプトプロパティ ANTHROPIC_WORKSPACE_ID（wrkspc_...）が必須。
 */
function claudeHeaders_(extra) {
  const headers = {
    'Authorization': 'Bearer ' + getApiKey_(),
    'anthropic-version': '2023-06-01',
  };
  const workspaceId = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_WORKSPACE_ID');
  if (workspaceId) headers['anthropic-workspace-id'] = workspaceId.trim();
  Object.keys(extra || {}).forEach(function (k) { headers[k] = extra[k]; });
  return headers;
}

/** 認識結果の JSON スキーマ。構造化出力で必ずこの形で返る。 */
const RECOGNITION_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['items', 'notes'],
  properties: {
    items: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['name', 'quantity', 'unit', 'remaining', 'confidence', 'is_new', 'product', 'note'],
        properties: {
          name: { type: 'string', description: '品目名（種類名）。品目マスタにあれば必ずその正規名。ブランド名や商品名は含めない。' },
          product: { type: 'string', description: 'パッケージから読み取れたブランド名・商品名（例: 「アタック 抗菌EX」）。読めなければ空文字。' },
          quantity: { type: 'integer', description: '写真から数えた個数。' },
          unit: { type: 'string', description: '単位（個・本・袋・ロール・箱など）。' },
          remaining: {
            type: 'string',
            enum: ['full', 'high', 'half', 'low', 'empty', 'unknown'],
            description: '開封済み容器の残量目安。未開封は full。',
          },
          confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
          is_new: { type: 'boolean', description: '品目マスタに無い新しい品目なら true。' },
          note: { type: 'string', description: '補足（見えにくい、前回と大きく違う理由など）。無ければ空文字。' },
        },
      },
    },
    notes: { type: 'string', description: '全体への補足や撮影方法の改善提案。無ければ空文字。' },
  },
};

function buildSystemPrompt_(drawer, master, previous) {
  const masterLines = master.length
    ? master.map(function (m) {
        return '- ' + m.name + ' | 単位: ' + (m.unit || '個') + (m.aliases.length ? ' | 別名: ' + m.aliases.join('、') : '');
      }).join('\n')
    : '（未登録）';

  const prevLines = previous.length
    ? previous.map(function (p) {
        return '- ' + p.name + ' × ' + p.quantity + ' ' + p.unit + (p.remaining ? '（' + p.remaining + '）' : '');
      }).join('\n')
    : '（記録なし）';

  return [
    'あなたは家庭の日用品在庫を写真から数える担当者です。',
    '渡される写真は「' + drawer.name + '」' + (drawer.location ? '（' + drawer.location + '）' : '') + 'という収納場所の中身です。',
    drawer.note ? 'この場所の説明: ' + drawer.note : '',
    '写真が複数ある場合は同じ場所を別の角度や区画で撮ったものなので、同じ物を二重に数えないでください。',
    '',
    'ルール:',
    '1. 写真にはっきり写っている物だけを数える。奥に隠れて見えない物は推測で足さず、見えている分だけ数えて confidence を low にする。',
    '2. 品目名（name）は種類名にする。下の品目マスタにある正規名を必ず使う（別名に一致した場合も正規名にする）。マスタに無い物は is_new を true にし、日本の家庭で通じる種類名を付ける（例: 「トイレットペーパー」「食器用洗剤 詰め替え」）。',
    '   ブランド名や商品名（例: 「アタック」「エリエール」）は name に入れず、パッケージから読めた範囲で product に書く。読めなければ空にする。',
    '3. 同じ品目は1行にまとめる。詰め替え用と本体ボトルは別の品目として扱う。',
    '4. quantity は個数（本・袋・ロール・箱・パックなど）。単位はマスタに合わせ、無い場合は写真から妥当な単位を選ぶ。',
    '5. 開封済みの容器は remaining に残量の目安を入れる。未開封は full。判断できなければ unknown。',
    '6. 前回の在庫は参考情報。写真に写っていない物は載せない。前回と大きく違う場合は note に理由を短く書く。',
    '7. 重ねてずらして置かれている物は、見えている端（エッジ）の数で数える。端が潰れて判別しにくい、5個以上重なっている、手前の物が奥の端を隠している場合は confidence を low にし、note に「重なりで端が数えにくい」などと書く。',
    '8. 暗い・ぼやけている・奥が見えないなど、撮り直した方がよい点があれば notes に短く書く。',
    '',
    '品目マスタ:',
    masterLines,
    '',
    '前回の在庫:',
    prevLines,
  ].filter(function (line) { return line !== null; }).join('\n');
}

/**
 * 写真を Claude に送って品目と数量を認識する。
 * @param {Array<{data: string, mediaType: string}>} images base64 と MIME タイプ
 * @param {Object} drawer 引き出し
 * @param {Array} master 品目マスタ
 * @param {Array} previous 前回の在庫
 * @return {{items: Array, notes: string, usage: Object, model: string}}
 */
function recognizeInventory(images, drawer, master, previous) {
  if (!images || !images.length) throw new Error('写真がありません。');
  if (images.length > CONFIG.MAX_PHOTOS) throw new Error('写真は ' + CONFIG.MAX_PHOTOS + ' 枚までです。');

  const content = images.map(function (img) {
    return {
      type: 'image',
      source: { type: 'base64', media_type: img.mediaType || 'image/jpeg', data: img.data },
    };
  });
  content.push({
    type: 'text',
    text: '上の写真' + (images.length > 1 ? images.length + '枚' : '') + 'に写っている日用品を数え、JSON で返してください。',
  });

  const body = {
    model: CONFIG.MODEL,
    max_tokens: CONFIG.MAX_TOKENS,
    thinking: { type: 'adaptive' },
    output_config: {
      effort: CONFIG.EFFORT,
      format: { type: 'json_schema', schema: RECOGNITION_SCHEMA },
    },
    // 安全分類器が拒否した場合に別モデルで自動再実行する（beta）。
    fallbacks: 'default',
    system: buildSystemPrompt_(drawer, master, previous),
    messages: [{ role: 'user', content: content }],
  };

  const json = callClaude_(body);
  const parsed = parseStructuredText_(json);

  const items = (parsed.items || []).map(function (it) {
    return {
      name: String(it.name || '').trim(),
      quantity: Math.max(0, Math.round(Number(it.quantity) || 0)),
      unit: String(it.unit || '').trim(),
      remaining: it.remaining || 'unknown',
      confidence: it.confidence || 'medium',
      isNew: it.is_new === true,
      product: String(it.product || '').trim(),
      note: String(it.note || '').trim(),
      needsReview: it.confidence === 'low' || it.is_new === true,
    };
  }).filter(function (it) { return it.name; });

  return {
    items: items,
    notes: String(parsed.notes || ''),
    usage: json.usage || {},
    model: json.model || CONFIG.MODEL,
  };
}

// ---------- 商品登録（品目マスタ作成用） ----------

const PRODUCT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['products', 'notes'],
  properties: {
    products: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['name', 'product', 'unit', 'category', 'aliases', 'exists_in_master', 'confidence', 'note'],
        properties: {
          name: { type: 'string', description: '品目名（種類名）。品目マスタにあればその正規名。ブランド名は含めない。' },
          product: { type: 'string', description: 'パッケージから読める正式な商品名。ブランド + 商品名 + 容量や枚数（例: 「キュキュット クリア除菌 詰替 770ml」）。' },
          unit: { type: 'string', description: '在庫を数えるときの単位（個・本・袋・ロール・箱・パックなど）。' },
          category: { type: 'string', enum: ['洗面', 'バス', 'トイレ', 'キッチン', '洗濯', '日用品', '電池', 'その他'] },
          aliases: { type: 'string', description: '種類名の言い換え候補をカンマ区切りで（例: 「食器洗剤,食器用洗剤 詰替」）。無ければ空文字。' },
          exists_in_master: { type: 'boolean', description: '同じ種類の品目が品目マスタに既にあるなら true。' },
          confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
          note: { type: 'string', description: '読み取れなかった部分や補足。無ければ空文字。' },
        },
      },
    },
    notes: { type: 'string' },
  },
};

function buildProductSystemPrompt_(master) {
  const masterLines = master.length
    ? master.map(function (m) {
        return '- ' + m.name + ' | 単位: ' + (m.unit || '個') + (m.product ? ' | 定番商品: ' + m.product : '') + (m.aliases.length ? ' | 別名: ' + m.aliases.join('、') : '');
      }).join('\n')
    : '（未登録）';
  return [
    'あなたは家庭の日用品の「品目マスタ」を作る担当者です。',
    '渡される写真には、登録したい商品がパッケージが読める距離で写っています。1枚に複数の商品が写っていることもあります。',
    '数量は数えなくてよい。写っている商品を1種類ずつ、登録用の情報として読み取ってください。',
    '',
    'ルール:',
    '1. name は種類名。既に品目マスタにある種類なら必ずその正規名を使い exists_in_master を true にする。無ければ日本の家庭で通じる種類名を付ける（例: 「食器用洗剤 詰め替え」）。詰め替え用と本体ボトルは別の種類にする。',
    '2. product はパッケージから読める正式な商品名。ブランド、商品名、容量や枚数まで含める。読めない部分は推測せず、読めた範囲だけ書いて note に「容量が読めない」などと書く。',
    '3. unit は在庫を数えるときの単位。category は最も近いものを選ぶ。',
    '4. aliases は、その種類名を人が呼びそうな別の言い方（略称、表記ゆれ）。',
    '5. 同じ商品が複数個写っていても1件にまとめる。',
    '',
    '品目マスタ:',
    masterLines,
  ].join('\n');
}

/**
 * 商品の写真から品目マスタ登録用の情報を読み取る。
 */
function recognizeProducts(images, master) {
  if (!images || !images.length) throw new Error('写真がありません。');
  if (images.length > CONFIG.MAX_PHOTOS) throw new Error('写真は ' + CONFIG.MAX_PHOTOS + ' 枚までです。');

  const content = images.map(function (img) {
    return { type: 'image', source: { type: 'base64', media_type: img.mediaType || 'image/jpeg', data: img.data } };
  });
  content.push({ type: 'text', text: '上の写真に写っている商品を、品目マスタ登録用に読み取って JSON で返してください。' });

  const body = {
    model: CONFIG.MODEL,
    max_tokens: CONFIG.MAX_TOKENS,
    thinking: { type: 'adaptive' },
    output_config: { effort: CONFIG.EFFORT, format: { type: 'json_schema', schema: PRODUCT_SCHEMA } },
    fallbacks: 'default',
    system: buildProductSystemPrompt_(master),
    messages: [{ role: 'user', content: content }],
  };

  const json = callClaude_(body);
  const parsed = parseStructuredText_(json);
  const products = (parsed.products || []).map(function (pr) {
    return {
      name: String(pr.name || '').trim(),
      product: String(pr.product || '').trim(),
      unit: String(pr.unit || '').trim() || '個',
      category: String(pr.category || '').trim(),
      aliases: String(pr.aliases || '').trim(),
      existsInMaster: pr.exists_in_master === true,
      confidence: pr.confidence || 'medium',
      note: String(pr.note || '').trim(),
    };
  }).filter(function (pr) { return pr.name; });

  return { products: products, notes: String(parsed.notes || ''), usage: json.usage || {}, model: json.model || CONFIG.MODEL };
}

/** Messages API を呼んで HTTP エラーと拒否を処理し、JSON を返す。 */
function callClaude_(body) {
  const response = UrlFetchApp.fetch(CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: claudeHeaders_({ 'anthropic-beta': 'server-side-fallback-2026-07-01' }),
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const status = response.getResponseCode();
  const text = response.getContentText();
  if (status !== 200) {
    throw new Error('Claude API エラー (HTTP ' + status + '): ' + summarizeApiError_(text));
  }
  const json = JSON.parse(text);
  if (json.stop_reason === 'refusal') {
    const detail = json.stop_details && json.stop_details.explanation ? json.stop_details.explanation : '';
    throw new Error('Claude が処理を拒否しました。' + detail);
  }
  if (json.stop_reason === 'max_tokens') {
    throw new Error('出力が長すぎて途中で切れました。写真の枚数を減らすか Config の MAX_TOKENS を増やしてください。');
  }
  return json;
}

/** 構造化出力の text ブロックを JSON として取り出す。 */
function parseStructuredText_(json) {
  const textBlock = (json.content || []).filter(function (b) { return b.type === 'text'; })[0];
  if (!textBlock) throw new Error('Claude から文字列の応答がありませんでした。');
  try {
    return JSON.parse(textBlock.text);
  } catch (e) {
    throw new Error('Claude の応答を JSON として読めませんでした: ' + textBlock.text.slice(0, 200));
  }
}

function summarizeApiError_(text) {
  try {
    const j = JSON.parse(text);
    if (j.error && j.error.message) {
      if (/anthropic-workspace-id/.test(j.error.message)) {
        return 'この API キーは複数ワークスペース対応のためワークスペースIDが必要です。'
          + 'キーを特定のワークスペースに限定して作り直すか、スクリプトプロパティ ANTHROPIC_WORKSPACE_ID に wrkspc_... を設定してください。';
      }
      if (j.error.type === 'authentication_error') {
        return 'API キーが無効か期限切れです。コンソールで新しいキーを発行して登録し直してください。';
      }
      return j.error.type + ': ' + j.error.message;
    }
  } catch (e) { /* JSON でなければそのまま */ }
  return text.slice(0, 300);
}

/** API キーとモデルの疎通確認。トークンを消費しない。 */
function testClaudeConnection() {
  const response = UrlFetchApp.fetch(CLAUDE_MODELS_URL + CONFIG.MODEL, {
    method: 'get',
    headers: claudeHeaders_(),
    muteHttpExceptions: true,
  });
  const status = response.getResponseCode();
  if (status !== 200) {
    throw new Error('HTTP ' + status + ': ' + summarizeApiError_(response.getContentText()));
  }
  const j = JSON.parse(response.getContentText());
  return { model: j.id, displayName: j.display_name };
}
