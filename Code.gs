/**
 * 1・2年目 秋旅 しおり（Google Apps Script 版）
 * データは、このスクリプトが入っているスプレッドシートの「data」シートに保存されます。
 * Gemini を使う場合は、プロジェクトの設定 → スクリプト プロパティに
 *   GEMINI_API_KEY : Google AI Studio で作った API キー
 * を追加してください。（任意で GEMINI_MODEL にモデル名を入れると変更できます）
 */

var SHEET_NAME = 'data';
var DAILY_LIMIT = 1000; // 念のための上限。実際の上限は Google の無料枠で決まる
// 上のモデルが今日の無料枠を使い切ったら、次のモデルで自動的にやり直す
var MODELS = ['gemini-2.5-flash', 'gemini-2.5-flash-lite'];

// 画面（index.html）は GitHub から読み込む。GitHub が更新されれば、貼り替えなしで反映される
var PAGE_URL = 'https://raw.githubusercontent.com/yoshimitsu-ao/akitabi-shiori/main/index.html';

function doGet() {
  return HtmlService.createHtmlOutput(page_())
    .setTitle('1・2年目 秋旅')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** GitHub の index.html を1分キャッシュして返す。取れないときは前回の分、それも無ければ手元の index を使う */
function page_() {
  var cache = CacheService.getScriptCache();
  var hit = getBig_(cache, 'page');
  if (hit) return hit;
  try {
    var res = UrlFetchApp.fetch(PAGE_URL + '?v=' + Math.floor(Date.now() / 60000), { muteHttpExceptions: true });
    var html = res.getResponseCode() === 200 ? res.getContentText('UTF-8') : '';
    if (html.indexOf('<html') >= 0) {
      putBig_(cache, 'page', html, 60);
      putBig_(cache, 'page-last', html, 21600);
      return html;
    }
  } catch (e) {}
  return getBig_(cache, 'page-last') || HtmlService.createHtmlOutputFromFile('index').getContent();
}

// キャッシュは1件100KBまでなので、2万5千文字ずつに分けてしまう
function putBig_(cache, key, s, sec) {
  var n = Math.ceil(s.length / 25000), o = {};
  for (var i = 0; i < n; i++) o[key + '_' + i] = s.substr(i * 25000, 25000);
  o[key + '_n'] = String(n);
  cache.putAll(o, sec);
}
function getBig_(cache, key) {
  var n = Number(cache.get(key + '_n') || 0);
  if (!n) return '';
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var got = cache.getAll(keys), out = '';
  for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return ''; out += got[keys[j]]; }
  return out;
}

function sheet_() {
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function load_() {
  var r = sheet_().getRange('B1:B5').getValues();
  var bbq = null, split = null, sched = null, memo = null, pack = null;
  try { pack = JSON.parse(r[4][0] || ''); } catch (e) {}
  if (!pack || !Array.isArray(pack.common)) pack = { common: ['厚手の上着', 'ニット帽・手袋', '焚き火用の上着', '着替え・部屋着', '充電器', '現金'], own: {}, done: {} };
  if (!pack.own || typeof pack.own !== 'object') pack.own = {};
  if (!pack.done || typeof pack.done !== 'object') pack.done = {};
  try { memo = JSON.parse(r[3][0] || ''); } catch (e) {}
  if (!memo || !Array.isArray(memo.notes)) memo = { album: '', notes: [] };
  if (!memo.avoid || typeof memo.avoid !== 'object') memo.avoid = {};
  try { bbq = JSON.parse(r[0][0] || ''); } catch (e) {}
  try { split = JSON.parse(r[1][0] || ''); } catch (e) {}
  try { sched = JSON.parse(r[2][0] || ''); } catch (e) {}
  if (!sched || !Array.isArray(sched.stops)) sched = { v: 3, stops: DEFAULT_STOPS.map(function (x, i) { var o = JSON.parse(JSON.stringify(x)); o.id = 's' + i; return o; }) };
  if (!sched.v) {
    // 2026/10/31 の変更：いちご狩りに会場名、チェックアウトもチェックインと同じ太字に
    sched.stops.forEach(function (x) {
      if (x.n === 'いちご狩り' && !x.s) x.s = '軽井沢ガーデンファーム';
      if (x.n === 'チェックアウト') x.k = 1;
    });
    sched.v = 2;
  }
  if (sched.v < 3) {
    // 2026/11/01 の変更：買い出しの地図をお店の名前で出す
    sched.stops.forEach(function (x) { if (/ツルヤ/.test(x.n) && !x.s) x.s = 'ツルヤ 軽井沢店'; });
    sched.v = 3;
  }
  if (!bbq || !Array.isArray(bbq.dishes)) bbq = { dishes: [] };
  if (!bbq.bought || typeof bbq.bought !== 'object') bbq.bought = {};
  if (!Array.isArray(bbq.wants)) bbq.wants = [{ n: 'クーラーボックス', ok: 0 }, { n: 'アウトドアチェア', ok: 0 }, { n: 'プロジェクター', ok: 0 }];
  if (!Array.isArray(bbq.extras)) bbq.extras = [];
  if (!bbq.adj || typeof bbq.adj !== 'object') bbq.adj = {};
  if (!split || !Array.isArray(split.members) || !Array.isArray(split.expenses)) split = { members: [], expenses: [] };
  if (!split.paid || typeof split.paid !== 'object') split.paid = {};
  return { bbq: bbq, split: split, sched: sched, memo: memo, pack: pack };
}

function save_(s) {
  sheet_().getRange('A1:B5').setValues([
    ['bbq', JSON.stringify(s.bbq)],
    ['split', JSON.stringify(s.split)],
    ['sched', JSON.stringify(s.sched)],
    ['memo', JSON.stringify(s.memo)],
    ['pack', JSON.stringify(s.pack)]
  ]);
}

function getState() {
  return load_();
}

function applyOp(o) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var s = load_();
    mutate_(s, o || {});
    save_(s);
    return s;
  } finally {
    lock.releaseLock();
  }
}

function str_(v, n) { return String(v == null ? '' : v).trim().slice(0, n); }
function servings_(n) { n = Math.round(Number(n) || 11); return Math.min(30, Math.max(1, n)); }
// スケジュールの最初の中身（スプレッドシートにまだ無いときだけ使う）
var DEFAULT_STOPS = [
  { d: 1, t: '7:00', n: '小田原 出発' },
  { d: 1, t: '11:15', n: 'いちご狩り', k: 1, s: '軽井沢ガーデンファーム', m: ['軽井沢ガーデンファームいちご園'] },
  { d: 1, t: '12:15', n: '軽井沢で昼食' },
  { d: 1, t: '14:00', n: '買い出し（ツルヤ）', s: 'ツルヤ 軽井沢店', m: ['ツルヤ 軽井沢店'] },
  { d: 1, t: '14:45', n: '白糸の滝', m: ['白糸の滝 軽井沢'] },
  { d: 1, t: '15:30', n: 'チェックイン', k: 1, m: ['Re to Li りとり 嬬恋村'] },
  { d: 1, t: '16:00', n: 'BBQ・焚き火', k: 1, s: '日没 16:40' },
  { d: 1, t: '21:00', n: '星空とまったり', s: '夜は0℃前後' },
  { d: 2, t: '', n: '朝ごはん？' },
  { d: 2, t: '10:00', n: 'チェックアウト', k: 1 },
  { d: 2, t: '10:10', n: '紅葉めぐり', k: 1, s: '候補', m: ['浅間牧場', '吾妻渓谷'] },
  { d: 2, t: '12:00', n: '昼食' },
  { d: 2, t: '13:00', n: '帰路' },
  { d: 2, t: '18:00', n: '小田原 着' }
];

function stop_(x) {
  var t = str_(x.t, 5);
  if (t && !/^\d{1,2}:\d{2}$/.test(t)) t = '';
  var m = (Array.isArray(x.m) ? x.m : []).map(function (q) { return str_(q, 300); }).filter(String).slice(0, 4);
  return { d: x.d == 2 ? 2 : 1, t: t, n: str_(x.n, 30), s: str_(x.s, 40), k: x.k ? 1 : 0, m: m };
}

var CATS = ['肉・魚', '野菜・果物', '主食・締め', '飲み物', '調味料', 'その他'];

function mutate_(s, o) {
  var members = s.split.members;
  switch (o.type) {
    case 'addDish': {
      var d = o.dish || {};
      var name = str_(d.name, 30);
      if (!name || s.bbq.dishes.some(function (x) { return x.name === name; })) return;
      var items = (Array.isArray(d.items) ? d.items : []).slice(0, 10).map(function (it) {
        return { n: str_(it.n, 30), q: str_(it.q, 12), c: CATS.indexOf(it.c) >= 0 ? it.c : 'その他' };
      }).filter(function (it) { return it.n; });
      var url = /^https?:\/\//.test(String(d.url || '')) ? str_(d.url, 500) : '';
      var by = str_(d.by, 20);
      s.bbq.dishes.push({ name: name, servings: servings_(d.servings), items: items, url: url, likes: 0, by: members.indexOf(by) >= 0 ? by : '', likedBy: [] });
      return;
    }
    case 'removeDish':
      s.bbq.dishes = s.bbq.dishes.filter(function (x) { return x.name !== o.name; });
      return;
    case 'setSteps': {
      var ss = (Array.isArray(o.steps) ? o.steps : []).map(function (t) { return str_(t, 120); }).filter(String).slice(0, 10);
      s.bbq.dishes.forEach(function (x) { if (x.name === o.name) x.steps = ss; });
      return;
    }
    case 'setServings':
      s.bbq.dishes.forEach(function (x) { if (x.name === o.name) x.servings = servings_(o.servings); });
      return;
    case 'addMember': {
      var m = str_(o.name, 20);
      if (m && members.indexOf(m) < 0) members.push(m);
      return;
    }
    case 'removeMember': {
      var used = s.split.expenses.some(function (x) { return x.p === o.name || (x.s || []).indexOf(o.name) >= 0; });
      if (!used) s.split.members = members.filter(function (x) { return x !== o.name; });
      return;
    }
    case 'addExpense': {
      var e = o.expense || {};
      var a = Math.round(Number(e.a));
      var p = str_(e.p, 20);
      var who = (Array.isArray(e.s) ? e.s : []).filter(function (x) { return members.indexOf(x) >= 0; });
      if (!(a > 0) || members.indexOf(p) < 0 || !who.length) return;
      s.split.expenses.push({ id: Utilities.getUuid(), t: str_(e.t, 30) || '支払い', a: a, p: p, s: who });
      return;
    }
    case 'setBought': {
      var n = str_(o.n, 30);
      if (!n) return;
      if (o.v) s.bbq.bought[n] = 1; else delete s.bbq.bought[n];
      return;
    }
    case 'addWant': {
      var wn = str_(o.n, 30);
      if (wn && s.bbq.wants.length < 30 && !s.bbq.wants.some(function (x) { return x.n === wn; })) s.bbq.wants.push({ n: wn, ok: 0 });
      return;
    }
    case 'removeWant':
      s.bbq.wants = s.bbq.wants.filter(function (x) { return x.n !== o.n; });
      return;
    case 'setWant':
      s.bbq.wants.forEach(function (x) { if (x.n === o.n) x.ok = o.v ? 1 : 0; });
      return;
    case 'wantBy': {
      var bw = str_(o.who, 20);
      if (members.indexOf(bw) < 0) return;
      s.bbq.wants.forEach(function (x) {
        if (x.n !== o.n) return;
        var by = (Array.isArray(x.by) ? x.by : []).filter(function (m) { return m !== bw; });
        if (o.v) by.push(bw);
        x.by = by; x.ok = by.length ? 1 : 0;
      });
      return;
    }
    case 'addExtra': {
      var en = str_(o.n, 30);
      if (en && s.bbq.extras.length < 40 && !s.bbq.extras.some(function (x) { return x.n === en; })) s.bbq.extras.push({ n: en, q: str_(o.q, 12), c: o.c === '飲み物' ? '飲み物' : 'その他' });
      return;
    }
    case 'setExtraQ':
      s.bbq.extras.forEach(function (x) { if (x.n === o.n) x.q = str_(o.q, 12); });
      return;
    case 'likeDish': {
      // 名前つきのいいね（who があるとき）。名前なしは前と同じく数だけ
      var lw = str_(o.who, 20);
      s.bbq.dishes.forEach(function (x) {
        if (x.name !== o.name) return;
        if (lw && members.indexOf(lw) >= 0) {
          var lb = (Array.isArray(x.likedBy) ? x.likedBy : []).filter(function (m) { return m !== lw; });
          if (o.d > 0) lb.push(lw);
          x.likedBy = lb;
        } else {
          x.likes = Math.max(0, (Number(x.likes) || 0) + (o.d > 0 ? 1 : -1));
        }
      });
      return;
    }
    case 'adjQ': {
      var an = str_(o.n, 30), ad = Math.round((Number(o.d) || 0) * 10) / 10;
      if (!an) return;
      ad = Math.max(-100000, Math.min(100000, ad));
      if (ad) s.bbq.adj[an] = ad; else delete s.bbq.adj[an];
      return;
    }
    case 'removeExtra':
      delete s.bbq.adj[o.n];
      s.bbq.extras = s.bbq.extras.filter(function (x) { return x.n !== o.n; });
      if (s.bbq.bought) delete s.bbq.bought[o.n];
      return;
    case 'addStop': {
      var ns = stop_(o.stop || {});
      if (!ns.n || s.sched.stops.length >= 40) return;
      ns.id = Utilities.getUuid();
      s.sched.stops.push(ns);
      return;
    }
    case 'updateStop': {
      var us = stop_(o.stop || {});
      if (!us.n) return;
      s.sched.stops.forEach(function (x, i) { if (x.id === o.id) { us.id = x.id; s.sched.stops[i] = us; } });
      return;
    }
    case 'removeStop':
      s.sched.stops = s.sched.stops.filter(function (x) { return x.id !== o.id; });
      return;
    case 'setPaid': {
      var pk = str_(o.k, 80);
      if (!pk) return;
      if (o.v) s.split.paid[pk] = 1; else delete s.split.paid[pk];
      return;
    }
    case 'setAvoid': {
      // 食べられないもの（a: アレルギー、d: 苦手）。その人の分をまるごと置き換える
      var aw = str_(o.who, 20);
      if (members.indexOf(aw) < 0) return;
      var clean = function (arr) { return (Array.isArray(arr) ? arr : []).map(function (t) { return str_(t, 20); }).filter(String).slice(0, 12); };
      var av = { a: clean(o.a), d: clean(o.d) };
      if (o.none) s.memo.avoid[aw] = { a: [], d: [], none: 1 };
      else if (av.a.length || av.d.length) s.memo.avoid[aw] = av; else delete s.memo.avoid[aw];
      return;
    }
    case 'setAlbum': {
      var au = str_(o.u, 500);
      s.memo.album = /^https:\/\//.test(au) ? au : '';
      return;
    }
    case 'addNote': {
      var nn = str_(o.n, 20), nt = str_(o.t, 80);
      if (!nn || !nt || s.memo.notes.length >= 60) return;
      s.memo.notes.push({ id: Utilities.getUuid(), n: nn, t: nt });
      return;
    }
    case 'removeNote':
      s.memo.notes = s.memo.notes.filter(function (x) { return x.id !== o.id; });
      return;
    case 'addPack': {
      var pn = str_(o.n, 30), pw = str_(o.who, 20);
      if (!pn) return;
      if (o.all) { if (s.pack.common.indexOf(pn) < 0 && s.pack.common.length < 40) s.pack.common.push(pn); return; }
      if (members.indexOf(pw) < 0) return;
      var own = s.pack.own[pw] = s.pack.own[pw] || [];
      if (own.indexOf(pn) < 0 && s.pack.common.indexOf(pn) < 0 && own.length < 30) own.push(pn);
      return;
    }
    case 'removePack':
      if (o.all) s.pack.common = s.pack.common.filter(function (x) { return x !== o.n; });
      else if (s.pack.own[o.who]) s.pack.own[o.who] = s.pack.own[o.who].filter(function (x) { return x !== o.n; });
      return;
    case 'setPack': {
      var dw = str_(o.who, 20), dn = str_(o.n, 30);
      if (!dw || !dn) return;
      var dd = s.pack.done[dw] = s.pack.done[dw] || {};
      if (o.v) dd[dn] = 1; else delete dd[dn];
      return;
    }
    case 'removeExpense':
      s.split.expenses = s.split.expenses.filter(function (x) { return x.id !== o.id; });
      return;
  }
}

function suggest(arg) {
  var cfg = gemini_();
  var name = str_(arg && arg.name, 30);
  var have = str_(arg && arg.have, 300);
  if (!name) return [];

  var prompt =
    'BBQで「' + name + '」を11人分作るための買い出しリストを作ってください。' +
    '宿にある塩・コショウ・炭・調理器具と、持参するお米（ご飯・パックご飯）は除きます。' +
    '調味料は、ほかの料理と使い回せる基本的なもの（醤油、油、砂糖、みりん、酒、バター、マヨネーズ、ケチャップなど）や、' +
    'すでにリストにあるもの（' + have + '）は入れず、その料理にしか使わないものだけにしてください。' +
    'JSONの配列だけを返してください。形式: [{"n":"食材名","q":"量（例: 1kg, 3本, 2パック）","c":"肉・魚|野菜・果物|主食・締め|調味料|その他 のどれか"}]。8品以内。';

  var arr = callGemini_(cfg, [{ text: prompt }]);
  return Array.isArray(arr) ? arr.slice(0, 8) : [];
}

/** Gemini の設定を読み、1日の回数を数える（材料・レシートの両方で共通） */
function gemini_() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY が設定されていません');
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var countKey = 'count_' + today;
  var count = Number(props.getProperty(countKey) || 0);
  if (count >= DAILY_LIMIT) throw new Error('今日の上限に達しました');
  props.setProperty(countKey, String(count + 1));
  var first = props.getProperty('GEMINI_MODEL');
  var models = first ? [first].concat(MODELS.filter(function (m) { return m !== first; })) : MODELS.slice();
  return { key: key, models: models };
}

/** Gemini に送り、返ってきた JSON を取り出す */
function callGemini_(cfg, parts) {
  var res = null;
  for (var k = 0; k < cfg.models.length; k++) {
    res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(cfg.models[k]) + ':generateContent',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-goog-api-key': cfg.key },
        muteHttpExceptions: true,
        payload: JSON.stringify({
          contents: [{ role: 'user', parts: parts }],
          generationConfig: { responseMimeType: 'application/json', temperature: 0.2 }
        })
      }
    );
    var code = res.getResponseCode();
    if (code === 200) break;
    if (code !== 429 && code !== 404 && code !== 503) throw new Error('Gemini エラー ' + code);
  }
  if (res.getResponseCode() === 429) throw new Error('今日の Gemini の無料枠を使い切りました');
  if (res.getResponseCode() !== 200) throw new Error('Gemini エラー ' + res.getResponseCode());
  var j = JSON.parse(res.getContentText());
  var out = (((j.candidates || [])[0] || {}).content || {}).parts || [];
  var text = out.map(function (p) { return p.text || ''; }).join('');
  try { return JSON.parse(text); } catch (e) {
    var m = text.match(/[\[{][\s\S]*[\]}]/);
    return m ? JSON.parse(m[0]) : null;
  }
}

/** レシートの写真から、店名と合計金額を読み取る */
function readReceipt(arg) {
  var data = String((arg && arg.data) || '');
  if (!data || data.length > 4000000) throw new Error('画像を読み込めませんでした');
  var mime = /^image\/(jpeg|png|webp)$/.test(arg.mime) ? arg.mime : 'image/jpeg';
  var cfg = gemini_();
  var prompt =
    'これは買い物のレシートの写真です。店名と、支払った合計金額（税込）を読み取ってください。' +
    'JSONだけを返してください。形式: {"t":"店名（短く。読めなければ空）","a":合計金額の整数（読めなければ0）}';
  var r = callGemini_(cfg, [{ inline_data: { mime_type: mime, data: data } }, { text: prompt }]) || {};
  var a = Math.round(Number(String(r.a == null ? '' : r.a).replace(/[^0-9.]/g, '')));
  return { t: str_(r.t, 30), a: a > 0 ? a : 0 };
}

/** レシピ（URL かスクショ）から、料理名と11人分の材料を出す */
function recipe(arg) {
  var url = str_(arg && arg.url, 500), data = String((arg && arg.data) || ''), have = str_(arg && arg.have, 300);
  var parts = [];
  if (url) {
    if (!/^https?:\/\//.test(url)) throw new Error('URLが正しくありません');
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (res.getResponseCode() >= 400) throw new Error('ページを開けませんでした');
    parts.push({ text: 'レシピページの内容:\n' + recipeText_(res.getContentText()) });
  } else if (data) {
    if (data.length > 6000000) throw new Error('画像が大きすぎます');
    parts.push({ inline_data: { mime_type: 'image/jpeg', data: data } });
  } else {
    return null;
  }
  var cfg = gemini_();
  parts.push({ text:
    'これは料理のレシピです。料理名と、BBQでこの料理を11人分作るための買い出しリストを作ってください。' +
    'レシピの材料の量を11人分に換算してください。宿にある塩・コショウ・炭・調理器具と、持参するお米（ご飯・パックご飯）は除きます。' +
    '調味料は、ほかの料理と使い回せる基本的なもの（醤油、油、砂糖、みりん、酒、バター、マヨネーズ、ケチャップなど）や、' +
    'すでにリストにあるもの（' + have + '）は入れないでください。' +
    'JSONだけを返してください。形式: {"name":"料理名（短く）","items":[{"n":"食材名","q":"量（例: 1kg, 3本, 2パック）","c":"肉・魚|野菜・果物|主食・締め|調味料|その他 のどれか"}]}。10品以内。' });
  var r = callGemini_(cfg, parts) || {};
  var items = (Array.isArray(r.items) ? r.items : []).slice(0, 10).map(function (it) {
    return { n: str_(it && it.n, 30), q: str_(it && it.q, 12), c: CATS.indexOf(it && it.c) >= 0 ? it.c : 'その他' };
  }).filter(function (it) { return it.n; });
  return { name: str_(r.name, 30), items: items };
}

function recipeText_(html) {
  var title = ((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || '').trim();
  var re = /<script[^>]*application\/ld\+json[^>]*>([\s\S]*?)<\/script>/gi, m, ld = [];
  while ((m = re.exec(html))) if (/Recipe/.test(m[1])) ld.push(m[1]);
  if (ld.length) return ('タイトル: ' + title + '\n' + ld.join('\n')).slice(0, 15000);
  var body = html.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ');
  return ('タイトル: ' + title + '\n' + body).slice(0, 15000);
}

/** 嬬恋村の天気（Open-Meteo・無料・キー不要）。1時間キャッシュ */
function weather() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('weather');
  if (hit) return JSON.parse(hit);
  var url = 'https://api.open-meteo.com/v1/forecast?latitude=36.47&longitude=138.53&timezone=Asia%2FTokyo' +
    '&daily=weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max&start_date=2026-11-07&end_date=2026-11-08';
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var out = [];
  if (res.getResponseCode() === 200) {
    var d = JSON.parse(res.getContentText()).daily || {};
    (d.time || []).forEach(function (t, i) {
      out.push({ d: t, c: d.weather_code[i], hi: d.temperature_2m_max[i], lo: d.temperature_2m_min[i], p: d.precipitation_probability_max[i] });
    });
  }
  cache.put('weather', JSON.stringify(out), 3600);
  return out;
}

/** 料理の作り方（手順）を Gemini に出してもらう */
function steps(arg) {
  var name = str_(arg && arg.name, 30), items = str_(arg && arg.items, 400);
  if (!name) return [];
  var cfg = gemini_();
  var prompt = '外のBBQコンロ（炭火・網）と焚き火、宿のキッチン（IHコンロ・鍋・フライパン）が使えます。' +
    '「' + name + '」を11人分作る手順を、短い日本語の文で4〜7個にまとめてください。材料: ' + items + '。' +
    '火加減や時間の目安も入れてください。JSONの文字列の配列だけを返してください。';
  var arr = callGemini_(cfg, [{ text: prompt }]);
  return Array.isArray(arr) ? arr.map(function (t) { return str_(t, 120); }).filter(String).slice(0, 8) : [];
}
