/*
 * MailPolish — 主程序：界面与流程
 *
 * 「边写边查」的原理（Outlook 不提供光标位置和打字事件，所以这样近似）：
 *   每 1 秒读一次草稿 → 和上一次对比，变化的位置 ≈ 你正在打字的地方 → 找到那一句
 *   → 你停顿约 2 秒后，把这句（连同上下文）发给 AI → 结果显示在侧边栏，点「替换」写回邮件
 *   查过的句子会记住结果，不会重复花钱。
 *
 * 每条建议是一张「卡片」，卡片记住自己的状态：
 *   待处理 → 点「替换」/「用这个」一键写回 → 已替换（可「撤销」，可「换个写法」再换）
 *
 * v1.2 整封邮件：
 *   - 全文检查：「全部改正」「全部用更专业的写法」「撤销全部」，一次写回邮件
 *   - 改写整封（或选中多段）：逐段改写，每段原地替换，加粗、项目符号、上标等格式不变
 *   - 数据行（Tab 对齐的 SKU 清单）和签名不送去改；数字或术语有变化的建议不会被「全部替换」自动套用
 */
(function () {
  'use strict';

  const VERSION = '1.2.0';
  const TU = window.TextUtil;
  const AI = window.MailPolishAI;
  const HOSTS = window.MailPolishHost;
  const MANIFEST = window.MailPolishManifest;

  // ---------------------------------------------------------------------------
  // 本机存储（Key、设置、用量）。存储不可用时也能正常运行，只是不会记住。
  // ---------------------------------------------------------------------------
  const store = {
    prefix: 'mailpolish.',
    init(partitionKey) { this.prefix = (partitionKey ? String(partitionKey) + '|' : '') + 'mailpolish.'; },
    get(key, fallback) {
      try {
        const v = window.localStorage.getItem(this.prefix + key);
        return v == null ? fallback : JSON.parse(v);
      } catch (e) { return fallback; }
    },
    set(key, value) {
      try { window.localStorage.setItem(this.prefix + key, JSON.stringify(value)); return true; } catch (e) { return false; }
    }
  };

  let settings = Object.assign({}, AI.DEFAULT_SETTINGS);
  let host = null;
  let inOutlook = false;

  // ---------------------------------------------------------------------------
  // 小工具
  // ---------------------------------------------------------------------------
  const $ = (sel, rootEl) => (rootEl || document).querySelector(sel);

  /** 创建元素。所有动态文字都用 textContent，不会把 AI 返回的内容当成 HTML 执行。 */
  function el(tag, props) {
    const node = document.createElement(tag);
    if (props) {
      for (const k of Object.keys(props)) {
        const v = props[k];
        if (v == null || v === false) continue;
        if (k === 'class') node.className = v;
        else if (k === 'text') node.textContent = v;
        else if (k.slice(0, 2) === 'on' && typeof v === 'function') node.addEventListener(k.slice(2), v);
        else node.setAttribute(k, v === true ? '' : String(v));
      }
    }
    for (let i = 2; i < arguments.length; i++) append(node, arguments[i]);
    return node;
  }
  function append(node, child) {
    if (child == null || child === false) return;
    if (Array.isArray(child)) { child.forEach(c => append(node, c)); return; }
    node.appendChild(child instanceof Node ? child : document.createTextNode(String(child)));
  }
  function clear(node) { while (node && node.firstChild) node.removeChild(node.firstChild); }

  /** 小按钮 */
  function btn(label, kind, onclick, disabled) {
    return el('button', { class: 'btn small' + (kind ? ' ' + kind : ''), onclick, disabled: disabled ? true : null }, label);
  }

  function diffNodes(a, b) {
    return TU.wordDiff(a, b).map(op => (op.op === 'eq' ? op.text : el(op.op === 'del' ? 'del' : 'ins', {}, op.text)));
  }

  let toastTimer = null;
  function toast(msg, ms) {
    const t = $('#toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, ms || 3500);
  }

  async function copyText(text) {
    try {
      if (navigator.clipboard && window.isSecureContext) { await navigator.clipboard.writeText(text); return true; }
    } catch (e) { /* 换用下面的方法 */ }
    try {
      const ta = el('textarea', { style: 'position:fixed;left:-9999px;top:0' });
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand('copy');
      ta.remove();
      return ok;
    } catch (e) { return false; }
  }
  async function copyWithToast(text, okMsg) {
    toast((await copyText(text)) ? (okMsg || '已复制') : '复制失败，请手动选择文字');
  }

  function setBusy(button, busyText) {
    const label = button.textContent;
    button.disabled = true;
    clear(button);
    append(button, [el('span', { class: 'spinner' }), ' ' + busyText]);
    return () => { button.disabled = false; button.textContent = label; };
  }

  function errText(e) { return e && e.message ? e.message : String(e); }
  /** 统一换行：Windows 的 \r\n、Word 的手动换行（\v）等都当成 \n */
  function toLF(s) { return String(s || '').replace(/\r\n?|[\v\u2028\u2029]/g, '\n'); }

  // ---------------------------------------------------------------------------
  // 用量与费用（估算）
  // ---------------------------------------------------------------------------
  function pad(n) { return String(n).padStart(2, '0'); }
  function localDay() { const d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }

  function loadUsage() {
    const day = localDay();
    const month = day.slice(0, 7);
    const saved = store.get('usage', {});
    const u = Object.assign({ day, calls: 0, cost: 0, month, monthCalls: 0, monthCost: 0 }, saved && typeof saved === 'object' ? saved : {});
    if (u.day !== day) { u.day = day; u.calls = 0; u.cost = 0; }
    if (u.month !== month) { u.month = month; u.monthCalls = 0; u.monthCost = 0; }
    return u;
  }
  function recordUsage(info) {
    const u = loadUsage();
    const c = Number(info && info.cost) || 0;
    u.calls += 1; u.cost += c; u.monthCalls += 1; u.monthCost += c;
    store.set('usage', u);
    renderUsage();
  }
  function fmtUSD(x) { return x < 0.01 ? '$' + x.toFixed(4) : '$' + x.toFixed(2); }
  function renderUsage() {
    const u = loadUsage();
    $('#usage').textContent = '今天 ' + u.calls + ' 次 · 约 ' + fmtUSD(u.cost);
    $('#usageDetail').textContent = '用量（估算）：今天 ' + u.calls + ' 次 ≈ ' + fmtUSD(u.cost) + '；本月 ' + u.monthCalls + ' 次 ≈ ' + fmtUSD(u.monthCost) +
      '。按 DeepSeek 公布的价格估算，实际以 DeepSeek 账单为准。';
  }

  async function callAI(opts) {
    const u = loadUsage();
    const cap = Number(settings.dailyCap) || AI.DEFAULT_SETTINGS.dailyCap;
    if (u.calls >= cap) throw AI.makeError('cap');
    return AI.chatJSON(Object.assign({}, opts, { settings, onUsage: recordUsage }));
  }

  // ---------------------------------------------------------------------------
  // 设置
  // ---------------------------------------------------------------------------
  function loadSettings() {
    const saved = store.get('settings', {});
    settings = Object.assign({}, AI.DEFAULT_SETTINGS, saved && typeof saved === 'object' ? saved : {});
    if (settings.replaceMode !== 'confirm') settings.replaceMode = 'direct';
  }
  function saveSettings() { store.set('settings', settings); }

  function fillSettingsForm() {
    const modelSel = $('#model');
    clear(modelSel);
    AI.MODELS.forEach(m => modelSel.appendChild(el('option', { value: m.id }, m.label)));
    const known = AI.MODELS.some(m => m.id === settings.model);
    modelSel.value = known ? settings.model : AI.MODELS[0].id;
    $('#customModel').value = known ? '' : (settings.model || '');
    $('#apiKey').value = settings.apiKey || '';
    $('#apiKey').type = 'password';
    $('#btnShowKey').textContent = '显示';
    $('#variant').value = settings.variant === 'UK' ? 'UK' : 'US';
    $('#replaceMode').value = settings.replaceMode === 'confirm' ? 'confirm' : 'direct';
    $('#about').value = settings.about || '';
    $('#terms').value = settings.terms || '';
    $('#pauseMs').value = String(settings.pauseMs || 2000);
    if (!$('#pauseMs').value) $('#pauseMs').value = '2000';
    $('#dailyCap').value = String(settings.dailyCap || AI.DEFAULT_SETTINGS.dailyCap);
    $('#baseUrl').value = settings.baseUrl || AI.DEFAULT_SETTINGS.baseUrl;
    $('#testResult').textContent = '';
    $('#versionInfo').textContent = 'MailPolish 版本 ' + VERSION;
  }

  /** 读取表单；格式不对时返回 { error } */
  function readSettingsForm() {
    const baseUrl = $('#baseUrl').value.trim() || AI.DEFAULT_SETTINGS.baseUrl;
    if (!/^https:\/\/[^\s/]+/i.test(baseUrl)) return { error: 'API 地址必须以 https:// 开头。' };
    let cap = parseInt($('#dailyCap').value, 10);
    if (!Number.isFinite(cap) || cap < 10) cap = 10;
    if (cap > 10000) cap = 10000;
    return {
      value: {
        apiKey: $('#apiKey').value.trim(),
        model: $('#customModel').value.trim() || $('#model').value,
        variant: $('#variant').value === 'UK' ? 'UK' : 'US',
        replaceMode: $('#replaceMode').value === 'confirm' ? 'confirm' : 'direct',
        about: $('#about').value.trim(),
        terms: $('#terms').value.trim(),
        pauseMs: parseInt($('#pauseMs').value, 10) || 2000,
        dailyCap: cap,
        baseUrl
      }
    };
  }

  function openSettings() {
    fillSettingsForm();
    renderUsage();
    $('#mainView').hidden = true;
    $('#settingsView').hidden = false;
  }
  function closeSettings() {
    $('#settingsView').hidden = true;
    $('#mainView').hidden = false;
  }
  function onSaveSettings() {
    const r = readSettingsForm();
    if (r.error) { toast(r.error, 5000); return; }
    Object.assign(settings, r.value);
    saveSettings();
    toast('已保存');
    if (settings.apiKey) hideNotice(); else showSetupNotice();
    closeSettings();
    // 设置变了（如模型、术语），之前出错的句子可以重新检查
    retryErrors(true);
    applyLiveToggle();
  }
  async function onTestConnection() {
    const r = readSettingsForm();
    const out = $('#testResult');
    if (r.error) { out.textContent = '✗ ' + r.error; return; }
    if (!r.value.apiKey) { out.textContent = '✗ 请先填写 API Key'; return; }
    out.textContent = '测试中…';
    try {
      const res = await AI.testConnection(Object.assign({}, settings, r.value));
      recordUsage({ cost: res.cost });
      out.textContent = '✓ 连接成功（模型：' + r.value.model + '）';
    } catch (e) {
      out.textContent = '✗ ' + AI.describeError(e);
    }
  }
  function onClearKey() {
    settings.apiKey = '';
    saveSettings();
    $('#apiKey').value = '';
    toast('已清除这台电脑上保存的 Key');
    showSetupNotice();
  }

  // ---------------------------------------------------------------------------
  // 顶部提示条（也用来做"没有卡片可挂靠"时的确认）
  // ---------------------------------------------------------------------------
  let noticeAskFinish = null; // 提示条上正在等回答的确认
  function showNotice(message, actions) {
    if (noticeAskFinish) { const f = noticeAskFinish; noticeAskFinish = null; f(-1); } // 新提示顶掉了旧的确认：当作取消
    const n = $('#notice');
    clear(n);
    n.appendChild(el('div', {}, message));
    if (actions && actions.length) {
      n.appendChild(el('div', { class: 'actions' }, actions.map(([label, fn], i) => el('button', { class: 'btn small' + (i === 0 ? ' primary' : ''), onclick: fn }, label))));
    }
    n.hidden = false;
  }
  function hideNotice() {
    if (noticeAskFinish) { const f = noticeAskFinish; noticeAskFinish = null; f(-1); return; } // 提示条上还有没回答的确认：当作取消
    $('#notice').hidden = true;
  }
  function showSetupNotice() {
    showNotice('还没有填写 DeepSeek API Key。填好后就可以开始检查。', [['去设置', openSettings]]);
  }

  // 等待用户回答的确认框（切换邮件时要全部取消，否则写入锁会一直占着）
  const pendingAsks = new Set();
  function cancelPendingAsks() {
    Array.from(pendingAsks).forEach(fn => fn(-1));
  }
  function askVia(show, hide) {
    return (message, labels) => new Promise(resolve => {
      const finish = i => { pendingAsks.delete(finish); hide(); resolve(i); };
      pendingAsks.add(finish);
      show(message, labels, finish);
    });
  }
  const askInNotice = askVia(
    (message, labels, finish) => { showNotice(message, labels.map((label, i) => [label, () => finish(i)])); noticeAskFinish = finish; },
    () => { noticeAskFinish = null; hideNotice(); });

  /**
   * 开始新的检查/改写前调用：还在等你确认的替换当作取消（那次还没写回邮件）；
   * 如果正在写回邮件，就请你稍等，免得旧卡片被清掉后写回的结果对不上。
   */
  function abandonPendingWrites() {
    if (pendingAsks.size) { cancelPendingAsks(); return true; }
    if (writeLock) { toast('正在写回邮件，请稍等一下再试'); return false; }
    return true;
  }

  // ---------------------------------------------------------------------------
  // 写入锁：同一时间只做一件写回邮件的事，避免两次写回互相覆盖
  // ---------------------------------------------------------------------------
  let writeLock = false;
  let itemSeq = 0; // 每切换一封邮件 +1，用来丢弃上一封邮件的迟到结果
  function beginWrite() {
    if (writeLock) { toast('上一个替换还没完成，请稍等一下'); return false; }
    writeLock = true;
    return true;
  }
  function endWrite() { writeLock = false; }

  // ---------------------------------------------------------------------------
  // 替换 / 撤销（统一处理"有图片时先确认"的设置）
  // ---------------------------------------------------------------------------
  const IMAGE_CONFIRM_TEXT = '这封邮件里有图片（比如签名或来信里的 logo）。直接替换需要把整个正文写回 Outlook。' +
    '替换后我会自动检查，图片变少就自动恢复原样。也可以先在邮件里选中这句话再替换，这样不会动到其它内容。';
  const IMAGE_CONFIRM_MANY = '这封邮件里有图片（比如签名或来信里的 logo）。一次改多处需要把整个正文写回 Outlook。' +
    '写回后我会自动检查，图片变少就自动恢复原样。';
  const PER_ITEM_REASONS = new Set(['not_found', 'multiple', 'crosses_paragraph', 'empty']);

  function reasonText(reason, action) {
    const verb = action === 'undo' ? '撤销' : '替换';
    return ({
      not_found: '在正文里没找到这句（可能刚刚又改过）。可以在邮件里选中它，再点「' + verb + '」。',
      multiple: '这句话在你写的部分里出现了不止一次。请先在邮件里选中要改的那一处，再点「' + verb + '」。',
      crosses_paragraph: '这段文字跨了段落，没法自动定位。请先在邮件里选中它，再点「' + verb + '」。',
      empty: '没有可' + verb + '的内容。',
      has_images: '请先在邮件里选中这句话，再点「' + verb + '」。',
      image_lost: '写回正文后检测到图片变少了，已经自动恢复原样。请先在邮件里选中这句话再' + verb + '，或者在设置里改成「有图片时先确认」。',
      text_lost: '写回正文后检测到内容不完整，已经自动恢复原样。请先在邮件里选中这句话再' + verb + '。',
      changed_facts: '这条建议改动了数字、日期或术语，批量替换时没有自动套用。请核对后点这张卡片上的「替换」。',
      suspicious: '这段改写和原文几乎没有相同的词，可能对应错了段落，批量替换时没有自动套用。请核对后点这张卡片上的「替换」。',
      item_changed: '你切换到了另一封邮件，这次没有' + verb + '。'
    })[reason] || verb + '没有成功。';
  }

  /** ask(message, labels) → Promise<index> ：由调用方决定在卡片里问还是在顶部提示条里问 */
  async function replaceWithPolicy(from, to, ask, ctx) {
    const direct = settings.replaceMode !== 'confirm';
    let r = await host.replaceSentence(from, to, { allowFullBody: direct, ctx });
    if (!r.ok && r.reason === 'has_images') {
      const choice = await ask(IMAGE_CONFIRM_TEXT, ['直接替换', '我先去选中']);
      if (choice !== 0) return { ok: false, cancelled: true };
      r = await host.replaceSentence(from, to, { allowFullBody: true, ctx });
    }
    return r;
  }

  async function revertWithPolicy(entry, ask) {
    const direct = settings.replaceMode !== 'confirm';
    let r = await host.revertEntry(entry, { allowFullBody: direct });
    if (!r.ok && r.reason === 'has_images') {
      const choice = await ask(IMAGE_CONFIRM_TEXT, ['直接撤销', '我先去选中']);
      if (choice !== 0) return { ok: false, cancelled: true };
      r = await host.revertEntry(entry, { allowFullBody: true });
    }
    return r;
  }

  /** 一次改多处：pairs = [{ original, replacement, occurrence? }] */
  async function replaceManyWithPolicy(pairs, ask, record, okLabel) {
    const direct = settings.replaceMode !== 'confirm';
    let r = await host.replaceMany(pairs, { allowFullBody: direct, record });
    if (!r.ok && r.reason === 'has_images') {
      const choice = await ask(IMAGE_CONFIRM_MANY, [okLabel, '取消']);
      if (choice !== 0) return { ok: false, cancelled: true };
      r = await host.replaceMany(pairs, { allowFullBody: true, record });
    }
    return r;
  }

  /** 替换成功后：新文字算"已处理"，不再重复检查 */
  function markApplied(text) {
    TU.segmentSentences(text).forEach(s => {
      const k = TU.sentenceKey(s.text);
      live.cache.set(k, { status: 'applied', result: null, text: s.text, at: Date.now(), source: 'live' });
      live.present.add(k);
    });
  }
  function markPresent(text) {
    TU.segmentSentences(text).forEach(s => live.present.add(TU.sentenceKey(s.text)));
  }
  function isPresent(text) {
    const segs = TU.segmentSentences(text);
    return segs.length > 0 && segs.every(s => live.present.has(TU.sentenceKey(s.text)));
  }

  // ---------------------------------------------------------------------------
  // 建议卡片
  // ---------------------------------------------------------------------------
  const cards = new Map();          // id -> 卡片
  const liveCardByKey = new Map();  // 原句 -> 卡片 id（边写边查）
  let fullCardIds = [];             // 全文检查的卡片
  let rewriteCardId = null;         // 改写结果卡片（选中一句/一段时的多版本）
  let cardSeq = 0;

  // 一组卡片（全文检查 / 逐段改写）的批量操作状态
  const groups = {
    full: { busy: false, confirm: null },
    rwb: { busy: false, confirm: null }
  };
  function groupOf(card) {
    if (card.source === 'full') return groups.full;
    if (card.source === 'rewriteBlock') return groups.rwb;
    return null;
  }
  function groupBusy(card) { const g = groupOf(card); return !!(g && g.busy); }

  function newCard(source, original, extra) {
    const card = Object.assign({
      id: 'c' + (++cardSeq), source, key: TU.sentenceKey(original), original, result: null,
      status: 'open', current: original, history: [], alts: null, altsSeen: [], altsRound: 0,
      altsLoading: false, altsError: '', busy: false, confirm: null, paragraph: '', at: Date.now(),
      ctx: null, failNote: ''
    }, extra || {});
    cards.set(card.id, card);
    return card;
  }

  function upsertLiveCard(text, result, paragraph) {
    const key = TU.sentenceKey(text);
    const existing = cards.get(liveCardByKey.get(key));
    if (existing) {
      if (existing.status === 'open') { existing.result = result; existing.at = Date.now(); }
      return existing;
    }
    const card = newCard('live', text, { result, paragraph: paragraph || '' });
    liveCardByKey.set(key, card.id);
    return card;
  }

  function rerender() {
    renderLive();
    renderFullCards();
    renderRewrite();
  }

  /** 在卡片里问一个问题（状态保存在卡片上，界面刷新也不会丢） */
  function askInCard(card) {
    return askVia((message, labels, finish) => { card.confirm = { message, labels, resolve: finish }; rerender(); },
      () => { card.confirm = null; rerender(); });
  }
  /** 在一组卡片上方的操作栏里问 */
  function askInGroup(g) {
    return askVia((message, labels, finish) => { g.confirm = { message, labels, resolve: finish }; rerender(); },
      () => { g.confirm = null; rerender(); });
  }

  /** 替换成功后更新卡片状态（ctx：新文字在邮件里的前后文，下次撤销或再换写法时用来定位） */
  function afterApplied(card, from, to, entry, ctx) {
    if (entry) entry.cardId = card.id;
    card.history.push({ from, to, entry });
    card.current = to;
    card.ctx = ctx || null;
    card.status = 'applied';
    card.failNote = '';
    card.at = Date.now();
    markApplied(to);
    const orig = live.cache.get(card.key);
    if (orig && orig.status !== 'applied') orig.status = 'replaced';
  }

  async function applyCard(card, newText) {
    if (card.busy || groupBusy(card)) return;
    const from = card.current;
    if (TU.normalize(from) === TU.normalize(newText)) { toast('邮件里已经是这个写法了'); return; }
    if (!beginWrite()) return;
    card.busy = true;
    card.failNote = '';
    rerender();
    try {
      const r = await replaceWithPolicy(from, newText, askInCard(card), card.ctx);
      if (r.cancelled) return;
      if (!r.ok) { toast(reasonText(r.reason), 8000); return; }
      afterApplied(card, from, newText, r.entry, r.ctx);
      if (card.source !== 'live') rebaseLive();
      updateUndo();
      toast(r.method === 'body' && inOutlook ? '已替换（光标可能会跳到正文开头）' : '已替换');
    } catch (e) {
      toast('替换失败：' + errText(e), 8000);
    } finally {
      endWrite();
      card.busy = false;
      rerender();
    }
  }

  async function undoCard(card) {
    const last = card.history[card.history.length - 1];
    if (!last || card.busy || groupBusy(card)) return;
    if (!beginWrite()) return;
    card.busy = true;
    rerender();
    try {
      const entry = last.entry || { original: last.from, replacement: last.to, method: 'body' };
      entry.ctx = card.ctx; // 用卡片最新记下的前后文找回这段文字
      const r = await revertWithPolicy(entry, askInCard(card));
      if (r.cancelled) return;
      if (!r.ok) { toast(reasonText(r.reason, 'undo'), 8000); return; }
      card.history.pop();
      card.current = last.from;
      card.ctx = r.ctx || null;
      card.status = card.history.length ? 'applied' : 'open';
      card.at = Date.now();
      markPresent(last.from);
      if (!card.history.length) {
        const orig = live.cache.get(card.key);
        if (orig) orig.status = card.result ? 'issues' : 'ok';
      }
      if (card.source !== 'live') rebaseLive();
      updateUndo();
      toast('已撤销');
    } catch (e) {
      toast('撤销失败：' + errText(e), 8000);
    } finally {
      endWrite();
      card.busy = false;
      rerender();
    }
  }

  function ignoreCard(card) {
    card.status = 'ignored';
    const orig = live.cache.get(card.key);
    if (orig && orig.status === 'issues') orig.status = 'ignored';
    rerender();
  }

  function dismissCard(card) {
    card.status = 'dismissed';
    rerender();
  }

  // ---------------------------------------------------------------------------
  // 一组卡片的批量操作：全部替换 / 撤销全部（只写回邮件一次）
  // ---------------------------------------------------------------------------
  let batchSeq = 0;
  function pickCorrected(card) { return card.result && card.result.changed ? card.result.corrected : ''; }
  function pickBetter(card) { return card.result ? (card.result.better || pickCorrected(card)) : ''; }

  /** 这张卡片用 pick 选出的写法：'ok' 可以批量替换；'review' 要你核对（改了数字/术语，或和原文差别很大）；'' 没有可替换的 */
  function batchKind(card, pick, terms) {
    if (card.status !== 'open') return '';
    const to = pick(card);
    if (!to || TU.normalize(to) === TU.normalize(card.current)) return '';
    if (isSuspicious(card, to)) return 'review';
    return TU.invariantWarnings(card.current, to, terms).length ? 'review' : 'ok';
  }
  function isSuspicious(card, to) {
    return !!(card.result && card.result.reviewOnly && to === card.result.corrected);
  }
  function countKind(list, pick, kind) {
    const terms = TU.parseTerms(settings.terms);
    return list.filter(c => batchKind(c, pick, terms) === kind).length;
  }

  async function applyGroup(g, list, pick, unit) {
    const u = unit || '处';
    if (g.busy) return;
    const terms = TU.parseTerms(settings.terms);
    const items = [];
    const skipped = [];
    list.forEach(card => {
      if (card.busy) return;
      const kind = batchKind(card, pick, terms);
      // 改了数字、日期或术语的建议不自动套用，留给你在卡片上核对
      if (kind === 'review') skipped.push(card);
      else if (kind === 'ok') items.push({ card, to: pick(card) });
    });
    skipped.forEach(card => { card.failNote = reasonText(isSuspicious(card, pick(card)) ? 'suspicious' : 'changed_facts'); });
    if (!items.length) {
      toast(skipped.length ? '有 ' + skipped.length + ' ' + u + '需要你核对，请在卡片上看过后单独替换。' : '没有需要替换的内容', 6000);
      rerender();
      return;
    }
    if (!beginWrite()) return;
    g.busy = true;
    items.forEach(it => { it.card.busy = true; it.card.failNote = ''; });
    rerender();
    try {
      const pairs = items.map(it => ({ original: it.card.current, replacement: it.to, ctx: it.card.ctx }));
      const r = await replaceManyWithPolicy(pairs, askInGroup(g), true, '全部替换');
      if (r.cancelled) return;
      if (!r.results) { toast(reasonText(r.reason), 8000); return; }
      const batchId = 'b' + (++batchSeq);
      let okN = 0;
      r.results.forEach((res, i) => {
        const it = items[i];
        if (res.ok) {
          okN++;
          if (res.entry) res.entry.batchId = batchId;
          afterApplied(it.card, it.card.current, it.to, res.entry, res.ctx);
        } else if (PER_ITEM_REASONS.has(res.reason)) {
          it.card.failNote = reasonText(res.reason);
        }
      });
      if (okN) { rebaseLive(); updateUndo(); }
      const notDone = items.length - okN + skipped.length;
      if (!okN) toast(PER_ITEM_REASONS.has(r.reason) ? '没能自动替换，原因写在卡片上。' : reasonText(r.reason), 8000);
      else if (notDone) toast('已替换 ' + okN + ' ' + u + '；还有 ' + notDone + ' ' + u + '没有自动替换，原因写在那几张卡片上。', 8000);
      else toast('已全部替换（' + okN + ' ' + u + '）' + (inOutlook ? '，光标可能会跳到正文开头' : '') + '。不满意可以点「撤销全部」。', 6000);
    } catch (e) {
      toast('替换失败：' + errText(e), 8000);
    } finally {
      endWrite();
      g.busy = false;
      items.forEach(it => { it.card.busy = false; });
      rerender();
    }
  }

  /** 一次撤销多张卡片：toOriginal=true 直接回到最初的原文；否则每张只退一步 */
  async function revertCards(g, list, toOriginal, unit) {
    const u = unit || '处';
    if (g.busy) return;
    const items = list.filter(c => c.status === 'applied' && c.history.length && !c.busy);
    if (!items.length) { toast('没有可撤销的替换'); return; }
    if (!beginWrite()) return;
    g.busy = true;
    items.forEach(c => { c.busy = true; });
    rerender();
    try {
      const targets = items.map(c => (toOriginal ? c.history[0] : c.history[c.history.length - 1]).from);
      const pairs = items.map((c, i) => ({ original: c.current, replacement: targets[i], ctx: c.ctx }));
      const r = await replaceManyWithPolicy(pairs, askInGroup(g), false, '全部撤销');
      if (r.cancelled) return;
      if (!r.results) { toast(reasonText(r.reason, 'undo'), 8000); return; }
      let okN = 0;
      r.results.forEach((res, i) => {
        if (!res.ok) return;
        okN++;
        const c = items[i];
        const removed = toOriginal ? c.history.splice(0) : [c.history.pop()];
        removed.forEach(h => { if (h.entry) host.removeUndo(h.entry); });
        c.current = targets[i];
        c.ctx = res.ctx || null;
        c.status = c.history.length ? 'applied' : 'open';
        c.at = Date.now();
        markPresent(targets[i]);
        if (!c.history.length) {
          const orig = live.cache.get(c.key);
          if (orig) orig.status = c.result ? 'issues' : 'ok';
        }
      });
      if (okN) { rebaseLive(); updateUndo(); }
      const failN = items.length - okN;
      if (!okN) toast(reasonText(r.reason, 'undo'), 8000);
      else if (failN) toast('已撤销 ' + okN + ' ' + u + '；还有 ' + failN + ' ' + u + '没找到（可能已经改过），可以在卡片上单独撤销。', 8000);
      else toast('已撤销 ' + okN + ' ' + u, 4000);
    } catch (e) {
      toast('撤销失败：' + errText(e), 8000);
    } finally {
      endWrite();
      g.busy = false;
      items.forEach(c => { c.busy = false; });
      rerender();
    }
  }

  /** 卡片上方的操作栏（全部替换 / 撤销全部 等）。buttons: [[文字, 样式, 点击]] */
  function renderGroupBar(box, g, buttons, info, extraDisabled) {
    const disabled = g.busy || !!extraDisabled;
    const sig = [buttons.map(b => b[0] + '|' + b[1]).join('\u0003'), disabled ? 1 : 0, g.busy ? 1 : 0,
      g.confirm ? g.confirm.message : '', info || ''].join('\u0001');
    if (box.dataset.sig === sig) return;
    box.dataset.sig = sig;
    clear(box);
    if (info) box.appendChild(el('div', { class: 'muted small' }, info));
    if (buttons.length) {
      box.appendChild(el('div', { class: 'actions' },
        buttons.map(([label, kind, fn]) => btn(label, kind, fn, disabled)),
        g.busy ? el('span', { class: 'muted small' }, el('span', { class: 'spinner' }), ' 处理中…') : null));
    }
    if (g.confirm) {
      const c = g.confirm;
      box.appendChild(el('div', { class: 'confirm' },
        el('div', {}, c.message),
        el('div', { class: 'actions' }, c.labels.map((label, i) => btn(label, i === 0 ? 'primary' : '', () => c.resolve(i))))));
    }
    box.hidden = !box.firstChild;
  }

  /** 「换个写法」：让 AI 给 3 种不同写法（避开之前给过的） */
  async function loadAlternatives(card) {
    if (card.altsLoading) return;
    if (!settings.apiKey) { showSetupNotice(); return; }
    card.altsLoading = true;
    card.altsError = '';
    const seq = itemSeq;
    rerender();
    try {
      const meta = await host.getMeta();
      const avoid = [];
      if (card.result) avoid.push(card.result.corrected, card.result.better);
      (card.alts || []).forEach(a => avoid.push(a.text));
      card.altsSeen.forEach(t => avoid.push(t));
      avoid.push(card.current);
      const uniq = Array.from(new Set(avoid.filter(Boolean).map(s => s.trim())))
        .filter(s => TU.normalize(s) !== TU.normalize(card.original));
      const ctx = { subject: meta.subject, composeType: meta.composeType, quoted: live.quoted || fullState.quoted || '', paragraph: card.paragraph };
      const req = AI.buildAlternativesRequest(card.original, ctx, settings, uniq);
      const r = await callAI(Object.assign({}, req, { timeoutMs: 45000 }));
      if (seq !== itemSeq) return; // 期间切换了邮件，结果作废
      const alts = AI.normalizeAlternatives(r.json, card.original, settings);
      card.altsSeen = card.altsSeen.concat((card.alts || []).map(a => a.text));
      card.alts = alts;
      card.altsRound += 1;
    } catch (e) {
      card.altsError = AI.describeError(e);
      if (e && e.kind === 'no_key') showSetupNotice();
    } finally {
      card.altsLoading = false;
      rerender();
    }
  }

  function cardSignature(card, focus) {
    return [card.id, card.status, card.busy ? 1 : 0, groupBusy(card) ? 1 : 0, card.current, card.history.length, card.altsLoading ? 1 : 0,
      card.altsRound, card.altsError, card.confirm ? card.confirm.message : '', focus ? 1 : 0, card.failNote,
      card.result ? card.result.corrected + card.result.better : '', card.versions ? card.versionsRound : '',
      card.source === 'rewrite' && rwRunning ? 'running' : ''].join('\u0001');
  }

  /** 只有内容真的变了才重画，避免你点按钮时界面刚好刷新 */
  function renderInto(container, items, renderFn) {
    const sig = items.map(it => cardSignature(it.card, it.focus)).join('\u0002');
    if (container.dataset.sig === sig) return;
    container.dataset.sig = sig;
    clear(container);
    items.forEach(it => container.appendChild(renderFn(it.card, it.focus)));
  }

  function renderConfirm(card) {
    if (!card.confirm) return null;
    const c = card.confirm;
    return el('div', { class: 'confirm' },
      el('div', {}, c.message),
      el('div', { class: 'actions' }, c.labels.map((label, i) => btn(label, i === 0 ? 'primary' : '', () => c.resolve(i)))));
  }

  function renderAlternatives(card) {
    if (!card.altsLoading && !card.alts && !card.altsError) return null;
    const box = el('div', { class: 'alts' }, el('div', { class: 'cap' }, '其他写法'));
    const busy = card.busy || card.altsLoading || groupBusy(card);
    if (card.altsError) box.appendChild(el('div', { class: 'warn' }, card.altsError));
    (card.alts || []).forEach(a => {
      const inUse = TU.normalize(a.text) === TU.normalize(card.current);
      box.appendChild(el('div', { class: 'alt' + (inUse ? ' in-use' : '') },
        el('div', { class: 'alt-head' }, el('span', { class: 'chip t-clarity' }, a.label), inUse ? el('span', { class: 'chip ok' }, '正在使用') : null),
        el('div', { class: 'text' }, a.text),
        a.warnings.map(w => el('div', { class: 'warn' }, '⚠ ' + w)),
        el('div', { class: 'actions' },
          btn(inUse ? '已使用' : '用这个', inUse ? '' : 'primary', () => applyCard(card, a.text), busy || inUse),
          btn('复制', 'ghost', () => copyWithToast(a.text)))));
    });
    if (card.altsLoading) box.appendChild(el('div', { class: 'muted small loading' }, el('span', { class: 'spinner' }), ' 正在生成新的写法…'));
    else box.appendChild(el('div', { class: 'actions' }, btn('再换一批', 'ghost', () => loadAlternatives(card), busy)));
    return box;
  }

  /** 一张建议卡片（边写边查、全文检查、逐段改写共用） */
  function renderCard(card, focus) {
    const res = card.result || { issues: [], warnings: [], changed: false, corrected: card.original, better: '', betterWarnings: [], betterExplain: '' };
    const applied = card.status === 'applied';
    const busy = card.busy || groupBusy(card);
    const box = el('div', { class: 'card' + (focus ? ' focus' : '') + (applied ? ' is-applied' : '') });
    const types = Array.from(new Set(res.issues.map(i => i.type)));

    box.appendChild(el('div', { class: 'card-head' },
      focus ? el('span', { class: 'chip badge' }, '正在写的句子') : null,
      applied ? el('span', { class: 'chip ok' }, '✓ 已替换') : types.map(t => el('span', { class: 'chip t-' + t }, AI.ISSUE_TYPES[t] || t)),
      !applied && !types.length ? el('span', { class: 'chip t-tone' }, card.source === 'rewriteBlock' ? '改写' : '表达') : null,
      card.busy ? el('span', { class: 'muted small' }, el('span', { class: 'spinner' }), ' 处理中') : null));

    if (applied) {
      box.appendChild(el('div', { class: 'diff' }, diffNodes(card.original, card.current)));
      const terms = TU.parseTerms(settings.terms);
      TU.invariantWarnings(card.original, card.current, terms).forEach(w => box.appendChild(el('div', { class: 'warn' }, '⚠ ' + w)));
      box.appendChild(el('div', { class: 'actions' },
        btn('撤销', 'primary', () => undoCard(card), busy),
        btn('换个写法', '', () => loadAlternatives(card), busy || card.altsLoading),
        btn('收起', 'ghost', () => dismissCard(card), busy)));
    } else {
      box.appendChild(el('div', { class: 'diff' }, res.changed ? diffNodes(card.original, res.corrected) : card.original));
      if (res.issues.length) {
        box.appendChild(el('ul', { class: 'issue-list' }, res.issues.map(i => el('li', {},
          el('span', { class: 'fix' }, (i.original || '∅') + ' → ' + (i.suggestion || '（删除）')),
          i.explain ? el('span', { class: 'why' }, '　' + i.explain) : null))));
      }
      res.warnings.forEach(w => box.appendChild(el('div', { class: 'warn' }, '⚠ ' + w)));
      if (card.failNote) box.appendChild(el('div', { class: 'warn' }, card.failNote));
      box.appendChild(el('div', { class: 'actions' },
        res.changed ? btn('替换', 'primary', () => applyCard(card, res.corrected), busy) : null,
        btn('换个写法', '', () => loadAlternatives(card), busy || card.altsLoading),
        btn('忽略', '', () => ignoreCard(card), busy),
        res.changed ? btn('复制', 'ghost', () => copyWithToast(res.corrected)) : null));
      if (res.better) {
        box.appendChild(el('div', { class: 'better' },
          el('div', { class: 'cap' }, '更专业的写法'),
          el('div', { class: 'text' }, res.better),
          res.betterExplain ? el('div', { class: 'note' }, res.betterExplain) : null,
          res.betterWarnings.map(w => el('div', { class: 'warn' }, '⚠ ' + w)),
          el('div', { class: 'actions' },
            btn('用这个', '', () => applyCard(card, res.better), busy),
            btn('复制', 'ghost', () => copyWithToast(res.better)))));
      }
    }
    append(box, renderConfirm(card));
    append(box, renderAlternatives(card));
    return box;
  }

  // ---------------------------------------------------------------------------
  // 边写边查
  // ---------------------------------------------------------------------------
  const live = {
    timer: null,
    busy: false,
    gen: 0,                  // 每次批量改动正文后 +1，丢弃改动前读到的旧内容
    prevMine: null,          // 上一次读到的「我写的部分」
    quoted: '',              // 引用的往来邮件（作为上下文）
    lastChangeAt: 0,
    dirty: new Map(),        // 改动过、还没检查的句子：key -> 原文
    inFlight: new Set(),     // 正在检查的句子
    cache: new Map(),        // 检查结果：key -> { status, result, text, at, source }
    present: new Set(),      // 目前正文里存在的句子
    focus: null,             // 正在编辑的句子 { key, text }
    focusPara: '',
    lastCallAt: 0,
    history: [],
    hostError: false,
    running: false,
    bulkHint: false          // 刚出现一大段新内容（如粘贴）时，提示用全文检查
  };

  function presentKeys(text) {
    return new Set(TU.segmentSentences(text).map(s => TU.sentenceKey(s.text)));
  }

  /**
   * 插件自己改了正文（全文检查、改写、撤销）之后调用：
   * 让「边写边查」重新记一次底稿，不把这些改动当成你新写的内容去检查（省钱，也不会冒出多余的卡片）。
   */
  function rebaseLive() {
    live.prevMine = null;
    const gen = ++live.gen;
    // 马上读一次新底稿（不等下一轮），之后你再打字或粘贴都能正常识别
    host.getDraft().then(draft => {
      if (gen !== live.gen || live.prevMine !== null) return;
      live.prevMine = draft.mine || '';
      live.present = presentKeys(live.prevMine);
      for (const k of Array.from(live.dirty.keys())) if (!live.present.has(k)) live.dirty.delete(k);
    }).catch(() => { /* 读不到就等下一轮定时读取 */ });
  }

  function resetLive() {
    live.prevMine = null;
    live.gen += 1;
    live.quoted = '';
    live.lastChangeAt = 0;
    live.dirty.clear();
    live.inFlight.clear();
    live.present = new Set();
    live.focus = null;
    live.focusPara = '';
    live.history = [];
    live.bulkHint = false;
    for (const [id, card] of Array.from(cards)) {
      if (card.source === 'live') cards.delete(id);
    }
    liveCardByKey.clear();
    renderLive();
  }

  /**
   * 定时读取草稿。正常每 1 秒一次；如果读取本身很慢（邮件很长），自动放慢，避免拖慢 Outlook。
   */
  function scheduleTick(delay) {
    clearTimeout(live.timer);
    live.timer = setTimeout(async () => {
      const t0 = Date.now();
      await liveTick();
      if (!live.running) return;
      const took = Date.now() - t0;
      const base = inOutlook ? 1000 : 700;
      scheduleTick(took > 400 ? Math.min(4000, base + took * 2) : base);
    }, delay);
  }
  function startLive() {
    if (live.running) return;
    live.running = true;
    scheduleTick(0);
  }
  function stopLive() {
    live.running = false;
    clearTimeout(live.timer);
    live.timer = null;
  }

  function applyLiveToggle() {
    $('#liveToggle').checked = !!settings.liveEnabled;
    const compose = host && host.isCompose();
    if (settings.liveEnabled && compose) {
      startLive();
      setLiveStatus(settings.apiKey ? 'idle' : 'nokey');
    } else {
      stopLive();
      setLiveStatus(!compose ? 'nocompose' : 'off');
    }
  }

  async function liveTick() {
    if (live.busy || writeLock) return; // 插件自己正在写回邮件时不读，免得把插件的改动当成你在打字
    live.busy = true;
    const gen = live.gen;
    try {
      const draft = await host.getDraft();
      if (gen !== live.gen || writeLock) return; // 读取期间插件改过（或正在改）正文，这次读到的内容不算
      if (live.hostError) { live.hostError = false; setLiveStatus('idle'); }
      const mine = draft.mine || '';
      live.quoted = draft.quoted || '';

      if (live.prevMine === null) { // 第一次读取（或刚批量改过正文）：记下现有内容，不主动检查
        live.prevMine = mine;
        live.present = presentKeys(mine);
        for (const k of Array.from(live.dirty.keys())) if (!live.present.has(k)) live.dirty.delete(k);
        renderLive();
        return;
      }

      if (mine !== live.prevMine) {
        const ch = TU.changedSentences(live.prevMine, mine);
        live.prevMine = mine;
        live.lastChangeAt = Date.now();
        const focus = ch.caret >= 0 ? TU.sentenceAt(mine, ch.caret) : null;
        const newCheckable = ch.sentences.filter(s => TU.isCheckableEnglish(s.text) && !live.cache.has(TU.sentenceKey(s.text)));
        if (newCheckable.length > 6) {
          // 一次出现很多新句子（例如粘贴了一大段）：不逐句自动检查，提示用「全文检查」
          live.bulkHint = true;
        } else {
          for (const s of ch.sentences) {
            const k = TU.sentenceKey(s.text);
            if (!live.cache.has(k)) live.dirty.set(k, s.text); // 已检查过/刚替换的句子不再排队
          }
        }
        live.focus = focus ? { key: TU.sentenceKey(focus.text), text: focus.text } : null;
        live.focusPara = focus ? TU.paragraphAt(mine, focus.start) : '';
        live.present = presentKeys(mine);
        for (const k of Array.from(live.dirty.keys())) if (!live.present.has(k)) live.dirty.delete(k);
        setLiveStatus(live.dirty.size || live.inFlight.size ? 'typing' : 'done');
        renderLive();
        return;
      }

      if (!live.dirty.size) return;
      const idle = Date.now() - live.lastChangeAt;
      if (idle < (Number(settings.pauseMs) || 2000)) return;
      const batch = pickBatch(idle);
      if (batch.length) runLiveCheck(batch); // 不等待结果，继续监听
      else if (live.focus && live.dirty.has(live.focus.key) && !live.inFlight.size) setLiveStatus(settings.apiKey ? 'waiting' : 'nokey');
    } catch (e) {
      if (!live.hostError) {
        live.hostError = true;
        setLiveStatus('error', '读取邮件内容失败：' + errText(e));
      }
    } finally {
      live.busy = false;
    }
  }

  /** 选出这一轮要检查的句子（正在写的那句优先，一次最多 4 句） */
  function pickBatch(idle) {
    if (!settings.apiKey) return [];
    const pause = Number(settings.pauseMs) || 2000;
    const items = [];
    for (const [key, text] of live.dirty) {
      if (live.inFlight.has(key)) continue;
      if (live.cache.has(key)) { live.dirty.delete(key); continue; }
      if (!TU.isCheckableEnglish(text)) { live.dirty.delete(key); continue; }
      const isFocus = !!live.focus && key === live.focus.key;
      if (isFocus && !TU.looksComplete(text) && idle < pause * 2) continue; // 可能还没写完，多等一会儿
      items.push({ key, text, isFocus });
    }
    if (!items.length) return [];
    if (Date.now() - live.lastCallAt < 1500) return [];
    items.sort((a, b) => Number(b.isFocus) - Number(a.isFocus));
    return items.slice(0, 4);
  }

  async function runLiveCheck(items) {
    live.lastCallAt = Date.now();
    items.forEach(it => { live.inFlight.add(it.key); live.dirty.delete(it.key); });
    setLiveStatus('checking', items.length);
    renderLive();
    const sentences = items.map((it, i) => ({ id: i + 1, text: it.text }));
    const paragraph = live.focusPara;
    const seq = itemSeq;
    try {
      const meta = await host.getMeta();
      const req = AI.buildCheckRequest(sentences, {
        subject: meta.subject, composeType: meta.composeType, quoted: live.quoted, paragraph
      }, settings, 'live');
      const r = await callAI(Object.assign({}, req, { timeoutMs: 30000 }));
      if (seq !== itemSeq) return; // 期间切换了邮件，结果作废
      const out = AI.normalizeCheckResult(r.json, sentences, settings);
      const byId = new Map(out.results.map(x => [x.id, x]));
      sentences.forEach((s, i) => {
        const res = byId.get(s.id) || null;
        live.cache.set(items[i].key, { status: res ? 'issues' : 'ok', result: res, text: s.text, at: Date.now(), source: 'live' });
        if (res) upsertLiveCard(s.text, res, paragraph);
        pushHistory(items[i].key);
      });
      pruneCache();
      setLiveStatus(live.dirty.size ? 'typing' : 'done');
    } catch (e) {
      const msg = AI.describeError(e);
      items.forEach(it => live.cache.set(it.key, { status: 'error', error: msg, text: it.text, at: Date.now(), source: 'live' }));
      setLiveStatus('error', msg);
      if (e && e.kind === 'no_key') showSetupNotice();
    } finally {
      items.forEach(it => live.inFlight.delete(it.key));
      renderLive();
    }
  }

  /** 结果缓存最多保留约 800 句，太多时删掉最旧的（侧边栏固定、连续写很多邮件时防止越积越多） */
  function pruneCache() {
    if (live.cache.size <= 800) return;
    const oldest = Array.from(live.cache.entries())
      .filter(([key]) => !live.present.has(key))
      .sort((a, b) => a[1].at - b[1].at)
      .slice(0, live.cache.size - 600);
    oldest.forEach(([key]) => live.cache.delete(key));
  }

  /** 把出错的句子重新放回待检查队列 */
  function retryErrors(silent) {
    let n = 0;
    for (const [key, entry] of Array.from(live.cache)) {
      if (entry.status === 'error') {
        live.cache.delete(key);
        if (live.present.has(key)) { live.dirty.set(key, entry.text); n++; }
      }
    }
    if (n) live.lastChangeAt = Date.now() - (Number(settings.pauseMs) || 2000) * 2;
    if (!silent) setLiveStatus(n ? 'typing' : 'idle');
  }

  function setLiveStatus(kind, arg) {
    const box = $('#liveStatus');
    box.className = 'status s-' + kind;
    const pause = ((Number(settings.pauseMs) || 2000) / 1000).toString();
    const text = {
      idle: '已开启：停顿 ' + pause + ' 秒自动检查',
      typing: '正在输入…',
      waiting: '等你写完这句…',
      checking: '正在检查' + (arg > 1 ? ' ' + arg + ' 句' : '') + '…',
      done: '已检查',
      off: '已关闭',
      nokey: '请先在设置里填写 API Key',
      nocompose: '请在写邮件或回复时使用',
      error: arg || '出错了'
    }[kind] || '';
    box.querySelector('.txt').textContent = text;
    const old = box.querySelector('.retry');
    if (old) old.remove();
    if (kind === 'error') box.appendChild(el('button', { class: 'btn small ghost retry', onclick: () => retryErrors(false) }, '重试'));
  }

  function pushHistory(key) {
    live.history = [key].concat(live.history.filter(k => k !== key)).slice(0, 15);
  }

  function visibleLiveCards() {
    const f = live.focus;
    const isFocus = c => !!f && TU.segmentSentences(c.current).some(s => TU.sentenceKey(s.text) === f.key);
    const list = [];
    for (const card of cards.values()) {
      if (card.source !== 'live' || card.status === 'ignored' || card.status === 'dismissed') continue;
      if (!isPresent(card.current) && !card.busy && !card.confirm) continue;
      list.push(card);
    }
    list.sort((a, b) => (Number(isFocus(b)) - Number(isFocus(a))) || (b.at - a.at));
    return list.slice(0, 6).map(card => ({ card, focus: isFocus(card) }));
  }

  function renderLive() {
    // 0) 粘贴了一大段时的提示
    const bulk = $('#liveBulk');
    clear(bulk);
    if (live.bulkHint) {
      append(bulk, [
        el('div', {}, '刚刚出现了一大段新内容（比如粘贴），为了省钱不逐句自动检查。可以用「全文检查」一次查完。'),
        el('div', { class: 'actions' },
          el('button', { class: 'btn small primary', onclick: () => { live.bulkHint = false; switchTab('full'); runFullCheck(); renderLive(); } }, '全文检查'),
          el('button', { class: 'btn small', onclick: () => { live.bulkHint = false; renderLive(); } }, '知道了'))
      ]);
    }
    bulk.hidden = !live.bulkHint;

    // 1) 正在写的句子
    const cur = $('#liveCurrent');
    clear(cur);
    const f = live.focus;
    const showCurrent = !!f && live.present.has(f.key) && TU.isCheckableEnglish(f.text);
    if (showCurrent) {
      const entry = live.cache.get(f.key);
      let cls = 'current';
      let cap = '正在写的句子';
      if (live.inFlight.has(f.key)) cap = '正在检查这句…';
      else if (entry && entry.status === 'ok') { cls += ' ok'; cap = '✓ 这句没问题'; }
      else if (entry && entry.status === 'applied') { cls += ' ok'; cap = '✓ 已替换'; }
      else if (entry && entry.status === 'issues') { cls += ' issues'; cap = '这句可以改进（见下方）'; }
      else if (entry && entry.status === 'ignored') cap = '已忽略这句的建议';
      else if (entry && entry.status === 'error') cap = '这句没检查成功：' + entry.error;
      cur.className = cls;
      append(cur, [el('div', { class: 'cap' }, cap), el('div', { class: 'sent' }, f.text)]);
    }
    cur.hidden = !showCurrent;

    // 2) 建议卡片
    const items = visibleLiveCards();
    renderInto($('#liveCards'), items, renderCard);
    $('#liveHint').hidden = items.length > 0 || showCurrent;

    // 3) 最近检查
    renderHistory();
  }

  function renderHistory() {
    const box = $('#liveHistory');
    clear(box);
    const items = live.history.map(k => live.cache.get(k)).filter(Boolean);
    $('#liveHistoryWrap').hidden = items.length === 0;
    $('#histCount').textContent = String(items.length);
    const marks = { ok: '✓', applied: '✓', replaced: '✓', issues: '!', ignored: '–', error: '×' };
    for (const e of items) {
      box.appendChild(el('div', { class: 'hist-item ' + e.status }, el('span', { class: 'mark' }, marks[e.status] || '·'), el('span', {}, e.text)));
    }
  }

  function updateUndo() { $('#btnUndo').hidden = !(host && host.canUndo()); }

  /** 底部「撤销上次替换」：撤销最近一次替换（属于哪张卡片就交给那张卡片处理；「全部替换」的一批就整批撤销） */
  async function onUndo() {
    const top = host.lastUndo();
    if (!top) { updateUndo(); return; }
    if (top.batchId) {
      const batch = Array.from(cards.values()).filter(c => {
        const h = c.history[c.history.length - 1];
        return (c.status === 'applied' || c.status === 'dismissed') && h && h.entry && h.entry.batchId === top.batchId;
      });
      if (batch.length) {
        const g = groupOf(batch[0]) || groups.full;
        if (g.busy) { toast('正在处理，请稍等一下'); return; }
        batch.forEach(c => { if (c.status === 'dismissed') c.status = 'applied'; });
        switchTab(g === groups.rwb ? 'rewrite' : 'full'); // 万一要确认，确认框在那个页面上
        await revertCards(g, batch, false, g === groups.rwb ? '段' : '处');
        return;
      }
    }
    const card = top.cardId ? cards.get(top.cardId) : null;
    if (card && card.history.length && card.history[card.history.length - 1].entry === top) {
      if (card.status === 'dismissed') card.status = 'applied';
      await undoCard(card);
      return;
    }
    if (!beginWrite()) return;
    try {
      const r = await revertWithPolicy(top, askInNotice);
      if (r.cancelled) return;
      if (r.ok) rebaseLive();
      toast(r.ok ? '已撤销' : reasonText(r.reason, 'undo'), r.ok ? 3000 : 8000);
    } catch (e) {
      toast('撤销失败：' + errText(e), 8000);
    } finally {
      endWrite();
      updateUndo();
      rerender();
    }
  }

  // ---------------------------------------------------------------------------
  // 全文检查
  // ---------------------------------------------------------------------------
  const fullState = { running: false, quoted: '' };

  async function runFullCheck() {
    if (fullState.running) return;
    if (!host.isCompose()) { toast('请在写邮件或回复时使用'); return; }
    if (!settings.apiKey) { showSetupNotice(); return; }
    if (!abandonPendingWrites()) return;
    const seq = itemSeq;
    fullState.running = true;
    const done = setBusy($('#btnFull'), '正在检查…');
    const ov = $('#fullOverall');
    clear(ov);
    fullCardIds.forEach(id => cards.delete(id));
    fullCardIds = [];
    groups.full.confirm = null;
    renderFullCards();
    try {
      const draft = await host.getDraft();
      const meta = await host.getMeta();
      fullState.quoted = draft.quoted || '';
      const mine = toLF(draft.mine);
      // 签名（Best Regards 之后）不检查
      const sigAt = TU.findSignatureStart(mine);
      const body = sigAt >= 0 ? mine.slice(0, sigAt) : mine;
      // 同一句只问 AI 一次；在邮件里出现几次，就有几张卡片（各改各的那一处）
      const uniq = [];
      const where = new Map();
      TU.segmentSentences(body).forEach(s => {
        if (!TU.isCheckableEnglish(s.text)) return;
        const k = TU.sentenceKey(s.text);
        if (!where.has(k)) { where.set(k, []); uniq.push(s); }
        where.get(k).push(s);
      });
      if (!uniq.length) { ov.appendChild(el('p', { class: 'hint' }, '没有找到需要检查的英文句子。')); return; }
      let note = '';
      let sentences = uniq;
      if (sentences.length > 60) { sentences = sentences.slice(0, 60); note = '内容较长，这次只检查了前 60 句。'; }
      const list = sentences.map((s, i) => ({ id: i + 1, text: s.text }));
      const req = AI.buildCheckRequest(list, { subject: meta.subject, composeType: meta.composeType, quoted: draft.quoted }, settings, 'full');
      const r = await callAI(Object.assign({}, req, { timeoutMs: 120000 }));
      if (seq !== itemSeq) return; // 期间切换了邮件，结果作废
      const out = AI.normalizeCheckResult(r.json, list, settings);
      // 记住结果，边写边查时就不会再重复检查这些句子
      list.forEach(s => {
        const res = out.results.find(x => x.id === s.id) || null;
        const key = TU.sentenceKey(s.text);
        const prev = live.cache.get(key);
        if (!prev || prev.source !== 'live') live.cache.set(key, { status: res ? 'issues' : 'ok', result: res, text: s.text, at: Date.now(), source: 'full' });
      });
      out.results.forEach(res => {
        const spots = where.get(TU.sentenceKey(res.original)) || [];
        if (!spots.length) { fullCardIds.push(newCard('full', res.original, { result: res }).id); return; }
        spots.forEach(s => fullCardIds.push(newCard('full', res.original, { result: res, ctx: TU.contextAt(mine, s.start, s.end) }).id));
      });
      ov.appendChild(el('div', { class: 'overall' },
        el('div', { class: 'cap' }, '整体评价'),
        el('div', {}, out.overall || '（AI 没有给出整体评价）'),
        el('div', { class: 'muted small mt' },
          '检查了 ' + list.length + ' 句，' + (out.results.length ? '有 ' + out.results.length + ' 句可以改进。' : '没有发现需要修改的地方。') +
          (sigAt >= 0 ? ' 签名部分没有检查。' : '') + (note ? ' ' + note : ''))));
      if (!out.results.length) ov.appendChild(el('div', { class: 'allgood' }, '✓ 没有发现问题'));
      renderFullCards();
      live.bulkHint = false;
      renderLive();
    } catch (e) {
      ov.appendChild(el('div', { class: 'warn' }, AI.describeError(e)));
      if (e && e.kind === 'no_key') showSetupNotice();
    } finally {
      fullState.running = false;
      done();
    }
  }

  function fullCards() { return fullCardIds.map(id => cards.get(id)).filter(Boolean); }

  function renderFullCards() {
    const list = fullCards();
    const nFix = countKind(list, pickCorrected, 'ok');
    const hasBetter = list.some(c => c.status === 'open' && c.result && c.result.better);
    const nBetter = countKind(list, pickBetter, 'ok');
    const terms = TU.parseTerms(settings.terms);
    const nReview = list.filter(c => batchKind(c, pickCorrected, terms) === 'review' || batchKind(c, pickBetter, terms) === 'review').length;
    const nApplied = list.filter(c => c.status === 'applied').length;
    const buttons = [];
    if (nFix) buttons.push(['全部改正（' + nFix + ' 处）', 'primary', () => applyGroup(groups.full, fullCards(), pickCorrected)]);
    if (hasBetter && nBetter) buttons.push(['全部用更专业的写法（' + nBetter + ' 处）', nFix ? '' : 'primary', () => applyGroup(groups.full, fullCards(), pickBetter)]);
    if (nApplied) buttons.push(['撤销全部（' + nApplied + ' 处）', '', () => revertCards(groups.full, fullCards(), true)]);
    const info = nReview && list.length ? nReview + ' 处建议需要你核对（改动了数字、日期或术语），不会被批量替换。' : '';
    renderGroupBar($('#fullBar'), groups.full, buttons, info);
    const items = list.filter(c => c.status !== 'ignored' && c.status !== 'dismissed').map(card => ({ card, focus: false }));
    renderInto($('#fullCards'), items, renderCard);
  }

  // ---------------------------------------------------------------------------
  // 改写
  //   选中一句或一段 → 给 2 个版本（可以从中文写成英文）
  //   选中多段，或「整封邮件」→ 逐段改写，每段一张卡片，「全部替换」原地改，格式不变
  // ---------------------------------------------------------------------------
  let rwTone = 'formal';
  let rwRunning = false;
  let rwBlock = null; // 逐段改写：{ source, blocks, cardByBlock, seen, notes, info }

  function buildRewriteControls() {
    const chips = $('#toneChips');
    clear(chips);
    AI.TONES.forEach(t => chips.appendChild(el('label', {},
      el('input', { type: 'radio', name: 'rwTone', value: t.id, checked: t.id === rwTone, onchange: () => { rwTone = t.id; } }),
      el('span', {}, t.label))));
    const sel = $('#scenario');
    clear(sel);
    AI.SCENARIOS.forEach(s => sel.appendChild(el('option', { value: s.id }, s.label)));
  }

  function rewriteNotice(msg) {
    const out = $('#rewriteMsg');
    clear(out);
    if (msg) out.appendChild(el('div', { class: 'warn' }, msg));
  }

  /** 清掉改写页上一次的结果（两种模式都清） */
  function clearRewriteResults() {
    if (rewriteCardId) cards.delete(rewriteCardId);
    rewriteCardId = null;
    if (rwBlock) rwBlock.cardByBlock.forEach(id => cards.delete(id));
    rwBlock = null;
    groups.rwb.confirm = null;
  }

  /** 值得改写的一段：不是数据行；至少 2 个英文单词，或者含中文 */
  function isRewritableBlock(text) {
    if (TU.isTableLike(text)) return false;
    const words = (String(text).match(/[A-Za-z][A-Za-z'’-]*/g) || []).length;
    return words >= 2 || /[㐀-鿿豈-﫿]/.test(text);
  }

  /** 生成改写。regenerate=true 时沿用上一次的原文，并要求 AI 写得和之前的不一样 */
  async function runRewrite(regenerate) {
    if (rwRunning) return;
    rewriteNotice('');
    if (!host.isCompose()) { rewriteNotice('请在写邮件或回复时使用。'); return; }
    if (!settings.apiKey) { showSetupNotice(); return; }
    if (groups.rwb.busy) { toast('正在写回邮件，请稍等一下再试'); return; }
    if (!abandonPendingWrites()) return;
    if (regenerate) {
      if (rwBlock) return runBlockRewrite('', rwBlock.source, true);
      const prevCard = rewriteCardId ? cards.get(rewriteCardId) : null;
      if (prevCard) return runVersionsRewrite(prevCard.original, prevCard.rwSource, true);
      return;
    }
    const sourceInput = document.querySelector('input[name="rwSource"]:checked');
    const source = sourceInput ? sourceInput.value : 'selection';
    let text = '';
    try {
      if (source === 'selection') {
        try { text = await host.getSelectedText(); } catch (e) { text = ''; } // 光标在收件人/主题栏时会读不到
        if (!text.trim()) { rewriteNotice('请先在邮件正文里选中要改写的文字（可以是中文），再点「生成改写」。想改整封邮件，请在上面选「整封邮件」。'); return; }
      } else {
        text = (await host.getDraft()).mine || '';
        if (!text.trim()) { rewriteNotice('邮件里还没有你写的内容。'); return; }
      }
    } catch (e) {
      rewriteNotice('读取邮件内容失败：' + errText(e));
      return;
    }
    if (source === 'all' || TU.splitBlocks(toLF(text)).length >= 2) return runBlockRewrite(text, source, false);
    return runVersionsRewrite(text.trim(), source, false);
  }

  /** 选中一句/一段：给 2 个版本 */
  async function runVersionsRewrite(text, source, regenerate) {
    const prevCard = rewriteCardId ? cards.get(rewriteCardId) : null;
    rwRunning = true;
    const seq = itemSeq;
    const done = setBusy($('#btnRewrite'), '正在改写…');
    if (prevCard && regenerate) prevCard.busy = true;
    renderRewrite();
    try {
      const draft = await host.getDraft();
      const meta = await host.getMeta();
      const avoid = regenerate && prevCard ? prevCard.versionsSeen.concat(prevCard.versions.map(v => v.text)) : [];
      const req = AI.buildRewriteRequest(text, { subject: meta.subject, quoted: draft.quoted }, settings, rwTone, $('#scenario').value, $('#rwExtra').value, avoid);
      const r = await callAI(Object.assign({}, req, { timeoutMs: 120000 }));
      if (seq !== itemSeq) return; // 期间切换了邮件，结果作废
      const res = AI.normalizeRewriteResult(r.json, text, settings);
      if (regenerate && prevCard) {
        prevCard.versionsSeen = avoid;
        prevCard.versions = res.versions;
        prevCard.notes = res.notes;
        prevCard.versionsRound += 1;
      } else {
        clearRewriteResults();
        const card = newCard('rewrite', text, { versions: res.versions, notes: res.notes, versionsSeen: [], versionsRound: 1, rwSource: source });
        rewriteCardId = card.id;
      }
    } catch (e) {
      rewriteNotice(AI.describeError(e));
      if (e && e.kind === 'no_key') showSetupNotice();
    } finally {
      if (prevCard) prevCard.busy = false;
      rwRunning = false;
      done();
      rerender();
    }
  }

  /**
   * 逐段改写：每一行（标题、项目符号、段落）一段，AI 一段对一段地改。
   * 重新生成时：已经替换或忽略的段落保持不变，其余段落重新改写，并告诉 AI 避开之前的写法。
   */
  async function runBlockRewrite(text, source, regenerate) {
    const prev = regenerate ? rwBlock : null;
    let list;
    let info = '';
    if (prev) {
      list = prev.blocks.filter(b => {
        const c = cards.get(prev.cardByBlock.get(b.id));
        return !(c && (c.status === 'applied' || c.status === 'ignored' || c.history.length));
      });
      if (!list.length) { rewriteNotice('所有建议都已经替换或忽略了。不满意可以先点「撤销全部」，再重新生成。'); return; }
    } else {
      const src = toLF(text);
      const candidates = TU.splitBlocks(src, { excludeSignature: source === 'all' }).filter(b => isRewritableBlock(b.text));
      if (!candidates.length) { rewriteNotice('没有找到需要改写的内容（数据行和签名不会改写）。'); return; }
      const kept = [];
      let total = 0;
      for (const b of candidates) {
        if (kept.length >= 80 || (kept.length && total + b.text.length > 9000)) break;
        kept.push(b);
        total += b.text.length;
      }
      if (kept.length < candidates.length) info = '内容较长，这次只改写了前 ' + kept.length + ' 段。';
      list = kept.map((b, i) => ({
        id: i + 1,
        text: b.text,
        ctx: source === 'all' ? TU.contextAt(src, b.start, b.end) : null,
        around: [kept[i - 1], b, kept[i + 1]].filter(Boolean).map(x => x.text).join('\n')
      }));
    }
    // 之前给过的写法（重新生成时避开）
    const avoidFor = b => {
      if (!prev) return [];
      const c = cards.get(prev.cardByBlock.get(b.id));
      return (prev.seen.get(b.id) || []).concat(c && c.result ? [c.result.corrected] : []);
    };

    rwRunning = true;
    const seq = itemSeq;
    const done = setBusy($('#btnRewrite'), '正在改写…');
    const lockedGroup = !!rwBlock; // 改写期间旧卡片不能点（结果回来会替换掉它们）
    if (lockedGroup) groups.rwb.busy = true;
    rerender();
    try {
      const draft = await host.getDraft();
      const meta = await host.getMeta();
      const avoid = [].concat.apply([], list.map(avoidFor));
      const req = AI.buildBlockRewriteRequest(list, { subject: meta.subject, quoted: draft.quoted }, settings, rwTone, $('#scenario').value, $('#rwExtra').value, avoid);
      const r = await callAI(Object.assign({}, req, { timeoutMs: 180000 }));
      if (seq !== itemSeq) return; // 期间切换了邮件，结果作废
      const out = AI.normalizeBlockRewrite(r.json, list, settings);
      const byId = new Map(out.results.map(x => [x.id, x]));
      let state = prev;
      if (!state) {
        clearRewriteResults();
        state = rwBlock = { source, blocks: list, cardByBlock: new Map(), seen: new Map(), notes: '', info };
      }
      list.forEach(b => {
        if (prev) state.seen.set(b.id, avoidFor(b).slice(-3));
        const old = cards.get(state.cardByBlock.get(b.id));
        if (old) cards.delete(old.id);
        state.cardByBlock.delete(b.id);
        const res = byId.get(b.id);
        if (res) state.cardByBlock.set(b.id, newCard('rewriteBlock', b.text, { result: res, ctx: b.ctx, paragraph: b.around }).id);
      });
      state.notes = out.notes;
      if (prev && !out.results.length) toast('AI 这次没有给出新的写法，可以换个语气或补充要求再试。', 6000);
    } catch (e) {
      rewriteNotice(AI.describeError(e));
      if (e && e.kind === 'no_key') showSetupNotice();
    } finally {
      if (lockedGroup) groups.rwb.busy = false;
      rwRunning = false;
      done();
      rerender();
    }
  }

  function rwBlockCards() {
    if (!rwBlock) return [];
    return rwBlock.blocks.map(b => cards.get(rwBlock.cardByBlock.get(b.id))).filter(Boolean);
  }

  function renderRewrite() {
    const card = rewriteCardId ? cards.get(rewriteCardId) : null;
    renderInto($('#rewriteResults'), card ? [{ card, focus: false }] : [], renderRewriteBody);
    renderBlockRewrite();
  }

  function renderBlockRewrite() {
    const wrap = $('#rwBlock');
    const head = $('#rwBlockHead');
    wrap.hidden = !rwBlock;
    if (!rwBlock) {
      clear(head);
      head.dataset.sig = '';
      renderGroupBar($('#rwBar'), groups.rwb, [], '');
      renderInto($('#rwCards'), [], renderCard);
      return;
    }
    const list = rwBlockCards();
    const skipNote = rwBlock.source === 'all' ? '数据行（如 SKU 清单）和签名不会改动。' : '数据行（如 SKU 清单）不会改动。';
    const summary = '共 ' + rwBlock.blocks.length + ' 段，' + (list.length ? list.length + ' 段有改写建议。' : 'AI 认为都不需要修改。') +
      skipNote + (rwBlock.info ? ' ' + rwBlock.info : '');
    const sig = rwBlock.notes + '\u0001' + summary;
    if (head.dataset.sig !== sig) {
      head.dataset.sig = sig;
      clear(head);
      head.appendChild(el('div', { class: 'overall' },
        el('div', { class: 'cap' }, '逐段改写'),
        rwBlock.notes ? el('div', {}, rwBlock.notes) : null,
        el('div', { class: 'muted small mt' }, summary)));
    }
    const nOpen = countKind(list, pickCorrected, 'ok');
    const nReview = countKind(list, pickCorrected, 'review');
    const nApplied = list.filter(c => c.status === 'applied').length;
    const buttons = [];
    if (nOpen) buttons.push(['全部替换（' + nOpen + ' 段）', 'primary', () => applyGroup(groups.rwb, rwBlockCards(), pickCorrected, '段')]);
    if (nApplied) buttons.push(['撤销全部（' + nApplied + ' 段）', '', () => revertCards(groups.rwb, rwBlockCards(), true, '段')]);
    buttons.push(['都不满意，重新生成', '', () => runRewrite(true)]);
    const info = nReview ? nReview + ' 段需要你核对（改动了数字、日期、术语，或和原文差别很大），不会被批量替换。' : '';
    renderGroupBar($('#rwBar'), groups.rwb, buttons, info, rwRunning);
    const items = list.filter(c => c.status !== 'ignored' && c.status !== 'dismissed').map(c => ({ card: c, focus: false }));
    renderInto($('#rwCards'), items, renderCard);
  }

  function renderRewriteBody(card) {
    const wrap = el('div', { class: 'rw-wrap' });
    const busy = card.busy || rwRunning;
    if (card.notes) wrap.appendChild(el('div', { class: 'overall' }, el('div', { class: 'cap' }, '改动说明'), el('div', {}, card.notes)));
    card.versions.forEach(v => {
      const inUse = card.status === 'applied' && TU.normalize(v.text) === TU.normalize(card.current);
      wrap.appendChild(el('div', { class: 'card' + (inUse ? ' is-applied' : '') },
        el('div', { class: 'card-head' }, el('span', { class: 'chip t-clarity' }, v.label), inUse ? el('span', { class: 'chip ok' }, '✓ 已替换') : null),
        el('div', { class: 'rw-text' }, v.text),
        v.warnings.map(w => el('div', { class: 'warn' }, '⚠ ' + w)),
        el('div', { class: 'actions' },
          inUse ? btn('撤销', 'primary', () => undoCard(card), busy) : btn(card.status === 'applied' ? '换成这个' : '替换选中内容', 'primary', () => applyRewrite(card, v.text), busy),
          btn('复制', 'ghost', () => copyWithToast(v.text, '已复制，可以粘贴到邮件里')))));
    });
    append(wrap, renderConfirm(card));
    wrap.appendChild(el('div', { class: 'actions' },
      btn('都不满意，重新生成', '', () => runRewrite(true), busy),
      card.status === 'applied' ? btn('撤销上一步', 'ghost', () => undoCard(card), busy) : null,
      card.busy ? el('span', { class: 'muted small' }, el('span', { class: 'spinner' }), ' 处理中…') : null));
    wrap.appendChild(el('details', { class: 'history' },
      el('summary', {}, '原文（想恢复时也可以复制回去）'),
      el('div', { class: 'rw-text' }, card.original),
      el('div', { class: 'actions' }, btn('复制原文', 'ghost', () => copyWithToast(card.original, '已复制原文')))));
    return wrap;
  }

  async function applyRewrite(card, text) {
    if (card.busy) return;
    if (!beginWrite()) return;
    card.busy = true;
    renderRewrite();
    try {
      const from = card.current;
      let r;
      if (card.status !== 'applied' && card.rwSource === 'selection') {
        // 第一次替换：优先替换当前选中的文字
        let sel = '';
        try { sel = await host.getSelectedText(); } catch (e) { sel = ''; }
        if (sel.trim() && TU.normalize(sel) !== TU.normalize(from)) {
          const choice = await askInCard(card)('你现在选中的文字和改写前的原文不一样。仍然用改写结果替换当前选中的内容吗？', ['替换', '取消']);
          if (choice !== 0) return;
          await host.replaceSelection(text);
          r = { ok: true, method: 'selection', entry: host.pushUndo({ original: sel, replacement: text, method: 'selection' }) };
        }
      }
      if (!r) r = await replaceWithPolicy(from, text, askInCard(card));
      if (r.cancelled) return;
      if (!r.ok) { toast(reasonText(r.reason), 8000); return; }
      if (r.entry) r.entry.cardId = card.id;
      card.history.push({ from: r.entry ? r.entry.original : from, to: text, entry: r.entry });
      card.current = text;
      card.ctx = r.ctx || null;
      card.status = 'applied';
      markApplied(text);
      rebaseLive();
      updateUndo();
      toast('已替换。不满意可以点「撤销」，或者换另一个版本。', 5000);
    } catch (e) {
      toast('替换失败：' + errText(e), 8000);
    } finally {
      endWrite();
      card.busy = false;
      renderRewrite();
    }
  }

  // ---------------------------------------------------------------------------
  // 浏览器模式：下载 Outlook 安装文件
  // ---------------------------------------------------------------------------
  function downloadManifest() {
    if (location.protocol !== 'https:') {
      toast('请先把网页发布到 https 网址（例如 GitHub Pages），再从那个网址打开本页下载安装文件。', 8000);
      return;
    }
    let xml;
    try { xml = MANIFEST.buildManifest(new URL('.', location.href).href); } catch (e) { toast(e.message, 6000); return; }
    const url = URL.createObjectURL(new Blob([xml], { type: 'application/xml' }));
    const a = el('a', { href: url, download: 'mailpolish-manifest.xml' });
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1500);
    toast('已下载 mailpolish-manifest.xml，安装方法见安装指南。', 6000);
  }

  // ---------------------------------------------------------------------------
  // 页面切换与启动
  // ---------------------------------------------------------------------------
  function switchTab(name) {
    if (['live', 'full', 'rewrite'].indexOf(name) < 0) name = 'live';
    document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
    ['live', 'full', 'rewrite'].forEach(n => { $('#tab-' + n).hidden = n !== name; });
    store.set('tab', name);
  }

  function onItemChanged() {
    itemSeq += 1;
    cancelPendingAsks();
    host.reset();
    resetLive();
    clear($('#fullOverall'));
    fullCardIds.forEach(id => cards.delete(id));
    fullCardIds = [];
    clearRewriteResults();
    groups.full.confirm = null;
    rewriteNotice('');
    rerender();
    updateUndo();
    applyLiveToggle();
  }

  function bindEvents() {
    document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => switchTab(t.dataset.tab)));
    $('#btnSettings').addEventListener('click', () => ($('#settingsView').hidden ? openSettings() : closeSettings()));
    $('#btnCloseSettings').addEventListener('click', closeSettings);
    $('#btnSave').addEventListener('click', onSaveSettings);
    $('#btnTest').addEventListener('click', onTestConnection);
    $('#btnClearKey').addEventListener('click', onClearKey);
    $('#btnShowKey').addEventListener('click', () => {
      const k = $('#apiKey');
      k.type = k.type === 'password' ? 'text' : 'password';
      $('#btnShowKey').textContent = k.type === 'password' ? '显示' : '隐藏';
    });
    $('#liveToggle').addEventListener('change', e => {
      settings.liveEnabled = !!e.target.checked;
      saveSettings();
      applyLiveToggle();
    });
    $('#btnFull').addEventListener('click', runFullCheck);
    $('#btnRewrite').addEventListener('click', () => runRewrite(false));
    $('#btnUndo').addEventListener('click', onUndo);
    $('#btnManifest').addEventListener('click', downloadManifest);
  }

  let booted = false;
  function boot(info) {
    if (booted) return;
    booted = true;
    const outlookType = window.Office && Office.HostType ? Office.HostType.Outlook : 'Outlook';
    inOutlook = !!(info && info.host === outlookType && Office.context && Office.context.mailbox);

    let partition = '';
    try { partition = inOutlook && Office.context.partitionKey ? Office.context.partitionKey : ''; } catch (e) { partition = ''; }
    store.init(partition);
    loadSettings();

    if (inOutlook) {
      host = new HOSTS.OutlookHost();
      document.body.classList.add('outlook-mode');
      try {
        if (Office.EventType && Office.EventType.ItemChanged && Office.context.mailbox.addHandlerAsync) {
          Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
        }
      } catch (e) { /* 旧版本不支持固定侧边栏事件，忽略 */ }
    } else {
      host = new HOSTS.StandaloneHost($('#saEditor'), $('#saContext'), $('#saSubject'));
      document.body.classList.add('standalone-mode');
      $('#standalone').hidden = false;
    }

    bindEvents();
    buildRewriteControls();
    fillSettingsForm();
    renderUsage();
    updateUndo();
    switchTab(store.get('tab', 'live'));
    if (!settings.apiKey) showSetupNotice();
    applyLiveToggle();
    rerender();
  }

  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady(info => boot(info));
    // office.js 万一迟迟没有响应（例如网络拦截），10 秒后按浏览器模式启动
    setTimeout(() => { if (!booted) boot(null); }, 10000);
  } else {
    boot(null);
  }
})();
