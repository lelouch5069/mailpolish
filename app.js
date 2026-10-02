/*
 * MailPolish — 主程序：界面与流程
 *
 * 「边写边查」的原理（Outlook 不提供光标位置和打字事件，所以这样近似）：
 *   每 1 秒读一次草稿 → 和上一次对比，变化的位置 ≈ 你正在打字的地方 → 找到那一句
 *   → 你停顿约 2 秒后，把这句（连同上下文）发给 AI → 结果显示在侧边栏，点「替换」写回邮件
 *   查过的句子会记住结果，不会重复花钱。
 */
(function () {
  'use strict';

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

  function setBusy(btn, busyText) {
    const label = btn.textContent;
    btn.disabled = true;
    clear(btn);
    append(btn, [el('span', { class: 'spinner' }), ' ' + busyText]);
    return () => { btn.disabled = false; btn.textContent = label; };
  }

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
    $('#about').value = settings.about || '';
    $('#terms').value = settings.terms || '';
    $('#pauseMs').value = String(settings.pauseMs || 2000);
    if (!$('#pauseMs').value) $('#pauseMs').value = '2000';
    $('#dailyCap').value = String(settings.dailyCap || AI.DEFAULT_SETTINGS.dailyCap);
    $('#baseUrl').value = settings.baseUrl || AI.DEFAULT_SETTINGS.baseUrl;
    $('#testResult').textContent = '';
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
  // 顶部提示条
  // ---------------------------------------------------------------------------
  function showNotice(message, actions) {
    const n = $('#notice');
    clear(n);
    n.appendChild(el('div', {}, message));
    if (actions && actions.length) {
      n.appendChild(el('div', { class: 'actions' }, actions.map(([label, fn]) => el('button', { class: 'btn small primary', onclick: fn }, label))));
    }
    n.hidden = false;
  }
  function hideNotice() { $('#notice').hidden = true; }
  function showSetupNotice() {
    showNotice('还没有填写 DeepSeek API Key。填好后就可以开始检查。', [['去设置', openSettings]]);
  }

  // ---------------------------------------------------------------------------
  // 边写边查
  // ---------------------------------------------------------------------------
  const live = {
    timer: null,
    busy: false,
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

  function resetLive() {
    live.prevMine = null;
    live.quoted = '';
    live.lastChangeAt = 0;
    live.dirty.clear();
    live.inFlight.clear();
    live.present = new Set();
    live.focus = null;
    live.focusPara = '';
    live.history = [];
    live.bulkHint = false;
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
    if (live.busy) return;
    live.busy = true;
    try {
      const draft = await host.getDraft();
      if (live.hostError) { live.hostError = false; setLiveStatus('idle'); }
      const mine = draft.mine || '';
      live.quoted = draft.quoted || '';

      if (live.prevMine === null) { // 第一次读取：记下现有内容，不主动检查
        live.prevMine = mine;
        live.present = presentKeys(mine);
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
          for (const s of ch.sentences) live.dirty.set(TU.sentenceKey(s.text), s.text);
        }
        live.focus = focus ? { key: TU.sentenceKey(focus.text), text: focus.text } : null;
        live.focusPara = focus ? TU.paragraphAt(mine, focus.start) : '';
        live.present = presentKeys(mine);
        for (const k of Array.from(live.dirty.keys())) if (!live.present.has(k)) live.dirty.delete(k);
        setLiveStatus('typing');
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
        setLiveStatus('error', '读取邮件内容失败：' + (e && e.message ? e.message : e));
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
    try {
      const meta = await host.getMeta();
      const req = AI.buildCheckRequest(sentences, {
        subject: meta.subject, composeType: meta.composeType, quoted: live.quoted, paragraph: live.focusPara
      }, settings, 'live');
      const r = await callAI(Object.assign({}, req, { timeoutMs: 30000 }));
      const out = AI.normalizeCheckResult(r.json, sentences, settings);
      const byId = new Map(out.results.map(x => [x.id, x]));
      sentences.forEach((s, i) => {
        const res = byId.get(s.id) || null;
        live.cache.set(items[i].key, { status: res ? 'issues' : 'ok', result: res, text: s.text, at: Date.now(), source: 'live' });
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

    // 2) 建议卡片：正文里还在、有问题、没被忽略或替换的句子；正在写的那句排最前
    const wrap = $('#liveCards');
    clear(wrap);
    const list = [];
    for (const [key, entry] of live.cache) {
      if (entry.status !== 'issues' || entry.source !== 'live' || !live.present.has(key) || !entry.result) continue;
      list.push({ key, entry, isFocus: !!f && f.key === key });
    }
    list.sort((a, b) => (Number(b.isFocus) - Number(a.isFocus)) || (b.entry.at - a.entry.at));
    list.slice(0, 6).forEach(({ key, entry, isFocus }) => {
      wrap.appendChild(renderResultCard(key, entry.result, { focus: isFocus, onDone: () => renderLive() }));
    });
    $('#liveHint').hidden = list.length > 0 || showCurrent;

    // 3) 最近检查
    renderHistory();
  }

  function renderHistory() {
    const box = $('#liveHistory');
    clear(box);
    const items = live.history.map(k => live.cache.get(k)).filter(Boolean);
    $('#liveHistoryWrap').hidden = items.length === 0;
    $('#histCount').textContent = String(items.length);
    const marks = { ok: '✓', applied: '✓', issues: '!', ignored: '–', error: '×' };
    for (const e of items) {
      box.appendChild(el('div', { class: 'hist-item ' + e.status }, el('span', { class: 'mark' }, marks[e.status] || '·'), el('span', {}, e.text)));
    }
  }

  // ---------------------------------------------------------------------------
  // 建议卡片（边写边查和全文检查共用）
  // ---------------------------------------------------------------------------
  function renderResultCard(key, res, opts) {
    const card = el('div', { class: 'card' + (opts.focus ? ' focus' : '') });
    const types = Array.from(new Set(res.issues.map(i => i.type)));
    card.appendChild(el('div', { class: 'card-head' },
      opts.focus ? el('span', { class: 'chip badge' }, '正在写的句子') : null,
      types.map(t => el('span', { class: 'chip t-' + t }, AI.ISSUE_TYPES[t] || t)),
      !types.length ? el('span', { class: 'chip t-tone' }, '表达') : null));

    if (res.changed) {
      card.appendChild(el('div', { class: 'diff' },
        TU.wordDiff(res.original, res.corrected).map(op => (op.op === 'eq' ? op.text : el(op.op === 'del' ? 'del' : 'ins', {}, op.text)))));
    } else {
      card.appendChild(el('div', { class: 'diff' }, res.original));
    }

    if (res.issues.length) {
      card.appendChild(el('ul', { class: 'issue-list' }, res.issues.map(i => el('li', {},
        el('span', { class: 'fix' }, (i.original || '∅') + ' → ' + (i.suggestion || '（删除）')),
        i.explain ? el('span', { class: 'why' }, '　' + i.explain) : null))));
    }
    res.warnings.forEach(w => card.appendChild(el('div', { class: 'warn' }, '⚠ ' + w)));

    const actions = el('div', { class: 'actions' });
    if (res.changed) actions.appendChild(el('button', { class: 'btn small primary', onclick: () => applyFix(card, key, res.original, res.corrected, opts) }, '替换'));
    actions.appendChild(el('button', { class: 'btn small', onclick: () => { markIgnored(key); card.remove(); if (opts.onDone) opts.onDone('ignored'); } }, '忽略'));
    if (res.changed) {
      actions.appendChild(el('button', {
        class: 'btn small ghost',
        onclick: async () => toast((await copyText(res.corrected)) ? '已复制' : '复制失败，请手动选择文字')
      }, '复制'));
    }
    card.appendChild(actions);

    if (res.better) {
      card.appendChild(el('div', { class: 'better' },
        el('div', { class: 'cap' }, '更专业的写法'),
        el('div', { class: 'text' }, res.better),
        res.betterExplain ? el('div', { class: 'note' }, res.betterExplain) : null,
        res.betterWarnings.map(w => el('div', { class: 'warn' }, '⚠ ' + w)),
        el('div', { class: 'actions' },
          el('button', { class: 'btn small', onclick: () => applyFix(card, key, res.original, res.better, opts) }, '用这个'),
          el('button', { class: 'btn small ghost', onclick: async () => toast((await copyText(res.better)) ? '已复制' : '复制失败，请手动选择文字') }, '复制'))));
    }
    return card;
  }

  function markIgnored(key) {
    const e = live.cache.get(key);
    if (e) e.status = 'ignored';
  }

  function confirmInCard(card, message, labels) {
    return new Promise(resolve => {
      const box = el('div', { class: 'confirm' }, el('div', {}, message));
      const acts = el('div', { class: 'actions' });
      labels.forEach((label, i) => acts.appendChild(el('button', {
        class: 'btn small' + (i === 0 ? ' primary' : ''),
        onclick: () => { box.remove(); resolve(i); }
      }, label)));
      box.appendChild(acts);
      card.appendChild(box);
    });
  }

  function reasonText(reason) {
    return ({
      not_found: '在正文里没找到这句（可能刚刚又改过）。可以在邮件里选中它，再点「替换」。',
      multiple: '这句话在正文里出现了不止一次。请先在邮件里选中要改的那一处，再点「替换」。',
      crosses_paragraph: '这句话跨了段落，没法自动定位。请先在邮件里选中它，再点「替换」。',
      empty: '没有可替换的内容。',
      has_images: '请先在邮件里选中这句话，再点按钮。'
    })[reason] || '替换没有成功。';
  }

  async function applyFix(card, key, original, replacement, opts) {
    const buttons = Array.from(card.querySelectorAll('.actions button'));
    buttons.forEach(b => { b.disabled = true; });
    try {
      let r = await host.replaceSentence(original, replacement, { allowFullBody: false });
      if (!r.ok && r.reason === 'has_images') {
        const choice = await confirmInCard(card,
          '这封邮件里有图片（比如签名里的 logo）。直接替换需要把整个正文写回 Outlook，我没法保证图片一定不受影响。更稳妥的做法：先在邮件里选中这句话，再点「替换」。',
          ['仍然直接替换', '我先去选中']);
        if (choice !== 0) return;
        r = await host.replaceSentence(original, replacement, { allowFullBody: true });
      }
      if (!r.ok) { toast(reasonText(r.reason), 6500); return; }

      // 记住：替换后的句子不需要再检查
      TU.segmentSentences(replacement).forEach(s => {
        live.cache.set(TU.sentenceKey(s.text), { status: 'applied', result: null, text: s.text, at: Date.now(), source: 'live' });
      });
      const e = live.cache.get(key);
      if (e) e.status = 'applied';
      card.classList.add('applied');
      updateUndo();
      toast(r.method === 'body' && inOutlook ? '已替换。注意：光标可能跳到正文开头或结尾。' : '已替换');
      if (opts && opts.onDone) opts.onDone('applied');
    } catch (err) {
      toast('替换失败：' + (err && err.message ? err.message : err), 6500);
    } finally {
      buttons.forEach(b => { b.disabled = false; });
    }
  }

  function updateUndo() { $('#btnUndo').hidden = !(host && host.canUndo()); }

  async function onUndo() {
    const top = host.undoStack[host.undoStack.length - 1];
    try {
      const r = await host.undo();
      if (r.ok) {
        if (top) {
          const e = live.cache.get(TU.sentenceKey(top.original));
          if (e && e.status === 'applied' && e.result) e.status = 'issues';
        }
        toast('已撤销');
        renderLive();
      } else if (r.reason === 'has_images') {
        toast('请先在邮件里选中刚才替换的那句话，再点「撤销上次替换」。', 6500);
      } else {
        toast(reasonText(r.reason).replace('「替换」', '「撤销上次替换」'), 6500);
      }
    } catch (e) {
      toast('撤销失败：' + (e && e.message ? e.message : e), 6500);
    }
    updateUndo();
  }

  // ---------------------------------------------------------------------------
  // 全文检查
  // ---------------------------------------------------------------------------
  let fullRunning = false;

  async function runFullCheck() {
    if (fullRunning) return;
    if (!host.isCompose()) { toast('请在写邮件或回复时使用'); return; }
    if (!settings.apiKey) { showSetupNotice(); return; }
    fullRunning = true;
    const done = setBusy($('#btnFull'), '正在检查…');
    const ov = $('#fullOverall');
    const cards = $('#fullCards');
    clear(ov);
    clear(cards);
    try {
      const draft = await host.getDraft();
      const meta = await host.getMeta();
      const seen = new Set();
      let sentences = TU.segmentSentences(draft.mine).filter(s => {
        if (!TU.isCheckableEnglish(s.text)) return false;
        const k = TU.sentenceKey(s.text);
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
      });
      if (!sentences.length) { ov.appendChild(el('p', { class: 'hint' }, '没有找到需要检查的英文句子。')); return; }
      let note = '';
      if (sentences.length > 60) { sentences = sentences.slice(0, 60); note = '内容较长，这次只检查了前 60 句。'; }
      const list = sentences.map((s, i) => ({ id: i + 1, text: s.text }));
      const req = AI.buildCheckRequest(list, { subject: meta.subject, composeType: meta.composeType, quoted: draft.quoted }, settings, 'full');
      const r = await callAI(Object.assign({}, req, { timeoutMs: 120000 }));
      const out = AI.normalizeCheckResult(r.json, list, settings);
      // 记住结果，边写边查时就不会再重复检查这些句子
      list.forEach(s => {
        const res = out.results.find(x => x.id === s.id) || null;
        const key = TU.sentenceKey(s.text);
        const prev = live.cache.get(key);
        if (!prev || prev.source !== 'live') live.cache.set(key, { status: res ? 'issues' : 'ok', result: res, text: s.text, at: Date.now(), source: 'full' });
      });
      renderFull(out, note, list.length);
      live.bulkHint = false;
      renderLive();
    } catch (e) {
      ov.appendChild(el('div', { class: 'warn' }, AI.describeError(e)));
      if (e && e.kind === 'no_key') showSetupNotice();
    } finally {
      fullRunning = false;
      done();
    }
  }

  function renderFull(out, note, count) {
    const ov = $('#fullOverall');
    clear(ov);
    ov.appendChild(el('div', { class: 'overall' },
      el('div', { class: 'cap' }, '整体评价'),
      el('div', {}, out.overall || '（AI 没有给出整体评价）'),
      el('div', { class: 'muted small mt' },
        '检查了 ' + count + ' 句，' + (out.results.length ? '有 ' + out.results.length + ' 句可以改进。' : '没有发现需要修改的地方。') + (note ? ' ' + note : ''))));
    const wrap = $('#fullCards');
    clear(wrap);
    if (!out.results.length) { wrap.appendChild(el('div', { class: 'allgood' }, '✓ 没有发现问题')); return; }
    out.results.forEach(res => {
      wrap.appendChild(renderResultCard(TU.sentenceKey(res.original), res, { focus: false }));
    });
  }

  // ---------------------------------------------------------------------------
  // 改写
  // ---------------------------------------------------------------------------
  let rwTone = 'formal';
  let rwRunning = false;

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

  async function runRewrite() {
    if (rwRunning) return;
    const out = $('#rewriteResults');
    clear(out);
    if (!host.isCompose()) { out.appendChild(el('div', { class: 'warn' }, '请在写邮件或回复时使用。')); return; }
    if (!settings.apiKey) { showSetupNotice(); return; }
    const sourceInput = document.querySelector('input[name="rwSource"]:checked');
    const source = sourceInput ? sourceInput.value : 'selection';
    let text = '';
    try {
      if (source === 'selection') {
        try { text = await host.getSelectedText(); } catch (e) { text = ''; } // 光标在收件人/主题栏时会读不到
        if (!text.trim()) {
          out.appendChild(el('div', { class: 'warn' }, '请先在邮件正文里选中要改写的文字（可以是中文），再点「生成改写」。'));
          return;
        }
      } else {
        text = ((await host.getDraft()).mine || '').trim();
        if (!text) { out.appendChild(el('div', { class: 'warn' }, '邮件里还没有你写的内容。')); return; }
      }
    } catch (e) {
      out.appendChild(el('div', { class: 'warn' }, '读取邮件内容失败：' + (e && e.message ? e.message : e)));
      return;
    }

    rwRunning = true;
    const done = setBusy($('#btnRewrite'), '正在改写…');
    try {
      const draft = await host.getDraft();
      const meta = await host.getMeta();
      const req = AI.buildRewriteRequest(text, { subject: meta.subject, quoted: draft.quoted }, settings, rwTone, $('#scenario').value, $('#rwExtra').value);
      const r = await callAI(Object.assign({}, req, { timeoutMs: 120000 }));
      renderRewrite(AI.normalizeRewriteResult(r.json, text, settings), text, source);
    } catch (e) {
      out.appendChild(el('div', { class: 'warn' }, AI.describeError(e)));
      if (e && e.kind === 'no_key') showSetupNotice();
    } finally {
      rwRunning = false;
      done();
    }
  }

  function renderRewrite(res, sourceText, source) {
    const out = $('#rewriteResults');
    clear(out);
    if (res.notes) out.appendChild(el('div', { class: 'overall' }, el('div', { class: 'cap' }, '改动说明'), el('div', {}, res.notes)));
    res.versions.forEach(v => {
      const card = el('div', { class: 'card' },
        el('div', { class: 'card-head' }, el('span', { class: 'chip t-clarity' }, v.label)),
        el('div', { class: 'rw-text' }, v.text),
        v.warnings.map(w => el('div', { class: 'warn' }, '⚠ ' + w)));
      card.appendChild(el('div', { class: 'actions' },
        el('button', { class: 'btn small primary', onclick: () => applyRewrite(card, v.text, sourceText) }, '替换选中内容'),
        el('button', { class: 'btn small', onclick: async () => toast((await copyText(v.text)) ? '已复制，可以粘贴到邮件里' : '复制失败，请手动选择文字') }, '复制')));
      out.appendChild(card);
    });
    out.appendChild(el('details', { class: 'history' },
      el('summary', {}, '原文（想恢复时可以复制回去）'),
      el('div', { class: 'rw-text' }, sourceText),
      el('div', { class: 'actions' }, el('button', { class: 'btn small ghost', onclick: async () => toast((await copyText(sourceText)) ? '已复制原文' : '复制失败') }, '复制原文'))));
    if (source === 'all') out.appendChild(el('p', { class: 'hint' }, '要整段替换：先在邮件里选中你写的部分，再点「替换选中内容」；也可以点「复制」后自己粘贴。'));
  }

  async function applyRewrite(card, text, sourceText) {
    let sel = '';
    try { sel = await host.getSelectedText(); } catch (e) { sel = ''; }
    if (!sel.trim()) { toast('请先在邮件里选中要替换的文字，再点「替换选中内容」。', 6500); return; }
    if (TU.normalize(sel) !== TU.normalize(sourceText)) {
      const choice = await confirmInCard(card, '你现在选中的文字和改写前的原文不一样。仍然用改写结果替换当前选中的内容吗？', ['替换', '取消']);
      if (choice !== 0) return;
    }
    try {
      await host.replaceSelection(text);
      TU.segmentSentences(text).forEach(s => {
        live.cache.set(TU.sentenceKey(s.text), { status: 'applied', result: null, text: s.text, at: Date.now(), source: 'live' });
      });
      toast('已替换。如果想恢复，下方「原文」里可以复制回去。', 5000);
    } catch (e) {
      toast('替换失败：' + (e && e.message ? e.message : e), 6500);
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
    host.reset();
    resetLive();
    clear($('#fullOverall'));
    clear($('#fullCards'));
    clear($('#rewriteResults'));
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
    $('#btnRewrite').addEventListener('click', runRewrite);
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
    renderLive();
  }

  if (window.Office && typeof Office.onReady === 'function') {
    Office.onReady(info => boot(info));
    // office.js 万一迟迟没有响应（例如网络拦截），10 秒后按浏览器模式启动
    setTimeout(() => { if (!booted) boot(null); }, 10000);
  } else {
    boot(null);
  }
})();
