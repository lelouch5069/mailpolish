/*
 * MailPolish — "宿主"适配层
 * 同一套界面可以跑在两种环境里：
 *  - OutlookHost：装进 Outlook 后，通过微软 Office.js 读写正在编辑的邮件
 *  - StandaloneHost：直接用浏览器打开网页时，读写页面上的文本框（用来试用/测试）
 *
 * 替换的策略：
 *  1. 如果你在邮件里选中的正好是这句话 → 只替换选中内容（不动其它内容）
 *  2. 否则在 HTML 正文里精确定位这句话，只改动不同的那几个字，再写回正文
 *     （微软接口限制：写回正文后光标会跳动；经典版 Outlook 不能用 Ctrl+Z 撤销，所以我们自己记录撤销）
 *  3. 一次替换多处（全部替换）：只读一次正文、逐处修改、只写回一次
 *  4. 同样的文字在下面引用的往来邮件里也有时，只改「我写的部分」里的那一处；只在引用里的文字一律不改
 *     同样的文字在我写的部分出现多次时，靠记下的前后文认出是哪一处
 *  5. 写回后立刻读回检查：图片变少或文字大量丢失 → 自动恢复原样，并报告失败
 *  6. 设置为「有图片时先确认」时，正文有图片会先返回 has_images，由界面询问用户
 */
(function (root) {
  'use strict';
  const TU = root.TextUtil;

  function officeCall(fn) {
    return new Promise((resolve, reject) => {
      try {
        fn(res => {
          if (res && res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
          else {
            const err = new Error((res && res.error && res.error.message) || 'Office 操作失败');
            err.kind = 'office';
            err.officeError = res && res.error;
            reject(err);
          }
        });
      } catch (e) {
        e.kind = e.kind || 'office';
        reject(e);
      }
    });
  }

  function countImages(html) {
    return (String(html || '').match(/<img\b/gi) || []).length;
  }

  /** 预计替换后文字变长（正数）或变短（负数）多少个字符 */
  function sizeDelta(pairs, results) {
    let d = 0;
    pairs.forEach((p, i) => {
      if (results[i] && results[i].ok) d += TU.normalize(p.replacement).length - TU.normalize(p.original).length;
    });
    return d;
  }

  function firstReason(results) {
    const bad = results.find(r => r && !r.ok);
    return (bad && bad.reason) || 'empty';
  }

  /** 撤销记录（两种宿主共用） */
  const undoMethods = {
    pushUndo(entry) {
      this.undoStack.push(entry);
      if (this.undoStack.length > 200) this.undoStack.shift();
      return entry;
    },
    removeUndo(entry) {
      const i = this.undoStack.lastIndexOf(entry);
      if (i >= 0) this.undoStack.splice(i, 1);
    },
    canUndo() { return this.undoStack.length > 0; },
    lastUndo() { return this.undoStack[this.undoStack.length - 1] || null; },
    /** 撤销某一次替换：把新文字换回原文（不会丢掉你之后写的其它内容） */
    async revertEntry(entry, opts) {
      if (!entry) return { ok: false, reason: 'empty' };
      const r = await this.replaceSentence(entry.replacement, entry.original,
        Object.assign({}, opts, { record: false, ctx: entry.ctx }));
      if (r.ok) this.removeUndo(entry);
      return r;
    },
    async undo(opts) { return this.revertEntry(this.lastUndo(), opts); },
    /** 把 replaceMany 成功的每一处记入撤销记录 */
    recordMany(pairs, results) {
      results.forEach((res, i) => {
        if (res.ok) res.entry = this.pushUndo({ original: pairs[i].original, replacement: pairs[i].replacement, method: 'body', ctx: res.ctx || null });
      });
    },
    reset() { this.undoStack = []; }
  };

  class OutlookHost {
    constructor() {
      this.kind = 'outlook';
      this.undoStack = [];
    }

    get item() { return Office.context.mailbox.item; }

    /** 当前是否在"写邮件/回复"界面 */
    isCompose() {
      const it = this.item;
      return !!it && typeof it.subject === 'object' && !!it.body && typeof it.body.setSelectedDataAsync === 'function';
    }

    platform() {
      try { return String((Office.context.diagnostics && Office.context.diagnostics.platform) || ''); } catch (e) { return ''; }
    }

    getBodyText(item) { const it = item || this.item; return officeCall(cb => it.body.getAsync(Office.CoercionType.Text, cb)); }
    getBodyHtml(item) { const it = item || this.item; return officeCall(cb => it.body.getAsync(Office.CoercionType.Html, cb)); }
    setBodyHtml(html, item) { const it = item || this.item; return officeCall(cb => it.body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb)); }

    async getDraft() {
      const text = await this.getBodyText();
      return TU.splitEmail(text || '');
    }

    async getMeta() {
      const it = this.item;
      let subject = '';
      let composeType = 'unknown';
      try { subject = await officeCall(cb => it.subject.getAsync(cb)); } catch (e) { /* 忽略 */ }
      if (typeof it.getComposeTypeAsync === 'function') {
        try { const v = await officeCall(cb => it.getComposeTypeAsync(cb)); composeType = (v && v.composeType) || 'unknown'; } catch (e) { /* 忽略 */ }
      }
      return { subject: subject || '', composeType };
    }

    async getSelectedText() {
      const v = await officeCall(cb => this.item.getSelectedDataAsync(Office.CoercionType.Text, cb));
      if (!v || v.sourceProperty !== 'body') return '';
      return v.data || '';
    }

    replaceSelection(text) {
      return officeCall(cb => this.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, cb));
    }

    /**
     * 把 original 改成 replacement。
     * opts.allowFullBody：正文有图片时也允许写回整个正文（一键替换模式）
     * opts.record：是否记入撤销记录（默认记）
     * opts.ctx：这段文字的前后文（同样的文字出现多次时用来认出是哪一处）
     * 返回 { ok, method: 'selection' | 'body', entry, ctx } 或 { ok:false, reason }；ctx 是新文字的前后文
     */
    async replaceSentence(original, replacement, opts) {
      const o = opts || {};
      const record = o.record !== false;
      let sel = '';
      try { sel = await this.getSelectedText(); } catch (e) { sel = ''; }
      if (sel && TU.normalize(sel) === TU.normalize(original)) {
        await this.replaceSelection(replacement);
        return { ok: true, method: 'selection', entry: record ? this.pushUndo({ original, replacement, method: 'selection', ctx: null }) : null, ctx: null };
      }
      const pairs = [{ original, replacement, ctx: o.ctx || null }];
      const r = await this.replaceMany(pairs, Object.assign({}, o, { record }));
      if (!r.ok) return { ok: false, reason: r.reason };
      return { ok: true, method: 'body', entry: r.results[0].entry || null, ctx: r.results[0].ctx || null };
    }

    /**
     * 一次替换多处（全部替换）：只读一次正文、逐处精确替换、只写回一次，再统一核对。
     * pairs: [{ original, replacement, ctx? }]
     * 返回 { ok, reason?, results: [{ ok, reason?, entry?, ctx? }] }；某几处找不到不影响其它处。
     */
    async replaceMany(pairs, opts) {
      const o = opts || {};
      const item = this.item; // 整个过程固定在这封邮件上
      const html = await this.getBodyHtml(item);
      if (!o.allowFullBody && countImages(html) > 0) return { ok: false, reason: 'has_images' };
      let mine = '';
      let quoted = '';
      try {
        const parts = TU.splitEmail((await this.getBodyText(item)) || '');
        mine = parts.mine;
        quoted = TU.normalize(parts.quoted);
      } catch (e) { mine = ''; quoted = ''; }
      let cur = html;
      const results = pairs.map(() => null);
      const steps = TU.planReplacements(mine, pairs);
      steps.forEach(step => {
        if (step.reason) { results[step.i] = { ok: false, reason: step.reason }; return; }
        const p = pairs[step.i];
        // 不在我写的部分、却在引用的往来邮件里：不改（可能是你刚改过这句，引用里恰好有一样的话）
        if (step.pos < 0 && mine && quoted.includes(TU.normalize(p.original))) { results[step.i] = { ok: false, reason: 'not_found' }; return; }
        const r = TU.replaceInHtml(cur, p.original, p.replacement, { occurrence: step.occurrence });
        if (!r.ok) { results[step.i] = { ok: false, reason: r.reason }; return; }
        cur = r.html;
        results[step.i] = { ok: true };
      });
      if (!results.some(r => r.ok)) return { ok: false, reason: firstReason(results), results };
      if (cur !== html) {
        if (this.item !== item) return { ok: false, reason: 'item_changed', results: pairs.map(() => ({ ok: false, reason: 'item_changed' })) };
        await this.setBodyHtml(cur, item);
        const check = await this.verifyAfterWrite(html, sizeDelta(pairs, results), item);
        if (!check.ok) return { ok: false, reason: check.reason, results: pairs.map(() => ({ ok: false, reason: check.reason })) };
      }
      const okIdx = new Set(results.map((r, i) => (r.ok ? i : -1)).filter(i => i >= 0));
      const ctxs = TU.contextsAfter(mine, pairs, steps, okIdx);
      results.forEach((r, i) => { if (r.ok) r.ctx = ctxs[i]; });
      if (o.record !== false) this.recordMany(pairs, results);
      return { ok: true, results };
    }

    /**
     * 写回正文后自动检查：图片数量没有变少、文字没有大量丢失。
     * delta：这次替换预计让文字变长（正数）或变短（负数）多少个字符。
     * 发现问题就把正文恢复成写回之前的样子，并返回失败。
     */
    async verifyAfterWrite(beforeHtml, delta, item) {
      let after;
      try { after = await this.getBodyHtml(item); } catch (e) { return { ok: true }; } // 读不回来就不判断
      const imgsBefore = countImages(beforeHtml);
      const imgsAfter = countImages(after);
      const lenBefore = TU.normalize(TU.htmlToText(beforeHtml)).length;
      const lenAfter = TU.normalize(TU.htmlToText(after)).length;
      const expected = lenBefore + (Number(delta) || 0);
      if (imgsAfter < imgsBefore || lenAfter < expected * 0.9 - 20) {
        try { await this.setBodyHtml(beforeHtml, item); } catch (e) { /* 尽力恢复 */ }
        return { ok: false, reason: imgsAfter < imgsBefore ? 'image_lost' : 'text_lost' };
      }
      return { ok: true };
    }
  }
  Object.assign(OutlookHost.prototype, undoMethods);

  class StandaloneHost {
    constructor(editor, contextEl, subjectEl) {
      this.kind = 'standalone';
      this.editor = editor;
      this.contextEl = contextEl;
      this.subjectEl = subjectEl;
      this.undoStack = [];
    }

    isCompose() { return true; }
    platform() { return 'Browser'; }

    async getDraft() {
      return { mine: this.editor.value, quoted: this.contextEl.value };
    }

    async getMeta() {
      return { subject: this.subjectEl.value, composeType: this.contextEl.value.trim() ? 'reply' : 'newMail' };
    }

    async getSelectedText() {
      const e = this.editor;
      return e.value.slice(e.selectionStart, e.selectionEnd);
    }

    async replaceSelection(text) {
      const e = this.editor;
      e.setRangeText(text, e.selectionStart, e.selectionEnd, 'end');
      e.dispatchEvent(new Event('input', { bubbles: true }));
    }

    async replaceSentence(original, replacement, opts) {
      const o = opts || {};
      const record = o.record !== false;
      const e = this.editor;
      const sel = e.value.slice(e.selectionStart, e.selectionEnd);
      if (sel && TU.normalize(sel) === TU.normalize(original)) {
        await this.replaceSelection(replacement);
        return { ok: true, method: 'selection', entry: record ? this.pushUndo({ original, replacement, method: 'selection', ctx: null }) : null, ctx: null };
      }
      const r = await this.replaceMany([{ original, replacement, ctx: o.ctx || null }], { record });
      if (!r.ok) return { ok: false, reason: r.reason };
      return { ok: true, method: 'body', entry: r.results[0].entry || null, ctx: r.results[0].ctx || null };
    }

    async replaceMany(pairs, opts) {
      const e = this.editor;
      const before = e.value;
      const results = pairs.map(() => null);
      const steps = TU.planReplacements(before, pairs);
      steps.forEach(step => {
        if (step.reason) { results[step.i] = { ok: false, reason: step.reason }; return; }
        const p = pairs[step.i];
        const r = TU.findInText(e.value, p.original, step.occurrence);
        if (!r.ok) { results[step.i] = { ok: false, reason: r.reason }; return; }
        e.setRangeText(p.replacement, r.start, r.end, 'preserve');
        results[step.i] = { ok: true };
      });
      if (!results.some(r => r.ok)) return { ok: false, reason: firstReason(results), results };
      e.dispatchEvent(new Event('input', { bubbles: true }));
      const okIdx = new Set(results.map((r, i) => (r.ok ? i : -1)).filter(i => i >= 0));
      const ctxs = TU.contextsAfter(before, pairs, steps, okIdx);
      results.forEach((r, i) => { if (r.ok) r.ctx = ctxs[i]; });
      if (!opts || opts.record !== false) this.recordMany(pairs, results);
      return { ok: true, results };
    }
  }
  Object.assign(StandaloneHost.prototype, undoMethods);

  root.MailPolishHost = { OutlookHost, StandaloneHost, countImages };
})(window);
