/*
 * MailPolish — "宿主"适配层
 * 同一套界面可以跑在两种环境里：
 *  - OutlookHost：装进 Outlook 后，通过微软 Office.js 读写正在编辑的邮件
 *  - StandaloneHost：直接用浏览器打开网页时，读写页面上的文本框（用来试用/测试）
 *
 * 替换一句话的策略：
 *  1. 如果你在邮件里选中的正好是这句话 → 只替换选中内容（不动其它内容）
 *  2. 否则在 HTML 正文里精确定位这句话，只改动不同的那几个字，再写回正文
 *     （微软接口限制：写回正文后光标会跳动；经典版 Outlook 不能用 Ctrl+Z 撤销，所以我们自己记录撤销）
 *  3. 写回后立刻读回检查：图片变少或文字大量丢失 → 自动恢复原样，并报告失败
 *  4. 设置为「有图片时先确认」时，正文有图片会先返回 has_images，由界面询问用户
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

  /** 撤销记录（两种宿主共用） */
  const undoMethods = {
    pushUndo(entry) {
      this.undoStack.push(entry);
      if (this.undoStack.length > 30) this.undoStack.shift();
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
      const r = await this.replaceSentence(entry.replacement, entry.original, Object.assign({}, opts, { record: false }));
      if (r.ok) this.removeUndo(entry);
      return r;
    },
    async undo(opts) { return this.revertEntry(this.lastUndo(), opts); },
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

    getBodyText() { return officeCall(cb => this.item.body.getAsync(Office.CoercionType.Text, cb)); }
    getBodyHtml() { return officeCall(cb => this.item.body.getAsync(Office.CoercionType.Html, cb)); }
    setBodyHtml(html) { return officeCall(cb => this.item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb)); }

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
     * 返回 { ok, method: 'selection' | 'body', entry } 或 { ok:false, reason }
     */
    async replaceSentence(original, replacement, opts) {
      const o = opts || {};
      const record = o.record !== false;
      let sel = '';
      try { sel = await this.getSelectedText(); } catch (e) { sel = ''; }
      if (sel && TU.normalize(sel) === TU.normalize(original)) {
        await this.replaceSelection(replacement);
        return { ok: true, method: 'selection', entry: record ? this.pushUndo({ original, replacement, method: 'selection' }) : null };
      }
      const html = await this.getBodyHtml();
      if (!o.allowFullBody && countImages(html) > 0) return { ok: false, reason: 'has_images' };
      const r = TU.replaceInHtml(html, original, replacement);
      if (!r.ok) return r;
      if (!r.unchanged) {
        await this.setBodyHtml(r.html);
        const check = await this.verifyAfterWrite(html, original, replacement);
        if (!check.ok) return check;
      }
      return { ok: true, method: 'body', entry: record ? this.pushUndo({ original, replacement, method: 'body' }) : null };
    }

    /**
     * 写回正文后自动检查：图片数量没有变少、文字没有大量丢失。
     * 发现问题就把正文恢复成写回之前的样子，并返回失败。
     */
    async verifyAfterWrite(beforeHtml, original, replacement) {
      let after;
      try { after = await this.getBodyHtml(); } catch (e) { return { ok: true }; } // 读不回来就不判断
      const imgsBefore = countImages(beforeHtml);
      const imgsAfter = countImages(after);
      const lenBefore = TU.normalize(TU.htmlToText(beforeHtml)).length;
      const lenAfter = TU.normalize(TU.htmlToText(after)).length;
      const expected = lenBefore - TU.normalize(original).length + TU.normalize(replacement).length;
      if (imgsAfter < imgsBefore || lenAfter < expected * 0.9 - 20) {
        try { await this.setBodyHtml(beforeHtml); } catch (e) { /* 尽力恢复 */ }
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
      const record = !opts || opts.record !== false;
      const e = this.editor;
      const sel = e.value.slice(e.selectionStart, e.selectionEnd);
      if (sel && TU.normalize(sel) === TU.normalize(original)) {
        await this.replaceSelection(replacement);
        return { ok: true, method: 'selection', entry: record ? this.pushUndo({ original, replacement, method: 'selection' }) : null };
      }
      const r = TU.findInText(e.value, original);
      if (!r.ok) return r;
      e.setRangeText(replacement, r.start, r.end, 'preserve');
      e.dispatchEvent(new Event('input', { bubbles: true }));
      return { ok: true, method: 'body', entry: record ? this.pushUndo({ original, replacement, method: 'body' }) : null };
    }
  }
  Object.assign(StandaloneHost.prototype, undoMethods);

  root.MailPolishHost = { OutlookHost, StandaloneHost, countImages };
})(window);
