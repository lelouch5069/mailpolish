/*
 * MailPolish — 文本工具（纯函数，不依赖 Outlook）
 * 浏览器里挂在 window.TextUtil；Node 测试时通过 module.exports 导出。
 *
 * 主要功能：
 *  - normalize / normalizeWithMap：统一空白、引号，方便"找同一句话"
 *  - splitEmail：把正文拆成「我写的部分」和「引用的往来邮件」
 *  - segmentSentences / sentenceAt：分句、找光标所在的句子
 *  - diffRegion：对比前后两次草稿，推算"正在编辑的位置"
 *  - replaceInHtml：在 HTML 正文里精确替换一句话（其余 HTML 原样保留）
 *  - wordDiff：逐词对比，用于显示删改
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------------------
  // 1. 规范化（用于比较/查找，不用于显示）
  // ---------------------------------------------------------------------------
  const ZERO_WIDTH = /[​-‍﻿­]/;
  const SPACE_LIKE = /[\s   　]/;
  const QUOTE_MAP = {
    '‘': "'", '’': "'", '‚': "'", '′': "'",
    '“': '"', '”': '"', '„': '"', '″': '"'
  };

  /**
   * 生成规范化字符串，并记录每个字符对应原文的位置。
   * 规则：去掉零宽字符；连续空白（含不换行空格）合并成一个空格；弯引号变直引号；首尾空白去掉。
   * 返回 { norm, map }，map[i] = norm[i] 在原文中的下标。
   */
  function normalizeWithMap(text) {
    const s = String(text == null ? '' : text);
    let norm = '';
    const map = [];
    let lastWasSpace = true; // 开头的空白直接丢弃
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ZERO_WIDTH.test(ch)) continue;
      if (SPACE_LIKE.test(ch)) {
        if (lastWasSpace) continue;
        norm += ' ';
        map.push(i);
        lastWasSpace = true;
        continue;
      }
      norm += QUOTE_MAP[ch] || ch;
      map.push(i);
      lastWasSpace = false;
    }
    if (norm.endsWith(' ')) { norm = norm.slice(0, -1); map.pop(); }
    return { norm, map };
  }

  function normalize(text) {
    return normalizeWithMap(text).norm;
  }

  /** 句子的身份标识：规范化后的文字（区分大小写，大小写错误也要能被检查出来） */
  function sentenceKey(text) {
    return normalize(text);
  }

  // ---------------------------------------------------------------------------
  // 2. 拆分「我写的部分」和「引用的往来邮件」
  // ---------------------------------------------------------------------------
  const REPLY_PATTERNS = [
    /^[ \t]*-{2,}[ \t]*Original Message[ \t]*-{2,}/im,
    /^[ \t]*-{2,}[ \t]*Forwarded message[ \t]*-{2,}/im,
    /^[ \t]*Begin forwarded message:/im,
    /^[ \t]*\*?From:\*?[ \t].*\r?\n(?:[ \t]*\r?\n)?[ \t]*\*?(?:Sent|Date):\*?[ \t]/im,
    /^[ \t]*发件人[:：].*\r?\n(?:[ \t]*\r?\n)?[ \t]*(?:发送时间|日期|时间)[:：]/m,
    /^[ \t]*On [^\r\n]{3,200}wrote:[ \t]*$/im,
    /^[ \t]*在 [^\r\n]{3,200}写道[:：][ \t]*$/m
  ];

  /**
   * 把邮件正文拆成：mine（自己写的部分）和 quoted（引用的历史邮件）。
   * 找最早出现的「回复/转发标头」作为分界；找不到就整封都算 mine。
   */
  function splitEmail(text) {
    const s = String(text == null ? '' : text);
    let cut = -1;
    for (const re of REPLY_PATTERNS) {
      const m = re.exec(s);
      if (m && (cut === -1 || m.index < cut)) cut = m.index;
    }
    if (cut === -1) return { mine: s, quoted: '' };
    let mine = s.slice(0, cut);
    // 网页版/新版 Outlook 会在标头前面加一行下划线，顺手去掉
    mine = mine.replace(/(?:\r?\n)[ \t]*_{8,}[ \t]*(?:\r?\n)?[\s]*$/, '\n');
    return { mine, quoted: s.slice(cut) };
  }

  // ---------------------------------------------------------------------------
  // 3. 分句
  // ---------------------------------------------------------------------------
  const ABBREVIATIONS = [
    'mr', 'mrs', 'ms', 'dr', 'prof', 'sr', 'jr', 'st', 'no', 'nos', 'inc', 'ltd', 'co',
    'corp', 'vs', 'approx', 'dept', 'est', 'jan', 'feb', 'mar', 'apr', 'jun', 'jul',
    'aug', 'sep', 'sept', 'oct', 'nov', 'dec', 'e.g', 'i.e', 'u.s', 'a.m', 'p.m', 'fig', 'ref', 'qty', 'pcs'
  ];
  const ABBR_RE = new RegExp('(?:^|[\\s(])(?:' + ABBREVIATIONS.map(a => a.replace(/\./g, '\\.')).join('|') + ')\\.$', 'i');

  function rawSegments(text) {
    const out = [];
    if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
      const seg = new Intl.Segmenter('en', { granularity: 'sentence' });
      for (const part of seg.segment(text)) {
        out.push({ start: part.index, end: part.index + part.segment.length });
      }
    } else {
      // 兜底：句末标点 + 空白，或换行
      const re = /[^.!?\r\n]*(?:[.!?]+["')\]]*[ \t]*|\r?\n|$)/g;
      let m;
      while ((m = re.exec(text)) !== null) {
        if (m[0].length === 0) { re.lastIndex++; if (re.lastIndex > text.length) break; continue; }
        out.push({ start: m.index, end: m.index + m[0].length });
      }
    }
    // 保证换行一定是分句点（Segmenter 通常已经这么做，这里再保险一次）
    const split = [];
    for (const r of out) {
      let start = r.start;
      for (let i = r.start; i < r.end; i++) {
        if (text[i] === '\n') {
          split.push({ start, end: i + 1 });
          start = i + 1;
        }
      }
      if (start < r.end) split.push({ start, end: r.end });
    }
    return split;
  }

  /**
   * 分句：返回 [{ text, start, end }]，text 已去掉首尾空白，start/end 是去空白后的位置。
   * 会把被缩写（如 "Mr."、"e.g."）错误切开的句子合并回来。
   */
  function segmentSentences(text) {
    const s = String(text == null ? '' : text);
    const raw = rawSegments(s);
    const merged = [];
    for (const r of raw) {
      const prev = merged[merged.length - 1];
      if (prev) {
        const prevText = s.slice(prev.start, prev.end);
        const prevTrim = prevText.replace(/\s+$/, '');
        if (!/\n\s*$/.test(prevText) && ABBR_RE.test(prevTrim)) {
          prev.end = r.end;
          continue;
        }
      }
      merged.push({ start: r.start, end: r.end });
    }
    const out = [];
    for (const r of merged) {
      const piece = s.slice(r.start, r.end);
      const lead = piece.length - piece.replace(/^\s+/, '').length;
      let trimmed = piece.trim();
      if (!trimmed) continue;
      let start = r.start + lead;
      // 行首的项目符号/编号（•、·、-、1.、a) 等）不算句子内容：不送检查，也不会被替换掉
      if (isLineStart(s, start)) {
        const m = LIST_MARKER_RE.exec(trimmed);
        if (m) {
          start += m[0].length;
          trimmed = trimmed.slice(m[0].length);
          if (!trimmed) continue;
        }
      }
      out.push({ text: trimmed, start, end: start + trimmed.length, segStart: r.start, segEnd: r.end });
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  // 3b. 项目符号、表格行、签名、逐段拆分
  // ---------------------------------------------------------------------------
  /**
   * 行首的项目符号或编号，后面必须跟空白，例如 "• " "· " "- " "1. " "(2) " "a) " "iv. "
   * 大写字母编号（"A."、"IV."）只在后面跟 Tab 或多个空格时才算，免得把 "J. Smith" 这类缩写当成编号
   * Word 第二、三级符号转成纯文字时是 "o"、"§"（Wingdings），这类也只在后面跟 Tab 或多个空格时才算
   */
  const LIST_MARKER_RE = /^(?:(?:[\u2022\u00B7\u25AA\u25E6\u25CF\u25CB\u25A0\u25A1\u25BA\u25B6\u27A2\u27A4\u2713\u2714\-\u2013\u2014*]|\(?\d{1,2}[.)]|\(?[a-z][.)]|\(?[ivx]{2,4}[.)]|\(?[A-Z]\)|\(?[IVX]{2,4}\))[\s\u00A0]+|(?:[A-Z]|[IVX]{2,4})\.(?:\t|[ \u00A0]{2,})[\s\u00A0]*|[o\u00A7\u00D8\u00FC](?:\t|[ \u00A0]{2,})[\s\u00A0]*)/;

  function isLineStart(s, pos) {
    for (let i = pos - 1; i >= 0; i--) {
      const ch = s[i];
      if (ch === '\n') return true;
      if (!/[ \t \r]/.test(ch)) return false;
    }
    return true;
  }

  /** 像表格/数据行（用 Tab 或一串空格对齐的列，例如 SKU 清单），不当成句子去改 */
  function isTableLike(text) {
    const t = String(text || '');
    return /\t/.test(t) || /\S[  ]{3,}\S/.test(t);
  }

  const SIGNOFF_RE = /^\s*(?:(?:best|kind|warm|warmest|with\s+best)\s+(?:regards|wishes)|regards|best|thanks|thank\s+you|thanks\s+again|many\s+thanks|cheers|sincerely|respectfully|yours\s+(?:truly|sincerely))\s*[,.!]?\s*$/i;

  /** 找签名的开始位置：最后 15 行里的结束语（Best Regards、Thanks 等）或 "--" 分隔线；找不到返回 -1 */
  function findSignatureStart(text) {
    const s = String(text || '');
    const starts = [0];
    for (let i = 0; i < s.length; i++) if (s[i] === '\n') starts.push(i + 1);
    const lines = starts.map((st, k) => s.slice(st, k + 1 < starts.length ? starts[k + 1] - 1 : s.length));
    let nonEmptySeen = 0;
    for (let k = lines.length - 1; k >= 0; k--) {
      const line = lines[k].replace(/\r$/, '');
      if (!line.trim()) continue;
      nonEmptySeen++;
      if (nonEmptySeen > 15) break;
      if (/^\s*--\s*$/.test(line)) return starts[k];
      if (SIGNOFF_RE.test(line)) {
        // "Thanks." 后面如果紧跟着正文句子（而不是名字、职位），它就不是结束语，不能把后面的内容当签名跳过
        const next = lines.slice(k + 1).map(l => l.replace(/\r$/, '')).find(l => l.trim());
        return next !== undefined && looksLikeSentence(next) ? -1 : starts[k];
      }
    }
    return -1;
  }

  /** 像一句正文（而不是名字、职位、电话这类签名行） */
  function looksLikeSentence(line) {
    const t = String(line || '').trim();
    const words = (t.match(/[A-Za-z][A-Za-z'’-]*/g) || []).length;
    return (words >= 4 && /[.!?]["'”’)]*$/.test(t)) || (words >= 6 && /[.!?]\s+[A-Z]/.test(t));
  }

  /**
   * 把邮件拆成"块"：每一行（标题、项目符号、段落）一块，去掉项目符号标记和首尾空白。
   * opts.excludeSignature：不包括签名部分。
   * 返回 [{ text, start, end }]
   */
  function splitBlocks(text, opts) {
    const s = String(text || '');
    let limit = s.length;
    if (opts && opts.excludeSignature) {
      const sig = findSignatureStart(s);
      if (sig >= 0) limit = sig;
    }
    const blocks = [];
    let pos = 0;
    while (pos < limit) {
      let nl = s.indexOf('\n', pos);
      if (nl === -1 || nl > limit) nl = limit;
      const line = s.slice(pos, nl);
      const lead = line.length - line.replace(/^\s+/, '').length;
      let content = line.trim();
      let start = pos + lead;
      const m = LIST_MARKER_RE.exec(content);
      if (m) { start += m[0].length; content = content.slice(m[0].length).trim(); }
      if (content) blocks.push({ text: content, start, end: start + content.length });
      pos = nl + 1;
    }
    return blocks;
  }

  /** 找到位置 pos（光标）所在或刚写完的那一句 */
  function sentenceAt(text, pos) {
    const sentences = segmentSentences(text);
    if (!sentences.length) return null;
    const p = Math.max(0, Math.min(pos, String(text).length));
    let best = null;
    for (const sen of sentences) {
      if (sen.segStart < p) best = sen; // 光标前一个字符所在的句子
      else break;
    }
    return best || sentences[0];
  }

  /** 光标所在段落（空行分隔），用作 AI 的上下文 */
  function paragraphAt(text, pos, maxLen) {
    const s = String(text);
    const limit = maxLen || 800;
    let start = s.lastIndexOf('\n\n', Math.max(0, pos - 1));
    start = start === -1 ? 0 : start + 2;
    let end = s.indexOf('\n\n', pos);
    end = end === -1 ? s.length : end;
    let para = s.slice(start, end).trim();
    if (para.length > limit) para = para.slice(0, limit) + '…';
    return para;
  }

  // ---------------------------------------------------------------------------
  // 4. 对比前后两次草稿 —— 推算正在编辑的位置
  // ---------------------------------------------------------------------------
  /**
   * 返回 { start, endOld, endNew }：变化区域。
   * 正在打字的光标通常就在 endNew（新插入内容的末尾）。
   */
  function diffRegion(prev, next) {
    const a = String(prev || '');
    const b = String(next || '');
    const minLen = Math.min(a.length, b.length);
    let p = 0;
    while (p < minLen && a.charCodeAt(p) === b.charCodeAt(p)) p++;
    let s = 0;
    while (s < minLen - p && a.charCodeAt(a.length - 1 - s) === b.charCodeAt(b.length - 1 - s)) s++;
    return { start: p, endOld: a.length - s, endNew: b.length - s };
  }

  /** 新文本中与变化区域重叠的句子 */
  function changedSentences(prev, next) {
    if (prev === next) return { sentences: [], caret: -1 };
    const r = diffRegion(prev, next);
    const all = segmentSentences(next);
    const hit = all.filter(sen => sen.segEnd >= r.start && sen.segStart <= r.endNew);
    return { sentences: hit, caret: r.endNew, region: r };
  }

  // ---------------------------------------------------------------------------
  // 5. 句子筛选
  // ---------------------------------------------------------------------------
  /** 是不是值得检查的英文句子（排除中文、太短的、纯数字的） */
  function isCheckableEnglish(sentence) {
    const s = String(sentence || '');
    if (isTableLike(s)) return false;
    const latin = (s.match(/[A-Za-z]/g) || []).length;
    const cjk = (s.match(/[㐀-鿿豈-﫿]/g) || []).length;
    const words = (s.match(/[A-Za-z][A-Za-z'’-]*/g) || []).length;
    if (latin < 8 || words < 3) return false;
    if (cjk > latin / 3) return false;
    return true;
  }

  /** 看起来是一句写完的话（以句末标点结尾） */
  function looksComplete(sentence) {
    return /[.!?。！？:;…)]["'”’)\]]*$/.test(String(sentence || '').trim());
  }

  // ---------------------------------------------------------------------------
  // 6. 在纯文本中查找一句话（容忍空白/引号差异）
  // ---------------------------------------------------------------------------
  /**
   * 找到 needle 在 text 里的位置。
   * occurrence：要第几处（从 0 开始）；不给时要求只出现一次，出现多次返回 multiple。
   */
  function findInText(text, needle, occurrence) {
    const { norm, map } = normalizeWithMap(text);
    const n = normalize(needle);
    if (!n) return { ok: false, reason: 'empty' };
    const idx = pickIndex(norm, n, occurrence);
    if (idx.reason) return { ok: false, reason: idx.reason };
    return { ok: true, start: map[idx.at], end: map[idx.at + n.length - 1] + 1 };
  }

  /** 规范化文字里 needle 出现的所有位置（允许重叠，最多 500 处） */
  function allIndexes(hay, needle) {
    const out = [];
    if (!needle) return out;
    for (let i = hay.indexOf(needle); i >= 0 && out.length < 500; i = hay.indexOf(needle, i + 1)) out.push(i);
    return out;
  }

  function pickIndex(norm, needle, occurrence) {
    if (Number.isInteger(occurrence) && occurrence >= 0) {
      const at = allIndexes(norm, needle)[occurrence];
      return at === undefined ? { reason: 'not_found' } : { at };
    }
    const at = norm.indexOf(needle);
    if (at < 0) return { reason: 'not_found' };
    if (norm.indexOf(needle, at + 1) >= 0) return { reason: 'multiple' };
    return { at };
  }

  // ---------------------------------------------------------------------------
  // 6b. 同样的文字出现不止一次时，定位到「我写的部分」里的那一处
  //     （回复时，下面引用的往来邮件里常有相同的标题或句子）
  // ---------------------------------------------------------------------------
  /** text 里 [start, end) 这段文字，是同样文字的第几处（从 0 开始） */
  function occurrenceAt(text, start, end) {
    const s = String(text || '');
    const n = normalize(s.slice(start, end));
    if (!n) return 0;
    return Math.max(0, allIndexes(normalize(s.slice(0, end)), n).length - 1);
  }

  const CTX = 60; // 记住前后各约 60 个字符

  /**
   * text 里 [start, end) 这段文字的"前后文"（规范化后前后各约 60 个字符）。
   * 同样的文字出现不止一处时，靠前后文认出原来是哪一处；你在别处改动邮件也不受影响。
   */
  function contextAt(text, start, end) {
    const s = String(text || '');
    const n = normalize(s.slice(start, end));
    const norm = normalize(s);
    const pos = allIndexes(norm, n)[occurrenceAt(s, start, end)];
    if (!n || pos === undefined) return null;
    return { before: norm.slice(Math.max(0, pos - CTX), pos), after: norm.slice(pos + n.length, pos + n.length + CTX) };
  }

  function contextScore(norm, pos, len, ctx) {
    const b = norm.slice(Math.max(0, pos - CTX), pos);
    const a = norm.slice(pos + len, pos + len + CTX);
    const cb = String(ctx.before || '');
    const ca = String(ctx.after || '');
    let i = 0;
    while (i < b.length && i < cb.length && b[b.length - 1 - i] === cb[cb.length - 1 - i]) i++;
    let j = 0;
    while (j < a.length && j < ca.length && a[j] === ca[j]) j++;
    return i + j;
  }

  /**
   * 要替换的文字在「我写的部分」(mine) 里怎么定位：
   *  - 只出现一次 → 用整封邮件里的第 1 处（引用的往来邮件在下面，里面相同的文字不会被改到）
   *  - 出现多次 → 用前后文 ctx 认出是哪一处；认不出来返回 multiple
   *  - 不在 mine 里 → pos 为 -1，由调用方决定（Outlook 里会拒绝改引用部分）
   * 返回 { occurrence?, pos, reason? }；pos 是在规范化 mine 里的位置（用来排序）
   */
  function locateInMine(mine, original, ctx) {
    const n = normalize(original);
    if (!n) return { reason: 'empty', pos: -1 };
    const norm = normalize(mine);
    const idx = allIndexes(norm, n);
    if (!idx.length) return { pos: -1 };
    if (idx.length === 1) return { occurrence: 0, pos: idx[0] };
    if (!ctx) return { reason: 'multiple', pos: idx[0] };
    const scores = idx.map(p => contextScore(norm, p, n.length, ctx));
    const best = Math.max.apply(null, scores);
    const k = scores.indexOf(best);
    if (best === 0 || scores.indexOf(best, k + 1) >= 0) return { reason: 'multiple', pos: idx[0] };
    return { occurrence: k, pos: idx[k] };
  }

  /**
   * 一次替换多处时的执行顺序：从邮件末尾往前替换，前面各处"是第几处"就不会受影响。
   * 两处要改到同一段文字（重叠）时，后面那处不改，交给你手动处理。
   * pairs: [{ original, replacement, ctx? }] → [{ i, occurrence?, pos, reason? }]
   */
  function planReplacements(mine, pairs) {
    const used = [];
    return pairs
      .map((p, i) => {
        const loc = Object.assign({ i }, locateInMine(mine, p.original, p.ctx));
        if (!loc.reason && loc.pos >= 0) {
          const end = loc.pos + normalize(p.original).length;
          if (used.some(r => loc.pos < r[1] && r[0] < end)) return { i, pos: loc.pos, reason: 'multiple' };
          used.push([loc.pos, end]);
        }
        return loc;
      })
      .sort((a, b) => b.pos - a.pos);
  }

  /**
   * 替换完成后，每处新文字在新的「我写的部分」里的前后文（以后撤销、再换写法时用来定位）。
   * steps: planReplacements 的结果；okIdx: 成功替换的 pairs 下标（Set）
   * 返回数组（按 pairs 下标）：{ before, after } 或 null（不在我写的部分里）
   */
  function contextsAfter(mine, pairs, steps, okIdx) {
    const done = steps.filter(st => okIdx.has(st.i) && st.pos >= 0 && !st.reason);
    let s = normalize(mine);
    done.forEach(st => { // 已经按位置从后往前排好，前面的位置不受影响
      const p = pairs[st.i];
      s = s.slice(0, st.pos) + normalize(p.replacement) + s.slice(st.pos + normalize(p.original).length);
    });
    const out = pairs.map(() => null);
    done.forEach(st => {
      let at = st.pos;
      done.forEach(o => { if (o.pos < st.pos) at += normalize(pairs[o.i].replacement).length - normalize(pairs[o.i].original).length; });
      const len = normalize(pairs[st.i].replacement).length;
      out[st.i] = { before: s.slice(Math.max(0, at - CTX), at), after: s.slice(at + len, at + len + CTX) };
    });
    return out;
  }

  // ---------------------------------------------------------------------------
  // 7. HTML 实体与精确替换
  // ---------------------------------------------------------------------------
  const NAMED_ENTITIES = {
    nbsp: ' ', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'",
    rsquo: '’', lsquo: '‘', rdquo: '”', ldquo: '“',
    ndash: '–', mdash: '—', hellip: '…', bull: '•', middot: '·',
    copy: '©', reg: '®', trade: '™', deg: '°', times: '×',
    eacute: 'é', egrave: 'è', aacute: 'á', agrave: 'à', ccedil: 'ç',
    uuml: 'ü', ouml: 'ö', auml: 'ä', ntilde: 'ñ', shy: '­', zwnj: '‌', zwj: '‍',
    ensp: ' ', emsp: ' ', thinsp: ' ', laquo: '«', raquo: '»', frac12: '½', plusmn: '±',
    euro: '\u20AC', pound: '\u00A3', yen: '\u00A5', cent: '\u00A2', sect: '\u00A7', para: '\u00B6', micro: '\u00B5',
    sup1: '\u00B9', sup2: '\u00B2', sup3: '\u00B3', frac14: '\u00BC', frac34: '\u00BE', ordm: '\u00BA', ordf: '\u00AA',
    divide: '\u00F7', minus: '\u2212', permil: '\u2030', larr: '\u2190', rarr: '\u2192', uarr: '\u2191', darr: '\u2193',
    Eacute: '\u00C9', Egrave: '\u00C8', Aacute: '\u00C1', Agrave: '\u00C0', Ccedil: '\u00C7', Ntilde: '\u00D1',
    Uuml: '\u00DC', Ouml: '\u00D6', Auml: '\u00C4', iacute: '\u00ED', oacute: '\u00F3', uacute: '\u00FA',
    acirc: '\u00E2', ecirc: '\u00EA', icirc: '\u00EE', ocirc: '\u00F4', ucirc: '\u00FB', szlig: '\u00DF'
  };

  /** 我们认得的实体名（大小写敏感；只有最常见的几个允许全大写写法） */
  function isKnownEntity(name) {
    return NAMED_ENTITIES[name] !== undefined || /^(AMP|LT|GT|QUOT|NBSP)$/.test(name);
  }

  function decodeEntities(str) {
    return String(str).replace(/&(#\d{1,7}|#[xX][0-9a-fA-F]{1,6}|[a-zA-Z][a-zA-Z0-9]{1,31});/g, (whole, body) => {
      if (body[0] === '#') {
        const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
        if (!Number.isFinite(code) || code < 0 || code > 0x10FFFF) return whole;
        try { return String.fromCodePoint(code); } catch (e) { return whole; }
      }
      if (!isKnownEntity(body)) return whole; // 认不出的实体原样保留
      return NAMED_ENTITIES[body] !== undefined ? NAMED_ENTITIES[body] : NAMED_ENTITIES[body.toLowerCase()];
    });
  }

  function encodeText(str) {
    return String(str)
      // 解码时认不出、原样保留下来的实体（如 &euro;）继续原样保留，不能变成 &amp;euro;
      .replace(/&([a-zA-Z][a-zA-Z0-9]{1,31};)?/g, (m, ent) => (ent && !isKnownEntity(ent.slice(0, -1)) ? m : '&amp;' + (ent || '')))
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/ /g, '&nbsp;');
  }

  const TOKEN_RE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<![^>]*>|<\?[\s\S]*?\?>|<\/?[A-Za-z][^>]*>/g;
  const SKIP_CONTENT_TAGS = new Set(['style', 'script', 'title', 'xml', 'template', 'noscript', 'head']);
  const BLOCK_TAGS = new Set([
    'p', 'div', 'br', 'li', 'ul', 'ol', 'table', 'tr', 'td', 'th', 'tbody', 'thead', 'tfoot',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'hr', 'section', 'article',
    'header', 'footer', 'body', 'html', 'center', 'dl', 'dt', 'dd', 'caption'
  ]);

  function tagName(raw) {
    const m = /^<\/?([A-Za-z][A-Za-z0-9:_-]*)/.exec(raw);
    return m ? m[1].toLowerCase() : '';
  }

  /** 把 HTML 拆成 tag / text / comment 片段，并标记哪些文字是可见正文 */
  function tokenizeHtml(html) {
    const s = String(html);
    const tokens = [];
    let last = 0;
    let m;
    TOKEN_RE.lastIndex = 0;
    while ((m = TOKEN_RE.exec(s)) !== null) {
      if (m.index > last) tokens.push({ type: 'text', raw: s.slice(last, m.index) });
      const raw = m[0];
      if (raw.startsWith('<!--') || raw.startsWith('<!') || raw.startsWith('<?')) tokens.push({ type: 'comment', raw });
      else tokens.push({ type: 'tag', raw, name: tagName(raw), closing: raw[1] === '/' });
      last = m.index + raw.length;
    }
    if (last < s.length) tokens.push({ type: 'text', raw: s.slice(last) });

    // 标记可见性：<head>/<style>/<script> 等里面的文字不算正文
    const skipStack = [];
    for (const t of tokens) {
      if (t.type === 'tag') {
        if (SKIP_CONTENT_TAGS.has(t.name)) {
          if (t.closing) {
            const at = skipStack.lastIndexOf(t.name);
            if (at >= 0) skipStack.length = at;
          } else if (!/\/>$/.test(t.raw)) {
            skipStack.push(t.name);
          }
        }
        t.block = BLOCK_TAGS.has(t.name);
      } else if (t.type === 'text') {
        t.visible = skipStack.length === 0;
      }
    }
    return tokens;
  }

  /**
   * 在 HTML 正文中把 original（一句话）替换成 replacement。
   * - 只改动命中的那段文字，其余 HTML（格式、图片、签名）原样保留
   * - 句子跨越多个格式标签（如部分加粗）也能处理：替换文字放进第一个片段
   * - 找不到、出现多次、或跨段落时返回 { ok:false, reason }，不做任何修改
   * - opts.occurrence：同样的文字出现多次时，改第几处（从 0 开始）
   */
  function replaceInHtml(html, original, replacement, opts) {
    const needle = normalize(original);
    if (!needle) return { ok: false, reason: 'empty' };
    const tokens = tokenizeHtml(html);

    let full = '';
    const segs = [];
    let pendingBreak = false;
    tokens.forEach((t, i) => {
      if (t.type === 'tag') { if (t.block) pendingBreak = true; return; }
      if (t.type !== 'text' || !t.visible) return;
      const dec = decodeEntities(t.raw);
      if (pendingBreak && full.length) full += '\n'; // 段落分隔（虚拟字符，不属于任何片段）
      pendingBreak = false;
      segs.push({ i, start: full.length, end: full.length + dec.length, dec });
      full += dec;
    });

    const { norm, map } = normalizeWithMap(full);
    const picked = pickIndex(norm, needle, opts && opts.occurrence);
    if (picked.reason) return { ok: false, reason: picked.reason };
    const idx = picked.at;

    // 整句的位置：用于确认它没有跨段落
    const sentFrom = map[idx];
    const sentTo = map[idx + needle.length - 1] + 1;
    const sentSegs = segs.filter(sg => sg.end > sentFrom && sg.start < sentTo);
    if (!sentSegs.length) return { ok: false, reason: 'not_found' };
    for (let k = 1; k < sentSegs.length; k++) {
      if (sentSegs[k].start !== sentSegs[k - 1].end) return { ok: false, reason: 'crosses_paragraph' };
    }

    // 按单词比较，只改真正不同的那几处；没改的词（以及它们的加粗、链接、上标）原样保留
    const repl = normalize(String(replacement).replace(/\s*\r?\n\s*/g, ' '));
    const hunks = diffHunks(needle, repl);
    if (!hunks.length) return { ok: true, html: String(html), unchanged: true };
    const work = new Map(); // 片段 → 改动后的文字
    for (let h = hunks.length - 1; h >= 0; h--) { // 从后往前改，前面的位置不受影响
      const hk = hunks[h];
      let from;
      let to;
      if (hk.to > hk.from) {
        from = map[idx + hk.from];
        to = hk.to < needle.length ? map[idx + hk.to] : map[idx + hk.to - 1] + 1;
      } else {
        from = to = hk.from > 0 ? map[idx + hk.from - 1] + 1 : map[idx];
      }
      if (!editSegments(segs, work, from, to, hk.text, hk.from === 0)) return { ok: false, reason: 'not_found' };
    }
    work.forEach((text, sg) => { tokens[sg.i].newRaw = encodeText(text); });
    return { ok: true, html: tokens.map(t => (t.newRaw !== undefined ? t.newRaw : t.raw)).join('') };
  }

  /** 逐词对比 a → b，合并成要修改的几处：[{ from, to, text }]（from/to 是在 a 里的位置） */
  function diffHunks(a, b) {
    const hunks = [];
    let pos = 0;
    let cur = null;
    for (const op of wordDiff(a, b)) {
      if (op.op === 'eq') {
        if (cur) { hunks.push(cur); cur = null; }
        pos += op.text.length;
        continue;
      }
      if (!cur) cur = { from: pos, to: pos, text: '' };
      if (op.op === 'del') { pos += op.text.length; cur.to = pos; } else cur.text += op.text;
    }
    if (cur) hunks.push(cur);
    return hunks;
  }

  /**
   * 在可见文字的 [from, to) 处删除并插入 text，只动涉及到的文字片段。
   * work 记录每个片段改动后的文字（同一片段可能被改多处）。
   */
  function editSegments(segs, work, from, to, text, atSentenceStart) {
    const cur = sg => (work.has(sg) ? work.get(sg) : sg.dec);
    if (to > from) {
      const hit = segs.filter(sg => sg.end > from && sg.start < to);
      if (!hit.length) return false;
      // 新文字放进第一个真正删掉了文字的片段（只删掉空格的片段不算），沿用那里的格式
      let anchor = hit.findIndex(sg => /\S/.test(sg.dec.slice(Math.max(0, from - sg.start), Math.min(sg.dec.length, to - sg.start))));
      if (anchor < 0) anchor = 0;
      for (let k = hit.length - 1; k >= 0; k--) {
        const sg = hit[k];
        const s = cur(sg);
        const lf = Math.max(0, from - sg.start);
        const lt = Math.min(sg.end - sg.start, to - sg.start);
        work.set(sg, s.slice(0, lf) + (k === anchor ? text : '') + s.slice(lt));
      }
      return true;
    }
    const before = segs.find(sg => sg.start < from && sg.end >= from); // 插入点前一个字所在的片段
    const after = segs.find(sg => sg.start <= from && sg.end > from);  // 插入点后一个字所在的片段
    let host = before || after;
    if (before && after && before !== after) {
      // 正好在两种格式的交界处：句首，或者新加的词以空格开头 → 放进后面的片段（不会并进前面的链接、加粗词）
      host = atSentenceStart || /^\s/.test(text) ? after : before;
    }
    if (!host) return false;
    const s = cur(host);
    const lp = from - host.start;
    work.set(host, s.slice(0, lp) + text + s.slice(lp));
    return true;
  }

  /** 从 HTML 中提取可见文字（测试和兜底用） */
  function htmlToText(html) {
    const tokens = tokenizeHtml(html);
    let out = '';
    for (const t of tokens) {
      if (t.type === 'tag' && t.block) out += '\n';
      else if (t.type === 'text' && t.visible) out += decodeEntities(t.raw).replace(/[ \t\r\n]+/g, ' ');
    }
    return out.replace(/[ \t]*\n[ \t]*/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  // ---------------------------------------------------------------------------
  // 8. 逐词对比（显示删改用）
  // ---------------------------------------------------------------------------
  function tokenizeWords(s) {
    return String(s).match(/[A-Za-z0-9À-ɏ]+(?:['’\-][A-Za-z0-9À-ɏ]+)*|\s+|[^\sA-Za-z0-9À-ɏ]/g) || [];
  }

  /** 返回 [{ op: 'eq'|'del'|'ins', text }] */
  function wordDiff(a, b) {
    const x = tokenizeWords(a);
    const y = tokenizeWords(b);
    const n = x.length;
    const m = y.length;
    if (n * m > 250000) return [{ op: 'del', text: String(a) }, { op: 'ins', text: String(b) }];
    const dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        dp[i][j] = x[i] === y[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
    const ops = [];
    const push = (op, text) => {
      const last = ops[ops.length - 1];
      if (last && last.op === op) last.text += text; else ops.push({ op, text });
    };
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (x[i] === y[j]) { push('eq', x[i]); i++; j++; }
      else if (dp[i + 1][j] >= dp[i][j + 1]) { push('del', x[i]); i++; }
      else { push('ins', y[j]); j++; }
    }
    while (i < n) push('del', x[i++]);
    while (j < m) push('ins', y[j++]);
    return ops;
  }

  // ---------------------------------------------------------------------------
  // 9. 安全检查：AI 是否改动了数字、日期或专业术语
  // ---------------------------------------------------------------------------
  function numberTokens(s) {
    return (String(s).match(/\d+(?:[.,:/-]\d+)*/g) || []).map(t => t.replace(/,/g, '')).sort();
  }

  /**
   * 返回中文警告列表，例如：["数字或日期有变化，请核对", "术语“ETA”被改动"]
   */
  function invariantWarnings(original, revised, terms) {
    const warnings = [];
    if (!revised) return warnings;
    const a = numberTokens(original).join('|');
    const b = numberTokens(revised).join('|');
    if (a !== b) warnings.push('数字或日期有变化，请核对');
    const lowerOrig = String(original).toLowerCase();
    const lowerRev = String(revised).toLowerCase();
    for (const term of terms || []) {
      const t = String(term).trim().toLowerCase();
      if (t.length < 2) continue;
      if (lowerOrig.includes(t) && !lowerRev.includes(t)) warnings.push('术语“' + term.trim() + '”被改动，请核对');
    }
    return warnings;
  }

  function parseTerms(text) {
    return String(text || '')
      .split(/[,，;；\n]+/)
      .map(s => s.trim())
      .filter(Boolean)
      .slice(0, 200);
  }

  const api = {
    normalize, normalizeWithMap, sentenceKey,
    splitEmail, segmentSentences, sentenceAt, paragraphAt,
    isTableLike, findSignatureStart, splitBlocks, LIST_MARKER_RE,
    diffRegion, changedSentences,
    isCheckableEnglish, looksComplete,
    findInText, occurrenceAt, contextAt, locateInMine, planReplacements, contextsAfter,
    decodeEntities, encodeText, tokenizeHtml, replaceInHtml, htmlToText,
    wordDiff, invariantWarnings, numberTokens, parseTerms
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.TextUtil = api;
})(typeof window !== 'undefined' ? window : globalThis);
