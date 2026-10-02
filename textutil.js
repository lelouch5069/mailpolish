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
      const trimmed = piece.trim();
      if (!trimmed) continue;
      const start = r.start + lead;
      out.push({ text: trimmed, start, end: start + trimmed.length, segStart: r.start, segEnd: r.end });
    }
    return out;
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
  function findInText(text, needle) {
    const { norm, map } = normalizeWithMap(text);
    const n = normalize(needle);
    if (!n) return { ok: false, reason: 'empty' };
    const idx = norm.indexOf(n);
    if (idx < 0) return { ok: false, reason: 'not_found' };
    if (norm.indexOf(n, idx + 1) >= 0) return { ok: false, reason: 'multiple' };
    return { ok: true, start: map[idx], end: map[idx + n.length - 1] + 1 };
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
    ensp: ' ', emsp: ' ', thinsp: ' ', laquo: '«', raquo: '»', frac12: '½', plusmn: '±'
  };

  function decodeEntities(str) {
    return String(str).replace(/&(#\d{1,7}|#[xX][0-9a-fA-F]{1,6}|[a-zA-Z][a-zA-Z0-9]{1,31});/g, (whole, body) => {
      if (body[0] === '#') {
        const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
        if (!Number.isFinite(code) || code < 0 || code > 0x10FFFF) return whole;
        try { return String.fromCodePoint(code); } catch (e) { return whole; }
      }
      const v = NAMED_ENTITIES[body] !== undefined ? NAMED_ENTITIES[body] : NAMED_ENTITIES[body.toLowerCase()];
      return v !== undefined ? v : whole;
    });
  }

  function encodeText(str) {
    return String(str)
      .replace(/&/g, '&amp;')
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
   */
  function replaceInHtml(html, original, replacement) {
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
    const idx = norm.indexOf(needle);
    if (idx < 0) return { ok: false, reason: 'not_found' };
    if (norm.indexOf(needle, idx + 1) >= 0) return { ok: false, reason: 'multiple' };

    // 整句的位置：用于确认它没有跨段落
    const sentFrom = map[idx];
    const sentTo = map[idx + needle.length - 1] + 1;
    const sentSegs = segs.filter(sg => sg.end > sentFrom && sg.start < sentTo);
    if (!sentSegs.length) return { ok: false, reason: 'not_found' };
    for (let k = 1; k < sentSegs.length; k++) {
      if (sentSegs[k].start !== sentSegs[k - 1].end) return { ok: false, reason: 'crosses_paragraph' };
    }

    // 只改真正不同的那一小段（保留其余文字的加粗/颜色等格式）
    const repl = normalize(String(replacement).replace(/\s*\r?\n\s*/g, ' '));
    let pre = 0;
    const maxPre = Math.min(needle.length, repl.length);
    while (pre < maxPre && needle[pre] === repl[pre]) pre++;
    let suf = 0;
    while (suf < Math.min(needle.length, repl.length) - pre &&
      needle[needle.length - 1 - suf] === repl[repl.length - 1 - suf]) suf++;
    const insertText = repl.slice(pre, repl.length - suf);
    const delStartN = idx + pre;                    // 规范化文本中的删除起点
    const delEndN = idx + needle.length - suf;      // 规范化文本中的删除终点（不含）
    if (delStartN === delEndN && !insertText) return { ok: true, html: String(html), unchanged: true };

    let from;
    let to;
    if (delEndN > delStartN) {
      from = map[delStartN];
      to = map[delEndN - 1] + 1;
    } else if (pre > 0) {
      // 纯插入：插在前一个字符之后，继承前文格式
      from = to = map[delStartN - 1] + 1;
    } else {
      // 插在句首
      from = to = map[delStartN];
    }

    let hit = segs.filter(sg => sg.end > from && sg.start < to);
    if (from === to) {
      const before = segs.find(sg => sg.start < from && sg.end >= from); // 含插入点前一个字符的片段
      const after = segs.find(sg => sg.start <= from && sg.end > from);  // 含插入点后一个字符的片段
      const host = pre > 0 ? (before || after) : (after || before);
      hit = host ? [host] : [];
    }
    if (!hit.length) return { ok: false, reason: 'not_found' };

    hit.forEach((sg, k) => {
      const localFrom = Math.max(0, from - sg.start);
      const localTo = Math.min(sg.dec.length, to - sg.start);
      const before = sg.dec.slice(0, localFrom);
      const after = sg.dec.slice(localTo);
      const middle = k === 0 ? insertText : '';
      tokens[sg.i].newRaw = encodeText(before + middle + after);
    });
    return { ok: true, html: tokens.map(t => (t.newRaw !== undefined ? t.newRaw : t.raw)).join('') };
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
    diffRegion, changedSentences,
    isCheckableEnglish, looksComplete,
    findInText, decodeEntities, encodeText, tokenizeHtml, replaceInHtml, htmlToText,
    wordDiff, invariantWarnings, numberTokens, parseTerms
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.TextUtil = api;
})(typeof window !== 'undefined' ? window : globalThis);
