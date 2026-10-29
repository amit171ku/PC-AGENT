(function () {
  "use strict";

  // ---------- helpers ----------
  function mk(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  // ---------- inline formatting ----------
  const INLINE_RE = new RegExp(
    "(" + [
      "`[^`\\n]+`",                                   // `code`
      "\\*\\*[^\\n]+?\\*\\*",                         // **bold**
      "~~[^~\\n]+?~~",                                // ~~strike~~
      "\\[[^\\]\\n]+\\]\\([^)\\s]+\\)",               // [text](url)
      "<br\\s*/?>",                                   // <br> <br/>
      "https?://[^\\s<>)\\]]+",                       // bare URL
      "\\*[^*\\s](?:[^*\\n]*[^*\\s])?\\*"             // *italic*
    ].join("|") + ")", "gi");

  const SAFE_URL = /^(https?:\/\/|mailto:)/i;

  function addText(parent, text) {
    // a newline inside text becomes a real line break
    text.split("\n").forEach((part, i) => {
      if (i > 0) parent.appendChild(document.createElement("br"));
      if (part) parent.appendChild(document.createTextNode(part));
    });
  }

  function addLink(parent, href, label) {
    const a = mk("a", "md-link");
    a.href = href;
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    a.textContent = label;
    parent.appendChild(a);
  }

  function inline(parent, text) {
    String(text).split(INLINE_RE).forEach((tok, i) => {
      if (!tok) return;
      if (i % 2 === 0) { addText(parent, tok); return; }   // plain text between tokens

      if (tok[0] === "`") {
        parent.appendChild(mk("code", "inline", tok.slice(1, -1)));
      } else if (/^<br/i.test(tok)) {
        parent.appendChild(document.createElement("br"));
      } else if (tok.startsWith("**")) {
        const b = document.createElement("strong");
        inline(b, tok.slice(2, -2));
        parent.appendChild(b);
      } else if (tok.startsWith("~~")) {
        const d = document.createElement("del");
        inline(d, tok.slice(2, -2));
        parent.appendChild(d);
      } else if (tok[0] === "[") {
        const m = /^\[([^\]]+)\]\(([^)\s]+)\)$/.exec(tok);
        if (m && SAFE_URL.test(m[2])) addLink(parent, m[2], m[1]);
        else addText(parent, tok);
      } else if (/^https?:/i.test(tok)) {
        const m = /^(.*?)([.,;:!?'"]*)$/.exec(tok);      // keep trailing punctuation outside the link
        addLink(parent, m[1], m[1]);
        if (m[2]) parent.appendChild(document.createTextNode(m[2]));
      } else if (tok[0] === "*") {
        const em = document.createElement("em");
        inline(em, tok.slice(1, -1));
        parent.appendChild(em);
      } else {
        addText(parent, tok);
      }
    });
  }

  // ---------- block patterns ----------
  const FENCE   = /^\s*```\s*([\w+#.-]*)\s*$/;
  const HEADING = /^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/;
  const HR      = /^\s{0,3}([-*_])(?:\s*\1){2,}\s*$/;
  const LIST    = /^(\s*)([-*+]|\d{1,3}[.)])\s+(.*)$/;
  const QUOTE   = /^\s{0,3}>\s?(.*)$/;
  const SEP     = /^\s*\|?\s*:?-{2,}:?\s*(?:\|\s*:?-{2,}:?\s*)*\|?\s*$/;

  function isTableStart(lines, i) {
    return i + 1 < lines.length &&
      lines[i].includes("|") &&
      lines[i + 1].includes("|") &&
      SEP.test(lines[i + 1]);
  }

  function startsBlock(lines, i) {
    const l = lines[i];
    return FENCE.test(l) || HEADING.test(l) || HR.test(l) ||
           QUOTE.test(l) || LIST.test(l) || isTableStart(lines, i);
  }

  function splitRow(line) {
    let s = line.trim();
    if (s.startsWith("|")) s = s.slice(1);
    if (s.endsWith("|") && !s.endsWith("\\|")) s = s.slice(0, -1);
    const cells = [];
    let cur = "";
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === "\\" && s[i + 1] === "|") { cur += "|"; i++; }
      else if (ch === "|") { cells.push(cur.trim()); cur = ""; }
      else cur += ch;
    }
    cells.push(cur.trim());
    return cells;
  }

  // ---------- block renderers ----------
  function renderCode(parent, lang, codeText) {
    const box = mk("div", "md-code");
    const bar = mk("div", "md-code-bar");
    bar.appendChild(mk("span", "md-code-lang", lang || "code"));
    const copy = mk("button", "md-code-copy", "Copy");
    copy.type = "button";
    copy.onclick = () => {
      if (!navigator.clipboard) return;
      navigator.clipboard.writeText(codeText).then(() => {
        copy.textContent = "Copied";
        setTimeout(() => (copy.textContent = "Copy"), 1500);
      });
    };
    bar.appendChild(copy);
    const pre = document.createElement("pre");
    pre.appendChild(mk("code", "", codeText));
    box.append(bar, pre);
    parent.appendChild(box);
  }

  function renderTable(parent, header, aligns, rows) {
    const wrap = mk("div", "md-table-wrap");
    const table = mk("table", "md-table");
    const n = header.length;

    const thead = document.createElement("thead");
    const hr = document.createElement("tr");
    header.forEach((h, c) => {
      const th = document.createElement("th");
      if (aligns[c]) th.style.textAlign = aligns[c];
      inline(th, h);
      hr.appendChild(th);
    });
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.forEach((r) => {
      const tr = document.createElement("tr");
      for (let c = 0; c < n; c++) {
        const td = document.createElement("td");
        if (aligns[c]) td.style.textAlign = aligns[c];
        inline(td, r[c] || "");
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
    parent.appendChild(wrap);
  }

  function renderList(parent, items) {
    const stack = [];   // { indent, list, lastLi }
    items.forEach((it) => {
      while (stack.length && it.indent < stack[stack.length - 1].indent) stack.pop();
      let top = stack[stack.length - 1];
      let forceNew = false;
      // switching between bullets and numbers at the same level starts a new list
      if (top && it.indent === top.indent && top.ordered !== it.ordered) {
        stack.pop();
        top = stack[stack.length - 1];
        forceNew = true;
      }
      if (!top || it.indent > top.indent || forceNew) {
        const list = document.createElement(it.ordered ? "ol" : "ul");
        if (it.ordered && it.num > 1 && !top) list.start = it.num;
        if (top && top.lastLi) top.lastLi.appendChild(list);
        else parent.appendChild(list);
        top = { indent: it.indent, ordered: it.ordered, list: list, lastLi: null };
        stack.push(top);
      }
      const li = document.createElement("li");
      inline(li, it.text);
      top.list.appendChild(li);
      top.lastLi = li;
    });
  }

  function renderBlocks(parent, lines) {
    let i = 0;
    while (i < lines.length) {
      const line = lines[i];
      if (!line.trim()) { i++; continue; }
      let m;

      // fenced code block
      if ((m = FENCE.exec(line))) {
        const lang = m[1];
        const buf = [];
        i++;
        while (i < lines.length && !/^\s*```\s*$/.test(lines[i])) buf.push(lines[i++]);
        i++; // closing fence
        renderCode(parent, lang, buf.join("\n"));
        continue;
      }

      // heading
      if ((m = HEADING.exec(line))) {
        const h = document.createElement("h" + m[1].length);
        inline(h, m[2]);
        parent.appendChild(h);
        i++;
        continue;
      }

      // horizontal rule
      if (HR.test(line)) {
        parent.appendChild(document.createElement("hr"));
        i++;
        continue;
      }

      // blockquote
      if (QUOTE.test(line)) {
        const inner = [];
        while (i < lines.length && QUOTE.test(lines[i])) inner.push(QUOTE.exec(lines[i++])[1]);
        const bq = document.createElement("blockquote");
        renderBlocks(bq, inner);
        parent.appendChild(bq);
        continue;
      }

      // table
      if (isTableStart(lines, i)) {
        const header = splitRow(lines[i]);
        const aligns = splitRow(lines[i + 1]).map((c) =>
          /^:-+:$/.test(c) ? "center" : /-:$/.test(c) ? "right" : /^:-/.test(c) ? "left" : "");
        i += 2;
        const rows = [];
        while (i < lines.length && lines[i].trim() && lines[i].includes("|")) {
          rows.push(splitRow(lines[i++]));
        }
        renderTable(parent, header, aligns, rows);
        continue;
      }

      // list (bullets, numbers, nested by indentation)
      if (LIST.test(line)) {
        const items = [];
        while (i < lines.length) {
          const l = lines[i];
          const lm = LIST.exec(l);
          if (lm && !HR.test(l)) {
            items.push({
              indent: lm[1].replace(/\t/g, "    ").length,
              ordered: /\d/.test(lm[2]),
              num: parseInt(lm[2], 10) || 1,
              text: lm[3]
            });
            i++;
          } else if (!l.trim()) {
            // a blank line keeps the list going only if another item follows
            let j = i + 1;
            while (j < lines.length && !lines[j].trim()) j++;
            if (j < lines.length && LIST.test(lines[j]) && !HR.test(lines[j])) i = j;
            else break;
          } else if (/^\s+\S/.test(l) && items.length && !FENCE.test(l)) {
            items[items.length - 1].text += "\n" + l.trim();   // wrapped item text
            i++;
          } else {
            break;
          }
        }
        renderList(parent, items);
        continue;
      }

      // paragraph (consecutive plain lines; single line breaks are kept)
      const buf = [line];
      i++;
      while (i < lines.length && lines[i].trim() && !startsBlock(lines, i)) buf.push(lines[i++]);
      const p = document.createElement("p");
      inline(p, buf.join("\n"));
      parent.appendChild(p);
    }
  }

  window.renderMarkdown = function (el, text) {
    const lines = String(text == null ? "" : text).replace(/\r\n?/g, "\n").split("\n");
    renderBlocks(el, lines);
  };
})();