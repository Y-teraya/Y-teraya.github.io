// ========= グローバル状態管理 =========
let rawEntriesStack = [];
const inputEl = document.getElementById("input");
const logOutputEl = document.getElementById("logOutput");
const styleSelectEl = document.getElementById("styleSelect");
const unifiedOutputEl = document.getElementById("unifiedOutput");

// ========= ユーティリティ =========
function sTrim(s) { return (s || "").replace(/^\s+|\s+$/g, ""); }

function convertAccentsToLatex(str) {
    if (!str) return "";

    const map = {
        "á": "\\'{a}", "Á": "\\'{A}", "é": "\\'{e}", "É": "\\'{E}", "í": "\\'{i}", "Í": "\\'{I}",
        "ó": "\\'{o}", "Ó": "\\'{O}", "ú": "\\'{u}", "Ú": "\\'{U}", "à": "\\`{a}", "À": "\\`{A}",
        "è": "\\`{e}", "È": "\\`{E}", "ì": "\\`{i}", "Ì": "\\`{I}", "ò": "\\`{o}", "Ò": "\\`{O}",
        "ù": "\\`{u}", "Ù": "\\`{U}", "ä": "\\\"{a}", "Ä": "\\\"{A}", "ë": "\\\"{e}", "Ë": "\\\"{E}",
        "ï": "\\\"{i}", "Ï": "\\\"{I}", "ö": "\\\"{o}", "Ö": "\\\"{O}", "ü": "\\\"{u}", "Ü": "\\\"{U}",
        "â": "\\^{a}", "Â": "\\^{A}", "ê": "\\^{e}", "Ê": "\\^{E}", "î": "\\^{i}", "Î": "\\^{I}",
        "ô": "\\^{o}", "Ô": "\\^{O}", "û": "\\^{u}", "Û": "\\^{U}", "ñ": "\\~{n}", "Ñ": "\\~{N}",
        "ç": "\\c{c}", "Ç": "\\c{C}", "ø": "\\o{}", "Ø": "\\O{}", "å": "\\aa{}", "Å": "\\AA{}"
    };

    let out = str;

    for (const [char, latex] of Object.entries(map)) {
        out = out.split(char).join(latex);   // ← 正規表現を使わないリテラル置換
    }

    return out;
}

function decodeLatexAccents(str) {
    if (!str) return "";
    let out = str.replace(/~/g, " ");

    const map = {
        "\\'{a}": "á", "\\'{A}": "Á", "\\'{e}": "é", "\\'{E}": "É", "\\'{i}": "í", "\\'{I}": "Í",
        "\\'{o}": "ó", "\\'{O}": "Ó", "\\'{u}": "ú", "\\'{U}": "Ú", "\\`{a}": "à", "\\`{A}": "À",
        "\\`{e}": "è", "\\`{E}": "È", "\\`{i}": "ì", "\\`{I}": "Ì", "\\`{o}": "ò", "\\`{O}": "Ò",
        "\\`{u}": "ù", "\\`{U}": "Ù", "\\\"{a}": "ä", "\\\"{A}": "Ä", "\\\"{e}": "ë", "\\\"{E}": "Ë",
        "\\\"{i}": "ï", "\\\"{I}": "Ï", "\\\"{o}": "ö", "\\\"{O}": "Ö", "\\\"{u}": "ü", "\\\"{U}": "Ü",
        "\\^{a}": "â", "\\^{A}": "Â", "\\^{e}": "ê", "\\^{E}": "Ê", "\\^{i}": "î", "\\^{I}": "Î",
        "\\^{o}": "ô", "\\^{O}": "Ô", "\\^{u}": "û", "\\^{U}": "Û", "\\~{n}": "ñ", "\\~{N}": "Ñ",
        "\\c{c}": "ç", "\\c{C}": "Ç", "\\o{}": "ø", "\\O{}": "Ø", "\\aa{}": "å", "\\AA{}": "Å"
    };

    // 長いキーから順に置換（部分一致を防ぐ）
    const keys = Object.keys(map).sort((a, b) => b.length - a.length);

    for (const k of keys) {
        out = out.split(k).join(map[k]);   // ← 正規表現を使わない
    }

    return out;
}

function monthToNumber(mstr) {
    if (!mstr) return "";
    const map = { jan:1, january:1, feb:2, february:2, mar:3, march:3, apr:4, april:4, may:5, jun:6, june:6, jul:7, july:7, aug:8, august:8, sep:9, september:9, oct:10, october:10, nov:11, november:11, dec:12, december:12 };
    return map[mstr.toLowerCase().substring(0,3)] ? String(map[mstr.toLowerCase().substring(0,3)]) : "";
}

function parseDateToYearMonth(datestr) {
    if (!datestr) return { year: "", month: "" };
    const s = datestr.trim();
    const ymddash = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (ymddash) return { year: ymddash[1], month: String(parseInt(ymddash[2], 10)) };
    const yMon = s.match(/^(\d{4})\s+([A-Za-z]+)/);
    if (yMon) return { year: yMon[1], month: monthToNumber(yMon[2]) };
    const y = s.match(/^(\d{4})/);
    return y ? { year: y[1], month: "" } : { year: "", month: "" };
}

function normalizePages(pages) {
    return pages ? pages.replace(/[–—−―‐–]+/g, "-").replace(/--+/g, "-") : "";
}

function parseAuthorString(authorStr) {
    if (!authorStr) return [];
    return authorStr.replace(/\r\n/g, "\n").split(/\n| and /).map(sTrim).filter(a => a.length > 0).map(p => {
        const m = p.split(",");
        if (m.length >= 2) return { last: sTrim(m[0]), first: sTrim(m.slice(1).join(",")) };
        const toks = p.trim().split(/\s+/);
        return toks.length === 1 ? { last: toks[0], first: "" } : { last: toks.pop(), first: toks.join(" ") };
    });
}

function formatAuthorsForBibtex(authors) {
    return (authors || []).map(a => {
        const last = convertAccentsToLatex(a.last || "");
        const first = convertAccentsToLatex(a.first || "");
        return first ? `${last}, ${first}` : last;
    }).join("\n      and ");
}

function formatAuthorsForCitation(authors) {
    if (!authors || authors.length === 0) return "";
    if (authors.length === 1) return authors[0].last;
    if (authors.length === 2) return `${authors[0].last} & ${authors[1].last}`;
    return authors[0].last;
}

// ========= フォーマット判定 & パース =========
function detectFormat(text) {
    const s = text.trim();
    if (!s) return "unknown";
    if (/^PMID-/.test(s)) return "nbib";
    if (/^TY\s+-/.test(s)) return "ris";
    if (/^@/.test(s)) return "bibtex";
    if (/\\bibitem/i.test(s)) return "bbl";
    return "txt";
}

// ========= nbib パーサ =========
function parseNbib(text) {
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  const entries = [];
  let current = null;

  function pushCurrent() {
    if (current) entries.push(current);
    current = {
      type: "article",
      raw: {},
      authors: [],
      title: "",
      journal: "",
      volume: "",
      number: "",
      pages: "",
      year: "",
      month: "",
      issn: "",
      doi: "",
      url: "",
      eprint: "",
      publisher: "",
      address: "",
      note: "",
      pmid: "",
      pmcid: ""
    };
  }

  for (let line of lines) {
    if (!line.trim()) {
      continue;
    }
    const m = line.match(/^([A-Z0-9]{2,4})\s*-\s*(.*)$/);
    if (!m) {
      if (current && current._lastTag) {
        current.raw[current._lastTag] =
          (current.raw[current._lastTag] || "") + " " + line.trim();
      }
      continue;
    }
    const tag = m[1];
    const val = m[2].trim();

    if (tag === "PMID") {
      if (current) pushCurrent();
      pushCurrent();
      current.pmid = val;
    }
    current._lastTag = tag;
    current.raw[tag] = (current.raw[tag] || "");
    if (current.raw[tag]) current.raw[tag] += " ";
    current.raw[tag] += val;
  }
  if (current) entries.push(current);

  for (const e of entries) {
    const r = e.raw;

    const auList = [];
    if (r["FAU"]) {
      const lines = text.match(/^FAU\s*-\s*(.*)$/gm) || [];
      for (const l of lines) {
        const v = l.split("-")[1].trim();
        auList.push(v);
      }
    }
    
    e.authors = auList.map(a => {
      const parts = a.split(",");
      if (parts.length >= 2) {
        return { last: sTrim(parts[0]), first: sTrim(parts.slice(1).join(",")) };
      } else {
        const toks = a.trim().split(/\s+/);
        return {
          last: toks[toks.length - 1],
          first: toks.slice(0, -1).join(" ")
        };
      }
    });

    e.title = r["TI"] || "";
    e.journal = r["JT"] || r["TA"] || "";
    e.volume = r["VI"] || "";
    e.number = r["IP"] || "";
    e.pages = normalizePages(r["PG"] || "");

    if (r["LID"]) {
      const mdoi = r["LID"].match(/10\.\S+/);
      if (mdoi) e.doi = mdoi[0].replace(/[.;]$/, "");
    }

    e.pmid = r["PMID"] || e.pmid;
    if (r["PMC"]) e.pmcid = r["PMC"];
    e.url = "";

    const dp = r["DP"] || "";
    const dObj = parseDateToYearMonth(dp);
    e.year = dObj.year;
    e.month = dObj.month;

    const notes = [];
    if (e.pmid) notes.push(`PMID: ${e.pmid}`);
    if (e.pmcid) notes.push(`PMCID: ${e.pmcid}`);
    e.note = notes.join("; ");
  }

  return entries;
}

// ========= RIS パーサ =========
function parseRis(text) {
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  const entries = [];
  let current = null;

  function pushCurrent() {
    if (current) entries.push(current);
    current = {
      type: "article",
      raw: {},
      authors: [],
      title: "",
      journal: "",
      volume: "",
      number: "",
      pages: "",
      year: "",
      month: "",
      issn: "",
      doi: "",
      url: "",
      eprint: "",
      publisher: "",
      address: "",
      note: ""
    };
  }

  for (let line of lines) {
    const m = line.match(/^([A-Z0-9]{2})\s*-\s*(.*)$/);
    if (!m) continue;
    const tag = m[1];
    const val = m[2].trim();

    if (tag === "TY") {
      pushCurrent();
      current.raw["TY"] = val;
      current._lastTag = tag;
      continue;
    }
    if (!current) continue;

    if (tag === "ER") {
      continue;
    }
    current.raw[tag] = current.raw[tag] || [];
    current.raw[tag].push(val);
    current._lastTag = tag;
  }
  if (current) entries.push(current);

  for (const e of entries) {
    const r = e.raw;

    const ty = (r["TY"] || [""])[0].toUpperCase();
    if (ty === "JOUR") e.type = "article";
    else if (ty === "BOOK") e.type = "book";
    else e.type = "article";

    const au = r["AU"] || [];
    e.authors = au.map(a => {
      const parts = a.split(",");
      if (parts.length >= 2) {
        return { last: sTrim(parts[0]), first: sTrim(parts.slice(1).join(",")) };
      } else {
        const toks = a.trim().split(/\s+/);
        return {
          last: toks[toks.length - 1],
          first: toks.slice(0, -1).join(" ")
        };
      }
    });

    e.title = (r["TI"] && r["TI"][0]) || (r["T1"] && r["T1"][0]) || "";
    e.journal = (r["T2"] && r["T2"][0]) || (r["JO"] && r["JO"][0]) || "";
    e.volume = (r["VL"] && r["VL"][0]) || "";
    e.number = (r["IS"] && r["IS"][0]) || "";

    const sp = (r["SP"] && r["SP"][0]) || "";
    const ep = (r["EP"] && r["EP"][0]) || "";
    let pages = "";
    if (sp && ep) pages = `${sp}-${ep}`;
    else if (sp) pages = sp;
    e.pages = normalizePages(pages);

    e.doi = (r["DO"] && r["DO"][0]) || "";
    e.url = (r["UR"] && r["UR"][0]) || "";
    e.publisher = (r["PB"] && r["PB"][0]) || "";
    e.address = (r["CY"] && r["CY"][0]) || "";

    const da = (r["DA"] && r["DA"][0]) || (r["Y1"] && r["Y1"][0]) || "";
    const dObj = parseDateToYearMonth(da);
    e.year = dObj.year;
    e.month = dObj.month;
  }

  return entries;
}

// ========= BibTeX パーサ =========
function parseBibtex(text) {
    const entries = [];
    const re = /@(\w+)\s*\{\s*([^,]+),([\s\S]*?)}\s*(?=@|$)/g;
    let m;

    while ((m = re.exec(text)) !== null) {
        const type = m[1].toLowerCase(), key = m[2].trim(), body = m[3];
        const fields = {};
        const fRe = /(\w+)\s*=\s*({(?:[^{}]*|{[^{}]*})*}|"[^"]*"|[^,\n]+)/g;
        let fm;

        while ((fm = fRe.exec(body)) !== null) {
            let val = fm[2].trim();
            // 括弧や引用符の除去
            if ((val[0]==='{' && val.slice(-1)==='}') || (val[0]==='"' && val.slice(-1)==='"')) {
                val = val.slice(1, -1);
            }
            // ★ ここでフィールド値をデコード！
            fields[fm[1].toLowerCase()] = decodeLatexAccents(val);
        }

        // 著者リストの各オブジェクトに対してもデコードを適用
        const parsedAuthors = parseAuthorString(fields.author || "").map(a => ({
            first: decodeLatexAccents(a.first),
            last: decodeLatexAccents(a.last)
        }));

        entries.push({ 
            type, 
            key, 
            authors: parsedAuthors, 
            title: fields.title, 
            journal: fields.journal, 
            volume: fields.volume, 
            pages: normalizePages(fields.pages), 
            year: fields.year, 
            doi: fields.doi 
        });
    }
    return entries;
}

// ========= bbl パーサ =========
function parseBbl(text) {
  const entries = [];
  const items = text.split(/\\bibitem\s*(?:\[[^\]]*\])?\s*\{/).slice(1);

  const clean = (str) => {
    if (!str) return "";
    return str
      .replace(/\\em\s+|\\it\s+|\\bf\s+|\\newblock/g, "")
      .replace(/\{|\}/g, "")
      .replace(/\s+/g, " ")
      .replace(/,\s*,/g, ",")
      .replace(/\.\s*\./g, ".")
      .replace(/[.,\s]+$/, "")
      .trim();
  };

  const normalizePages = (p) => p.replace(/[-–—]+/g, "-");

  for (const item of items) {
    const braceIndex = item.indexOf('}');
    if (braceIndex === -1) continue;

    let body = item.substring(braceIndex + 1).trim();
    body = body.replace(/\r?\n/g, " ");

    const common = { 
      type: "article", 
      authors: [], 
      title: "", 
      journal: "", 
      year: "", 
      volume: "", 
      pages: "",
      hasEtAl: false 
    };

    const yearMatch = body.match(/(.*?)\s*\((\d{4})\)/);
    if (yearMatch) {
      common.year = yearMatch[2];
      const rawAuthors = yearMatch[1];
      
      if (rawAuthors.includes("et al.")) {
        common.hasEtAl = true;
      }

      const authorList = rawAuthors
        .split(/\s+and\s+|,\s+and\s+|,/)
        .map(a => a.trim())
        .filter(a => a.length > 0 && a !== "et al.");

      common.authors = authorList.map(a => {
        const parts = a.split(",");
        return parts.length >= 2 
          ? { last: clean(parts[0]), first: clean(parts.slice(1).join(" ")) }
          : { last: clean(a), first: "" };
      });
    }

    const blocks = body.split(/\\newblock\s*/i).map(b => b.trim());

    if (blocks[1]) {
      common.title = clean(blocks[1]);
    }

    if (blocks[2]) {
      const journalPart = blocks[2];
      const jm = journalPart.match(/(.*?)\s*(\d+)\s*:\s*([\d\-–—]+)/);
      if (jm) {
        common.journal = clean(jm[1]);
        common.volume = jm[2];
        common.pages = normalizePages(jm[3]);
      } else {
        common.journal = clean(journalPart);
      }
    }
    
    entries.push(common);
  }

  return entries;
}

// ========= 変換 & 出力 =========
function toBibtex(entries) {
    // 1行目に意図的に追加するコメント
    const headerComment = "@comment{}\n";

    const bibContent = entries.map((e, i) => {
        const key = (formatAuthorsForCitation(e.authors).replace(/\s+/g, "") + (e.year || "xxxx")).toLowerCase();
        
        // emタグを \textit{ } に変換する処理
        const processText = (str) => {
            if (!str) return "";
            return convertAccentsToLatex(str)
                .replace(/<em>/g, "\\textit{")
                .replace(/<\/em>/g, "}");
        };

        const f = [
            ["author",  formatAuthorsForBibtex(e.authors)],
            ["title",   processText(e.title)],
            ["journal", processText(e.journal)],
            ["volume",  e.volume],
            ["pages",   e.pages],
            ["year",    e.year],
            ["doi",     e.doi],
        ];
        
        const body = f.map(([n, v]) => `  ${n.padEnd(9)} = {${v || ""}}`).join(",\n");
        return `@${e.type || "article"}{${key},\n${body}\n}`;
    }).join("\n\n");

    // コメントを先頭に結合して返す
    return headerComment + bibContent;
}

function escHtml(s) {
  if (!s) return "";

  // --- 1. <i>…</i> を退避 ---
  const placeholders = [];
  s = s.replace(/<i>[\s\S]*?<\/i>/g, m => {
    const key = `__I_TAG_${placeholders.length}__`;
    placeholders.push(m);
    return key;
  });

  // --- 2. 通常の HTML エスケープ ---
  s = s.replace(/[&<>"']/g, c => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[c]));

  // --- 3. プレースホルダを元に戻す ---
  placeholders.forEach((orig, i) => {
    s = s.replace(`__I_TAG_${i}__`, orig);
  });

  return s;
}

function formatAuthorList(authors) {
  if (!authors || authors.length === 0) return "";
  return authors.map(a => {
    const first = a.first ? a.first.trim() : "";
    const last = a.last ? a.last.trim() : "";
    return `${last} ${first}`.trim();
  });
}

/**
 * 参考文献リストを生成するメイン関数
 * @param {Array} entries - 文献データの配列
 * @param {string} style - フォーマットスタイル ("nogyokagaku", "seibutsu", "apa", "shokubutsu", "dojo")
 * @param {number} startIndex - 開始番号
 */
function toReferenceList(entries, style, startIndex) {
  return entries.map((e, i) => {
    const idx = startIndex + i;
    
    // 各学会・スタイルごとの関数に振り分け
    switch (style) {
      case "nogyokagaku":
        return formatRefNogyokagaku(e, idx);
      
      case "seibutsu":
        return formatRefSeibutsu(e, idx);
      
      case "apa":
        return formatRefAPA(e);
      
      case "shokubutsu":
        return formatRefShokubutsu(e);

      case "dojo":
        return formatRefdojo(e, idx);
      
      default:
        // デフォルト（APA風）
        return formatRefAPA(e);
    }
  }).join("\n").replace(/\\textit\{([^}]*)\}/g, "<i>$1</i>");
}

// --- 以下、提供された各フォーマット関数 ---

function formatRefNogyokagaku(e, index) {
  let rawAuthors = formatAuthorList(e.authors);
  let authors = rawAuthors.map(name => {
    const parts = name.split(/\s+/);
    const last = parts[0];
    const firstParts = parts.slice(1);
    const initials = firstParts
      .filter(x => x)
      .map(x => x[0] + ".")
      .join("");
    return `${last} ${initials}`;
  });
  if (authors.length === 1) {
    // 1人 → そのまま
    authors = [authors[0]];
  } else if (authors.length === 2) {
    // 2人 → A and B
    authors = [`${authors[0]} and ${authors[1]}`];
  } else if (authors.length < 10) {
    // 3〜9人 → A, B, C and D
    const last = authors.pop();
    authors = [`${authors.join(", ")} and ${last}`];
  } else {
    // 10人以上 → 5人 + et al.
    authors = authors.slice(0, 5);
    authors.push("<i>et al.</i>");
  }
  const authorsStr = authors.join(", ");
  const title = e.title ? `${e.title}.` : "";
  const journal = e.journal ? `<i>${escHtml(e.journal)}</i>` : "";
  const vol = e.volume ? ` <b>${escHtml(e.volume)}</b>` : "";
  const pages = e.pages ? `, ${escHtml(e.pages)}` : "";
  const yearBlock = e.year ? ` (${escHtml(e.year)}).` : "";
  return `${index}) ${escHtml(authorsStr)}, ${title} ${journal},${vol}${pages}${yearBlock}`;
}

function formatRefSeibutsu(e, index) {
  let authors = e.authors.map(a => {
  const initials = (a.first || "")
    .split(/\s+/)
    .filter(x => x)
    .map(x => x[0] + ".")
    .join("");
  return `${a.last} ${initials}`;
});
if (authors.length === 1) {
  authors = [authors[0]];
} else if (authors.length === 2) {
  authors = [`${authors[0]} and ${authors[1]}`];
} else {
  const last = authors.pop();
  authors = [`${authors.join(", ")} and ${last}`];
}
  const authorsStr = authors.join(", ");
  const journal = e.journal ? `<i>${escHtml(e.journal)}</i>` : "";
  const vol = e.volume ? `, <b>${escHtml(e.volume)}</b>` : "";
  const pages = e.pages ? `, ${escHtml(e.pages)}` : "";
  const year = e.year ? ` (${escHtml(e.year)})` : "";
  return `${index}) ${escHtml(authorsStr)}: ${journal}${vol}${pages}${year}.`;
}

function formatRefAPA(e) {
  const rawAuthors = formatAuthorList(e.authors);
  const authors = (() => {
    if (!rawAuthors || rawAuthors.length === 0) return "";
    const apaAuthors = rawAuthors.map(name => {
      const [last, ...firstParts] = name.split(/\s+/);
      const initials = firstParts
        .filter(x => x)
        .map(x => x[0].toUpperCase() + ".")
        .join(" ");
      return `${last}${initials ? ", " + initials : ""}`;
    });
    if (apaAuthors.length === 1) return apaAuthors[0];
    const lastAuthor = apaAuthors.pop();
    return apaAuthors.join(", ") + ", & " + lastAuthor;
  })();
  const year = e.year ? ` (${escHtml(e.year)}).` : " (n.d.).";
  const title = e.title ? ` ${e.title}.` : "";
  const journal = e.journal ? ` <i>${escHtml(e.journal)}</i>` : "";
  const vol = e.volume ? `, <i>${escHtml(e.volume)}</i>` : "";
  const num = e.number ? `(${escHtml(e.number)})` : "";
  const pages = e.pages ? `, ${escHtml(e.pages)}` : "";
  const doi = e.doi ? ` https://doi.org/${escHtml(e.doi)}` : "";
  return `${escHtml(authors)}${year}${title}${journal}${vol}${num}${pages}.${doi}`;
}

function formatRefShokubutsu(e) {
let authors = e.authors.map(a => {
  const initials = (a.first || "")
    .split(/\s+/)
    .filter(x => x)
    .map(x => x[0] + ".")
    .join("");
  return `${a.last} ${initials}`;
});
if (authors.length === 1) {
  // 1人 → そのまま
  authors = [authors[0]];
} else if (authors.length === 2) {
  // 2人 → A and B
  authors = [`${authors[0]} and ${authors[1]}`];
} else if (authors.length < 5) {
  // 3〜5人 → A, B, C and D
  const last = authors.pop();
  authors = [`${authors.join(", ")} and ${last}`];
} else {
  // 6人以上 → 6人 + et al.
  authors = authors.slice(0, 6);
  authors.push("et al.");
}
  const authorsStr = authors.join(", ");
  const year = e.year ? ` (${escHtml(e.year)})` : "";
  const title = e.title ? ` ${e.title}.` : "";
  const journal = e.journal ? ` <i>${escHtml(e.journal)}</i>` : "";
  const vol = e.volume ? ` <b>${escHtml(e.volume)}</b>` : "";
  const pages = e.pages ? `: ${escHtml(e.pages)}` : "";
  return `${escHtml(authorsStr)}${year}${title}${journal}${vol}${pages}.`;
}

function formatRefdojo(e) {
  // formatAuthorList で基本形（"Last First"）を取得
  let rawAuthors = formatAuthorList(e.authors);

  // 欧文イニシャル化 + 和文対応
  let authors = rawAuthors.map(name => {
    const parts = name.split(/\s+/);
    const last = parts[0];
    const firstParts = parts.slice(1);

    // 和文判定（非ASCIIが含まれる場合）
    const isJapanese = /[^\x00-\x7F]/.test(last);

    if (isJapanese) {
      // 和文 → 姓・名
      return `${last}・${firstParts.join("")}`;
    }

    // 欧文 → Last, Initials
    const initials = firstParts
      .filter(x => x)
      .map(x => x[0].toUpperCase() + ".")
      .join("");
    return initials ? `${last}, ${initials}` : last;
  });

  if (authors.length === 1) {
    authors = [authors[0]];
  } else if (authors.length === 2) {
    authors = [`${authors[0]} and ${authors[1]}`];
  } else if (authors.length < 10) {
    const last = authors.pop();
    authors = [`${authors.join(", ")} and ${last}`];
  } else {
    authors = authors.slice(0, 5);
    authors.push("et al.");
  }

  const authorsStr = authors.join(", ");

  // 各フィールド
  const title = e.title ? `${e.title}.` : "";
  const journal = e.journal ? `<i>${escHtml(e.journal)}</i>` : "";
  const vol = e.volume ? ` <b>${escHtml(e.volume)}</b>` : "";
  const pages = e.pages ? `, ${escHtml(e.pages)}` : "";
  const yearBlock = e.year ? ` ${escHtml(e.year)}.` : "";

  // 最終整形
  return `${escHtml(authorsStr)} ${yearBlock} ${title} ${journal},${vol}${pages}`;
}

// ========= UI 制御 =========
function renderStack() {
    const style = styleSelectEl.value;
    unifiedOutputEl.innerHTML = rawEntriesStack.length ? "" : "<p style='color:#999; padding:20px;'>スタックは空です．</p>";
    
    rawEntriesStack.forEach((item, idx) => {
        const card = document.createElement("div");
        card.className = "entry-card";
        card.style = "background:#fff; border:1px solid #ddd; border-radius:8px; padding:10px; margin-bottom:15px;";
        
        let content = "";
        if (style === "bibtex_mode") {
            // BibTeXを生成。個別にラップして確実に改行コードを維持する
            content = item.editedBib || toBibtex([item.data]);
            card.innerHTML = `<div style="display:flex; justify-content:space-between; margin-bottom:8px;">
                                <span style="font-weight:bold;">#${idx+1}</span>
                                <button onclick="removeEntryAt(${idx})" class="btn-danger" style="padding:2px 8px;">削除</button>
                              </div>
                              <div contenteditable="true" class="bib-editor" style="font-family:monospace; font-size:13px; background:#f9f9f9; padding:8px; outline:none; white-space: pre; overflow-x: auto;"></div>`;
            // innerText を使うことで \n を改行として流し込む
            card.querySelector(".bib-editor").innerText = content;
        } else {
            // 通常の引用リスト形式
            content = item.editedRef || toReferenceList([item.data], style, idx + 1);
            card.innerHTML = `<div style="display:flex; justify-content:space-between; margin-bottom:8px;">
                                <span style="font-weight:bold;">#${idx+1}</span>
                                <button onclick="removeEntryAt(${idx})" class="btn-danger" style="padding:2px 8px;">削除</button>
                              </div>
                              <div contenteditable="true" class="ref-editor" style="font-family:serif; line-height:1.6; outline:none;">${content.replace(/\n/g, "<br>")}</div>`;
        }
        
        const editor = card.querySelector('[contenteditable="true"]');
        editor.addEventListener("input", () => {
            if (style === "bibtex_mode") item.editedBib = editor.innerText;
            else item.editedRef = editor.innerHTML;
        });

        unifiedOutputEl.appendChild(card);
    });
    updateDownloadButtonState();
}

function removeEntryAt(index) {
    if (confirm(`エントリー #${index + 1} を削除しますか？`)) {
        rawEntriesStack.splice(index, 1);
        renderStack();
        logOutputEl.value = `削除しました．合計: ${rawEntriesStack.length} 件`;
    }
}

function updateDownloadButtonState() {
    const downloadBtn = document.getElementById("mainDownloadBtn");
    const copyBtn = document.getElementById("copyAllBtn");
    const isEmpty = rawEntriesStack.length === 0;
    
    downloadBtn.disabled = isEmpty;
    copyBtn.disabled = isEmpty;
    downloadBtn.innerText = styleSelectEl.value === "bibtex_mode" ? ".bib で一括保存" : ".rtf でリスト保存";
}

// ========= イベントリスナー登録 =========
document.getElementById("convertGenericBtn").onclick = () => {
    const text = inputEl.value.trim();
    if (!text) return;
    const fmt = detectFormat(text);
    let entries = [];
    if (fmt === "nbib") entries = parseNbib(text);
    else if (fmt === "ris") entries = parseRis(text);
    else if (fmt === "bibtex") entries = parseBibtex(text);
    else if (fmt === "bbl") entries = parseBbl(text);
    
    if (entries.length) {
        entries.forEach(e => rawEntriesStack.push({ data: e, editedBib: null, editedRef: null }));
        inputEl.value = "";
        renderStack();
        logOutputEl.value = `${entries.length} 件追加．合計: ${rawEntriesStack.length} 件`;
    } else alert("文献データを検出できませんでした．");
};

document.getElementById("removeComment").onclick = () => {
    inputEl.value = inputEl.value.replace(/@comment\{.*?\}/g, "").replace(/\\begin\{thebibliography\}\{.*?\}/g, "").replace(/\\end\{thebibliography\}/g, "");
};

document.getElementById("clearInputBtn").onclick = () => inputEl.value = "";

document.getElementById("clearAllBtn").onclick = () => {
    if (confirm("スタックをすべて削除しますか？")) { rawEntriesStack = []; renderStack(); logOutputEl.value = ""; }
};

document.getElementById("copyAllBtn").onclick = () => {
    const range = document.createRange();
    range.selectNodeContents(unifiedOutputEl);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
    document.execCommand("copy");
    alert("コピーしました");
};

document.getElementById("fileInput").onchange = (e) => {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (ev) => inputEl.value = ev.target.result;
        reader.readAsText(file, "UTF-8");
    }
};

styleSelectEl.onchange = renderStack;

document.getElementById("mainDownloadBtn").onclick = () => {
    const style = styleSelectEl.value;
    let content = "";

    if (style === "bibtex_mode") {
        content = rawEntriesStack.map(item => item.editedBib || toBibtex([item.data])).join("\n\n");
        downloadFile(content, "references.bib", "text/plain;charset=utf-8");
    } else {
        const entries = rawEntriesStack.map(item => item.data);
        const htmlContent = toReferenceList(entries, style, 1);
        
        // RTF変換を実行
        const rtfContent = convertToRTF(htmlContent);
        // RTFはバイナリに近い扱いのため、Blob作成時に明示的なエンコード指定はせずに出力
        downloadFile(rtfContent, "references.rtf", "application/rtf");
    }
};

/**
 * HTMLをRTFに変換し、特殊文字のエスケープも行う
 */
function convertToRTF(html) {
    if (!html) return "";

    // --- 追加: HTML特殊文字をデコードする処理 ---
    let decodedHtml = html
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&#039;/g, "'");

    let rtf = decodedHtml
        // emタグとiタグの両方を斜体に
        .replace(/<(i|em)>(.*?)<\/\1>/g, "{\\i $2}")
        // bタグを太字に
        .replace(/<b>(.*?)<\/b>/g, "{\\b $1}")
        // 改行処理
        .replace(/<br\s*\/?>/g, "\\line ")
        .replace(/\n/g, "\\line ");

    // アクセント付き文字などのUnicode文字をRTFエスケープ形式に変換
    rtf = rtf.replace(/[^\x00-\x7F]/g, function(c) {
        return "\\u" + c.charCodeAt(0).toString() + "?";
    });

    // RTFヘッダー
    const header = "{\\rtf1\\ansi\\ansicpg932\\deff0{\\fonttbl{\\f0\\fnil\\fcharset128 MS Mincho;}{\\f1\\fnil\\fcharset0 Times New Roman;}}\n";
    const footer = "\n}";

    return header + "{\\f1 " + rtf + "}" + footer;
}

function downloadFile(content, fileName, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const a = document.createElement("a");
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }, 0);
}

// 初期化
window.addEventListener("pageshow", renderStack);