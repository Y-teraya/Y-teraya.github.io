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
    return str.replace(/[^ -~]/g, ch => map[ch] || ch);
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
    const keys = Object.keys(map).sort((a, b) => b.length - a.length);
    for (const k of keys) { out = out.replace(new RegExp(k.replace(/\\/g, "\\\\"), "g"), map[k]); }
    return out.replace(/\\/g, "");
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

function parseNbib(text) {
    const entries = [];
    let current = null;
    text.split(/\n/).forEach(line => {
        const m = line.match(/^([A-Z0-9]{2,4})\s*-\s*(.*)$/);
        if (m) {
            const tag = m[1], val = m[2].trim();
            if (tag === "PMID") { if(current) entries.push(current); current = { type:"article", authors:[], raw:{} }; }
            if (current) { current.raw[tag] = (current.raw[tag] || "") + (current.raw[tag] ? " " : "") + val; current._lt = tag; }
        } else if (current && current._lt) {
            current.raw[current._lt] += " " + line.trim();
        }
    });
    if (current) entries.push(current);
    return entries.map(e => {
        const r = e.raw;
        const fau = text.match(/^FAU\s*-\s*(.*)$/gm) || []; // 簡易取得
        e.authors = fau.map(l => { const p = l.split("-")[1].split(","); return { last: sTrim(p[0]), first: sTrim(p[1]) }; });
        e.title = r["TI"] || ""; e.journal = r["JT"] || r["TA"] || ""; e.volume = r["VI"] || ""; e.pages = normalizePages(r["PG"] || "");
        const d = parseDateToYearMonth(r["DP"] || ""); e.year = d.year; e.month = d.month; e.doi = (r["LID"] || "").match(/10\.\S+/)?.[0] || "";
        return e;
    });
}

function parseRis(text) {
    const entries = []; let current = null;
    text.split(/\n/).forEach(line => {
        const m = line.match(/^([A-Z0-9]{2})\s*-\s*(.*)$/);
        if (!m) return;
        const tag = m[1], val = m[2].trim();
        if (tag === "TY") { if(current) entries.push(current); current = { type:"article", authors:[], raw:{} }; }
        if (!current) return;
        if (tag === "AU") current.authors.push(...parseAuthorString(val));
        else current.raw[tag] = val;
    });
    if (current) entries.push(current);
    return entries.map(e => {
        e.title = e.raw["TI"] || e.raw["T1"] || ""; e.journal = e.raw["T2"] || e.raw["JO"] || "";
        e.volume = e.raw["VL"] || ""; e.pages = normalizePages(`${e.raw["SP"] || ""}-${e.raw["EP"] || ""}`);
        const d = parseDateToYearMonth(e.raw["DA"] || e.raw["Y1"] || ""); e.year = d.year; e.doi = e.raw["DO"] || "";
        return e;
    });
}

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
            if ((val[0]==='{' && val.slice(-1)==='}') || (val[0]==='"' && val.slice(-1)==='"')) val = val.slice(1, -1);
            fields[fm[1].toLowerCase()] = val;
        }
        entries.push({ type, key, authors: parseAuthorString(fields.author), title: fields.title, journal: fields.journal, volume: fields.volume, pages: normalizePages(fields.pages), year: fields.year, doi: fields.doi });
    }
    return entries;
}

function parseBbl(text) {
    return text.split(/\\bibitem/).slice(1).map(item => {
        const body = item.replace(/\r?\n/g, " ").trim();
        const yearM = body.match(/\((\d{4})\)/);
        return { type: "article", authors: parseAuthorString(body.split("(")[0]), title: body.match(/\\newblock\s+(.*?)\\newblock/)?.[1] || "", year: yearM?.[1] || "" };
    });
}

// ========= 変換 & 出力 =========
function toBibtex(entries) {
    return entries.map((e, i) => {
        const key = (formatAuthorsForCitation(e.authors).replace(/\s+/g, "") + (e.year || "xxxx")).toLowerCase();
        const f = [["author", formatAuthorsForBibtex(e.authors)], ["title", e.title], ["journal", convertAccentsToLatex(e.journal)], ["volume", e.volume], ["pages", e.pages], ["year", e.year], ["doi", e.doi]];
        const body = f.map(([n, v]) => `  ${n.padEnd(9)} = {${v || ""}}`).join(",\n");
        return `@${e.type || "article"}{${key},\n${body}\n}`;
    }).join("\n\n");
}

function escHtml(s) { return (s || "").replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c])); }

function formatAuthorList(authors, style) {
    if (!authors || authors.length === 0) return "";
    const fmt = authors.map(a => {
        if (style === "APA") {
            const init = (a.first || "").split(/\s+/).filter(x => x).map(x => x[0].toUpperCase() + ".").join(" ");
            return `${a.last}${init ? ", " + init : ""}`;
        }
        return `${a.last} ${a.first}`.trim();
    });
    return style === "APA" && fmt.length > 1 ? fmt.slice(0, -1).join(", ") + ", & " + fmt.slice(-1) : fmt.join(", ");
}

function toReferenceList(entries, style, startIndex) {
    return entries.map((e, i) => {
        const idx = startIndex + i;
        const au = formatAuthorList(e.authors, style === "apa" ? "APA" : "default");
        const title = e.title ? ` ${decodeLatexAccents(e.title)}.` : "";
        const jr = e.journal ? ` <i>${escHtml(decodeLatexAccents(e.journal))}</i>` : "";
        const vol = e.volume ? ` <b>${escHtml(e.volume)}</b>` : "";
        const pg = e.pages ? `, ${escHtml(e.pages)}` : "";
        const yr = e.year ? ` (${escHtml(e.year)})` : "";

        if (style === "nogyokagaku") return `${idx}) ${escHtml(au)}, ${title}${jr},${vol}${pg}${yr}.`;
        if (style === "seibutsu") return `${idx}) ${escHtml(au)}: ${jr}${vol}${pg}${yr}.`;
        if (style === "apa") return `${escHtml(au)}${yr}.${title}${jr}${vol}${pg}.`;
        return `${escHtml(au)}${yr}.${title}${jr}${vol}${pg}.`;
    }).join("\n");
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
            content = item.editedBib || toBibtex([item.data]);
            card.innerHTML = `<div style="display:flex; justify-content:space-between; margin-bottom:8px;"><span style="font-weight:bold;">#${idx+1} [BibTeX]</span><button onclick="removeEntryAt(${idx})" class="btn-danger" style="padding:2px 8px;">削除</button></div>
                              <div contenteditable="true" class="bib-editor" style="font-family:monospace; font-size:13px; background:#f9f9f9; padding:8px; outline:none;">${escHtml(content)}</div>`;
        } else {
            content = item.editedRef || toReferenceList([item.data], style, idx + 1).replace(/\n/g, "<br>");
            card.innerHTML = `<div style="display:flex; justify-content:space-between; margin-bottom:8px;"><span style="font-weight:bold;">#${idx+1}</span><button onclick="removeEntryAt(${idx})" class="btn-danger" style="padding:2px 8px;">削除</button></div>
                              <div contenteditable="true" class="ref-editor" style="font-family:serif; line-height:1.6; outline:none;">${content}</div>`;
        }
        
        // 編集時の保存処理
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
    const btn = document.getElementById("mainDownloadBtn");
    btn.disabled = rawEntriesStack.length === 0;
    btn.innerText = styleSelectEl.value === "bibtex_mode" ? ".bib で一括保存" : ".rtf でリスト保存";
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
        const blob = new Blob([content], { type: "text/plain" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "references.bib";
        a.click();
    } else {
        alert("リスト保存機能（RTF/TXT）は現在プレビュー版です。画面からコピーしてご利用ください。");
    }
};

// 初期化
window.addEventListener("pageshow", renderStack);