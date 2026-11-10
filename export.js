/* ============================================================
   EXPORT ▾  (Faik, 2026-11-10)
   One Export button beside Fast lookup. It lists what is on the tab
   in front of the user and exports the chosen part:
     • Entire tab   → one workbook, one sheet per pane
     • Numbers      → the KPI tiles
     • any table    → its rows, numbers as real numbers; drill-down
                      trees go out complete, grouped so Excel's own
                      +/− buttons open and close each level
     • any chart    → the numbers behind it + the picture, or the
                      picture alone as a PNG
   Pop-ups ("Show more…") get their own Export button, because the
   header is covered while one is open.

   Who sees it: EXPORT_VIEWERS in index.html (applyExportAccess shows
   or hides everything with class exportUi). Presentation-level only,
   like the other restricted items in this tool.

   Everything is read from the page as it stands (filters, sort and
   search applied), so new tabs and panes are covered without any
   per-tab code. Workbooks are written with ExcelJS, fetched the first
   time someone exports.
   ============================================================ */
(function () {
    "use strict";

    const EXCELJS_URL = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    const FMT = { money: '"$"#,##0.00;[Red]-"$"#,##0.00', pct: "0.00%", num: "#,##0.##" };
    const TRANSPARENT = /^(transparent|rgba\(.*,\s*0\))$/;
    const byId = id => document.getElementById(id);
    const vis = el => !!el && el.getClientRects().length > 0 && getComputedStyle(el).visibility !== "hidden";
    const clean = s => (s || "").replace(/[▲▼↑↓⇅↗]/g, "").replace(/[ \s]+/g, " ").trim();
    const txt = el => el ? clean(el.textContent) : "";
    const allowed = () => typeof exportAllowed === "function" && exportAllowed();

    // ---------- what is on screen ----------
    function scope() {
        const tab = txt(byId("headerTitle")) || "Sales";
        const modal = [...document.querySelectorAll(".modal-overlay")].filter(vis).pop();
        if (modal) return { root: modal, title: txt(modal.querySelector(".modal-header h2")) || tab };
        const drill = [...document.querySelectorAll(".drill")].find(vis);
        if (drill) return { root: drill, title: tab + " drill-down" };
        return { root: [...document.querySelectorAll("section[data-tab-content]")].find(vis) || null, title: tab };
    }
    function paneTitle(el, fallback) {
        const box = el.closest(".panel, .card, .modal-card");
        const h = box && box.querySelector(".panel-header h2, .head h3, .modal-header h2, h3, h2");
        return txt(h) || fallback;
    }
    function collect(sc) {
        sc = sc || scope();
        const out = { title: sc.title, root: sc.root, tiles: [], tables: [], charts: [], pictures: [] };
        if (!sc.root) return out;
        out.tiles = [...sc.root.querySelectorAll(".kpi")].filter(vis);
        const seen = new Map();
        const named = (el, fb) => { let n = paneTitle(el, fb); const k = (seen.get(n) || 0) + 1; seen.set(n, k); return k > 1 ? n + " (" + k + ")" : n; };
        sc.root.querySelectorAll("table").forEach(t => {
            if (!vis(t)) return;
            const rows = t.querySelectorAll("tbody tr");
            if (!rows.length || (rows.length === 1 && rows[0].querySelector("td.empty"))) return;
            out.tables.push({ el: t, name: named(t, "Table") });
        });
        sc.root.querySelectorAll("canvas").forEach(c => {
            if (!vis(c)) return;
            const ch = window.Chart && Chart.getChart(c);
            if (ch) out.charts.push({ el: c, name: named(c, "Chart") });
            else if (c.clientWidth > 150 && c.clientHeight > 100) out.pictures.push({ el: c, name: named(c, "Visual") });
        });
        sc.root.querySelectorAll("svg").forEach(s => {
            const r = s.getBoundingClientRect();
            if (vis(s) && r.width > 200 && r.height > 120) out.pictures.push({ el: s, name: named(s, "Visual") });
        });
        return out;
    }

    // ---------- cells ----------
    // A cell becomes a number when it sits in a numeric column or
    // carries a $ or % sign; codes such as "000123" stay text.
    function parseValue(text, numeric) {
        const t = clean(text);
        if (t === "") return { v: null };
        if (t === "—" || t === "-" || t === "n/a") return { v: numeric ? null : t };
        const m = /^(\(|[+\-−]?)\s*(\$?)\s*(\d[\d,]*\.?\d*|\.\d+)\s*(%?)\s*\)?$/.exec(t);
        if (!m || !(numeric || m[2] || m[4])) return { v: t };
        let n = parseFloat(m[3].replace(/,/g, ""));
        if (isNaN(n)) return { v: t };
        if (m[1] === "(" || m[1] === "-" || m[1] === "−") n = -n;
        if (m[4]) return { v: n / 100, z: FMT.pct };
        return { v: n, z: m[2] ? FMT.money : FMT.num };
    }
    function cellText(td) {
        const input = td.querySelector("input, select");
        if (input && !td.textContent.trim()) return input.value || "";
        const parts = [];
        const w = document.createTreeWalker(td, NodeFilter.SHOW_TEXT);
        for (let n = w.nextNode(); n; n = w.nextNode()) {
            if (n.parentElement && n.parentElement.closest("button")) continue;   // action buttons are not data
            const t = clean(n.nodeValue);
            if (t && !/^[▾▸·]$/.test(t)) parts.push(t);
        }
        return parts.join(" ");
    }

    // ---------- drill-down trees: every level, not only what is open ----------
    // Each tree draws a row's children only when its "expanded" set holds
    // the row. For the export the sets are swapped for one that says yes
    // to everything, the tree is drawn once in full and read, then the
    // real sets go back and it is drawn again as the user had it. No
    // paint happens in between, so nothing flickers.
    const TREE_RENDER = {
        customerDrillTable:    () => renderCustomerDrillTree(),
        branchDrillTable:      () => renderBranchDrillTree(),
        eomDrillTable:         () => renderEOMDrillTree(),
        salespersonDrillTable: () => renderSalespersonDrillTree(),
        decompTable:           () => renderDecomposition(),
    };
    function expandAll(table) {
        const rerender = TREE_RENDER[table.id] || (state.drill ? null : () => renderCurrentTab());
        if (!rerender) return null;
        const slots = [[state, "expandedMatrix"], [state, "expandedQuoteMatrix"], [state, "expandedBoMatrix"]];
        if (typeof decompState !== "undefined") slots.push([decompState, "expanded"]);
        const saved = slots.map(([o, k]) => o[k]);
        const ALL = { has: () => true, add() { return this; }, delete() { return false; }, clear() {}, size: 0, forEach() {}, [Symbol.iterator]: function* () {} };
        const restore = () => slots.forEach(([o, k], i) => { o[k] = saved[i]; });
        slots.forEach(([o, k]) => { o[k] = ALL; });
        try { rerender(); } catch (e) { restore(); try { rerender(); } catch (e2) { /* leave as is */ } console.warn("[Export] full tree not available, exporting the open rows:", e); return null; }
        return () => { restore(); try { rerender(); } catch (e) { console.warn("[Export] tree redraw failed:", e); } };
    }

    // ---------- table → rows ----------
    function scrapeTable(table) {
        const isTree = !!table.querySelector("tbody tr[data-expandable]");
        const headRow = [...table.querySelectorAll("thead tr")].pop();
        const headers = [];
        if (headRow) headRow.querySelectorAll("th, td").forEach(th => { for (let i = 0; i < (th.colSpan || 1); i++) headers.push(txt(th)); });
        const rows = [];
        table.querySelectorAll("tbody tr, tfoot tr").forEach(tr => {
            const tds = [...tr.children].filter(c => c.tagName === "TD" || c.tagName === "TH");
            if (!tds.length || (tds.length === 1 && tds[0].classList.contains("empty"))) return;
            let level = 0;
            if (isTree) { const lead = /^[ \s]*/.exec(tds[0].textContent)[0]; level = Math.round((lead.match(/ /g) || []).length / 4); }
            const cells = [];
            tds.forEach((td, i) => {
                cells.push(i === 0 && isTree ? { v: cellText(td) } : parseValue(cellText(td), td.classList.contains("num")));
                for (let k = 1; k < (td.colSpan || 1); k++) cells.push({ v: null });
            });
            const fw = tr.style.fontWeight || getComputedStyle(tr).fontWeight;
            rows.push({ cells, level, bold: fw === "bold" || parseInt(fw, 10) >= 600 });
        });
        const width = Math.max(headers.length, ...rows.map(r => r.cells.length), 1);
        while (headers.length < width) headers.push("");
        for (let c = width - 1; c >= 0; c--) {
            if (headers[c] || rows.some(r => r.cells[c] && r.cells[c].v != null && r.cells[c].v !== "")) continue;
            headers.splice(c, 1); rows.forEach(r => r.cells.splice(c, 1));
        }
        return { headers, rows, tree: isTree };
    }
    function tableModel(table) {
        const isTree = !!table.querySelector("tbody tr[data-expandable]");
        const undo = isTree ? expandAll(table) : null;
        try { return scrapeTable(table.id ? (byId(table.id) || table) : table); }
        finally { if (undo) undo(); }
    }
    function tilesModel(tiles) {
        const rows = [];
        const push = (label, valueEl, noteEl) => {
            const vals = valueEl ? [...valueEl.childNodes].map(n => clean(n.textContent)).filter(Boolean) : [];
            const first = parseValue(vals[0] || "", true), second = vals[1] ? parseValue(vals[1], true) : { v: null };
            rows.push({ cells: [{ v: label }, first, second, { v: noteEl ? txt(noteEl) || null : null }], level: 0 });
        };
        for (const k of tiles) {
            push(txt(k.querySelector(".label")), k.querySelector(".value"), k.querySelector(":scope > .sub"));
            k.querySelectorAll(".ds-part").forEach(p => push(txt(p.querySelector(".ds-part-lbl")), p.querySelector(".ds-part-val"), p.querySelector(".sub")));
        }
        return { headers: ["Tile", "Value", "Second value", "Note"], rows, tree: false };
    }

    // ---------- chart → rows ----------
    function chartModel(chart) {
        const data = chart.data || {}, labels = data.labels || [];
        const sets = (data.datasets || []).map((d, i) => ({ d, i })).filter(x => chart.isDatasetVisible(x.i));
        const horiz = (chart.options && chart.options.indexAxis) === "y";
        const prim = v => (v == null || typeof v === "number" || typeof v === "string" || typeof v === "boolean");
        const cell = v => (typeof v === "number" ? { v, z: FMT.num } : { v: v == null ? null : String(v) });
        const val = v => {
            if (v == null || typeof v === "number") return v;
            if (Array.isArray(v)) return typeof v[1] === "number" && typeof v[0] === "number" ? v[1] - v[0] : null;
            if (typeof v === "object") { const p = horiz ? v.x : v.y; return typeof p === "number" ? p : (typeof v.v === "number" ? v.v : (typeof v.value === "number" ? v.value : null)); }
            return typeof v === "string" && v !== "" && !isNaN(v) ? Number(v) : null;
        };
        const flat = labels.length && sets.every(s => (s.d.data || []).length <= labels.length);
        if (flat) {
            const headers = ["Label"].concat(sets.map(s => clean(s.d.label) || (sets.length === 1 ? "Value" : "Series " + (s.i + 1))));
            const rows = labels.map((l, r) => ({ level: 0, cells: [{ v: Array.isArray(l) ? l.join(" ") : (l == null ? "" : String(l)) }].concat(sets.map(s => cell(val((s.d.data || [])[r])))) }));
            return { headers, rows, tree: false };
        }
        // points without shared labels (scatter, bubble): one row per point
        const keys = [];
        sets.forEach(s => (s.d.data || []).forEach(p => { if (p && typeof p === "object" && !Array.isArray(p)) Object.keys(p).forEach(k => { if (prim(p[k]) && !keys.includes(k)) keys.push(k); }); }));
        if (!keys.length) keys.push("value");
        const rows = [];
        sets.forEach(s => (s.d.data || []).forEach(p => rows.push({ level: 0, cells: [{ v: clean(s.d.label) || "Series " + (s.i + 1) }].concat(keys.map(k => cell(p && typeof p === "object" ? p[k] : p))) })));
        return { headers: ["Series"].concat(keys), rows, tree: false };
    }

    // ---------- pictures ----------
    function backgroundOf(el) {
        for (let e = el; e && e.nodeType === 1; e = e.parentElement) { const b = getComputedStyle(e).backgroundColor; if (b && !TRANSPARENT.test(b)) return b; }
        return "#ffffff";
    }
    function canvasPng(c) {
        const o = document.createElement("canvas"); o.width = c.width; o.height = c.height;
        const g = o.getContext("2d"); g.fillStyle = backgroundOf(c); g.fillRect(0, 0, o.width, o.height); g.drawImage(c, 0, 0);
        return { url: o.toDataURL("image/png"), w: c.clientWidth || c.width, h: c.clientHeight || c.height };
    }
    const SVG_STYLE = ["fill", "fill-opacity", "stroke", "stroke-width", "stroke-opacity", "stroke-dasharray", "opacity", "font-family", "font-size", "font-weight", "text-anchor", "dominant-baseline", "visibility", "display"];
    function svgPng(svg) {
        return new Promise((resolve, reject) => {
            const r = svg.getBoundingClientRect(), w = Math.round(r.width), h = Math.round(r.height);
            const copy = svg.cloneNode(true);
            const src = [svg, ...svg.querySelectorAll("*")], dst = [copy, ...copy.querySelectorAll("*")];
            src.forEach((s, i) => { const cs = getComputedStyle(s); SVG_STYLE.forEach(p => { const v = cs.getPropertyValue(p); if (v) dst[i].style.setProperty(p, v); }); });
            copy.setAttribute("xmlns", "http://www.w3.org/2000/svg"); copy.setAttribute("width", w); copy.setAttribute("height", h);
            const img = new Image();
            img.onload = () => {
                const o = document.createElement("canvas"); o.width = w * 2; o.height = h * 2;
                const g = o.getContext("2d"); g.fillStyle = backgroundOf(svg); g.fillRect(0, 0, o.width, o.height); g.drawImage(img, 0, 0, o.width, o.height);
                try { resolve({ url: o.toDataURL("image/png"), w, h }); } catch (e) { reject(e); }
            };
            img.onerror = () => reject(new Error("could not draw this visual"));
            img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new XMLSerializer().serializeToString(copy));
        });
    }
    const picture = el => (el.tagName === "CANVAS" ? Promise.resolve(canvasPng(el)) : svgPng(el));

    // ---------- workbook ----------
    let excelLoading = null;
    function loadExcel() {
        if (window.ExcelJS) return Promise.resolve();
        if (!excelLoading) excelLoading = new Promise((res, rej) => {
            const s = document.createElement("script"); s.src = EXCELJS_URL;
            s.onload = () => res(); s.onerror = () => { excelLoading = null; rej(new Error("the Excel writer could not be downloaded — check the connection and try again")); };
            document.head.appendChild(s);
        });
        return excelLoading;
    }
    function sheetName(wb, wanted) {
        const base = (wanted || "Sheet").replace(/[\[\]:*?\/\\]/g, " ").replace(/\s+/g, " ").trim().slice(0, 31) || "Sheet";
        let name = base, n = 2;
        const taken = new Set(wb.worksheets.map(w => w.name.toLowerCase()));
        while (taken.has(name.toLowerCase())) { const suf = " (" + n++ + ")"; name = base.slice(0, 31 - suf.length) + suf; }
        return name;
    }
    function addSheet(wb, wanted, model, pic) {
        const ws = wb.addWorksheet(sheetName(wb, wanted), { views: [{ state: "frozen", ySplit: 1 }] });
        if (model) {
            if (model.tree) ws.properties.outlineProperties = { summaryBelow: false, summaryRight: false };
            ws.addRow(model.headers).font = { bold: true };
            const widths = model.headers.map(h => String(h).length);
            // Excel reads "this group is closed" from the row that owns the
            // group (the one above its children). ExcelJS derives that flag
            // from the level alone and gets it wrong for summary-above trees,
            // so it is set per row here: closed = has children underneath.
            if (model.tree) ws.properties.outlineLevelRow = Math.min(7, Math.max(0, ...model.rows.map(r => r.level)));
            for (let ri = 0; ri < model.rows.length; ri++) {
                const r = model.rows[ri];
                const row = ws.addRow(r.cells.map(c => c.v));
                if (model.tree) { const parent = ri + 1 < model.rows.length && model.rows[ri + 1].level > r.level; Object.defineProperty(row, "collapsed", { get: () => parent, configurable: true }); }
                r.cells.forEach((c, i) => {
                    if (c.z) row.getCell(i + 1).numFmt = c.z;
                    const len = c.v == null ? 0 : (typeof c.v === "number" ? (c.z === FMT.pct ? 8 : Math.round(Math.abs(c.v)).toString().length + 8) : String(c.v).length);
                    if (len > (widths[i] || 0)) widths[i] = len;
                });
                if (r.bold) row.font = { bold: true };
                if (model.tree) {
                    if (r.level) { row.outlineLevel = Math.min(7, r.level); row.hidden = true; row.getCell(1).alignment = { indent: Math.min(15, r.level * 2) }; }
                    else row.font = { bold: true };
                }
            }
            widths.forEach((w, i) => { ws.getColumn(i + 1).width = Math.min(60, Math.max(8, w + 2 + (model.tree && i === 0 ? 8 : 0))); });
            if (!model.tree && model.rows.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: model.headers.length } };
        }
        if (pic) {
            const scale = Math.min(1, 900 / pic.w);
            const id = wb.addImage({ base64: pic.url, extension: "png" });
            ws.addImage(id, { tl: { col: model ? model.headers.length + 1 : 0, row: model ? 1 : 0 }, ext: { width: Math.round(pic.w * scale), height: Math.round(pic.h * scale) } });
        }
        return ws.name;
    }
    function aboutSheet(wb, title, contents) {
        const ws = wb.addWorksheet("About");
        const f = state.filters || {}, lines = [
            ["Sales Analytics export", ""],
            ["Tab", title],
            ["Period", (byId("fromDate").value || "") + " to " + (byId("toDate").value || "")],
            ["Branch filter", f.location || "All"], ["Salesperson filter", f.salesperson || "All"],
            ["Customer filter", f.customerName || f.customer || "All"], ["Item filter", f.item || "All"],
            ["Rentals", state.includeRentals ? "included" : "excluded"],
            ["Figures as of", state.dataAsOf ? state.dataAsOf + " (snapshot)" : "live load"],
            ["Exported", new Date().toLocaleString("en-AU")],
            ["Exported by", typeof currentSignInName === "function" ? currentSignInName() : ""],
            ["", ""], ["Sheet", "What it holds"],
        ].concat(contents);
        lines.forEach((l, i) => { const row = ws.addRow(l); if (i === 0 || l[0] === "Sheet") row.font = { bold: true }; });
        ws.getColumn(1).width = 34; ws.getColumn(2).width = 70;
    }
    function save(blob, name) {
        const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
        document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    }
    const fileBase = (title, part) => ("sales-tool " + title + (part && part !== title ? " - " + part : "")).replace(/[\\\/:*?"<>|·]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 110) + " " + new Date().toISOString().slice(0, 10);
    async function writeBook(wb, base) {
        const buf = await wb.xlsx.writeBuffer();
        save(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), base + ".xlsx");
    }

    // ---------- the export itself ----------
    // job: { kind: "tab" | "tiles" | "table" | "chart" | "picture", el, name, as: "xlsx" | "png" }
    async function run(job) {
        if (!allowed()) return;
        const sc = scope(), found = collect(sc);
        const card = typeof showBusy === "function";
        const step = async (frac, what) => { if (card) { setBusyProgress(frac); setBusyDetail(what); await nextPaint(); } };
        if (card) showBusy("Preparing the export", job.name || found.title);
        try {
            if (job.as === "png") {
                await step(0.3, job.name);
                const pic = await picture(job.el);
                const bin = atob(pic.url.split(",")[1]), bytes = new Uint8Array(bin.length);
                for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
                save(new Blob([bytes], { type: "image/png" }), fileBase(found.title, job.name) + ".png");
                if (card) await finishBusy();
                return;
            }
            await step(0.05, "Getting the Excel writer…");
            await loadExcel();
            const wb = new ExcelJS.Workbook(); wb.creator = "Sales Analytics";
            const contents = [];
            const parts = job.kind === "tab"
                ? [].concat(found.tiles.length ? [{ kind: "tiles", name: "Numbers (tiles)" }] : [],
                            found.tables.map(t => ({ kind: "table", el: t.el, name: t.name })),
                            found.charts.map(c => ({ kind: "chart", el: c.el, name: c.name })),
                            found.pictures.map(p => ({ kind: "picture", el: p.el, name: p.name })))
                : [job];
            if (!parts.length) { alert("There is nothing on this tab to export yet."); return; }
            // Trees are read last: reading one redraws the tab, which can replace
            // other panes' elements. Sheets still come out in on-screen order.
            const isTree = p => p.kind === "table" && !!p.el.querySelector("tbody tr[data-expandable]");
            const order = parts.map((p, i) => i).sort((a, b) => (isTree(parts[a]) ? 1 : 0) - (isTree(parts[b]) ? 1 : 0) || a - b);
            const built = [];
            let done = 0;
            for (const i of order) {
                const p = parts[i];
                await step(0.1 + 0.7 * (done++ / parts.length), p.name);
                try {
                    if (p.kind === "tiles") built[i] = { model: tilesModel(found.tiles), note: "The KPI tiles" };
                    else if (p.kind === "table") { const m = tableModel(p.el); built[i] = { model: m, note: m.tree ? "Drill-down: use the + buttons at the left to open each level" : "Table" }; }
                    else if (p.kind === "chart") { const ch = Chart.getChart(p.el); built[i] = { model: ch ? chartModel(ch) : null, pic: canvasPng(p.el), note: "Chart: the numbers behind it, with the picture beside them" }; }
                    else built[i] = { model: null, pic: await picture(p.el), note: "Visual (picture only)" };
                } catch (e) { console.warn("[Export] skipped " + p.name + ":", e); built[i] = { skipped: (e && e.message) || String(e) }; }
            }
            await step(0.85, "Writing the workbook…");
            if (job.kind === "tab") aboutSheet(wb, found.title, contents);
            parts.forEach((p, i) => {
                const b = built[i];
                if (!b || b.skipped) { contents.push(["(not exported)", p.name + (b ? " — " + b.skipped : "")]); return; }
                contents.push([addSheet(wb, p.name, b.model, b.pic), p.name + " — " + b.note]);
            });
            if (job.kind === "tab") { const ws = wb.getWorksheet("About"); contents.forEach(c => ws.addRow(c)); }
            if (!wb.worksheets.length) { alert("Nothing could be exported from this pane."); return; }
            await writeBook(wb, fileBase(found.title, job.kind === "tab" ? "" : job.name));
            if (card) { setBusyDetail("Done"); await finishBusy(); }
        } catch (e) {
            alert("Export failed: " + ((e && e.message) || e));
        } finally {
            if (card) hideBusy();
        }
    }

    // ---------- menu ----------
    const btn = byId("exportBtn"), menu = byId("exportMenu");
    let marked = null, hoverTimer = null;
    function mark(el) {
        if (marked) { marked.style.outline = marked.dataset.exOutline || ""; marked.style.outlineOffset = ""; delete marked.dataset.exOutline; marked = null; }
        clearTimeout(hoverTimer);
        if (!el) return;
        const box = el.closest(".panel, .card, .kpi-row, .ds-top") || el;
        box.dataset.exOutline = box.style.outline || ""; box.style.outline = "3px solid #ff9800"; box.style.outlineOffset = "2px"; marked = box;
        hoverTimer = setTimeout(() => { const r = box.getBoundingClientRect(); if (r.top < 70 || r.bottom > window.innerHeight) box.scrollIntoView({ block: "center", behavior: "smooth" }); }, 350);
    }
    function closeMenu() { if (menu) menu.style.display = "none"; mark(null); }
    function openMenu() {
        const found = collect();
        const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
        const jobs = [];
        const row = (label, job, extra) => { jobs.push(job); return '<div class="ex-row" data-job="' + (jobs.length - 1) + '"><span class="ex-name">' + esc(label) + "</span>" + (extra || '<span class="ex-act">Excel</span>') + "</div>"; };
        const head = t => '<div class="ex-head">' + esc(t) + "</div>";
        let html = '<div class="ex-title">Export from ' + esc(found.title) + "</div>";
        const total = found.tiles.length + found.tables.length + found.charts.length + found.pictures.length;
        if (!total) html += '<div class="ex-empty">Nothing to export on this tab yet. Load the figures first.</div>';
        else {
            html += row("Entire tab — every pane on its own sheet", { kind: "tab", as: "xlsx", name: found.title });
            if (found.tiles.length) html += head("Numbers") + row("Numbers (tiles)", { kind: "tiles", as: "xlsx", name: "Numbers (tiles)", el: found.tiles[0] });
            if (found.tables.length) html += head("Tables") + found.tables.map(t => row(t.name, { kind: "table", as: "xlsx", name: t.name, el: t.el })).join("");
            if (found.charts.length) html += head("Charts") + found.charts.map(c => {
                jobs.push({ kind: "chart", as: "png", name: c.name, el: c.el });
                return row(c.name, { kind: "chart", as: "xlsx", name: c.name, el: c.el }, '<span class="ex-act">Excel</span><span class="ex-act" data-job="' + (jobs.length - 1) + '">Picture</span>');
            }).join("");
            if (found.pictures.length) html += head("Visuals (picture only)") + found.pictures.map(p => row(p.name, { kind: "picture", as: "png", name: p.name, el: p.el }, '<span class="ex-act">Picture</span>')).join("");
        }
        menu.innerHTML = html;
        const r = btn.getBoundingClientRect();
        menu.style.left = Math.max(8, Math.min(r.left, window.innerWidth - 440)) + "px"; menu.style.top = (r.bottom + 6) + "px";
        menu.style.maxHeight = Math.max(200, window.innerHeight - r.bottom - 24) + "px";
        menu.style.display = "block";
        menu.onmouseover = e => { const el = e.target.closest(".ex-row"); mark(el && jobs[+el.dataset.job] ? jobs[+el.dataset.job].el : null); };
        menu.onmouseleave = () => mark(null);
        menu.onclick = e => {
            const act = e.target.closest(".ex-act[data-job]"), el = act || e.target.closest(".ex-row");
            if (!el) return;
            const job = jobs[+el.dataset.job];
            closeMenu();
            if (job) run(job);
        };
    }
    if (btn && menu) {
        btn.addEventListener("click", e => { e.stopPropagation(); if (!allowed()) return; if (menu.style.display === "block") closeMenu(); else openMenu(); });
        document.addEventListener("click", e => { if (menu.style.display === "block" && !menu.contains(e.target)) closeMenu(); });
        document.addEventListener("keydown", e => { if (e.key === "Escape") closeMenu(); });
        window.addEventListener("resize", closeMenu);
    }
    // Pop-ups cover the header, so each gets its own button.
    document.querySelectorAll(".modal-overlay .modal-header").forEach(h => {
        const b = document.createElement("button");
        b.type = "button"; b.className = "more-btn exportUi"; b.textContent = "Export"; b.style.display = "none"; b.style.marginLeft = "auto";
        b.title = "Export this list to Excel";
        b.addEventListener("click", e => {
            e.stopPropagation();
            const found = collect();
            const parts = found.tables.length + found.charts.length;
            if (parts === 1 && found.tables.length) run({ kind: "table", as: "xlsx", name: found.tables[0].name, el: found.tables[0].el });
            else run({ kind: "tab", as: "xlsx", name: found.title });
        });
        const close = h.querySelector(".modal-close, [id$='Close'], button:last-child");
        if (close && close.parentNode === h) h.insertBefore(b, close); else h.appendChild(b);
    });
    if (typeof applyExportAccess === "function") applyExportAccess();
    window.__exportTest = { collect, run, scrapeTable, tableModel, chartModel, parseValue };
})();
