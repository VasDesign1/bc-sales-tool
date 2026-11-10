// ============================================================
// summary-mock.js — MOCK-UP emails for Faik (2026-11-11): what a
// weekly and a monthly sales summary could look like next to the daily
// one. Sends three test emails to DAILY_EMAIL_TO, built from a month
// that is already closed:
//   1. the Daily Snapshot as it stood at the close of that month
//   2. a weekly summary for one week of it (default: the 2nd)
//   3. the monthly summary
// Run by hand only (workflow summary-mock.yml). Nothing here is on a
// schedule and the live morning email (daily-email.js) is not touched.
//
// Like the daily email, the figures come from the tool's own page in
// headless Chrome: the newest snapshot is poured in, the page's clock is
// wound back to the date wanted, and the page's own functions
// (mtdTotals, lineNetAmount, the targets, the selling-day calendar) do
// the sums. The weekly and monthly blocks are drawn in the page with the
// tool's own styles and photographed, so they look like the tool.
// What cannot be wound back: Shipped Not Invoiced (the snapshot only
// holds today's open orders) and invoices back-dated after the event.
//
// Env: as daily-email.js, plus
//   MOCK_MONTH  "YYYY-MM"   month to show (default: last month)
//   MOCK_WEEK   1..6        which week of it (default: 2)
// Local layout test (no sign-in, nothing sent):
//   node scripts/summary-mock.js --local <payload.json> <targets.json> <out folder>
// ============================================================
"use strict";
const fs = require("fs"), os = require("os"), path = require("path"), http = require("http"), { spawn } = require("child_process");
const E = require("./daily-email.js");

const REPO = path.join(__dirname, "..");
const THEME = (process.env.EMAIL_THEME || "linen").toLowerCase().replace(/[^a-z]/g, "") || "linen";
const SHOT_WIDTH = 1560, SHOW_WIDTH = 1200;       // as daily-email.js
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ---------- in the page ----------
// A clock that can be set: `new Date()` and Date.now() answer with the
// pretend moment while one is set; dates built from values are untouched.
const FAKE_CLOCK = `(() => {
    const D = Date; let fake = null;
    class F extends D { constructor(...a) { if (a.length) super(...a); else super(fake == null ? D.now() : fake); } static now() { return fake == null ? D.now() : fake; } }
    window.Date = F;
    window.__setNow = (y, m, d, h) => { fake = y == null ? null : new D(y, m - 1, d, h == null ? 12 : h).getTime(); };
})();`;
// Same loading steps as daily-email.js: snapshot in, targets workbooks through the page's own parser.
const PAGE_SETUP = `(async () => {
    const overlay = document.getElementById("signinOverlay"); if (overlay) overlay.style.display = "none";
    window.alert = m => console.error("ALERT " + m);
    const payload = await (await fetch("/__email/payload.json")).json();
    hydrateSnapshot(payload);
    const t = await (await fetch("/__email/targets.json")).json();
    if (t.direct) Object.assign(spTargets, t.direct);
    else {
        spTargets.all = []; spTargets.skipped = [];
        for (const f of t.files) {
            try {
                const parsed = parseTargetsWorkbook(XLSX.read(await (await fetch("/__email/" + f.file)).arrayBuffer(), { type: "array" }));
                if (parsed.blocks) spTargets.all.push({ name: f.name, modified: f.modified, months: [...new Set([...parsed.months, ...parsed.branch.months])].sort(), parsed });
                else spTargets.skipped.push(f.name);
            } catch (e) { spTargets.skipped.push(f.name); }
        }
        applyTargetsSelection();
    }
    return JSON.stringify({ invoices: state.invoices.length, asOf: state.dataAsOf || "", targetFiles: (spTargets.files || []).length });
})()`;

// Everything below runs inside the tool's page (it is turned into text and
// evaluated there), so it may use the page's functions and nothing of Node.
function pageLib(cfg) {
    const M = window.__m = { orig: state.invoices };
    const WD = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], WDL = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const MNL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    const D = s => { const a = s.split("-").map(Number); return new Date(a[0], a[1] - 1, a[2]); };
    const addDays = (s, n) => { const d = D(s); d.setDate(d.getDate() + n); return localISO(d); };
    const monthEnd = mk => localISO(new Date(+mk.slice(0, 4), +mk.slice(5), 0));
    const prevMk = mk => localISO(new Date(+mk.slice(0, 4), +mk.slice(5) - 2, 1)).slice(0, 7);
    const monthLbl = mk => MNL[+mk.slice(5) - 1] + " " + mk.slice(0, 4);
    const dLbl = s => { const d = D(s); return WD[d.getDay()] + " " + d.getDate() + " " + MN[d.getMonth()]; };
    const longLbl = s => { const d = D(s); return WDL[d.getDay()] + " " + d.getDate() + " " + MNL[d.getMonth()] + " " + d.getFullYear(); };
    const sell = (a, b) => sellingDaysInRange(D(a), D(b));
    const sdCache = {}, sdTot = mk => sdCache[mk] || (sdCache[mk] = sell(mk + "-01", monthEnd(mk)));
    const nextSell = s => { let d = addDays(s, 1); while (!isSellingDay(D(d))) d = addDays(d, 1); return d; };
    // "Mon 5 – Fri 9 Oct": first to last selling day inside a window
    const span = (a, b) => {
        let f = a, l = b;
        while (f < b && !isSellingDay(D(f))) f = addDays(f, 1);
        while (l > f && !isSellingDay(D(l))) l = addDays(l, -1);
        return f === l ? dLbl(f) : (f.slice(0, 7) === l.slice(0, 7) ? WD[D(f).getDay()] + " " + D(f).getDate() : dLbl(f)) + " – " + dLbl(l);
    };

    // ---- winding the page back ----
    // The clock goes to the end of `day`, and invoices dated after `cap` are put aside.
    M.asAt = (day, cap) => {
        const a = day.split("-").map(Number); window.__setNow(a[0], a[1], a[2], 23);
        state.invoices = cap ? M.orig.filter(i => String(i.postingDate || "").slice(0, 10) <= cap) : M.orig;
    };
    M.real = () => { window.__setNow(null); state.invoices = M.orig; };
    // Sales, gross profit and invoice count per day, from the Month-to-Date
    // tab's own function with the clock on the last day wanted.
    const dayRows = last => { M.asAt(last, last); const m = new Map(); for (const r of mtdTotals().daily) m.set(r.day, { sales: r.sales, gp: r.gp, inv: r.invCount }); return m; };
    const book = asAt => { const cache = {}; return day => { if (day > asAt) return null; const mk = day.slice(0, 7); if (!cache[mk]) cache[mk] = dayRows(mk === asAt.slice(0, 7) ? asAt : monthEnd(mk)); return cache[mk].get(day) || null; }; };
    const sumWin = (get, from, to) => {
        let sales = 0, gp = 0, inv = 0;
        for (let s = from; s <= to; s = addDays(s, 1)) { const r = get(s); if (r) { sales += r.sales; gp += r.gp; inv += r.inv; } }
        const days = sell(from, to);
        return { from, to, label: span(from, to), days, sales, gp, gpPct: sales ? gp / sales : 0, inv, perDay: days ? sales / days : 0 };
    };

    // ---- targets for any stretch of days: each selling day carries its month's target ÷ that month's selling days ----
    const tgt = (entity, code, from, to) => {
        let t = 0;
        for (let s = from; s <= to; s = addDays(s, 1)) { if (!isSellingDay(D(s))) continue; const mk = s.slice(0, 7), n = sdTot(mk); t += n ? (salespersonBudgetTarget("sales", mk, code, entity) || 0) / n : 0; }
        return t;
    };
    const codesOf = entity => Object.keys(spBudgetTable("sales", entity));
    const tgtAll = (from, to) => codesOf("branch").reduce((s, c) => s + tgt("branch", c, from, to), 0);

    // ---- sales by branch / salesperson and kit units for a stretch of days (same sums as the run-rate tables) ----
    const kitCodes = new Map(); DS_KIT_TARGETS.forEach(g => g.items.forEach(c => kitCodes.set(c.toUpperCase(), c)));
    const ent = (from, to) => {
        const br = new Map(), sp = new Map(), kits = {};
        for (const inv of M.orig) {
            const d = String(inv.postingDate || "").slice(0, 10);
            if (d < from || d > to) continue;
            const s = inv.salesperson || "(unassigned)";
            sp.set(s, (sp.get(s) || 0) + invoiceSalesAmount(inv));
            for (const ln of (inv.salesInvoiceLines || [])) {
                const b = locationCode(ln.locationId);
                br.set(b, (br.get(b) || 0) + lineNetAmount(ln));
                const c = kitCodes.get(String(ln.lineObjectNumber || "").trim().toUpperCase());
                if (c) kits[c] = (kits[c] || 0) + num(ln.quantity);
            }
        }
        return { br, sp, kits };
    };
    const rowsFor = (entity, cur, prev, from, to) => {
        const names = spBudgetNames(entity), key = entity === "branch" ? "br" : "sp", a = cur[key], b = prev[key];
        const rows = [...new Set([...a.keys(), ...b.keys(), ...codesOf(entity)])].map(code => {
            const sales = a.get(code) || 0, pv = b.get(code) || 0, target = tgt(entity, code, from, to);
            return { code, name: names[code] || "", sales, prev: pv, target, pct: target ? sales / target : null };
        }).filter(r => Math.abs(r.sales) > 0.005 || Math.abs(r.prev) > 0.005 || r.target).sort((x, y) => y.sales - x.sales);
        const total = rows.reduce((t, r) => { t.sales += r.sales; t.prev += r.prev; t.target += r.target; return t; }, { code: "Total", name: "", sales: 0, prev: 0, target: 0 });
        total.pct = total.target ? total.sales / total.target : null;
        return { rows, total };
    };
    const kitTgt = (g, from, to) => { let t = 0; for (let s = from; s <= to; s = addDays(s, 1)) if (isSellingDay(D(s))) t += g.target / sdTot(s.slice(0, 7)); return t; };
    const kitSum = (g, e) => g.items.reduce((s, c) => s + (e.kits[c] || 0), 0);

    // ---- the month cut into weeks (Mon–Sun, clipped to the month; a weekend-only stub joins its neighbour) ----
    const mk = cfg.month, pmk = prevMk(mk), mStart = mk + "-01", mEnd = monthEnd(mk);
    let weeks = [];
    for (let s = mStart; s <= mEnd;) { const dow = D(s).getDay(); let e = addDays(s, dow === 0 ? 0 : 7 - dow); if (e > mEnd) e = mEnd; weeks.push({ from: s, to: e }); s = addDays(e, 1); }
    if (weeks.length > 1 && !sell(weeks[0].from, weeks[0].to)) { weeks[1].from = weeks[0].from; weeks.shift(); }
    const lastW = weeks.length - 1;
    if (lastW > 0 && !sell(weeks[lastW].from, weeks[lastW].to)) { weeks[lastW - 1].to = weeks[lastW].to; weeks.pop(); }
    const monthTarget = tgtAll(mStart, mEnd);
    const strip = (get, asAt) => {
        let cum = 0;
        return weeks.filter(w => w.from <= asAt).map((w, i) => {
            const s = sumWin(get, w.from, w.to > asAt ? asAt : w.to);
            s.n = i + 1; s.target = tgtAll(s.from, s.to); s.pct = s.target ? s.sales / s.target : null;
            cum += s.sales; s.cum = cum; s.cumPct = monthTarget ? cum / monthTarget : null;
            return s;
        });
    };

    // ---- WEEKLY: the full Mon–Sun week holding the chosen week of the month, against the week before ----
    const wkNo = Math.max(1, Math.min(weeks.length, cfg.week)), bucket = weeks[wkNo - 1];
    const cwFrom = addDays(bucket.from, -((D(bucket.from).getDay() + 6) % 7)), cwTo = addDays(cwFrom, 6);
    const pwFrom = addDays(cwFrom, -7), pwTo = addDays(cwFrom, -1);
    const getW = book(cwTo);
    const W = sumWin(getW, cwFrom, cwTo);
    W.n = wkNo; W.monthLabel = monthLbl(mk); W.asAt = cwTo; W.arrives = longLbl(nextSell(cwTo)); W.asAtLabel = longLbl(cwTo);
    W.prev = sumWin(getW, pwFrom, pwTo);
    W.target = tgtAll(cwFrom, cwTo); W.pct = W.target ? W.sales / W.target : null;
    {
        const mtdTo = cwTo > mEnd ? mEnd : cwTo;
        const eC = ent(cwFrom, cwTo), eP = ent(pwFrom, pwTo), eM = ent(mStart, mtdTo);
        W.branch = rowsFor("branch", eC, eP, cwFrom, cwTo);
        W.sp = rowsFor("salesperson", eC, eP, cwFrom, cwTo);
        W.kits = DS_KIT_TARGETS.map(g => ({ title: g.title, monthTarget: g.target, weekTarget: kitTgt(g, cwFrom, cwTo), qty: kitSum(g, eC), prev: kitSum(g, eP), mtd: kitSum(g, eM),
            codes: g.items.map(c => ({ code: c, qty: eC.kits[c] || 0, prev: eP.kits[c] || 0, mtd: eM.kits[c] || 0 })) }));
        W.strip = strip(getW, mtdTo);
        W.month = { target: monthTarget, mtd: W.strip.length ? W.strip[W.strip.length - 1].cum : 0, sdTotal: sdTot(mk), sdPassed: sell(mStart, mtdTo) };
        W.month.pct = monthTarget ? W.month.mtd / monthTarget : null;
    }

    // ---- MONTHLY: the closed month against its target and the month before ----
    const getM = book(mEnd);
    const Mo = sumWin(getM, mStart, mEnd);
    Mo.monthLabel = monthLbl(mk); Mo.prevLabel = monthLbl(pmk); Mo.asAt = mEnd; Mo.arrives = longLbl(nextSell(mEnd));
    Mo.prev = sumWin(getM, pmk + "-01", monthEnd(pmk));
    Mo.target = monthTarget; Mo.balance = Mo.sales - monthTarget; Mo.pct = monthTarget ? Mo.sales / monthTarget : null;
    {
        M.real();
        const eC = ent(mStart, mEnd), eP = ent(pmk + "-01", monthEnd(pmk));
        const closed = r => { r.status = r.target ? (r.sales >= r.target ? "TARGET MET" : "TARGET MISSED") : "NO PLAN"; return r; };
        Mo.branch = rowsFor("branch", eC, eP, mStart, mEnd); Mo.branch.rows.forEach(closed); closed(Mo.branch.total);
        Mo.sp = rowsFor("salesperson", eC, eP, mStart, mEnd); Mo.sp.rows.forEach(closed); closed(Mo.sp.total);
        // Margin by branch: the run-rate table's own per-branch margin, with cost brought to the month's headline cost (as the tool's tabs do).
        const marg = new Map(runRateRows({ metric: "margin", entity: "branch", budget: true }, mk, sellingDaysForMonthKey(mk, new Date())).map(r => [r.entity, r.cyVal]));
        const rawCost = Mo.branch.rows.reduce((s, r) => s + (r.sales - (marg.get(r.code) || 0)), 0), k = rawCost > 0 ? (Mo.sales - Mo.gp) / rawCost : 1;
        Mo.branch.rows.forEach(r => { r.gpPct = r.sales ? (r.sales - (r.sales - (marg.get(r.code) || 0)) * k) / r.sales : null; });
        Mo.branch.total.gpPct = Mo.gpPct;
        Mo.kits = DS_KIT_TARGETS.map(g => { const qty = kitSum(g, eC); return { title: g.title, target: g.target, qty, prev: kitSum(g, eP), balance: qty - g.target, pct: qty / g.target, status: qty >= g.target ? "TARGET MET" : "TARGET MISSED",
            codes: g.items.map(c => ({ code: c, qty: eC.kits[c] || 0, prev: eP.kits[c] || 0 })) }; });
        Mo.strip = strip(getM, mEnd);
        // Financial year to date (July → this month) against the budget, month by month
        const fyStart = (+mk.slice(5) >= 7 ? +mk.slice(0, 4) : +mk.slice(0, 4) - 1) + "-07";
        let fs = 0, ft = 0, n = 0, first = "";
        for (let m = fyStart; m <= mk; m = localISO(new Date(+m.slice(0, 4), +m.slice(5), 1)).slice(0, 7)) {
            const t = tgtAll(m + "-01", monthEnd(m)); if (!t) continue;
            let s = 0; ent(m + "-01", monthEnd(m)).br.forEach(v => s += v);
            fs += s; ft += t; n++; if (!first) first = m;
        }
        Mo.fy = { label: first ? MN[+first.slice(5) - 1] + " – " + MN[+mk.slice(5) - 1] + " " + mk.slice(0, 4) : "", months: n, sales: fs, target: ft, balance: fs - ft, pct: ft ? fs / ft : null };
    }
    M.real();
    M.weekly = W; M.monthly = Mo;

    // ---- drawing, with the tool's own styles ----
    const money = fmtMoneyFull, esc = escapeHtml, m$ = v => Math.abs(v) > 0.005 ? money(v) : "—";
    const sgn = v => (v >= 0 ? "+" : "−") + money(Math.abs(v));
    const sgnPct = (c, p) => p ? (c - p >= 0 ? "+" : "−") + fmtPct(Math.abs((c - p) / p)) : "";
    const tone = (v, html) => '<i style="font-style:normal; color:' + (v >= 0 ? "#2e7d32" : "#c62828") + ';">' + html + "</i>";
    const q = n => Number.isInteger(n) ? fmtInt(n) : fmtQty(n), one = n => (Math.round(n * 10) / 10).toLocaleString("en-AU", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
    const bar = p => p == null ? "—" : `<span class="rr-bar b${rrPctBand(p)}"><span class="rr-fill" style="width:${Math.max(0, Math.min(100, p * 100)).toFixed(1)}%;"></span><span class="rr-lbl">${fmtPct(p)}</span></span>`;
    const chgTd = (c, p) => Math.abs(c) < 0.005 && Math.abs(p) < 0.005 ? '<td class="num">—</td>' : `<td class="num ${c - p >= 0 ? "pos" : "neg"}">${sgn(c - p)}${p ? `<span class="rr-name">${sgnPct(c, p)}</span>` : ""}</td>`;
    const balTd = (c, t) => !t ? '<td class="num">—</td>' : `<td class="num ${c - t >= 0 ? "pos" : "neg"}">${sgn(c - t)}</td>`;
    const th = (l, n) => `<th${n ? ' class="num"' : ""}>${l}</th>`;
    const nameTd = r => `<td>${esc(r.code)}${r.name ? `<span class="rr-name">${esc(r.name)}</span>` : ""}</td>`;
    const legend = '<span class="rr-legend"><span class="b0">&lt; 50%</span><span class="b1">50–99%</span><span class="b2">100%+</span></span>';
    const panel = (title, head, body, extra) => `<div class="panel"><div class="panel-header"><h2>${title}</h2>${extra || ""}</div><div class="table-scroll"><table class="ds-table"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div></div>`;
    const tile = (cls, label, value, sub) => `<div class="kpi ${cls}"><div class="label">${label}</div><div class="value">${value}</div><div class="sub">${sub || "&nbsp;"}</div></div>`;
    const part = (label, value, sub) => `<div class="ds-part"><div class="ds-part-lbl">${label}</div><div class="ds-part-val">${value}</div><div class="sub">${sub || "&nbsp;"}</div></div>`;
    const stat = (lbl, val) => `<div class="ds-kit-stat"><span>${lbl}</span><b>${val}</b></div>`;
    const stripPanel = (title, rows, markN) => {
        const tot = rows.reduce((t, s) => { t.days += s.days; t.sales += s.sales; t.gp += s.gp; t.target += s.target; return t; }, { days: 0, sales: 0, gp: 0, target: 0 });
        return panel(title, th("Week") + th("Selling Days", 1) + th("Sales", 1) + th("A Selling Day", 1) + th("GP %", 1) + th("Week Target", 1) + th("% of Week Target", 1) + th("Month So Far", 1) + th("% of Month Target", 1),
            rows.map(s => `<tr${s.n === markN ? ' style="font-weight:700;"' : ""}><td>Week ${s.n}<span class="rr-name">${esc(s.label)}</span></td><td class="num">${s.days}</td><td class="num">${money(s.sales)}</td><td class="num">${money(s.perDay)}</td><td class="num">${fmtPct(s.gpPct)}</td><td class="num">${m$(s.target)}</td><td class="num">${bar(s.pct)}</td><td class="num">${money(s.cum)}</td><td class="num">${bar(s.cumPct)}</td></tr>`).join("")
            + `<tr class="ds-total"><td>Total</td><td class="num">${tot.days}</td><td class="num">${money(tot.sales)}</td><td class="num">${money(tot.days ? tot.sales / tot.days : 0)}</td><td class="num">${fmtPct(tot.sales ? tot.gp / tot.sales : 0)}</td><td class="num">${m$(tot.target)}</td><td class="num">${bar(tot.target ? tot.sales / tot.target : null)}</td><td class="num">${money(tot.sales)}</td><td class="num">${bar(monthTarget ? tot.sales / monthTarget : null)}</td></tr>`, legend);
    };

    const weeklyHtml = () => {
        const w = W, p = w.prev, d = w.sales - p.sales;
        const row = (r, tot) => `<tr${tot ? ' class="ds-total"' : ""}>${nameTd(r)}<td class="num">${m$(r.prev)}</td><td class="num">${tot ? money(r.sales) : m$(r.sales)}</td>${chgTd(r.sales, r.prev)}<td class="num">${m$(r.target)}</td>${balTd(r.sales, r.target)}<td class="num">${bar(r.pct)}</td></tr>`;
        const table = (title, col, t) => panel(title, th(col) + th("Week Before", 1) + th("This Week", 1) + th("Change", 1) + th("Week Target", 1) + th("Balance", 1) + th("% of Week Target", 1), t.rows.map(r => row(r)).join("") + row(t.total, true), legend);
        const kit = k => `<div class="panel"><div class="panel-header"><h2>${esc(k.title)}</h2></div>
            <div class="ds-kit-sum">${stat("Sold this week", q(k.qty))}${stat("Week before", q(k.prev))}${stat("Week's share of target", one(k.weekTarget))}${stat("Month so far", q(k.mtd))}${stat("Month target", fmtInt(k.monthTarget))}${stat("% of month target", bar(k.mtd / k.monthTarget))}</div>
            <div class="table-scroll"><table class="ds-table"><thead><tr>${th("Item code")}${th("This Week", 1)}${th("Week Before", 1)}${th("Month So Far", 1)}</tr></thead><tbody>
            ${k.codes.map(c => `<tr><td>${esc(c.code)}</td><td class="num">${q(c.qty)}</td><td class="num">${q(c.prev)}</td><td class="num">${q(c.mtd)}</td></tr>`).join("")}
            <tr class="ds-total"><td>Total</td><td class="num">${q(k.qty)}</td><td class="num">${q(k.prev)}</td><td class="num">${q(k.mtd)}</td></tr></tbody></table></div></div>`;
        return `<div class="hint rr-hint"><span><strong>Week ${w.n} of ${esc(w.monthLabel)}</strong> · ${esc(w.label)} · ${w.days} selling days. Compared with the week before, ${esc(p.label)} (${p.days} selling days).</span></div>
            <div class="ds-top">
                <div class="kpi k-rev ds-pane"><div class="label">Sales (excl. tax) · the week</div>
                    <div class="value ds-big">${money(w.sales)}<span class="ds-pct">${p.sales ? tone(d, sgnPct(w.sales, p.sales)) + " on the week before" : ""}</span></div>
                    <div class="ds-parts">${part("Week before", money(p.sales), esc(p.label))}${part("Change", tone(d, sgn(d)), fmtInt(w.inv) + " invoices · " + fmtInt(p.inv) + " the week before")}${part("A selling day", money(w.perDay), "week before " + money(p.perDay))}</div></div>
                <div class="kpi k-gp ds-pane"><div class="label">Gross Profit · the week</div>
                    <div class="value ds-big">${money(w.gp)}<span class="ds-pct">${fmtPct(w.gpPct)}</span></div><div class="sub">on invoiced sales</div>
                    <div class="ds-parts">${part("Week before", money(p.gp) + '<span class="ds-pct">' + fmtPct(p.gpPct) + "</span>")}</div></div>
            </div>
            <div class="kpi-row ds-days">
                ${tile("k-rev", "Week Target", money(w.target), "month target ÷ " + sdTot(mk) + " selling days × " + w.days)}
                ${tile("k-inv", "Balance", w.target ? tone(w.sales - w.target, sgn(w.sales - w.target)) : "—", "sales − week target")}
                ${tile("k-gpp", "% of Week Target", bar(w.pct))}
                ${tile("k-cust", esc(w.monthLabel) + " So Far", money(w.month.mtd), fmtPct(w.month.pct) + " of the month target · " + w.month.sdPassed + " of " + w.month.sdTotal + " selling days gone")}
            </div>
            ${table("The Week · by Branch", "Branch", w.branch)}
            ${table("The Week · by Salesperson", "Salesperson", w.sp)}
            <div class="ds-pair">${w.kits.map(kit).join("")}</div>
            ${stripPanel(esc(w.monthLabel) + " · week by week so far", w.strip, w.n)}`;
    };

    const monthlyHtml = () => {
        const m = Mo, p = m.prev, d = m.sales - p.sales, pl = esc(m.prevLabel.split(" ")[0]);
        const row = (r, tot, gp) => `<tr${tot ? ' class="ds-total"' : ""}>${nameTd(r)}<td class="num">${m$(r.prev)}</td><td class="num">${m$(r.target)}</td><td class="num">${tot ? money(r.sales) : m$(r.sales)}</td>${balTd(r.sales, r.target)}${chgTd(r.sales, r.prev)}${gp ? `<td class="num">${r.gpPct == null ? "—" : fmtPct(r.gpPct)}</td>` : ""}<td>${statusBadgeRunRate(r.status)}</td><td class="num">${bar(r.pct)}</td></tr>`;
        const table = (title, col, t, gp) => panel(title, th(col) + th(pl + " Sales", 1) + th("Sales Target", 1) + th("Sales For Month", 1) + th("Balance", 1) + th("Change on " + pl, 1) + (gp ? th("GP %", 1) : "") + th("Status") + th("% Achieved", 1),
            t.rows.map(r => row(r, false, gp)).join("") + row(t.total, true, gp), legend);
        const kit = k => `<div class="panel"><div class="panel-header"><h2>${esc(k.title)}</h2></div>
            <div class="ds-kit-sum">${stat("Sold in the month", q(k.qty))}${stat("Target", fmtInt(k.target))}${stat("Balance", tone(k.balance, (k.balance >= 0 ? "+" : "−") + q(Math.abs(k.balance))))}${stat(pl, q(k.prev))}${stat("Status", statusBadgeRunRate(k.status))}${stat("% Achieved", bar(k.pct))}</div>
            <div class="table-scroll"><table class="ds-table"><thead><tr>${th("Item code")}${th(pl, 1)}${th(esc(m.monthLabel.split(" ")[0]), 1)}${th("Change", 1)}</tr></thead><tbody>
            ${k.codes.map(c => `<tr><td>${esc(c.code)}</td><td class="num">${q(c.prev)}</td><td class="num">${q(c.qty)}</td><td class="num ${c.qty - c.prev >= 0 ? "pos" : "neg"}">${(c.qty - c.prev >= 0 ? "+" : "−") + q(Math.abs(c.qty - c.prev))}</td></tr>`).join("")}
            <tr class="ds-total"><td>Total</td><td class="num">${q(k.prev)}</td><td class="num">${q(k.qty)}</td><td class="num">${(k.qty - k.prev >= 0 ? "+" : "−") + q(Math.abs(k.qty - k.prev))}</td></tr></tbody></table></div></div>`;
        return `<div class="hint rr-hint"><span><strong>${esc(m.monthLabel)}</strong> · the month as it closed, ${m.days} selling days. Compared with ${esc(m.prevLabel)} (${p.days} selling days). Preliminary: late invoices, credits and cost adjustments can still move these figures.</span></div>
            <div class="ds-top">
                <div class="kpi k-rev ds-pane"><div class="label">Sales (excl. tax) · the month</div>
                    <div class="value ds-big">${money(m.sales)}<span class="ds-pct">${m.pct == null ? "" : fmtPct(m.pct) + " of target"}</span></div>
                    <div class="ds-parts">${part("Target", m$(m.target), fmtInt(m.inv) + " invoices")}${part("Balance", m.target ? tone(m.balance, sgn(m.balance)) : "—", "sales − target")}${part(pl, money(p.sales), p.sales ? tone(d, sgn(d) + " · " + sgnPct(m.sales, p.sales)) : "")}</div></div>
                <div class="kpi k-gp ds-pane"><div class="label">Gross Profit · the month</div>
                    <div class="value ds-big">${money(m.gp)}<span class="ds-pct">${fmtPct(m.gpPct)}</span></div><div class="sub">on invoiced sales · cost ${money(m.sales - m.gp)}</div>
                    <div class="ds-parts">${part(pl, money(p.gp) + '<span class="ds-pct">' + fmtPct(p.gpPct) + "</span>", tone(m.gp - p.gp, sgn(m.gp - p.gp)))}</div></div>
            </div>
            <div class="kpi-row ds-days">
                ${tile("k-rev", "Selling Days", fmtInt(m.days), pl + " had " + p.days)}
                ${tile("k-inv", "Sales a Selling Day", money(m.perDay), pl + " " + money(p.perDay))}
                ${tile("k-cust", "Financial Year So Far · " + esc(m.fy.label), money(m.fy.sales), "budget " + money(m.fy.target) + " · " + tone(m.fy.balance, sgn(m.fy.balance)))}
                ${tile("k-gpp", "% of Year-to-Date Budget", bar(m.fy.pct), m.fy.months + " months")}
            </div>
            ${table("The Month · by Branch", "Branch", m.branch, true)}
            ${table("The Month · by Salesperson", "Salesperson", m.sp, false)}
            <div class="ds-pair">${m.kits.map(kit).join("")}</div>
            ${stripPanel(esc(m.monthLabel) + " · week by week", m.strip, 0)}`;
    };

    // ---- showing one thing at a time and measuring it for the picture ----
    const daily = () => document.querySelector('section[data-tab-content="dailySnapshot"]');
    const mine = () => { let s = document.getElementById("mockSum"); if (!s) { s = document.createElement("section"); s.id = "mockSum"; daily().parentElement.appendChild(s); } return s; };
    M.render = async kind => {
        const out = {};
        if (!M.styled) {      // as daily-email.js: the tab alone, none of the tool's chrome or controls
            const st = document.createElement("style");
            st.textContent = ".sidebar, .header, .toolbar, #tabStrip, #filterChips, .filter-chips, #themePicker, .progress, #busyWrap, #exportMenu, #helpTip { display: none !important; }"
                + " .app { display: block !important; } .main { margin: 0 !important; width: auto !important; } .container { padding: 18px 20px !important; max-width: none !important; }"
                + " section[data-tab-content='dailySnapshot'] .rr-tgt, section[data-tab-content='dailySnapshot'] [data-ds-goto], #dsEmpty, #dsKits .panel-header .meta { display: none !important; }"
                + " #mockSum .ds-pct i { font-weight: 700; }";
            document.head.appendChild(st); M.styled = true;
        }
        if (kind.t === "daily") {
            M.asAt(kind.asAt, kind.asAt);
            mine().style.display = "none";
            setActiveTab("dailySnapshot"); renderDailySnapshot();
            daily().style.display = "";
            document.getElementById("dsAsOf").textContent = kind.note || "";
            out.daily = { f: dsSnapshotFigures(), names: { branch: spBudgetNames("branch"), salesperson: spBudgetNames("salesperson") }, asOf: kind.asAt + " (rewound)", today: kind.asAt };
        } else {
            M.real();
            document.querySelectorAll("section[data-tab-content]").forEach(s => s.style.display = "none");
            const s = mine(); s.innerHTML = kind.t === "weekly" ? weeklyHtml() : monthlyHtml(); s.style.display = "";
        }
        window.scrollTo(0, 0);
        if (document.fonts && document.fonts.ready) await document.fonts.ready;
        await new Promise(r => setTimeout(r, 400));
        const box = document.querySelector(".container"), c = box.getBoundingClientRect();
        let bg = "rgb(255, 255, 255)";
        for (let e = box; e; e = e.parentElement) { const b = getComputedStyle(e).backgroundColor; if (b && b !== "transparent" && b !== "rgba(0, 0, 0, 0)") { bg = b; break; } }
        out.clip = { x: 0, y: Math.max(0, Math.floor(c.top + window.scrollY)), width: Math.ceil(document.documentElement.clientWidth), height: Math.ceil(c.height) };
        out.bg = bg; out.ink = getComputedStyle(box).color;
        return JSON.stringify(out);
    };
    return JSON.stringify({ weekly: W, monthly: Mo });
}

// ---------- headless Chrome with the tool's page open ----------
async function openTool(dataDir) {
    const mime = { ".html": "text/html", ".js": "application/javascript", ".json": "application/json" };
    const srv = http.createServer((q, r) => {
        const u = decodeURIComponent(q.url.split("?")[0]);
        const f = u.startsWith("/__email/") ? path.join(dataDir, path.basename(u)) : path.join(REPO, u === "/" ? "index.html" : u);
        if (!f.startsWith(dataDir) && !f.startsWith(REPO)) { r.writeHead(403); return r.end(); }
        fs.stat(f, (e, st) => {
            if (e || !st.isFile()) { r.writeHead(404); return r.end(); }
            r.writeHead(200, { "Content-Type": mime[path.extname(f)] || "application/octet-stream", "Content-Length": st.size });
            fs.createReadStream(f).pipe(r);
        });
    });
    await new Promise(r => srv.listen(0, "127.0.0.1", r));
    const port = srv.address().port, dbg = 9400 + Math.floor(Math.random() * 400);
    const chromePath = process.env.CHROME_PATH || (process.platform === "win32" ? "C:/Program Files/Google/Chrome/Application/chrome.exe" : "google-chrome");
    const chrome = spawn(chromePath, ["--headless=new", "--no-sandbox", "--disable-gpu", "--remote-debugging-port=" + dbg,
        "--user-data-dir=" + path.join(dataDir, "chrome"), "--window-size=" + SHOT_WIDTH + ",1200", "--hide-scrollbars", "about:blank"], { stdio: "ignore" });
    const close = () => { try { chrome.kill(); } catch (e) { /* gone */ } srv.close(); };
    try {
        let tabs;
        for (let i = 0; i < 80 && !tabs; i++) { try { tabs = await (await fetch("http://127.0.0.1:" + dbg + "/json")).json(); } catch (e) { await sleep(250); } }
        if (!tabs) throw new Error("headless Chrome did not start (" + chromePath + ")");
        const ws = new WebSocket(tabs.find(t => t.type === "page").webSocketDebuggerUrl);
        await new Promise((res, rej) => { ws.addEventListener("open", res); ws.addEventListener("error", () => rej(new Error("cannot reach headless Chrome"))); });
        let id = 0; const pend = new Map(), pageErrors = [];
        ws.addEventListener("message", ev => {
            const m = JSON.parse(ev.data);
            if (m.id && pend.has(m.id)) { pend.get(m.id)(m); pend.delete(m.id); }
            if (m.method === "Runtime.exceptionThrown") pageErrors.push(String((m.params.exceptionDetails.exception || {}).description || m.params.exceptionDetails.text).slice(0, 300));
            if (m.method === "Runtime.consoleAPICalled" && m.params.type === "error") pageErrors.push(m.params.args.map(a => a.value || a.description).join(" ").slice(0, 300));
        });
        const send = (method, params) => new Promise(r => { const i = ++id; pend.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
        // Runs an expression in the page and hands back the text it returns.
        const ev = async (expression, what) => {
            const r = await send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
            if (!r.result || !r.result.result || typeof r.result.result.value !== "string") {
                const d = r.result && r.result.exceptionDetails;
                throw new Error("the page could not " + what + ": " + String((d && ((d.exception && d.exception.description) || d.text)) || JSON.stringify(r).slice(0, 300)).slice(0, 500) + (pageErrors.length ? " | page errors: " + pageErrors.slice(0, 4).join(" || ") : ""));
            }
            return r.result.result.value;
        };
        const shoot = async (clip, dsf) => {
            await send("Emulation.setDeviceMetricsOverride", { width: SHOT_WIDTH, height: 1200, deviceScaleFactor: dsf, mobile: false });
            await sleep(300);
            const shot = await send("Page.captureScreenshot", { format: "png", captureBeyondViewport: true, clip: Object.assign({ scale: 1 }, clip) });
            if (!shot.result || !shot.result.data) throw new Error("screenshot failed: " + JSON.stringify(shot).slice(0, 300));
            return Buffer.from(shot.result.data, "base64");
        };
        await send("Runtime.enable"); await send("Page.enable");
        await send("Emulation.setTimezoneOverride", { timezoneId: "Australia/Melbourne" });
        await send("Page.addScriptToEvaluateOnNewDocument", { source: FAKE_CLOCK + "try{localStorage.setItem('bcSalesTheme','" + THEME + "');}catch(e){}" });
        await send("Emulation.setDeviceMetricsOverride", { width: SHOT_WIDTH, height: 1200, deviceScaleFactor: 2, mobile: false });
        await send("Page.navigate", { url: "http://127.0.0.1:" + port + "/index.html" });
        for (let i = 0; i < 120; i++) {
            const r = await send("Runtime.evaluate", { expression: "typeof dsSnapshotFigures === 'function' && typeof hydrateSnapshot === 'function' && typeof XLSX !== 'undefined' && typeof Chart !== 'undefined'", returnByValue: true });
            if (r.result && r.result.result && r.result.result.value === true) break;
            if (i === 119) throw new Error("the tool page did not finish loading in headless Chrome");
            await sleep(500);
        }
        return { ev, shoot, close };
    } catch (e) { close(); throw e; }
}

// ---------- AI dot points (optional, as in the daily email) ----------
const RULES = [
    "Use only the figures provided. Quote them as given, rounded sensibly (for example $401k, 19% of target); do not derive new totals or forecasts of your own. If something needed is missing, leave it out rather than guessing.",
    "In the figures: sales = invoiced sales excluding tax; prev = the same figure for the period before; target = the sales target for the period; pct and every field ending in Pct are percentages already; perDay = sales per selling day; days = selling days; cum = month to date; codes under kits are item codes with units sold.",
    "This is read by the whole team, so name a salesperson only to recognise a good result; describe shortfalls at branch or team level, not by person.",
    "Australian English, matter-of-fact and encouraging without hype. Output format: each dot point on its own line starting with a hyphen and a space, and nothing else: no greeting, no intro line, no sign-off, no headings, no emojis. Each point is one plain sentence of at most about 25 words, and no two points make the same observation.",
    "So the points can be skimmed, put the one or two things that matter most in each point in bold by wrapping them in double asterisks, like **$401k** or **Dandenong** or **19% of target**. Bold short phrases only, never a whole sentence, and use no other markdown.",
];
const WEEKLY_SYSTEM = ["You write the short opening dot points of a weekly sales email at an Australian HVAC wholesaler. It goes to internal staff on Monday morning, above a picture of last week's figures.",
    "You are given last week's figures as JSON: the week's sales, gross profit and target (the week's share of the monthly target), the same for the week before, each branch and each salesperson, two product quantity targets (AirTouch kits and Omniflow units), the month's weeks so far, and the month-to-date position.",
    "Write 5 to 7 dot points: how the week went against its target and against the week before; margin; which branch moved most; a salesperson result worth recognising; the kit and Omniflow units; and where the week leaves the month against its target."].concat(RULES).join("\n\n");
const MONTHLY_SYSTEM = ["You write the short opening dot points of a monthly sales email at an Australian HVAC wholesaler. It goes to internal staff on the first working day of the new month, above a picture of the closed month's figures.",
    "You are given the closed month's figures as JSON: sales, gross profit and target, the same for the month before, each branch and each salesperson with target met or missed, two product quantity targets (AirTouch kits and Omniflow units), the month week by week, and the financial year to date against budget (fy).",
    "Write 7 to 9 dot points: the month's sales against target; against the month before; gross profit and margin; each branch in turn; salespeople who met target; the kit and Omniflow targets; the strongest and weakest week; and the financial year to date against budget."].concat(RULES).join("\n\n");
// Rounds everything for the model: dollars to whole numbers, ratios to percentages with one decimal.
function slim(v, key) {
    if (Array.isArray(v)) return v.map(x => slim(x));
    if (v && typeof v === "object") { const o = {}; for (const k in v) if (!/^(from|to|asAt|arrives|asAtLabel)$/.test(k)) o[k] = slim(v[k], k); return o; }
    if (typeof v === "number") return /pct$/i.test(key || "") ? +(v * 100).toFixed(1) : /^(qty|prev|mtd|weekTarget)$/.test(key || "") && Math.abs(v) < 1000 ? +v.toFixed(1) : Math.round(v);
    return v;
}
async function ask(system, figures, what) {
    if (!process.env.ANTHROPIC_API_KEY) return "";
    try {
        const mod = require("@anthropic-ai/sdk"), Anthropic = mod.default || mod;
        const response = await new Anthropic().beta.messages.create({
            model: "claude-opus-5-5", max_tokens: 8000,
            betas: ["server-side-fallback-2026-07-01"], fallbacks: "default",      // as daily-email.js
            system, messages: [{ role: "user", content: "The figures:\n\n" + JSON.stringify(slim(figures), null, 1) }],
        });
        if (response.stop_reason === "refusal") { console.warn("Commentary (" + what + "): declined, email goes without it."); return ""; }
        const text = response.content.filter(b => b.type === "text").map(b => b.text).join("\n").trim();
        console.log("Commentary (" + what + "): " + text.length + " characters · tokens in " + response.usage.input_tokens + " / out " + response.usage.output_tokens);
        return text;
    } catch (e) { console.warn("Commentary (" + what + "): skipped (" + String((e && e.message) || e).slice(0, 200) + ")"); return ""; }
}

// ---------- the emails ----------
// Same frame as the daily email: header and dot points on a pale card, then the picture(s), colours taken from the tool's page.
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const money0 = n => "$" + Math.round(n).toLocaleString("en-AU");
const pct0 = p => p == null ? "" : " (" + (p * 100).toFixed(0) + "% of target)";
const hex = (rgb, fallback) => { const m = /(\d+)\D+(\d+)\D+(\d+)/.exec(rgb || ""); return m ? "#" + [1, 2, 3].map(i => (+m[i]).toString(16).padStart(2, "0")).join("") : fallback; };
function emailHtml(mail, src) {
    const mix = (a, b, t) => "#" + [1, 3, 5].map(i => Math.round(parseInt(a.slice(i, i + 2), 16) * (1 - t) + parseInt(b.slice(i, i + 2), 16) * t).toString(16).padStart(2, "0")).join("");
    const page = hex(mail.pics[0].bg, "#ffffff"), ink = hex(mail.pics[0].ink, "#1f2430");
    const light = [1, 3, 5].reduce((s, i) => s + parseInt(page.slice(i, i + 2), 16), 0) / 3 > 128;
    const outer = mix(page, ink, 0.16), card = light ? mix(page, "#ffffff", 0.7) : mix(page, ink, 0.07), line = mix(page, ink, 0.28);
    const FONT = "font-family:Calibri,Segoe UI,Arial,sans-serif;font-size:11pt;color:" + ink + ";";
    const points = mail.points ? mail.points.split(/\n+/).map(p => p.replace(/^\s*(?:[-•–]|\*(?!\*)|\d+[.)])\s*/, "").trim()).filter(Boolean) : [];
    const rich = p => esc(p).replace(/\*\*([^*]+)\*\*/g, '<strong style="font-weight:700;">$1</strong>').replace(/\*/g, "");
    const list = points.length ? '<ul style="' + FONT + 'line-height:1.5;margin:0;padding:0 0 0 26px;">' + points.map(p => '<li style="' + FONT + 'margin:0 0 6px;">' + rich(p) + "</li>").join("") + "</ul>" : "";
    const gap = h => '<tr><td height="' + h + '" style="height:' + h + 'px;font-size:1px;line-height:1px;">&nbsp;</td></tr>';
    const cardRow = inner => '<tr><td align="left" bgcolor="' + card + '" style="background-color:' + card + ';border:1px solid ' + line + ';padding:20px 24px 18px;">' + inner + "</td></tr>";
    return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>'
        + '<body bgcolor="' + outer + '" style="margin:0;padding:0;background-color:' + outer + ';">'
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="' + outer + '" style="background-color:' + outer + ';"><tr><td align="center" style="padding:22px 14px 28px;">'
        + '<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" width="' + SHOW_WIDTH + '" style="width:' + SHOW_WIDTH + 'px;max-width:100%;margin:0 auto;">'
        + cardRow('<div style="' + FONT + 'font-size:14pt;font-weight:700;letter-spacing:.3px;">' + esc(mail.header) + "</div>"
            + '<div style="' + FONT + 'font-size:10pt;padding-top:3px;">' + esc(mail.sub) + "</div>"
            + (list ? '<div style="height:16px;line-height:16px;font-size:1px;">&nbsp;</div>' + list : ""))
        + mail.pics.map((p, i) => gap(20)
            + (p.caption ? '<tr><td align="left" bgcolor="' + card + '" style="background-color:' + card + ';border:1px solid ' + line + ';border-bottom:0;padding:10px 24px;"><div style="' + FONT + 'font-weight:700;">' + esc(p.caption) + "</div></td></tr>" : "")
            + '<tr><td align="left" bgcolor="' + page + '" style="background-color:' + page + ';border:1px solid ' + line + ';"><img src="' + src(i) + '" width="' + SHOW_WIDTH + '" alt="' + esc(p.alt) + '" style="display:block;width:100%;max-width:' + SHOW_WIDTH + 'px;height:auto;border:0;margin:0 auto;"></td></tr>').join("")
        + "</table></td></tr></table></body></html>";
}
async function sendMail(mail, to) {
    const message = {
        subject: mail.subject,
        body: { contentType: "HTML", content: emailHtml(mail, i => "cid:mockpic" + i) },
        toRecipients: to.map(a => ({ emailAddress: { address: a } })),
        attachments: mail.pics.map((p, i) => ({ "@odata.type": "#microsoft.graph.fileAttachment", name: mail.key + "-" + (i + 1) + ".png", contentType: "image/png", contentBytes: p.png.toString("base64"), contentId: "mockpic" + i, isInline: true })),
    };
    const resp = await E.graph("/me/sendMail", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ message, saveToSentItems: true }) });
    if (resp.status !== 202) throw new Error("sendMail (" + mail.key + ") → HTTP " + resp.status + ": " + (await resp.text()).slice(0, 400));
}

if (require.main === module) (async () => {
    const local = process.argv.indexOf("--local");
    let meta = null;
    if (local < 0) {
        for (const k of ["BC_TENANT", "BC_CLIENT_ID", "BC_REFRESH_TOKEN", "SNAPSHOT_PASSPHRASE"]) if (!process.env[k]) throw new Error("Missing env " + k);
        meta = (await E.chooseSnapshot()).meta;          // run by hand: the newest snapshot
    }
    // Default month: the one before today's, Melbourne time.
    const melb = new Intl.DateTimeFormat("en-CA", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit" }).format(new Date()).slice(0, 7);
    const lastMonth = new Date(Date.UTC(+melb.slice(0, 4), +melb.slice(5) - 2, 1)).toISOString().slice(0, 7);
    const cfg = { month: /^\d{4}-\d{2}$/.test(process.env.MOCK_MONTH || "") ? process.env.MOCK_MONTH : lastMonth, week: parseInt(process.env.MOCK_WEEK, 10) || 2 };
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "summary-mock-"));
    let tool = null;
    try {
        if (local >= 0) {
            fs.copyFileSync(process.argv[local + 1], path.join(dir, "payload.json"));
            fs.copyFileSync(process.argv[local + 2], path.join(dir, "targets.json"));
        } else await E.collectFromSharePoint(dir, meta);
        tool = await openTool(dir);
        const setup = JSON.parse(await tool.ev(PAGE_SETUP, "load the snapshot"));
        const figs = JSON.parse(await tool.ev("(" + pageLib.toString() + ")(" + JSON.stringify(cfg) + ")", "work out the figures"));
        const W = figs.weekly, Mo = figs.monthly;
        console.log("Mock-up for " + cfg.month + ", week " + W.n + ": " + setup.invoices + " invoices in the snapshot · " + setup.targetFiles + " targets file(s) · "
            + W.branch.rows.length + "/" + Mo.branch.rows.length + " branch rows · " + W.sp.rows.length + "/" + Mo.sp.rows.length + " salesperson rows · " + Mo.strip.length + " weeks");
        // Takes the pictures an email needs; steps the sharpness down if together they are too heavy for one Graph call (~3 MB).
        const pictures = async kinds => {
            for (const dsf of [2, 1.5, 1]) {
                const out = [];
                for (const k of kinds) { const r = JSON.parse(await tool.ev("window.__m.render(" + JSON.stringify(k) + ")", "draw the " + k.t + " block")); r.png = await tool.shoot(r.clip, dsf); out.push(Object.assign(r, k)); }
                if (dsf === 1 || out.reduce((s, p) => s + p.png.length, 0) <= 2.6 * 1048576) return out;
            }
        };
        const snapLbl = meta ? "the " + (E.SLOT_LABEL[meta.slot] || meta.slot) + " snapshot taken " + String(meta.fetchedAtMelbourne || "") : "test data";
        const rewound = day => "Mock-up: figures rewound to the close of " + day + ". Shipped Not Invoiced is today's and cannot be rewound.";
        const mails = [];
        {   // 1. the daily email as it stood when the month closed
            const pics = await pictures([{ t: "daily", asAt: Mo.asAt, note: rewound(Mo.to.slice(8) + "/" + Mo.to.slice(5, 7) + "/" + Mo.to.slice(0, 4)), alt: "Daily Snapshot at the close of " + Mo.monthLabel }]);
            const f = pics[0].daily.f, tot = f.branch.total;
            mails.push({ key: "daily", pics, points: await E.writeCommentary(pics[0].daily),
                subject: "[MOCK-UP] Daily Snapshot · close of " + Mo.monthLabel + " · sales " + money0(f.sales) + pct0(tot.pctAchieved),
                header: "DAILY SALES SUMMARY · MOCK-UP · CLOSE OF " + Mo.monthLabel.toUpperCase(),
                sub: "Test only. " + Mo.monthLabel + " laid out exactly as the daily email, rewound to the last day of the month from " + snapLbl + "." });
        }
        {   // 2. the weekly summary, with the usual month-to-date picture under it
            const pics = await pictures([{ t: "weekly", alt: "Weekly summary, week " + W.n + " of " + W.monthLabel },
                { t: "daily", asAt: W.asAt, note: rewound(W.asAtLabel), alt: "Daily Snapshot after that week", caption: W.monthLabel.toUpperCase() + " SO FAR · THE USUAL DAILY PICTURE, AS IT STOOD AFTER THAT WEEK" }]);
            mails.push({ key: "weekly", pics, points: await ask(WEEKLY_SYSTEM, W, "weekly"),
                subject: "[MOCK-UP] Weekly Summary · week " + W.n + " of " + W.monthLabel + " · " + W.label + " · sales " + money0(W.sales) + pct0(W.pct).replace("of target", "of week target"),
                header: "WEEKLY SALES SUMMARY · MOCK-UP · WEEK " + W.n + " OF " + W.monthLabel.toUpperCase() + " · " + W.label.toUpperCase(),
                sub: "Test only. As it would have arrived on " + W.arrives + ": the week on top, then the usual month-to-date picture underneath." });
        }
        {   // 3. the monthly summary
            const pics = await pictures([{ t: "monthly", alt: "Monthly summary for " + Mo.monthLabel }]);
            mails.push({ key: "monthly", pics, points: await ask(MONTHLY_SYSTEM, Mo, "monthly"),
                subject: "[MOCK-UP] Monthly Summary · " + Mo.monthLabel + " · sales " + money0(Mo.sales) + pct0(Mo.pct),
                header: "MONTHLY SALES SUMMARY · MOCK-UP · " + Mo.monthLabel.toUpperCase() + " · PRELIMINARY",
                sub: "Test only. As it would have arrived on " + Mo.arrives + ", the first working day of the new month." });
        }
        for (const m of mails) console.log("  " + m.key + ": " + m.pics.length + " picture(s), " + (m.pics.reduce((s, p) => s + p.png.length, 0) / 1024).toFixed(0) + " KB" + (m.points ? ", with dot points" : ", no dot points"));
        if (local >= 0) {
            const out = process.argv[local + 3]; fs.mkdirSync(out, { recursive: true });
            for (const m of mails) {
                m.pics.forEach((p, i) => fs.writeFileSync(path.join(out, m.key + "-" + (i + 1) + ".png"), p.png));
                fs.writeFileSync(path.join(out, m.key + ".html"), "<!-- " + m.subject + " -->\n" + emailHtml(m, i => m.key + "-" + (i + 1) + ".png"));
            }
            fs.writeFileSync(path.join(out, "figures.json"), JSON.stringify(figs, null, 1));
            console.log("Wrote " + mails.length + " emails to " + out);
            return;
        }
        const to = (process.env.DAILY_EMAIL_TO || "").split(/[,;\s]+/).filter(Boolean);
        if (!to.length) throw new Error("DAILY_EMAIL_TO is empty — no recipients");
        if (to.length > 1 && process.env.MOCK_ALLOW_MANY !== "1") throw new Error("DAILY_EMAIL_TO now holds " + to.length + " addresses — the mock-ups are for Faik only, nothing sent");
        for (const m of mails) { await sendMail(m, to); console.log("Sent the " + m.key + " mock-up."); }
    } finally {
        if (tool) tool.close();
        try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) { /* temp dir */ }
    }
})().catch(e => { console.error("SUMMARY MOCK-UP FAILED: " + e.message); process.exit(1); });
