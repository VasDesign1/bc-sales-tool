// ============================================================
// daily-email.js — the morning sales email (Faik, 2026-11-10; weekly and
// monthly summaries added 2026-11-11).
//
// One email each weekday morning:
//   Monday                        last week's summary, then the usual
//                                 month-so-far picture with its own points
//   first working day of a month  the closed month's summary (with last
//                                 week's under it when that day is a Monday)
//   any other weekday             the Daily Snapshot picture
//
// The figures are NOT recalculated here. The robot opens the tool's own
// index.html in headless Chrome, pours the newest snapshot into it the
// same way Fast lookup does, reads the targets workbooks with the page's
// own parser, and photographs what the page draws: the Daily Snapshot tab
// as it is, and the weekly / monthly blocks drawn in that page with the
// tool's own sums and styles (pageLib). So tool and email can never
// disagree in figures or in looks. The same figures feed the subject line
// and the optional AI dot points above each picture.
//
// Runs in GitHub Actions on Node 22 (needs the built-in WebSocket).
// Env:
//   BC_TENANT, BC_CLIENT_ID, BC_REFRESH_TOKEN   token exchange (Graph:
//                                               Files.ReadWrite.All + Mail.Send, delegated)
//   SNAPSHOT_PASSPHRASE                         opens the snapshot
//   DAILY_EMAIL_TO, DAILY_EMAIL_CC              recipients, comma separated
//   ANTHROPIC_API_KEY                           optional: adds the dot points
//                                               (needs `npm install @anthropic-ai/sdk`)
//   EMAIL_THEME                                 optional: tool theme for the pictures (default linen)
//   DRY_RUN=1                                   build everything, send nothing
//   CHROME_PATH                                 optional, default google-chrome
// Runs by hand only:
//   EMAIL_KIND    daily | weekly | monthly | monthly+weekly (default: what the day calls for)
//   EMAIL_AS_OF   YYYY-MM-DD: build the email as if it were that morning; marked
//                 TEST and sent to the first DAILY_EMAIL_TO address only
// Local layout test (no Microsoft sign-in, nothing sent):
//   node scripts/daily-email.js --local <payload.json> <targets.json> <out.html>
// ============================================================
"use strict";
const fs = require("fs"), os = require("os"), path = require("path"), http = require("http");
const crypto = require("crypto"), zlib = require("zlib"), { spawn } = require("child_process");
const { SP_CONFIG } = require(path.join(__dirname, "..", "bc-fetchers.js"));

const REPO = path.join(__dirname, "..");
const GRAPH = "https://graph.microsoft.com/v1.0";
const TOOL_URL = "https://vasdesign1.github.io/bc-sales-tool/";
const TARGETS_FOLDER = "Sales Budgets-Targets";      // keep in step with SP_TARGETS_FOLDER in index.html
const SLOTS = ["0500", "1100", "1700"];
const THEME = (process.env.EMAIL_THEME || "linen").toLowerCase().replace(/[^a-z]/g, "") || "linen";
const SHOT_WIDTH = 1560;      // CSS px the tab is laid out at for the picture: wide enough for all ten columns and the two kit panels side by side
const SHOW_WIDTH = 1200;      // px the email shows the picture at (it is taken at 2x, so it stays sharp)
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ---------- Microsoft Graph ----------
let _tok = null;
async function graphToken() {
    if (_tok) return _tok;
    const body = new URLSearchParams({
        grant_type: "refresh_token", client_id: process.env.BC_CLIENT_ID, refresh_token: process.env.BC_REFRESH_TOKEN,
        scope: "https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/Files.ReadWrite.All offline_access",
    });
    const resp = await fetch("https://login.microsoftonline.com/" + process.env.BC_TENANT + "/oauth2/v2.0/token",
        { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body.toString() });
    const tok = await resp.json();
    if (!tok.access_token) throw new Error("Graph token exchange failed: " + JSON.stringify(tok).slice(0, 500));
    console.log("  [graph] token scopes: " + (tok.scope || "(not reported)"));
    if (!/Mail\.Send/i.test(tok.scope || "")) console.warn("  [graph] WARNING: Mail.Send is not in the token — sending will be refused until the robot's sign-in is renewed with that permission.");
    return (_tok = tok.access_token);
}
async function graph(url, opts) {
    const resp = await fetch(url.startsWith("http") ? url : GRAPH + url, Object.assign({}, opts, { headers: Object.assign({ Authorization: "Bearer " + await graphToken() }, (opts && opts.headers) || {}) }));
    return resp;
}
async function graphJson(url) {
    const resp = await graph(url); const text = await resp.text();
    if (!resp.ok) throw new Error("Graph GET " + url + " → HTTP " + resp.status + ": " + text.slice(0, 300));
    return JSON.parse(text);
}
const enc = p => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");

// ---------- SharePoint: snapshots, targets, and the "sent today" note ----------
let _drive = null;
async function spItem(p) {
    if (!_drive) {
        const site = await graphJson("/sites/" + SP_CONFIG.host + ":/?$select=id");
        _drive = (await graphJson("/sites/" + site.id + "/drive?$select=id")).id;
    }
    return "/drives/" + _drive + "/root:/" + enc(p);
}
async function spReadJson(p) {
    const r = await graph(await spItem(p) + ":/content");
    if (!r.ok) return null;
    try { return JSON.parse(await r.text()); } catch (e) { return null; }
}
const SLOT_LABEL = { "0500": "Morning", "1100": "Midday", "1700": "Evening" };   // keep in step with SLOTS in snapshot.js
const SENT_NOTE = "/snapshots/daily-email.sent.json";
function melbourneNow() {
    const parts = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit",
        hour: "2-digit", minute: "2-digit", weekday: "short", hourCycle: "h23" }).formatToParts(new Date());
    const g = t => parts.find(x => x.type === t).value;
    return { date: g("year") + "-" + g("month") + "-" + g("day"), minutes: (parseInt(g("hour"), 10) % 24) * 60 + parseInt(g("minute"), 10), weekday: g("weekday") };
}
// Which snapshot goes out, and whether anything goes out at all.
//   By hand (Run workflow): always sends, from the newest snapshot (or the
//   forced SLOT), and leaves the "sent today" note alone.
//   On the schedule (Faik, 2026-11-10): one email each weekday morning, as
//   soon as today's Morning snapshot exists. If it still has not arrived by
//   FALLBACK_AT, the email goes out once from the newest snapshot there is
//   (normally yesterday's Evening one). A note on SharePoint records the
//   day, so later firings — and a late Morning snapshot — send nothing more.
const WINDOW_FROM = 5 * 60, FALLBACK_AT = 7 * 60 + 30, WINDOW_TO = 11 * 60;   // Melbourne minutes
async function chooseSnapshot() {
    const metas = {};
    for (const slot of SLOTS) metas[slot] = await spReadJson(SP_CONFIG.folder + "/snapshots/" + slot + ".meta.json");
    const newest = SLOTS.map(s => metas[s]).filter(Boolean).sort((a, b) => String(b.fetchedAtMelbourne).localeCompare(String(a.fetchedAtMelbourne)))[0] || null;
    if (process.env.GITHUB_EVENT_NAME !== "schedule") {
        const forced = (process.env.SLOT || "").trim();
        const meta = forced ? metas[forced] : newest;
        if (!meta) throw new Error("no snapshot found on SharePoint" + (forced ? " for slot " + forced : ""));
        return { meta, mark: false };
    }
    const mel = melbourneNow();
    if (mel.weekday === "Sat" || mel.weekday === "Sun") { console.log("Weekend in Melbourne — no email."); return null; }
    if (mel.minutes < WINDOW_FROM || mel.minutes >= WINDOW_TO) { console.log("Outside the morning window — nothing to do."); return null; }
    const note = await spReadJson(SP_CONFIG.folder + SENT_NOTE);
    if (note && note.date === mel.date) { console.log("Today's email already went out (" + (SLOT_LABEL[note.slot] || note.slot) + " snapshot) — nothing to do."); return null; }
    const morning = metas["0500"];
    if (morning && String(morning.fetchedAtMelbourne).slice(0, 10) === mel.date) return { meta: morning, mark: mel.date };
    if (mel.minutes >= FALLBACK_AT) {
        if (!newest) throw new Error("no snapshot found on SharePoint");
        console.log("Morning snapshot still missing — sending from the newest snapshot instead.");
        return { meta: newest, mark: mel.date };
    }
    console.log("Morning snapshot not there yet — waiting for the next firing.");
    return null;
}
async function markSent(date, meta) {
    const r = await graph(await spItem(SP_CONFIG.folder + SENT_NOTE) + ":/content?@microsoft.graph.conflictBehavior=replace",
        { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ date, slot: meta.slot, snapshot: meta.fetchedAtMelbourne, sentAtUtc: new Date().toISOString() }) });
    if (!r.ok) console.warn("Could not record that today's email was sent (HTTP " + r.status + ") — a later firing may send it again.");
}

// ---------- collect the inputs into a folder the page can read ----------
async function collectFromSharePoint(dir, best) {
    const folderItem = p => spItem(SP_CONFIG.folder + p);
    console.log("Snapshot: slot " + best.slot + " captured " + best.fetchedAtMelbourne + " Melbourne");
    const r = await graph(await folderItem("/snapshots/" + best.slot + ".bin") + ":/content");
    if (!r.ok) throw new Error("snapshot download → HTTP " + r.status);
    const bin = Buffer.from(await r.arrayBuffer());
    const salt = bin.subarray(0, 16), iv = bin.subarray(16, 28), ct = bin.subarray(28, bin.length - 16), tag = bin.subarray(bin.length - 16);
    const key = crypto.pbkdf2Sync(process.env.SNAPSHOT_PASSPHRASE, salt, 600000, 32, "sha256");   // keep in sync with snapshot.js
    const d = crypto.createDecipheriv("aes-256-gcm", key, iv); d.setAuthTag(tag);
    const json = zlib.gunzipSync(Buffer.concat([d.update(ct), d.final()]));
    fs.writeFileSync(path.join(dir, "payload.json"), json);
    console.log("  payload " + (json.length / 1048576).toFixed(1) + " MB");
    // targets workbooks, oldest first (a newer file wins where months overlap — as in the tool)
    const list = await graphJson(await folderItem("/" + TARGETS_FOLDER) + ":/children?$select=name,lastModifiedDateTime,file&$top=200");
    const files = (list.value || []).filter(f => f.file && /\.xls[xm]?$/i.test(f.name) && !/^~\$/.test(f.name))
        .sort((a, b) => (a.lastModifiedDateTime || "").localeCompare(b.lastModifiedDateTime || ""));
    const out = [];
    for (let i = 0; i < files.length; i++) {
        const fr = await graph(await folderItem("/" + TARGETS_FOLDER + "/" + files[i].name) + ":/content");
        if (!fr.ok) { console.warn("  targets file skipped (HTTP " + fr.status + ")"); continue; }
        fs.writeFileSync(path.join(dir, "t" + i + ".xlsx"), Buffer.from(await fr.arrayBuffer()));
        out.push({ file: "t" + i + ".xlsx", name: files[i].name, modified: files[i].lastModifiedDateTime || "" });
    }
    console.log("  targets workbooks: " + out.length);
    fs.writeFileSync(path.join(dir, "targets.json"), JSON.stringify({ files: out }));
}

// ---------- in the tool's own page ----------
// A clock that can be set: `new Date()` and Date.now() answer with the
// pretend moment while one is set; dates built from values are untouched.
// The summaries need it to ask the page's month-to-date function about a
// month that is already over; a test run (EMAIL_AS_OF) uses it to pretend
// the whole email is being sent on another morning.
const FAKE_CLOCK = `(() => {
    const D = Date; let fake = null;
    class F extends D { constructor(...a) { if (a.length) super(...a); else super(fake == null ? D.now() : fake); } static now() { return fake == null ? D.now() : fake; } }
    window.Date = F;
    window.__setNow = (y, m, d, h) => { fake = y == null ? null : new D(y, m - 1, d, h == null ? 12 : h).getTime(); };
})();`;
// Snapshot in, targets workbooks through the page's own parser.
const PAGE_SETUP = `(async () => {
    const overlay = document.getElementById("signinOverlay"); if (overlay) overlay.style.display = "none";
    window.alert = m => console.error("ALERT " + m);
    const payload = await (await fetch("/__email/payload.json")).json();
    hydrateSnapshot(payload);
    const t = await (await fetch("/__email/targets.json")).json();
    if (t.direct) Object.assign(spTargets, t.direct);                       // local layout test only
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
    return JSON.stringify({ invoices: state.invoices.length, asOf: state.dataAsOf || "", targetFiles: (spTargets.files || []).length, skipped: spTargets.skipped || [] });
})()`;

// Everything in pageLib runs inside the tool's page (it is turned into
// text and evaluated there), so it may use the page's functions and
// nothing of Node. It decides which email today is, works out the weekly
// and monthly figures with the page's own sums, and draws each block with
// the tool's own styles so it can be photographed.
//   cfg.kind   "auto" | "daily" | "weekly" | "monthly" | "monthly+weekly"
//   cfg.asOf   "" or "YYYY-MM-DD": pretend it is that morning (test runs)
// Which email (Faik, 2026-11-11):
//   Monday                        → last week's summary, then the usual
//                                   month-so-far picture with its own points
//   first working day of a month  → the closed month's summary (and last
//                                   week's under it when that day is a Monday)
//   any other weekday             → the daily picture, as before
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
    // "Mon 5 – Fri 9 Oct": first to last selling day inside a window
    const span = (a, b) => {
        let f = a, l = b;
        while (f < b && !isSellingDay(D(f))) f = addDays(f, 1);
        while (l > f && !isSellingDay(D(l))) l = addDays(l, -1);
        return f === l ? dLbl(f) : (f.slice(0, 7) === l.slice(0, 7) ? WD[D(f).getDay()] + " " + D(f).getDate() : dLbl(f)) + " – " + dLbl(l);
    };

    // ---- the page's clock ----
    // base(): the moment the email is for — now, or 5 AM on cfg.asOf with
    // later invoices put aside. asAt(day): the close of an earlier day.
    M.base = () => {
        if (!cfg.asOf) { window.__setNow(null); state.invoices = M.orig; return; }
        const a = cfg.asOf.split("-").map(Number); window.__setNow(a[0], a[1], a[2], 5);
        state.invoices = M.orig.filter(i => String(i.postingDate || "").slice(0, 10) < cfg.asOf);
    };
    M.asAt = day => { const a = day.split("-").map(Number); window.__setNow(a[0], a[1], a[2], 23); state.invoices = M.orig.filter(i => String(i.postingDate || "").slice(0, 10) <= day); };
    M.base();
    const today = localISO(new Date()), yesterday = addDays(today, -1), todayMk = today.slice(0, 7);

    // ---- which email ----
    let firstSell = todayMk + "-01"; while (!isSellingDay(D(firstSell))) firstSell = addDays(firstSell, 1);
    let kind = cfg.kind || "auto";
    if (kind === "auto") { const mo = today === firstSell, wk = D(today).getDay() === 1; kind = mo && wk ? "monthly+weekly" : mo ? "monthly" : wk ? "weekly" : "daily"; }
    const kinds = kind === "monthly+weekly" ? ["monthly", "weekly"] : kind === "monthly" ? ["monthly"] : kind === "weekly" ? ["weekly", "daily"] : ["daily"];
    const out = { today, kinds };

    // Sales, gross profit and invoice count per day, from the Month-to-Date
    // tab's own function with the clock on the last day wanted.
    const dayRows = last => { M.asAt(last); const m = new Map(); for (const r of mtdTotals().daily) m.set(r.day, { sales: r.sales, gp: r.gp, inv: r.invCount }); return m; };
    const bookCache = {};
    const get = day => { if (day > yesterday) return null; const mk = day.slice(0, 7); if (!bookCache[mk]) bookCache[mk] = dayRows(mk === yesterday.slice(0, 7) ? yesterday : monthEnd(mk)); return bookCache[mk].get(day) || null; };
    const sumWin = (from, to) => {
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
            if (d < from || d > to || d > yesterday) continue;
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

    // ---- a month cut into weeks (Mon–Sun, clipped to the month; a weekend-only stub joins its neighbour) ----
    const weeksOf = mk => {
        const mStart = mk + "-01", mEnd = monthEnd(mk), weeks = [];
        for (let s = mStart; s <= mEnd;) { const dow = D(s).getDay(); let e = addDays(s, dow === 0 ? 0 : 7 - dow); if (e > mEnd) e = mEnd; weeks.push({ from: s, to: e }); s = addDays(e, 1); }
        if (weeks.length > 1 && !sell(weeks[0].from, weeks[0].to)) { weeks[1].from = weeks[0].from; weeks.shift(); }
        const last = weeks.length - 1;
        if (last > 0 && !sell(weeks[last].from, weeks[last].to)) { weeks[last - 1].to = weeks[last].to; weeks.pop(); }
        return weeks;
    };
    const strip = (mk, asAt) => {
        let cum = 0; const mt = tgtAll(mk + "-01", monthEnd(mk));
        return weeksOf(mk).filter(w => w.from <= asAt).map((w, i) => {
            const s = sumWin(w.from, w.to > asAt ? asAt : w.to);
            s.n = i + 1; s.target = tgtAll(s.from, s.to); s.pct = s.target ? s.sales / s.target : null;
            cum += s.sales; s.cum = cum; s.cumPct = mt ? cum / mt : null;
            return s;
        });
    };

    // ---- WEEKLY: the last full Mon–Sun week against the week before ----
    let W = null;
    if (kinds.includes("weekly")) {
        const cwFrom = addDays(today, -((D(today).getDay() + 6) % 7) - 7), cwTo = addDays(cwFrom, 6);
        const pwFrom = addDays(cwFrom, -7), pwTo = addDays(cwFrom, -1);
        const mk = addDays(cwFrom, 4).slice(0, 7), mStart = mk + "-01", mEnd = monthEnd(mk);      // the month the week's Friday falls in
        const mtdTo = cwTo > mEnd ? mEnd : cwTo, fri = addDays(cwFrom, 4);
        W = sumWin(cwFrom, cwTo);
        W.n = Math.max(1, weeksOf(mk).findIndex(w => w.from <= fri && fri <= w.to) + 1); W.monthLabel = monthLbl(mk); W.sdMonth = sdTot(mk);
        W.prev = sumWin(pwFrom, pwTo);
        W.target = tgtAll(cwFrom, cwTo); W.pct = W.target ? W.sales / W.target : null;
        const eC = ent(cwFrom, cwTo), eP = ent(pwFrom, pwTo), eM = ent(mStart, mtdTo);
        W.branch = rowsFor("branch", eC, eP, cwFrom, cwTo);
        W.sp = rowsFor("salesperson", eC, eP, cwFrom, cwTo);
        W.kits = DS_KIT_TARGETS.map(g => ({ title: g.title, monthTarget: g.target, weekTarget: kitTgt(g, cwFrom, cwTo), qty: kitSum(g, eC), prev: kitSum(g, eP), mtd: kitSum(g, eM),
            codes: g.items.map(c => ({ code: c, qty: eC.kits[c] || 0, prev: eP.kits[c] || 0, mtd: eM.kits[c] || 0 })) }));
        W.strip = strip(mk, mtdTo);
        const mt = tgtAll(mStart, mEnd), mtd = W.strip.length ? W.strip[W.strip.length - 1].cum : 0;
        W.month = { target: mt, mtd, pct: mt ? mtd / mt : null };
        out.weekly = W;
    }

    // ---- MONTHLY: the month that has just closed against its target and the month before ----
    let Mo = null;
    if (kinds.includes("monthly")) {
        const mk = prevMk(todayMk), pmk = prevMk(mk), mStart = mk + "-01", mEnd = monthEnd(mk), monthTarget = tgtAll(mStart, mEnd);
        Mo = sumWin(mStart, mEnd);
        Mo.monthLabel = monthLbl(mk); Mo.prevLabel = monthLbl(pmk);
        Mo.prev = sumWin(pmk + "-01", monthEnd(pmk));
        Mo.target = monthTarget; Mo.balance = Mo.sales - monthTarget; Mo.pct = monthTarget ? Mo.sales / monthTarget : null;
        M.base();
        const eC = ent(mStart, mEnd), eP = ent(pmk + "-01", monthEnd(pmk));
        const closed = r => { r.status = r.target ? (r.sales >= r.target ? "TARGET MET" : "TARGET MISSED") : "NO PLAN"; return r; };
        // Margin by branch / salesperson: the run-rate tables' own margin per row, with cost brought to the month's headline cost (as the tool's tabs do).
        const withGp = (entity, t) => {
            t.rows.forEach(closed); closed(t.total);
            const marg = new Map(runRateRows({ metric: "margin", entity, budget: true }, mk, sellingDaysForMonthKey(mk, new Date())).map(r => [r.entity, r.cyVal]));
            const rawCost = t.rows.reduce((s, r) => s + (r.sales - (marg.get(r.code) || 0)), 0), k = rawCost > 0 ? (Mo.sales - Mo.gp) / rawCost : 1;
            t.rows.forEach(r => { r.gpPct = r.sales ? (r.sales - (r.sales - (marg.get(r.code) || 0)) * k) / r.sales : null; });
            t.total.gpPct = Mo.gpPct;
            return t;
        };
        state.invoices = M.orig;       // runRateRows reads the whole month whatever the clock says
        Mo.branch = withGp("branch", rowsFor("branch", eC, eP, mStart, mEnd));
        Mo.sp = withGp("salesperson", rowsFor("salesperson", eC, eP, mStart, mEnd));
        Mo.kits = DS_KIT_TARGETS.map(g => { const qty = kitSum(g, eC); return { title: g.title, target: g.target, qty, prev: kitSum(g, eP), balance: qty - g.target, pct: qty / g.target, status: qty >= g.target ? "TARGET MET" : "TARGET MISSED",
            codes: g.items.map(c => ({ code: c, qty: eC.kits[c] || 0, prev: eP.kits[c] || 0 })) }; });
        Mo.strip = strip(mk, mEnd);
        // Financial year to date (July → this month) against the budget, month by month
        const fyStart = (+mk.slice(5) >= 7 ? +mk.slice(0, 4) : +mk.slice(0, 4) - 1) + "-07";
        let fs = 0, ft = 0, n = 0, first = "";
        for (let m = fyStart; m <= mk; m = localISO(new Date(+m.slice(0, 4), +m.slice(5), 1)).slice(0, 7)) {
            const t = tgtAll(m + "-01", monthEnd(m)); if (!t) continue;
            let s = 0; ent(m + "-01", monthEnd(m)).br.forEach(v => s += v);
            fs += s; ft += t; n++; if (!first) first = m;
        }
        Mo.fy = { label: first ? MN[+first.slice(5) - 1] + " – " + MN[+mk.slice(5) - 1] + " " + mk.slice(0, 4) : "", months: n, sales: fs, target: ft, balance: fs - ft, pct: ft ? fs / ft : null };
        out.monthly = Mo;
    }
    M.base();

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
    const stripPanel = (title, rows, markN, monthTarget) => {
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
                ${tile("k-rev", "Week Target", money(w.target), "month target ÷ " + w.sdMonth + " selling days × " + w.days)}
                ${tile("k-inv", "Balance", w.target ? tone(w.sales - w.target, sgn(w.sales - w.target)) : "—", "sales − week target")}
                ${tile("k-gpp", "% of Week Target", bar(w.pct))}
                ${tile("k-cust", esc(w.monthLabel) + " After That Week", money(w.month.mtd), fmtPct(w.month.pct) + " of the month target")}
            </div>
            ${table("The Week · by Branch", "Branch", w.branch)}
            ${table("The Week · by Salesperson", "Salesperson", w.sp)}
            <div class="ds-pair">${w.kits.map(kit).join("")}</div>
            ${stripPanel(esc(w.monthLabel) + " · week by week so far", w.strip, w.n, w.month.target)}`;
    };

    const monthlyHtml = () => {
        const m = Mo, p = m.prev, d = m.sales - p.sales, pl = esc(m.prevLabel.split(" ")[0]);
        const row = (r, tot) => `<tr${tot ? ' class="ds-total"' : ""}>${nameTd(r)}<td class="num">${m$(r.prev)}</td><td class="num">${m$(r.target)}</td><td class="num">${tot ? money(r.sales) : m$(r.sales)}</td>${balTd(r.sales, r.target)}${chgTd(r.sales, r.prev)}<td class="num">${r.gpPct == null ? "—" : fmtPct(r.gpPct)}</td><td>${statusBadgeRunRate(r.status)}</td><td class="num">${bar(r.pct)}</td></tr>`;
        const table = (title, col, t) => panel(title, th(col) + th(pl + " Sales", 1) + th("Sales Target", 1) + th("Sales For Month", 1) + th("Balance", 1) + th("Change on " + pl, 1) + th("GP %", 1) + th("Status") + th("% Achieved", 1),
            t.rows.map(r => row(r, false)).join("") + row(t.total, true), legend);
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
            ${table("The Month · by Branch", "Branch", m.branch)}
            ${table("The Month · by Salesperson", "Salesperson", m.sp)}
            <div class="ds-pair">${m.kits.map(kit).join("")}</div>
            ${stripPanel(esc(m.monthLabel) + " · week by week", m.strip, 0, m.target)}`;
    };

    // ---- showing one block at a time and measuring it for the picture ----
    const daily = () => document.querySelector('section[data-tab-content="dailySnapshot"]');
    const mine = () => { let s = document.getElementById("mailSum"); if (!s) { s = document.createElement("section"); s.id = "mailSum"; daily().parentElement.appendChild(s); } return s; };
    M.render = async kind => {
        const r = {};
        if (!M.styled) {      // the block alone: none of the tool's chrome, and none of the tab's controls that mean nothing in an email
            const st = document.createElement("style");
            st.textContent = ".sidebar, .header, .toolbar, #tabStrip, #filterChips, .filter-chips, #themePicker, .progress, #busyWrap, #exportMenu, #helpTip { display: none !important; }"
                + " .app { display: block !important; } .main { margin: 0 !important; width: auto !important; } .container { padding: 18px 20px !important; max-width: none !important; }"
                + " section[data-tab-content='dailySnapshot'] .rr-tgt, section[data-tab-content='dailySnapshot'] [data-ds-goto], #dsEmpty, #dsKits .panel-header .meta { display: none !important; }"
                + " #mailSum .ds-pct i { font-weight: 700; }";
            document.head.appendChild(st); M.styled = true;
        }
        M.base();
        if (kind === "daily") {
            mine().style.display = "none";
            setActiveTab("dailySnapshot"); renderDailySnapshot();          // paints the tab with the targets now in place
            daily().style.display = "";
            if (cfg.asOf) document.getElementById("dsAsOf").textContent = "Test: figures rewound to the morning of " + longLbl(cfg.asOf) + ". Shipped Not Invoiced is today's and cannot be rewound.";
            r.daily = { f: dsSnapshotFigures(), names: { branch: spBudgetNames("branch"), salesperson: spBudgetNames("salesperson") }, asOf: state.dataAsOf || "", today };
        } else {
            document.querySelectorAll("section[data-tab-content]").forEach(s => s.style.display = "none");
            const s = mine(); s.innerHTML = kind === "weekly" ? weeklyHtml() : monthlyHtml(); s.style.display = "";
        }
        window.scrollTo(0, 0);
        if (document.fonts && document.fonts.ready) await document.fonts.ready;
        await new Promise(res => setTimeout(res, 400));
        const box = document.querySelector(".container"), c = box.getBoundingClientRect();
        // The page's own background and text colour, so the email around the picture can match the theme.
        let bg = "rgb(255, 255, 255)";
        for (let e = box; e; e = e.parentElement) { const b = getComputedStyle(e).backgroundColor; if (b && b !== "transparent" && b !== "rgba(0, 0, 0, 0)") { bg = b; break; } }
        r.clip = { x: 0, y: Math.max(0, Math.floor(c.top + window.scrollY)), width: Math.ceil(document.documentElement.clientWidth), height: Math.ceil(c.height) };
        r.bg = bg; r.ink = getComputedStyle(box).color;
        return JSON.stringify(r);
    };
    return JSON.stringify(out);
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
        // The page works on the local clock: make that Melbourne, wherever the robot runs.
        await send("Emulation.setTimezoneOverride", { timezoneId: "Australia/Melbourne" });
        // The settable clock and the chosen theme, from the first paint.
        await send("Page.addScriptToEvaluateOnNewDocument", { source: FAKE_CLOCK + "try{localStorage.setItem('bcSalesTheme','" + THEME + "');}catch(e){}" });
        await send("Emulation.setDeviceMetricsOverride", { width: SHOT_WIDTH, height: 1200, deviceScaleFactor: 2, mobile: false });
        await send("Page.navigate", { url: "http://127.0.0.1:" + port + "/index.html" });
        for (let i = 0; i < 120; i++) {
            const r = await send("Runtime.evaluate", { expression: "typeof dsSnapshotFigures === 'function' && typeof hydrateSnapshot === 'function' && typeof XLSX !== 'undefined' && typeof Chart !== 'undefined'", returnByValue: true });
            if (r.result && r.result.result && r.result.result.value === true) break;
            if (i === 119) throw new Error("the tool page did not finish loading in headless Chrome");
            await sleep(500);
        }
        return { ev, shoot, close, pageErrors };
    } catch (e) { close(); throw e; }
}

// ---------- AI dot points (optional) ----------
// A few dot points above each picture. The model is handed the figures
// the picture already shows and only describes them; it calculates
// nothing that lands in the email. Skipped quietly when there is no API
// key, or when the call fails or is declined: the email must still go out.
const STYLE_RULES = [
    "Use only the figures provided. Quote them as given, rounded sensibly (for example $401k, 19% of target); do not derive new totals or forecasts of your own. If something needed is missing, leave it out rather than guessing.",
    "This is read by the whole team, so name a salesperson only to recognise a result that is ahead of pace or on target; describe shortfalls at branch or team level, not by person.",
    "Australian English, matter-of-fact and encouraging without hype. Output format: each dot point on its own line starting with a hyphen and a space, and nothing else: no greeting, no intro line, no sign-off, no headings, no emojis. One point per idea, each a single plain sentence of at most about 25 words, and no two points making the same observation. If the figures do not support that many distinct points, write fewer rather than padding.",
    "So the points can be skimmed, put the one or two things that matter most in each point in bold by wrapping them in double asterisks, like **$401k** or **Dandenong** or **19% of target**: the key figure, and the branch, person or product it belongs to. Bold short phrases only, never a whole sentence, and use no other markdown.",
];
const SUMMARY_KEYS = "In the figures: sales = invoiced sales excluding tax; prev = the same figure for the period before; target = the sales target for the period; pct and every field ending in Pct are percentages already; perDay = sales per selling day; days = selling days; cum = month to date; codes under kits are item codes with units sold.";
const COMMENTARY_SYSTEM = [
    "You write the short opening dot points of a daily sales email at an Australian HVAC wholesaler. It goes to internal staff every morning, above a picture of the sales dashboard they already know.",
    "You are given the dashboard's own figures as JSON: month-to-date sales, goods shipped but not yet invoiced, gross profit, selling days, each branch and each salesperson against their monthly sales target, and two product quantity targets (AirTouch kits and Omniflow units).",
    "Write 10 to 12 dot points that walk a busy reader through where the month stands and what matters today. Cover, in roughly this order: sales so far against the month's target and against where they should be at this point of the month; goods shipped but not yet invoiced and what they add; gross profit and margin; the daily rate achieved so far and the rate now required; each branch in turn; the salespeople picture; the AirTouch kit target; the Omniflow target; and what would most help close the gap.",
].concat(STYLE_RULES).join("\n\n");
const WEEKLY_SYSTEM = [
    "You write the short opening dot points of a weekly sales email at an Australian HVAC wholesaler. It goes to internal staff on Monday morning, above a picture of last week's figures.",
    "You are given last week's figures as JSON: the week's sales, gross profit and target (the week's share of the monthly target), the same for the week before, each branch and each salesperson, two product quantity targets (AirTouch kits and Omniflow units), the month's weeks so far, and where the week left the month.",
    "Write 6 to 8 dot points: how the week went against its target and against the week before; margin; each branch; salesperson results worth recognising; the AirTouch kit and Omniflow units; and where the week leaves the month against its target.",
    SUMMARY_KEYS,
].concat(STYLE_RULES).join("\n\n");
const MONTHLY_SYSTEM = [
    "You write the short opening dot points of a monthly sales email at an Australian HVAC wholesaler. It goes to internal staff on the first working day of the new month, above a picture of the closed month's figures.",
    "You are given the closed month's figures as JSON: sales, gross profit and target, the same for the month before, each branch and each salesperson with target met or missed and margin (gpPct), two product quantity targets (AirTouch kits and Omniflow units), the month week by week, and the financial year to date against budget (fy).",
    "Write 8 to 10 dot points: the month's sales against target; against the month before; gross profit and margin; each branch in turn; salespeople who met target; the AirTouch kit and Omniflow targets; the strongest and weakest week; and the financial year to date against budget.",
    SUMMARY_KEYS,
].concat(STYLE_RULES).join("\n\n");
function commentaryInput(res) {
    const f = res.f;
    const rows = (t, names) => t.rows.filter(r => r.target > 0).map(r => ({ code: r.entity, name: names[r.entity] || undefined, target: Math.round(r.target), sales_so_far: Math.round(r.cyVal),
        pct_of_target: +(r.pctAchieved * 100).toFixed(1), status: r.status, required_daily: Math.round(r.requiredDaily) }));
    const tot = t => ({ target: Math.round(t.total.target), sales_so_far: Math.round(t.total.cyVal), pct_of_target: t.total.pctAchieved == null ? null : +(t.total.pctAchieved * 100).toFixed(1),
        status: t.total.status, target_per_day: Math.round(t.total.targetDaily), actual_per_day: Math.round(t.total.actualDaily), required_per_day: Math.round(t.total.requiredDaily) });
    return {
        date: res.today, month: f.monthLabel, figures_as_of: res.asOf,
        selling_days: { total: f.sd.total, passed: f.sd.passed, remaining: f.sd.remaining, pct_of_month_passed: f.sd.total ? +(f.sd.passed / f.sd.total * 100).toFixed(1) : null },
        month_to_date: { sales_excl_tax: Math.round(f.sales), shipped_not_invoiced: Math.round(f.sni), sales_plus_shipped_not_invoiced: Math.round(f.total),
                         gross_profit: Math.round(f.gp), gross_profit_pct: +(f.gpPct * 100).toFixed(2), invoices: f.invoices },
        branches: { total: tot(f.branch), rows: rows(f.branch, res.names.branch || {}) },
        salespeople: { total: tot(f.salesperson), rows: rows(f.salesperson, res.names.salesperson || {}) },
        quantity_targets: (f.kits || []).map(k => ({ name: k.title, monthly_target_units: k.target, units_sold: k.qty, pct_of_target: +(k.pctAchieved * 100).toFixed(1), status: k.status,
                                                    required_per_day: +Math.max(0, k.requiredDaily).toFixed(1), by_item: k.rows.map(r => ({ item: r.label, units: r.qty })) })),
    };
}
// Rounds the weekly / monthly figures for the model: dollars to whole numbers, ratios to percentages with one decimal.
function slim(v, key) {
    if (Array.isArray(v)) return v.map(x => slim(x));
    if (v && typeof v === "object") { const o = {}; for (const k in v) if (k !== "from" && k !== "to") o[k] = slim(v[k], k); return o; }
    if (typeof v === "number") return /pct$/i.test(key || "") ? +(v * 100).toFixed(1) : /^(qty|prev|mtd|weekTarget)$/.test(key || "") && Math.abs(v) < 1000 ? +v.toFixed(1) : Math.round(v);
    return v;
}
async function writeCommentary(system, figures, what) {
    if (!process.env.ANTHROPIC_API_KEY) { console.log("Commentary (" + what + "): no ANTHROPIC_API_KEY — email goes without it."); return ""; }
    try {
        const mod = require("@anthropic-ai/sdk");
        const Anthropic = mod.default || mod;
        const client = new Anthropic();
        const response = await client.beta.messages.create({
            model: "claude-opus-5-5",
            max_tokens: 8000,
            // If a safety classifier declines, the API re-runs the request on its default fallback model.
            betas: ["server-side-fallback-2026-07-01"],
            fallbacks: "default",
            system,
            messages: [{ role: "user", content: "The figures:\n\n" + JSON.stringify(figures, null, 1) }],
        });
        if (response.stop_reason === "refusal") { console.warn("Commentary (" + what + "): request declined — email goes without it."); return ""; }
        const text = response.content.filter(b => b.type === "text").map(b => b.text).join("\n").trim();
        console.log("Commentary (" + what + "): " + text.length + " characters · model " + response.model + " · tokens in " + response.usage.input_tokens + " / out " + response.usage.output_tokens);
        return text;
    } catch (e) {
        console.warn("Commentary (" + what + "): skipped (" + String((e && e.message) || e).slice(0, 200) + ") — email goes without it.");
        return "";
    }
}

// ---------- the email ----------
// Each block is a picture drawn by the tool's own page (Faik, 2026-11-10:
// same colours, shape and format as the tool) with its dot points as real
// text above it. Layout: a bold header and the first block's points on a
// pale card, a gap, that block's picture; every further block gets its
// own card (a caption and its points) and its picture. Calibri 11pt,
// centred, exactly as wide as the pictures.
// Colours come from the tab's own page colour, in three tones so the
// parts stay distinct: a deeper tone behind the whole message, a paler
// card for the text, and the pictures in the page colour itself.
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const money0 = n => "$" + Math.round(n).toLocaleString("en-AU");
const pct0 = (p, of) => p == null ? "" : " (" + (p * 100).toFixed(0) + "% of " + of + ")";
const hex = (rgb, fallback) => { const m = /(\d+)\D+(\d+)\D+(\d+)/.exec(rgb || ""); return m ? "#" + [1, 2, 3].map(i => (+m[i]).toString(16).padStart(2, "0")).join("") : fallback; };
function emailHtml(mail, src) {
    const mix = (a, b, t) => "#" + [1, 3, 5].map(i => Math.round(parseInt(a.slice(i, i + 2), 16) * (1 - t) + parseInt(b.slice(i, i + 2), 16) * t).toString(16).padStart(2, "0")).join("");
    const page = hex(mail.blocks[0].bg, "#ffffff"), ink = hex(mail.blocks[0].ink, "#1f2430");
    const light = [1, 3, 5].reduce((s, i) => s + parseInt(page.slice(i, i + 2), 16), 0) / 3 > 128;
    const outer = mix(page, ink, 0.16), card = light ? mix(page, "#ffffff", 0.7) : mix(page, ink, 0.07), line = mix(page, ink, 0.28);
    const FONT = "font-family:Calibri,Segoe UI,Arial,sans-serif;font-size:11pt;color:" + ink + ";";
    // **like this** in a point becomes bold; any stray asterisks are dropped.
    const rich = p => esc(p).replace(/\*\*([^*]+)\*\*/g, '<strong style="font-weight:700;">$1</strong>').replace(/\*/g, "");
    const list = text => {
        const points = text ? text.split(/\n+/).map(p => p.replace(/^\s*(?:[-•–]|\*(?!\*)|\d+[.)])\s*/, "").trim()).filter(Boolean) : [];
        return points.length ? '<ul style="' + FONT + 'line-height:1.5;margin:0;padding:0 0 0 26px;">' + points.map(p => '<li style="' + FONT + 'margin:0 0 6px;">' + rich(p) + "</li>").join("") + "</ul>" : "";
    };
    const space = '<div style="height:16px;line-height:16px;font-size:1px;">&nbsp;</div>';
    const gap = h => '<tr><td height="' + h + '" style="height:' + h + 'px;font-size:1px;line-height:1px;">&nbsp;</td></tr>';
    const cardRow = inner => '<tr><td align="left" bgcolor="' + card + '" style="background-color:' + card + ';border:1px solid ' + line + ';padding:20px 24px 18px;">' + inner + "</td></tr>";
    const title = t => '<div style="' + FONT + 'font-size:14pt;font-weight:700;letter-spacing:.3px;">' + esc(t) + "</div>";
    const picRow = (b, i) => '<tr><td align="left" bgcolor="' + page + '" style="background-color:' + page + ';border:1px solid ' + line + ';"><img src="' + src(i) + '" width="' + SHOW_WIDTH + '" alt="' + esc(b.alt) + '" style="display:block;width:100%;max-width:' + SHOW_WIDTH + 'px;height:auto;border:0;margin:0 auto;"></td></tr>';
    return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>'
        + '<body bgcolor="' + outer + '" style="margin:0;padding:0;background-color:' + outer + ';">'
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="' + outer + '" style="background-color:' + outer + ';"><tr><td align="center" style="padding:22px 14px 28px;">'
        + '<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" width="' + SHOW_WIDTH + '" style="width:' + SHOW_WIDTH + 'px;max-width:100%;margin:0 auto;">'
        + mail.blocks.map((b, i) => {
            const pts = list(b.points);
            const head = i === 0
                ? title(mail.header) + (mail.sub ? '<div style="' + FONT + 'font-size:10pt;padding-top:3px;">' + mail.sub + "</div>" : "") + (pts ? space + pts : "")
                : title(b.caption) + (pts ? space + pts : "");
            return (i ? gap(28) : "") + cardRow(head) + gap(20) + picRow(b, i);
        }).join("")
        + "</table></td></tr></table></body></html>";
}

async function sendMail(mail, to, cc) {
    const message = {
        subject: mail.subject,
        body: { contentType: "HTML", content: emailHtml(mail, i => "cid:salespic" + i) },
        toRecipients: to.map(a => ({ emailAddress: { address: a } })),
        ccRecipients: (cc || []).map(a => ({ emailAddress: { address: a } })),
        attachments: mail.blocks.map((b, i) => ({ "@odata.type": "#microsoft.graph.fileAttachment", name: b.kind + "-summary.png", contentType: "image/png",
                                                  contentBytes: b.png.toString("base64"), contentId: "salespic" + i, isInline: true })),
    };
    const resp = await graph("/me/sendMail", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ message, saveToSentItems: true }) });
    if (resp.status !== 202) throw new Error("sendMail → HTTP " + resp.status + ": " + (await resp.text()).slice(0, 400));
}

// Puts the email together from what the page decided and drew.
async function buildMail(tool, plan, snapshot, asOf) {
    // The pictures; the sharpness steps down if together they are too heavy for one Graph call (about 3 MB).
    let blocks = [];
    for (const dsf of [2, 1.5, 1]) {
        blocks = [];
        for (const kind of plan.kinds) {
            const r = JSON.parse(await tool.ev("window.__m.render(" + JSON.stringify(kind) + ")", "draw the " + kind + " block"));
            r.kind = kind; r.png = await tool.shoot(r.clip, dsf); r.scale = dsf;
            blocks.push(r);
        }
        if (blocks.reduce((s, b) => s + b.png.length, 0) <= 2.6 * 1048576) break;
    }
    const W = plan.weekly, Mo = plan.monthly, first = plan.kinds[0];
    const longDate = iso => new Date(iso + "T12:00:00Z").toLocaleDateString("en-AU", { weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "UTC" });
    for (const b of blocks) {
        if (b.kind === "daily") {
            const f = b.daily.f;
            b.points = await writeCommentary(COMMENTARY_SYSTEM, commentaryInput(b.daily), "daily");
            b.caption = f.monthLabel.toUpperCase() + " SO FAR";
            b.alt = "Daily Snapshot for " + f.monthLabel + ": sales plus shipped not invoiced " + money0(f.total) + ", gross profit " + money0(f.gp) + " (" + (f.gpPct * 100).toFixed(2) + "%).";
        } else if (b.kind === "weekly") {
            b.points = await writeCommentary(WEEKLY_SYSTEM, slim(W), "weekly");
            b.caption = "LAST WEEK · " + W.label.toUpperCase();
            b.alt = "Weekly summary, " + W.label + ": sales " + money0(W.sales) + ", gross profit " + money0(W.gp) + ".";
        } else {
            b.points = await writeCommentary(MONTHLY_SYSTEM, slim(Mo), "monthly");
            b.caption = Mo.monthLabel.toUpperCase();
            b.alt = "Monthly summary for " + Mo.monthLabel + ": sales " + money0(Mo.sales) + ", gross profit " + money0(Mo.gp) + ".";
        }
    }
    // Header says which snapshot the figures come from, and its day when it is not today's.
    const snap = snapshot || {};
    const snapDay = String(snap.fetchedAtMelbourne || "").slice(0, 10), snapTime = String(snap.fetchedAtMelbourne || "").slice(11).replace(/^24:/, "00:");
    const slotName = snap.slot ? (SLOT_LABEL[snap.slot] || snap.slot) : "";
    const snapLbl = slotName ? slotName + " snapshot" + (snapDay && snapDay !== plan.today ? " of " + longDate(snapDay) : "") + (snapTime ? ", taken " + snapTime : "") : "";
    const slotTag = slotName ? " · " + slotName.toUpperCase() + " SNAPSHOT" : "";
    const day = new Date(plan.today + "T12:00:00Z").toLocaleDateString("en-AU", { weekday: "short", day: "numeric", month: "short", timeZone: "UTC" });
    let subject, header;
    if (first === "daily") {
        const f = blocks[0].daily.f;
        subject = "Daily Snapshot · " + day + " · " + f.monthLabel + " sales " + money0(f.sales) + pct0(f.branch.total.pctAchieved, "target");
        header = "DAILY SALES SUMMARY · " + longDate(plan.today).toUpperCase() + slotTag;
    } else if (first === "weekly") {
        subject = "Weekly Summary · " + W.label + " · sales " + money0(W.sales) + pct0(W.pct, "week target");
        header = "WEEKLY SALES SUMMARY · " + W.label.toUpperCase() + slotTag;
    } else {
        subject = "Monthly Summary · " + Mo.monthLabel + " · sales " + money0(Mo.sales) + pct0(Mo.pct, "target");
        header = "MONTHLY SALES SUMMARY · " + Mo.monthLabel.toUpperCase() + " · PRELIMINARY" + slotTag;
    }
    const sub = (asOf ? "<strong style=\"font-weight:700;\">TEST</strong>, as it would have arrived on " + esc(longDate(asOf)) + ". " : "")
        + (snapLbl ? "Figures from the <strong style=\"font-weight:700;\">" + esc(snapLbl) + "</strong>" : "");
    return { subject: (asOf ? "[TEST] " : "") + subject, header, sub, blocks };
}

// Run as a script: build and send.
if (require.main === module) (async () => {
    const local = process.argv.indexOf("--local");
    let chosen = null;
    if (local < 0) {
        for (const k of ["BC_TENANT", "BC_CLIENT_ID", "BC_REFRESH_TOKEN", "SNAPSHOT_PASSPHRASE"]) if (!process.env[k]) throw new Error("Missing env " + k);
        chosen = await chooseSnapshot();
        if (!chosen) return;            // scheduled firing with nothing to send
    }
    // A run by hand may ask for a particular email (EMAIL_KIND) or pretend to be another morning (EMAIL_AS_OF); the schedule never does.
    const byHand = process.env.GITHUB_EVENT_NAME !== "schedule";
    const kindIn = byHand ? (process.env.EMAIL_KIND || "").trim().toLowerCase() : "", asOfIn = byHand ? (process.env.EMAIL_AS_OF || "").trim() : "";
    const cfg = { kind: ["daily", "weekly", "monthly", "monthly+weekly"].includes(kindIn) ? kindIn : "auto", asOf: /^\d{4}-\d{2}-\d{2}$/.test(asOfIn) ? asOfIn : "" };
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "daily-email-"));
    let tool = null;
    try {
        if (local >= 0) {
            fs.copyFileSync(process.argv[local + 1], path.join(dir, "payload.json"));
            fs.copyFileSync(process.argv[local + 2], path.join(dir, "targets.json"));
        } else await collectFromSharePoint(dir, chosen.meta);
        tool = await openTool(dir);
        const setup = JSON.parse(await tool.ev(PAGE_SETUP, "load the snapshot"));
        const plan = JSON.parse(await tool.ev("(" + pageLib.toString() + ")(" + JSON.stringify(cfg) + ")", "work out the figures"));
        if (tool.pageErrors.length) console.warn("  page messages: " + tool.pageErrors.slice(0, 5).join(" || "));
        const mail = await buildMail(tool, plan, chosen ? chosen.meta : { slot: "0500", fetchedAtMelbourne: setup.asOf }, cfg.asOf);
        console.log("Email for " + plan.today + (cfg.asOf ? " (test date)" : "") + ": " + plan.kinds.join(" + ") + " · " + setup.invoices + " invoices in the snapshot · " + setup.targetFiles + " targets file(s)"
            + (setup.skipped.length ? " · skipped " + setup.skipped.length : "") + " · " + mail.blocks.map(b => b.kind + " picture " + Math.round(b.clip.width * b.scale) + "×" + Math.round(b.clip.height * b.scale) + " px, " + (b.png.length / 1024).toFixed(0) + " KB"
            + (b.points ? ", with points" : ", no points")).join(" · ") + " (" + THEME + " theme)");
        if (local >= 0) {
            const outHtml = process.argv[local + 3], base = outHtml.replace(/\.html?$/i, "");
            mail.blocks.forEach((b, i) => fs.writeFileSync(base + "-" + (i + 1) + ".png", b.png));
            fs.writeFileSync(outHtml, "<!-- " + mail.subject + " -->\n" + emailHtml(mail, i => path.basename(base) + "-" + (i + 1) + ".png"));
            console.log("Wrote " + outHtml + " + " + mail.blocks.length + " picture(s)");
            return;
        }
        const addrs = v => (v || "").split(/[,;\s]+/).filter(Boolean);
        const to = addrs(process.env.DAILY_EMAIL_TO);
        const cc = addrs(process.env.DAILY_EMAIL_CC).filter(a => !to.some(t => t.toLowerCase() === a.toLowerCase()));
        if (process.env.DRY_RUN === "1" || process.env.DRY_RUN === "true") { await graphToken(); console.log("DRY RUN — nothing sent (" + to.length + " to, " + cc.length + " cc configured)."); return; }
        if (!to.length) throw new Error("DAILY_EMAIL_TO is empty — no recipients");
        // A test for another morning goes to the first address only (Faik), never to the whole list.
        if (cfg.asOf) { await sendMail(mail, to.slice(0, 1), []); console.log("Sent the test to 1 recipient."); return; }
        await sendMail(mail, to, cc);
        console.log("Sent to " + to.length + " recipient(s)" + (cc.length ? " + " + cc.length + " cc" : "") + ".");
        if (chosen.mark) await markSent(chosen.mark, chosen.meta);
    } finally {
        if (tool) tool.close();
        try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) { /* temp dir */ }
    }
})().catch(e => { console.error("DAILY EMAIL FAILED: " + e.message); process.exit(1); });
module.exports = { openTool, PAGE_SETUP, pageLib };
