/* ============================================================
   STAFF ACTIVITY  (Operations → Staff Activity; Faik, 2026-11-25)
   Who raised, changed, shipped, invoiced and received what — by
   employee, branch, customer / supplier, salesperson or period.

   OWNER-ONLY: the nav link carries class restrictedNav and the tab
   is in RESTRICTED_TABS (index.html). Nothing here loads or draws
   for anybody else.

   Data: ONE file, snapshots/activity.bin on SharePoint, built by the
   robot scripts/activity.js (see the notes there for where each
   figure comes from). The tab never touches the sales data the rest
   of the tool loads and is not affected by the toolbar filters.
     who RAISED / CHANGED an order   Wiise Change Log (first line of
                                     a document = the person who raised it)
     who POSTED shipments, receipts  the user on the stock movement
     who POSTED invoices, credits    the user on the customer / vendor ledger
   Live reads, both with the viewer's own sign-in:
     "Get latest from Wiise"   everything newer than the file, added on
                               screen only (activity-engine.js, the same
                               rules the robot runs; the file is untouched)
     "full change history"     of one order, from the Change Log.

   Keep help.js in step when tiles, columns or controls change.
   ============================================================ */
(function () {
    "use strict";
    const $ = id => document.getElementById(id);
    const esc = s => escapeHtml(s == null ? "" : s);

    // What is counted. side: sales | purch. stamp = the event carries a time of day.
    const MEAS = [
        { k: "soC", side: "sales", label: "Sales orders raised",        short: "Orders raised",      tint: "k-rev",  stamp: true },
        { k: "soM", side: "sales", label: "Sales orders changed",       short: "Orders changed",     tint: "k-qty",  stamp: true },
        { k: "soD", side: "sales", label: "Sales orders deleted",       short: "Orders deleted",     tint: "k-gpp",  stamp: true, gone: true },
        { k: "sqC", side: "sales", label: "Quotes raised",              short: "Quotes raised",      tint: "k-loc",  stamp: true },
        { k: "sqV", side: "sales", label: "Quotes turned into orders",  short: "Quotes → orders",    tint: "k-avg",  stamp: true },
        { k: "sqD", side: "sales", label: "Quotes deleted",             short: "Quotes deleted",     tint: "k-gpp",  stamp: true, gone: true },
        { k: "sbC", side: "sales", label: "Blanket orders raised",      short: "Blankets raised",    tint: "k-loc",  stamp: true },
        { k: "sbV", side: "sales", label: "Blanket orders turned into orders", short: "Blankets → orders", tint: "k-avg", stamp: true },
        { k: "sh",  side: "sales", label: "Shipments posted",           short: "Shipments",          tint: "k-inv" },
        { k: "si",  side: "sales", label: "Sales invoices posted",      short: "Invoices",           tint: "k-gp" },
        { k: "sc",  side: "sales", label: "Sales credit memos posted",  short: "Credit memos",       tint: "k-gpp" },
        { k: "poC", side: "purch", label: "Purchase orders raised",     short: "POs raised",         tint: "k-rev",  stamp: true },
        { k: "poM", side: "purch", label: "Purchase orders changed",    short: "POs changed",        tint: "k-qty",  stamp: true },
        { k: "poD", side: "purch", label: "Purchase orders deleted",    short: "POs deleted",        tint: "k-gpp",  stamp: true, gone: true },
        { k: "pr",  side: "purch", label: "Purchase receipts posted",   short: "Receipts",           tint: "k-inv" },
        { k: "pi",  side: "purch", label: "Purchase invoices posted",   short: "Purch. invoices",    tint: "k-gp" },
        { k: "pc",  side: "purch", label: "Purchase credit memos posted", short: "Purch. credits",   tint: "k-gpp" },
    ];
    const MBY = Object.fromEntries(MEAS.map(m => [m.k, m]));
    // Header fields Wiise rewrites by itself while a document is released or
    // posted. A change made up only of these is the system's own
    // bookkeeping, not somebody editing the order, and is left out of
    // "changed". (Shipping and invoicing are counted from the posted
    // documents instead.)
    const NOISE = new Set(["Status", "Ship", "Invoice", "Receive", "Shipping No.", "Posting No.", "Receiving No.", "Last Shipping No.", "Last Posting No.",
        "Last Receiving No.", "Last Return Receipt No.", "Last Return Shipment No.", "Return Receipt No.", "Return Shipment No.", "Print Posted Documents",
        "Doc. No. Occurrence", "No. of Archived Versions", "Job Queue Status", "Job Queue Entry ID", "Dimension Set ID", "Shipping No. Series", "Posting No. Series",
        "Receiving No. Series", "Posting Description", "Shipped", "Completely Shipped", "Completely Received", "Last Email Sent Time", "Last Email Sent Status",
        "Sent as Email", "Last Email Notif Cleared", "No. Printed", "Invoice Discount Value", "Invoice Discount Calculation", "Invoice Discount Amount",
        "Posting Date", "Document Date", "VAT Date", "Due Date", "Pmt. Discount Date", "Prepmt. Pmt. Discount Date", "Prepayment Due Date", "Posted Tax Document",
        "Shortcut Dimension 1 Code", "Shortcut Dimension 2 Code",
        "Work Description", "Recalculate Invoice Disc.", "Prepayment No.", "Last Prepayment No.", "Outstanding Amount ($)", "Coupled to Dataverse", "IC Status"]);
    const VIEWS = [["user", "Employee"], ["branch", "Branch"], ["party", "Customer / supplier"], ["sp", "Salesperson / purchaser"], ["period", "Period"]];
    const CHARTS = [
        ["trend", "Over time — stacked by the View by choice"],
        ["rank",  "Ranking — every activity side by side"],
        ["share", "Share of the total"],
        ["party", "Who works for which customer / supplier"],
        ["heat",  "Busiest day and hour (orders and quotes raised or changed)"],
    ];
    const PURCH_PAGES = { order: 50, receipt: 136, invoice: 138, credit: 140 };

    const S = {
        status: "idle", err: "", data: null, meta: null, ev: [],
        range: "month", from: "", to: "", grain: "day", side: "sales", show: "count",
        branch: "", user: "", sp: "", party: "",
        view: "user", thens: ["", "", ""], sort: { key: "__total", dir: "desc" },      // thens: up to three further levels under View by
        chart: "trend", cm: "",
        pick: null, docQ: "", docLimit: 300, hist: null,
        docSort: { key: "date", dir: "desc" }, histSort: { key: "when", dir: "asc" },
        othOpen: false, othUser: "", othSort: { key: "v", dir: "desc" },      // the "what is in Others" list under the chart
        live: "", liveErr: "",      // "" | loading | done | error  — "Get latest from Wiise"
    };
    const money0 = n => n == null || isNaN(n) ? "—" : new Intl.NumberFormat("en-AU", { style: "currency", currency: "AUD", maximumFractionDigits: 0 }).format(n);
    const plural = (n, w) => fmtInt(n) + " " + w + (n === 1 ? "" : "s");
    const allowed = () => typeof currentSignInName === "function" && currentSignInName() === RESTRICTED_NAV_OWNER;

    // ---------- dates ----------
    const today = () => (typeof melbourneTodayISO === "function" ? melbourneTodayISO() : new Date().toISOString().slice(0, 10));
    const D = iso => { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
    const iso = d => d.toISOString().slice(0, 10);
    const addDays = (s, n) => { const d = D(s); d.setUTCDate(d.getUTCDate() + n); return iso(d); };
    const monday = s => { const d = D(s); const wd = (d.getUTCDay() + 6) % 7; d.setUTCDate(d.getUTCDate() - wd); return iso(d); };
    const WD = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const dayLabel = s => { const d = D(s); return WD[(d.getUTCDay() + 6) % 7] + " " + d.getUTCDate() + " " + MON[d.getUTCMonth()]; };
    const dateLabel = s => { if (!s) return "—"; const d = D(s); return String(d.getUTCDate()).padStart(2, "0") + " " + MON[d.getUTCMonth()] + " " + d.getUTCFullYear(); };
    const bucket = s => S.grain === "month" ? s.slice(0, 7) : S.grain === "week" ? monday(s) : s;
    const bucketLabel = b => S.grain === "month" ? MON[+b.slice(5, 7) - 1] + " " + b.slice(0, 4) : S.grain === "week" ? "Wk of " + D(b).getUTCDate() + " " + MON[D(b).getUTCMonth()] : dayLabel(b);
    function setRange(r) {
        const t = today(); S.range = r;
        if (r === "today") { S.from = S.to = t; }
        else if (r === "yesterday") { S.from = S.to = addDays(t, -1); }
        else if (r === "week") { S.from = monday(t); S.to = t; }
        else if (r === "lastweek") { S.from = addDays(monday(t), -7); S.to = addDays(monday(t), -1); }
        else if (r === "month") { S.from = t.slice(0, 8) + "01"; S.to = t; }
        else if (r === "lastmonth") { const d = D(t.slice(0, 8) + "01"); d.setUTCDate(0); S.to = iso(d); S.from = S.to.slice(0, 8) + "01"; }
        else if (r === "fy") { const y = +t.slice(0, 4), m = +t.slice(5, 7); S.from = (m >= 7 ? y : y - 1) + "-07-01"; S.to = t; }
        else if (r === "all") { S.from = (S.meta && S.meta.postedFrom) || "2026-07-01"; S.to = t; }
        // a sensible step for the length of the period
        const days = (D(S.to) - D(S.from)) / 86400000 + 1;
        if (r !== "custom") S.grain = days > 120 ? "month" : days > 35 ? "week" : "day";
    }

    // ---------- names ----------
    const titleCase = s => String(s).toLowerCase().replace(/(^|[\s._-])([a-z])/g, (m, a, b) => a + b.toUpperCase());
    function userLabel(u) {
        if (!u) return "(no user recorded)";
        const n = S.data && S.data.users && S.data.users[u];
        return n && n.toUpperCase() !== u && !/^user_/i.test(n) ? n : (/^USER_[0-9A-F]{20,}$/.test(u) ? "System account" : titleCase(u));
    }
    function branchLabel(b) {
        if (!b) return "(no branch)";
        const n = S.data && S.data.locs && S.data.locs[b];
        return n && n !== b ? b + " · " + n : b;
    }
    const partyName = e => (e.kind === "S" ? S.data.cust[e.party] : S.data.vend[e.party]) || e.party || "(not recorded)";

    // ---------- file → events ----------
    function build(d) {
        const ev = [];
        const byOrder = bag => { const m = new Map(); for (const no in bag) { const x = bag[no]; if (!x.o) continue; const t = m.get(x.o) || { a: 0, c: x.c, b: x.b, s: x.s }; t.a += x.a || 0; m.set(x.o, t); } return m; };
        const siO = byOrder(d.si), piO = byOrder(d.pi), shO = byOrder(d.sh), prO = byOrder(d.pr);
        // Which quote an order was made from. Wiise writes that on the order as it
        // is created (no Change Log line), so it is known for certain only once the
        // order has been removed (x.q, read from its removal line). Until then: a
        // quote removed by the same person within two minutes of the order's first
        // line, for the same customer, is the quote it was made from.
        const mins = ts => Date.UTC(+ts.slice(0, 4), +ts.slice(5, 7) - 1, +ts.slice(8, 10), +ts.slice(11, 13), +ts.slice(14, 16)) / 60000;
        const quoteOf = new Map(), turned = new Set(), goneQuotes = new Map();      // order no → quote no · quotes that became orders · user → removed quotes
        for (const key in d.docs) { const x = d.docs[key]; if (key.startsWith("S1|") && x.q) { quoteOf.set(key.slice(3), x.q); turned.add(x.q); } }
        for (const key in d.docs) { const x = d.docs[key]; if (key.startsWith("S0|") && x.d && !turned.has(key.slice(3))) { const l = goneQuotes.get(x.d[1]) || []; l.push({ no: key.slice(3), m: mins(x.d[0]), cu: x.cu || "" }); goneQuotes.set(x.d[1], l); } }
        for (const key in d.docs) {
            const x = d.docs[key];
            if (!key.startsWith("S1|") || x.q || !x.c) continue;
            const m = mins(x.f[0]); let best = null;
            for (const qd of (goneQuotes.get(x.f[1]) || [])) {
                if (turned.has(qd.no) || Math.abs(qd.m - m) > 2 || (qd.cu && x.cu && qd.cu !== x.cu)) continue;
                if (!best || Math.abs(qd.m - m) < Math.abs(best.m - m)) best = qd;
            }
            if (best) { quoteOf.set(key.slice(3), best.no); turned.add(best.no); }
        }
        for (const key in d.docs) {
            const x = d.docs[key], kind = key[0], type = key[1], no = key.slice(3);
            if (type !== "1" && !(kind === "S" && (type === "0" || type === "4"))) continue;       // orders, sales quotes and blanket sales orders
            const inv = (kind === "S" ? siO : piO).get(no), goods = (kind === "S" ? shO : prO).get(no), linked = inv || goods || {};
            const base = { kind, no, key, party: x.cu || linked.c || "", br: x.br || linked.b || "", sp: x.sp != null ? x.sp : (linked.s || ""),
                amt: x.a != null ? x.a : (inv ? Math.round(inv.a * 100) / 100 : null), au: x.au || "", open: !!x.o, closed: x.d || null, doc: type === "1" ? "order" : type === "4" ? "blanket" : "quote", dropped: !!x.d && !inv && !goods };
            if (x.c) {
                const stamp = { date: x.f[0].slice(0, 10), hour: +x.f[0].slice(11, 13), time: x.f[0].slice(11, 16), user: x.f[1] };
                ev.push(Object.assign({ m: type === "0" ? "sqC" : type === "4" ? "sbC" : kind === "S" ? "soC" : "poC" }, base, stamp));
                if (kind === "S" && type === "1" && x.bo) ev.push(Object.assign({ m: "sbV", blanket: x.bo }, base, stamp));
                if (kind === "S" && type === "1" && quoteOf.has(no)) ev.push(Object.assign({ m: "sqV", quote: quoteOf.get(no), quoteGuess: !x.q }, base, stamp));
            }
            // Deleted = removed from Wiise without ever being shipped, received or
            // invoiced (an order) or without becoming an order (a quote). Counted
            // on the day it was removed, for the person who removed it.
            if (x.d && type !== "4" && (type === "1" ? base.dropped : !turned.has(no))) {
                ev.push(Object.assign({ m: type === "0" ? "sqD" : kind === "S" ? "soD" : "poD", raised: x.c ? x.f : null }, base,
                    { date: x.d[0].slice(0, 10), hour: +x.d[0].slice(11, 13), time: x.d[0].slice(11, 16), user: x.d[1] }));
            }
            if (type === "1" && x.e) for (const ek in x.e) {
                const cut = ek.indexOf("|"), dh = ek.slice(0, cut), fields = x.e[ek][1].map(i => d.fields[i]).filter(f => f && !NOISE.has(f));
                if (!fields.length) continue;
                ev.push(Object.assign({ m: kind === "S" ? "soM" : "poM", date: dh.slice(0, 10), hour: +dh.slice(11, 13), time: dh.slice(11, 13) + ":00", user: ek.slice(cut + 1), n: x.e[ek][0], fields }, base));
            }
        }
        const posted = (bag, m, kind, doc) => {
            for (const no in bag) {
                const x = bag[no];
                if (!x.d) continue;                       // a stub without its ledger line yet
                let amt = x.a, approx = false;
                if (amt == null && x.g != null) { amt = Math.round(x.g / 1.1 * 100) / 100; approx = true; }
                ev.push({ m, kind, doc, no, key: m + "|" + no, date: x.d, hour: -1, time: "", user: x.u || "", party: x.c || "", br: x.b || "", sp: x.s || "",
                    amt: amt == null ? null : amt, approx, qty: x.q, lines: x.l, order: x.o || "", via: x.i || "" });
            }
        };
        posted(d.sh, "sh", "S", "shipment"); posted(d.si, "si", "S", "invoice"); posted(d.sc, "sc", "S", "credit");
        posted(d.pr, "pr", "P", "receipt"); posted(d.pi, "pi", "P", "invoice"); posted(d.pc, "pc", "P", "credit");
        return ev;
    }

    // ---------- loading ----------
    // One download serves this tab and Item Changes (item-changes.js): both wait on the same promise.
    let loadP = null;
    function load(force) {
        if (S.status === "loading" && loadP) return loadP;
        return (loadP = doLoad(force));
    }
    async function doLoad(force) {
        S.status = "loading"; S.err = ""; paint();
        try {
            let meta = {};
            try { const r = await spFetch("snapshots/activity.meta.json"); if (r.ok) meta = await r.json(); else if (r.status === 404) throw new Error("NOFILE"); } catch (e) { if (e.message === "NOFILE") throw e; }
            const own = meta.lock === "activity", store = own ? "activityPassphrase" : "snapshotPassphrase";
            const resp = await spFetch("snapshots/activity.bin");
            if (!resp.ok) throw new Error(resp.status === 404 ? "NOFILE" : "download failed (HTTP " + resp.status + ")");
            const buf = await resp.arrayBuffer();
            let data = null;
            for (let attempt = 0; attempt < 3 && !data; attempt++) {
                let pass = ""; try { pass = (attempt === 0 && localStorage.getItem(store)) || ""; } catch (e) { /* storage blocked */ }
                if (!pass) pass = (prompt(own ? "Enter the Staff Activity passphrase:" : "Enter the snapshot passphrase:") || "").trim();
                if (!pass) throw new Error("no passphrase entered");
                try { data = await decryptSnapshot(buf, pass); try { localStorage.setItem(store, pass); } catch (e) { /* storage blocked */ } }
                catch (e) { try { localStorage.removeItem(store); } catch (e2) { /* storage blocked */ } }
            }
            if (!data) throw new Error("the passphrase did not open the file");
            inject(data);
        } catch (e) {
            S.status = "error";
            S.err = e.message === "NOFILE" ? "The staff activity file has not been built yet. It is written by the robot (staff-activity workflow)." : "Could not load the staff activity file: " + (e.message || e);
            paint();
        }
    }
    // ---------- live top-up: what Wiise holds beyond the robot's file ----------
    // Runs the robot's own rules (activity-engine.js) on a COPY of the file,
    // reading only what is newer than its cursors, and shows the result.
    // Nothing is written anywhere; ↻ goes back to the file as published.
    let liveP = null;
    function live() {
        if (S.status !== "ready") return Promise.resolve();
        if (S.live === "loading" && liveP) return liveP;
        return (liveP = doLive());
    }
    async function doLive() {
        S.live = "loading"; S.liveErr = ""; paint();
        try {
            if (typeof activityTopUp !== "function") throw new Error("activity-engine.js did not load");
            const co = await bcGetCompanyInternalName(), cid = await bcGetCompanyId(), enc = encodeURIComponent, t0 = Date.now();
            const st = typeof structuredClone === "function" ? structuredClone(S.data) : JSON.parse(JSON.stringify(S.data));
            await activityTopUp(st, { light: true,
                odata: (set, filter, select, label) => bcFetchAll(BC_ODATA_URL + "/Company('" + enc(co) + "')/" + set + "?$filter=" + enc(filter) + (select ? "&$select=" + select : ""), label),
                api: (p, label) => bcFetchAll(BC_API_URL + "/companies(" + cid + ")/" + p, label),
                warn: m => console.warn("[Staff activity live]", m) });
            st.meta.liveAt = actMel(new Date()).replace("T", " ");
            st.meta.liveSecs = Math.round((Date.now() - t0) / 100) / 10;
            S.data = st; S.meta = st.meta; S.ev = build(st); S.live = "done";
        } catch (e) {
            S.live = "error"; S.liveErr = "Could not read the latest from Wiise: " + (e.message || e);
        }
        const prog = document.getElementById("bcProgress"); if (prog) prog.style.display = "none";    // the corner progress line bcFetchAll switches on
        paint();
    }
    function inject(data) {
        S.live = ""; S.liveErr = "";
        S.data = data; S.meta = data.meta || {}; S.ev = build(data); S.status = "ready";
        setRange(S.range === "custom" ? "month" : S.range);
        paint();
    }

    // ---------- filtering + aggregation ----------
    const sideOf = e => MBY[e.m].side;
    // A filter list also offers "the ones with nothing" (no branch, no salesperson).
    // That choice needs a value of its own: an empty value means "all".
    const NONE = "\u2205";
    const hit = (sel, v) => sel === NONE ? !v : v === sel;
    // the levels chosen under View by: no repeats, none equal to the view, stops at the first blank
    function thenList() {
        const out = [];
        for (const t of S.thens) { if (!t || t === S.view || out.includes(t)) break; out.push(t); }
        return out;
    }
    function filtered(opts) {
        opts = opts || {};
        const q = S.party.trim().toLowerCase();
        return S.ev.filter(e => {
            if (e.date < S.from || e.date > S.to) return false;
            if (S.side !== "all" && sideOf(e) !== S.side) return false;
            if (!opts.noBranch && S.branch && !hit(S.branch, e.br)) return false;
            if (!opts.noUser && S.user && !hit(S.user, e.user)) return false;
            if (!opts.noSp && S.sp && !hit(S.sp, e.sp)) return false;
            if (q && !(String(e.party).toLowerCase().includes(q) || partyName(e).toLowerCase().includes(q))) return false;
            return true;
        });
    }
    const measures = () => MEAS.filter(m => S.side === "all" || m.side === S.side);
    function keyOf(e, dim) {
        switch (dim) {
            case "user":   return [e.user || "", userLabel(e.user)];
            case "branch": return [e.br || "", branchLabel(e.br)];
            case "party":  return [e.kind + "|" + e.party, partyName(e) + (e.kind === "P" ? " (supplier)" : "")];
            case "sp":     return [e.sp || "", e.sp || "(none)"];
            case "period": return [bucket(e.date), bucketLabel(bucket(e.date))];
        }
        return ["", ""];
    }
    function node(key, label) { return { key, label, m: {}, parties: new Set(), kids: new Map(), n: 0 }; }
    function add(n, e) {
        const b = n.m[e.m] || (n.m[e.m] = new Map());
        if (!b.has(e.key)) b.set(e.key, e.amt || 0);      // one document counts once per row, however many times it was touched
        if (e.party) n.parties.add(e.kind + "|" + e.party);
        n.n++;
    }
    const cnt = (n, k) => n.m[k] ? n.m[k].size : 0;
    const amt = (n, k) => { let s = 0; if (n.m[k]) for (const v of n.m[k].values()) s += v; return s; };
    const val = (n, k) => S.show === "amount" ? amt(n, k) : cnt(n, k);
    const totalOf = n => measures().reduce((s, m) => s + (m.gone ? 0 : val(n, m.k)), 0);      // deleted documents are shown, not added to the total
    const fmtVal = v => S.show === "amount" ? (v ? money0(v) : "—") : (v ? fmtInt(v) : "—");

    // ---------- one-off page build ----------
    function scaffold() {
        const root = $("saRoot");
        if (!root || root.dataset.built) return;
        if (!$("saStyle")) {
            const st = document.createElement("style"); st.id = "saStyle";
            st.textContent = `
                /* every tile on ONE line: as many equal columns as there are tiles; the figure shrinks with its tile */
                #saTiles { grid-template-columns: repeat(var(--sa-n, 6), minmax(0, 1fr)); gap: 10px; }
                #saTiles .kpi { padding: 11px 12px; cursor: pointer; transition: transform .15s, opacity .15s; display: flex; flex-direction: column; container-type: inline-size; min-width: 0; }
                #saTiles .kpi:hover { transform: translateY(-1px); }
                #saTiles .kpi .label { font-size: 9.5px; letter-spacing: .4px; line-height: 1.3; margin-bottom: auto; }
                #saTiles .kpi .value { font-size: 22px; font-size: clamp(15px, 17cqw, 26px); white-space: nowrap; }
                #saTiles .kpi .sub { font-size: 10.5px; line-height: 1.35; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
                @media (max-width: 1100px) { #saTiles { grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); } }
                .sa-tiles.has-pick .kpi:not(.active) { opacity: .5; }
                .sa-tiles .kpi.active { outline: 3px solid #ffd54f; outline-offset: 1px; }
                /* panel headings: title and its small print side by side on one line of text, controls at the right */
                #saRoot .panel-header:not(.sa-bar) { justify-content: flex-start; align-items: baseline; }
                .sa-bar { display: flex; flex-wrap: wrap; align-items: center; justify-content: flex-start; gap: 8px 10px; }
                #saDocs th:last-child, #saDocs td:last-child { min-width: 220px; }
                #saDocs td:nth-child(6) { max-width: 210px; overflow: hidden; text-overflow: ellipsis; }
                .sa-bar .lbl { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .5px; opacity: .7; margin-left: 6px; }
                .sa-bar .lbl:first-child { margin-left: 0; }
                .sa-bar input[type=date], .sa-bar input[type=search] { font: inherit; font-size: 12px; }
                .sa-ctl { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; margin-left: auto; }
                .sa-ctl .lbl { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .5px; opacity: .7; margin-left: 8px; }
                /* the activity table on one page: tight rows, headings on two lines, no sideways scroll; it only scrolls inside its box once it is taller than the screen */
                #saPanel .table-scroll { max-height: calc(100vh - 110px); }
                #saTable { width: 100%; }
                #saTable th { white-space: normal; font-size: 9.5px; letter-spacing: .3px; line-height: 1.25; padding: 6px 6px; vertical-align: bottom; }
                #saTable td { padding: 2px 6px; font-size: 12px; line-height: 1.4; white-space: nowrap; }
                #saTable th:first-child, #saTable td:first-child { max-width: 190px; overflow: hidden; text-overflow: ellipsis; padding-left: 12px; }
                #saTable td.sa-cell { cursor: pointer; }
                #saTable td.sa-cell:hover { text-decoration: underline; }
                #saTable td.sa-on, #saTable th.sa-on { outline: 2px solid #ffd54f; outline-offset: -2px; }
                #saDocs td, #saHistTable td { white-space: nowrap; }
                #saDocs th, #saHistTable th { cursor: pointer; }
                #saDocs td.wrap, #saHistTable td.wrap { white-space: normal; }
                .sa-link { cursor: pointer; text-decoration: underline dotted; text-underline-offset: 3px; }
                .sa-note { font-size: 11px; opacity: .75; padding: 8px 14px 12px; }
                .sa-heat { border-collapse: collapse; width: 100%; table-layout: fixed; }
                .sa-heat th, .sa-heat td { font-size: 11px; text-align: center; padding: 6px 2px; border: 1px solid rgba(128,128,128,.18); }
                .sa-heat th { font-weight: 700; opacity: .8; }
                .sa-heat td.sa-on { outline: 2px solid #ffd54f; outline-offset: -2px; }
                .sa-chip { display: inline-block; font-size: 10px; padding: 1px 6px; border-radius: 8px; background: rgba(128,128,128,.18); margin-left: 6px; }
                .sa-empty { padding: 28px 16px; text-align: center; opacity: .75; }
            `;
            document.head.appendChild(st);
        }
        const opt = (v, l) => `<option value="${v}">${l}</option>`;
        root.innerHTML = `
            <div class="hint rr-hint"><span id="saHint">Who raised, changed, shipped, invoiced and received what.</span>
                <span class="rr-tgt"><span class="rr-tgt-lbl">Activity file</span><span id="saAsOf">not loaded</span><button type="button" class="more-btn" id="saReload" title="Read the latest staff activity file again">↻</button><button type="button" class="more-btn" id="saLive" title="Read everything newer than the file straight from Wiise, with your own sign-in">Get latest from Wiise</button></span></div>
            <div class="panel"><div class="panel-header sa-bar">
                <span class="lbl">Period</span>
                <select id="saRange" class="mini-select" title="Quick period">
                    ${opt("today", "Today")}${opt("yesterday", "Yesterday")}${opt("week", "This week")}${opt("lastweek", "Last week")}${opt("month", "This month")}${opt("lastmonth", "Last month")}${opt("fy", "This financial year")}${opt("all", "Everything held")}${opt("custom", "Custom dates")}
                </select>
                <input type="date" id="saFrom" class="mini-select" title="From"><span>to</span><input type="date" id="saTo" class="mini-select" title="To">
                <span class="lbl">Step</span>
                <button type="button" class="gp-filter-btn" data-sa-grain="day">Daily</button><button type="button" class="gp-filter-btn" data-sa-grain="week">Weekly</button><button type="button" class="gp-filter-btn" data-sa-grain="month">Monthly</button>
                <span class="lbl">Side</span>
                <button type="button" class="gp-filter-btn" data-sa-side="sales">Sales</button><button type="button" class="gp-filter-btn" data-sa-side="purch">Purchasing</button><button type="button" class="gp-filter-btn" data-sa-side="all">Both</button>
                <span class="lbl">Show</span>
                <button type="button" class="gp-filter-btn" data-sa-show="count">Documents</button><button type="button" class="gp-filter-btn" data-sa-show="amount">Amounts</button>
            </div><div class="panel-header sa-bar">
                <span class="lbl">Filter</span>
                <select id="saBranch" class="mini-select" title="Branch"></select>
                <select id="saUser" class="mini-select" title="Employee"></select>
                <select id="saSp" class="mini-select" title="Salesperson or purchaser code on the document"></select>
                <input type="search" id="saParty" class="mini-select" placeholder="Customer or supplier…" title="Part of a customer or supplier name or number">
                <button type="button" class="more-btn" id="saClear" title="Clear the filters (keeps the period)">Clear filters</button>
            </div><div class="panel-header sa-bar" style="border-bottom:0;" id="saViewTop">
                <span class="lbl">View by</span>${VIEWS.map(v => `<button type="button" class="gp-filter-btn" data-sa-view="${v[0]}">${v[1]}</button>`).join("")}
                <span class="lbl">Then by</span><select id="saThen" data-sa-then="0" class="mini-select" title="Second level: click a row to open it"></select><select id="saThen2" data-sa-then="1" class="mini-select" title="Third level: open a row, then open one of its rows"></select><select id="saThen3" data-sa-then="2" class="mini-select" title="Fourth level"></select>
            </div></div>
            <div class="kpi-row sa-tiles" id="saTiles"></div>
            <div class="panel" id="saPanel">
                <div class="panel-header"><h2 id="saTableTitle">Activity by employee</h2><div class="meta" id="saMeta"></div>
                </div>
                <div class="table-scroll"><table id="saTable" class="ds-table"><thead></thead><tbody></tbody></table></div>
                <div class="sa-note" id="saTableNote"></div>
            </div>
            <div class="panel">
                <div class="panel-header"><h2 id="saChartTitle">Charts</h2>
                    <div class="sa-ctl">
                        <select id="saChartSel" class="mini-select" title="Chart">${CHARTS.map(c => opt(c[0], c[1])).join("")}</select>
                        <select id="saChartMeasure" class="mini-select" title="Activity shown in the chart"></select>
                    </div>
                </div>
                <div class="chart-wrap" id="saChartWrap" style="position:relative; min-height:380px; height:380px; padding:10px 14px 14px;"><canvas id="saChart"></canvas></div>
                <div id="saHeat" style="display:none; padding:10px 14px 14px;"></div>
                <div class="sa-note" id="saChartNote"></div>
                <div id="saOthers" style="padding:0 14px 12px;"></div>
            </div>
            <div class="panel" id="saDocsPanel">
                <div class="panel-header"><h2 id="saDocsTitle">Documents</h2><div class="meta" id="saDocsMeta"></div>
                    <div class="sa-ctl"><button type="button" class="more-btn" id="saPickClear" style="display:none;">Show all documents</button>
                        <input type="search" id="saDocQ" class="mini-select" placeholder="Search document, employee, customer…" title="Search the list below"></div>
                </div>
                <div class="table-scroll"><table id="saDocs" class="ds-table"><thead></thead><tbody></tbody></table></div>
                <div class="sa-note" id="saDocsNote"></div>
            </div>
            <div class="panel" id="saHistPanel" style="display:none;">
                <div class="panel-header"><h2 id="saHistTitle">Order history</h2><div class="meta" id="saHistMeta"></div>
                    <div class="sa-ctl"><button type="button" class="more-btn" id="saHistLive">Read every change from Wiise</button><button type="button" class="more-btn" id="saHistClose">Close</button></div>
                </div>
                <div class="table-scroll"><table id="saHistTable" class="ds-table"><thead></thead><tbody></tbody></table></div>
                <div class="sa-note" id="saHistNote"></div>
            </div>`;
        root.dataset.built = "1";

        // ---- wiring ----
        $("saReload").addEventListener("click", () => load(true));
        $("saLive").addEventListener("click", () => live());
        $("saRange").addEventListener("change", e => { if (e.target.value !== "custom") setRange(e.target.value); else S.range = "custom"; S.pick = null; paint(); });
        const dateChange = () => { const f = $("saFrom").value, t = $("saTo").value; if (!f || !t) return; S.from = f <= t ? f : t; S.to = f <= t ? t : f; S.range = "custom"; S.pick = null; paint(); };
        $("saFrom").addEventListener("change", dateChange); $("saTo").addEventListener("change", dateChange);
        $("saBranch").addEventListener("change", e => { S.branch = e.target.value; S.pick = null; paint(); });
        $("saUser").addEventListener("change", e => { S.user = e.target.value; S.pick = null; paint(); });
        $("saSp").addEventListener("change", e => { S.sp = e.target.value; S.pick = null; paint(); });
        $("saParty").addEventListener("input", e => { S.party = e.target.value; S.pick = null; paint(); });
        $("saClear").addEventListener("click", () => { S.branch = S.user = S.sp = S.party = ""; S.pick = null; paint(); });
        root.addEventListener("change", e => {
            const t = e.target.closest("[data-sa-then]"); if (!t) return;
            const i = +t.dataset.saThen; S.thens[i] = t.value;
            for (let j = i + 1; j < S.thens.length; j++) if (!t.value || S.thens[j] === t.value) S.thens[j] = "";      // a blank level ends the chain; no level twice
            S.thens = thenList().concat(["", "", ""]).slice(0, 3);
            paint();
        });
        $("saChartSel").addEventListener("change", e => { S.chart = e.target.value; paint(); });
        $("saChartMeasure").addEventListener("change", e => { S.cm = e.target.value; paint(); });
        $("saDocQ").addEventListener("input", e => { S.docQ = e.target.value; S.docLimit = 300; paintDocs(); });
        $("saPickClear").addEventListener("click", () => { S.pick = null; paint(); });
        $("saHistClose").addEventListener("click", () => { S.hist = null; paintHist(); });
        $("saHistLive").addEventListener("click", histLive);
        root.addEventListener("click", e => {
            if (e.target.closest("a")) return;
            let b;
            if ((b = e.target.closest("[data-sa-grain]"))) { S.grain = b.dataset.saGrain; paint(); return; }
            if ((b = e.target.closest("[data-sa-side]"))) { S.side = b.dataset.saSide; S.pick = null; S.cm = ""; S.sort = { key: "__total", dir: "desc" }; paint(); return; }
            if ((b = e.target.closest("[data-sa-show]"))) { S.show = b.dataset.saShow; paint(); return; }
            if ((b = e.target.closest("[data-sa-view]"))) { S.view = b.dataset.saView; S.thens = S.thens.filter(t => t && t !== S.view).concat(["", "", ""]).slice(0, 3); S.pick = null; S.sort = { key: S.view === "period" ? "__key" : "__total", dir: S.view === "period" ? "asc" : "desc" }; paint(); return; }
            if ((b = e.target.closest("[data-sa-tile]"))) { const k = b.dataset.saTile; S.pick = S.pick && S.pick.m === k && !S.pick.row ? null : { m: k, label: MBY[k].label }; S.cm = k; paint(); return; }
            if ((b = e.target.closest("th[data-sa-dsort]"))) {
                const k = b.dataset.saDsort;
                S.docSort = S.docSort.key === k ? { key: k, dir: S.docSort.dir === "asc" ? "desc" : "asc" } : { key: k, dir: k === "date" || k === "qty" || k === "amt" ? "desc" : "asc" };
                paintDocs(); return;
            }
            if ((b = e.target.closest("th[data-sa-hsort]"))) {
                const k = b.dataset.saHsort;
                S.histSort = S.histSort.key === k ? { key: k, dir: S.histSort.dir === "asc" ? "desc" : "asc" } : { key: k, dir: "asc" };
                paintHist(); return;
            }
            if ((b = e.target.closest("th[data-sa-osort]"))) {
                const k = b.dataset.saOsort;
                S.othSort = S.othSort.key === k ? { key: k, dir: S.othSort.dir === "asc" ? "desc" : "asc" } : { key: k, dir: k === "v" ? "desc" : "asc" };
                paint(); return;
            }
            if ((b = e.target.closest("[data-sa-oth-toggle]"))) { S.othOpen = !S.othOpen; if (!S.othOpen) S.othUser = ""; paint(); return; }
            if ((b = e.target.closest("[data-sa-oth-all]"))) { S.othUser = ""; paint(); return; }
            if ((b = e.target.closest("tr[data-sa-oth-row]"))) {
                S.pick = { id: "oth|" + b.dataset.saOthRow, m: S.cm, row: JSON.parse(b.dataset.saOthRow), label: b.dataset.label + " · " + MBY[S.cm].label };
                paint(); $("saDocsPanel").scrollIntoView({ behavior: "smooth", block: "nearest" }); return;
            }
            if ((b = e.target.closest("th[data-sa-sort]"))) {
                const k = b.dataset.saSort;
                if (S.sort.key === k) S.sort.dir = S.sort.dir === "asc" ? "desc" : "asc"; else S.sort = { key: k, dir: k === "__label" || k === "__key" ? "asc" : "desc" };
                paint(); return;
            }
            if ((b = e.target.closest("td[data-sa-heat]"))) {      // a day-and-hour figure: list what was raised or changed then
                const id = "heat|" + b.dataset.saHeat, same = S.pick && S.pick.id === id, [wd, h] = b.dataset.saHeat.split("|");
                S.pick = same ? null : { id, m: "", row: [], heat: [+wd, h === "" ? -1 : +h], label: b.dataset.label + " · raised or changed" };
                paint();
                if (!same) $("saDocsPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
                return;
            }
            if ((b = e.target.closest("td.sa-cell"))) {
                const tr = b.closest("tr"), m = b.dataset.m || "";
                const same = S.pick && S.pick.id === tr.dataset.key + "|" + m;
                S.pick = same ? null : { id: tr.dataset.key + "|" + m, m, row: tr.dataset.row ? JSON.parse(tr.dataset.row) : [], label: (tr.dataset.label || "All") + (m ? " · " + MBY[m].label : "") };
                paint();
                if (!same) $("saDocsPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
                return;
            }
            if ((b = e.target.closest("[data-sa-hist]"))) { S.hist = { kind: b.dataset.saKind, no: b.dataset.saHist, live: null }; paintHist(); $("saHistPanel").scrollIntoView({ behavior: "smooth", block: "nearest" }); return; }
            if ((b = e.target.closest("[data-sa-more]"))) { S.docLimit += 500; paintDocs(); return; }
            if ((b = e.target.closest("tr[data-key]"))) {
                if (b.dataset.expandable !== "1") return;
                const k = b.dataset.key;
                if (state.expandedMatrix.has(k)) state.expandedMatrix.delete(k); else state.expandedMatrix.add(k);
                paint();
            }
        });
    }

    // ---------- painting ----------
    let filteredCache = [];
    function paint() {
        if (!allowed()) return;
        scaffold();
        if (!$("saRoot") || !$("saRoot").dataset.built) return;
        const ready = S.status === "ready";
        const fileAsOf = () => "as of " + String(S.meta.builtAtMelbourne || "").slice(11) + " on " + dateLabel(String(S.meta.builtAtMelbourne || "").slice(0, 10));
        $("saAsOf").textContent = S.status === "loading" ? "loading…" : S.status === "error" ? "not loaded" : !ready ? "not loaded"
            : S.live === "loading" ? fileAsOf() + " · reading Wiise…"
            : S.live === "error" ? fileAsOf() + " · live read failed"
            : S.meta.liveAt ? "LIVE from Wiise at " + S.meta.liveAt.slice(11) + " (" + S.meta.liveSecs + " s)" : fileAsOf();
        $("saAsOf").title = S.liveErr || (S.meta && S.meta.liveAt ? "The robot's file " + fileAsOf() + ", plus everything Wiise holds since. On screen only; ↻ goes back to the file." : "");
        $("saLive").disabled = !ready || S.live === "loading";
        $("saLive").textContent = S.live === "loading" ? "Reading Wiise…" : "Get latest from Wiise";
        $("saHint").innerHTML = S.status === "error" ? esc(S.err)
            : !ready ? "Loading the staff activity file…"
            : (S.live === "error" ? "<strong>" + esc(S.liveErr) + "</strong> Showing the robot's file. " : "") + "Who raised, changed, shipped, invoiced and received what. <strong>Raised</strong> and <strong>changed</strong> come from the Wiise Change Log, held from "
              + dateLabel(String(S.meta.logFrom || "").slice(0, 10)) + "; <strong>posted</strong> documents are held from " + dateLabel(S.meta.postedFrom || "") + ". The toolbar filters above do not apply here; use the ones below. Click a tile or any figure to list its documents.";
        ["saRange", "saFrom", "saTo", "saBranch", "saUser", "saSp", "saParty", "saThen", "saThen2", "saThen3", "saChartSel", "saChartMeasure", "saDocQ"].forEach(id => { $(id).disabled = !ready; });
        if (!ready) {
            $("saTiles").innerHTML = ""; $("saTable").querySelector("thead").innerHTML = ""; $("saTable").querySelector("tbody").innerHTML = `<tr><td class="empty">${esc(S.status === "error" ? S.err : "Loading…")}</td></tr>`;
            $("saDocs").querySelector("tbody").innerHTML = ""; destroyChart("saChart"); return;
        }
        $("saRange").value = S.range; $("saFrom").value = S.from; $("saTo").value = S.to;
        const root = $("saRoot");
        root.querySelectorAll("[data-sa-grain]").forEach(b => b.classList.toggle("active", b.dataset.saGrain === S.grain));
        root.querySelectorAll("[data-sa-side]").forEach(b => b.classList.toggle("active", b.dataset.saSide === S.side));
        root.querySelectorAll("[data-sa-show]").forEach(b => b.classList.toggle("active", b.dataset.saShow === S.show));
        root.querySelectorAll("[data-sa-view]").forEach(b => b.classList.toggle("active", b.dataset.saView === S.view));
        if ($("saParty").value !== S.party) $("saParty").value = S.party;

        // filter lists: whoever / whatever has something in the period without that filter
        const fill = (id, rows, dim, cur, allLabel) => {
            const seen = new Map(); for (const e of rows) { const [k, l] = keyOf(e, dim); if (!seen.has(k)) seen.set(k, l); }
            if (cur && cur !== NONE && !seen.has(cur)) seen.set(cur, cur);
            $(id).innerHTML = `<option value=""${cur ? "" : " selected"}>${allLabel}</option>` + [...seen].sort((a, b) => a[1].localeCompare(b[1], undefined, { numeric: true }))
                .map(([v, l]) => `<option value="${escapeAttr(v || NONE)}"${(v || NONE) === cur ? " selected" : ""}>${esc(l)}</option>`).join("");
        };
        fill("saBranch", filtered({ noBranch: true }), "branch", S.branch, "All branches");
        fill("saUser", filtered({ noUser: true }), "user", S.user, "All employees");
        fill("saSp", filtered({ noSp: true }), "sp", S.sp, "All salespeople / purchasers");
        const levels = thenList();
        root.querySelectorAll("[data-sa-then]").forEach(el => {
            const i = +el.dataset.saThen, taken = levels.slice(0, i);
            el.disabled = false;
            el.style.display = i === 0 || levels.length >= i ? "" : "none";       // the next level appears once the one before is chosen
            el.innerHTML = `<option value="">${i === 0 ? "(nothing)" : "(no further level)"}</option>` + VIEWS.filter(v => v[0] !== S.view && !taken.includes(v[0])).map(v => `<option value="${v[0]}"${v[0] === levels[i] ? " selected" : ""}>${v[1]}</option>`).join("");
        });
        const ms = measures();
        if (!ms.some(m => m.k === S.cm)) S.cm = ms[0].k;
        $("saChartMeasure").innerHTML = ms.map(m => `<option value="${m.k}"${m.k === S.cm ? " selected" : ""}>${m.label}</option>`).join("");
        $("saChartSel").value = S.chart;

        const rows = filteredCache = filtered();
        paintTiles(rows, ms); paintTable(rows, ms); paintChart(rows, ms); paintDocs(); paintHist();
    }

    function paintTiles(rows, ms) {
        const tot = node("tot", "Total"); rows.forEach(e => add(tot, e));
        const tilePick = S.pick && !S.pick.row ? S.pick.m : "";
        $("saTiles").classList.toggle("has-pick", !!tilePick); $("saTiles").style.setProperty("--sa-n", ms.length > 11 ? Math.ceil(ms.length / 2) : ms.length);      // Sales + Purchasing together: two lines
        $("saTiles").innerHTML = ms.map(m => {
            const c = cnt(tot, m.k), a = amt(tot, m.k), people = new Set(rows.filter(e => e.m === m.k).map(e => e.user)).size;
            return `<div class="kpi ${m.tint}${tilePick === m.k ? " active" : ""}" data-sa-tile="${m.k}" title="Click to list these documents">
                <div class="label">${esc(m.label)}</div><div class="value">${S.show === "amount" ? money0(a) : fmtInt(c)}</div>
                <div class="sub">${S.show === "amount" ? plural(c, "document") : (a ? money0(a) : "no value held")}<br>${plural(people, "employee")}</div></div>`;
        }).join("");
    }

    function paintTable(rows, ms) {
        const viewName = (VIEWS.find(v => v[0] === S.view) || [])[1] || "";
        const levels = thenList(), dims = [S.view].concat(levels);
        $("saTableTitle").textContent = "Activity by " + viewName.toLowerCase() + levels.map(t => ", then " + (VIEWS.find(v => v[0] === t) || [])[1].toLowerCase()).join("");
        const top = new Map(), tot = node("tot", "Total");
        for (const e of rows) {
            add(tot, e);
            let map = top, path = "sa", row = [];
            for (const dim of dims) {                      // one node per level, each inside the one above
                const [k, l] = keyOf(e, dim);
                path += "|" + dim + "|" + k; row = row.concat([[dim, k]]);
                let n = map.get(path);
                if (!n) map.set(path, n = Object.assign(node(path, l), { row, sortKey: k }));
                add(n, e); map = n.kids;
            }
        }
        const sv = n => S.sort.key === "__label" ? n.label.toLowerCase() : S.sort.key === "__key" ? n.sortKey : S.sort.key === "__total" ? totalOf(n) : S.sort.key === "__parties" ? n.parties.size : val(n, S.sort.key);
        const ordered = list => list.slice().sort((a, b) => {
            const x = sv(a), y = sv(b), m = S.sort.dir === "asc" ? 1 : -1;
            const c = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true });
            return c * m || a.label.localeCompare(b.label);
        });
        const th = (k, l, numCol, title) => `<th data-sa-sort="${k}" style="cursor:pointer;"${title ? ` title="${escapeAttr(title)}"` : ""} class="${numCol ? "num " : ""}${S.sort.key === k ? "sorted-" + S.sort.dir : ""}">${esc(l)}</th>`;
        $("saTable").querySelector("thead").innerHTML = "<tr>" + th(S.view === "period" ? "__key" : "__label", viewName, false)
            + ms.map(m => th(m.k, m.short, true, m.label + (S.show === "amount" ? " — value excl. tax" : " — number of documents"))).join("")
            + th("__total", S.show === "amount" ? "All value" : "All documents", true, "The columns added together, leaving out the deleted ones")
            + th("__parties", S.side === "purch" ? "Suppliers" : S.side === "sales" ? "Customers" : "Customers / suppliers", true, "How many different customers or suppliers these documents were for") + "</tr>";
        const cells = (n) => ms.map(m => { const v = val(n, m.k); return `<td class="num${v ? " sa-cell" : ""}${S.pick && S.pick.id === n.key + "|" + m.k ? " sa-on" : ""}" data-m="${m.k}">${fmtVal(v)}</td>`; }).join("")
            + `<td class="num sa-cell${S.pick && S.pick.id === n.key + "|" ? " sa-on" : ""}" data-m="">${fmtVal(totalOf(n))}</td><td class="num">${fmtInt(n.parties.size)}</td>`;
        const acc = [];
        const emit = (n, depth) => {
            const kids = [...n.kids.values()], expandable = kids.length > 0, open = state.expandedMatrix.has(n.key);
            acc.push(`<tr data-key="${escapeAttr(n.key)}" data-expandable="${expandable ? 1 : 0}" data-label="${escapeAttr(n.label)}" data-row='${escapeAttr(JSON.stringify(n.row))}' style="${treeRowStyle(depth)}cursor:${expandable ? "pointer" : "default"};">
                <td title="${escapeAttr(n.label)}">${"&nbsp;".repeat(depth * 4)}<span style="display:inline-block; width:14px; color:var(--tree-tint);">${expandable ? (open ? "▾" : "▸") : ""}</span> ${esc(n.label)}</td>${cells(n)}</tr>`);
            if (expandable && open) ordered(kids).forEach(c => emit(c, depth + 1));
        };
        ordered([...top.values()]).forEach(n => emit(n, 0));
        $("saTable").querySelector("tbody").innerHTML = acc.length
            ? acc.join("") + `<tr class="ds-total" data-key="sa|total" data-label="All" data-row="[]"><td>Total</td>${cells(tot)}</tr>`
            : `<tr><td colspan="${ms.length + 3}" class="empty">Nothing recorded for this period and these filters.</td></tr>`;
        $("saMeta").textContent = rows.length ? fmtInt(top.size) + " " + ({ user: ["employee", "employees"], branch: ["branch", "branches"], party: ["customer / supplier", "customers / suppliers"], sp: ["salesperson / purchaser", "salespeople / purchasers"], period: ["period", "periods"] })[S.view][top.size === 1 ? 0 : 1] + " · " + dateLabel(S.from) + " to " + dateLabel(S.to) : "";
        $("saTableNote").textContent = "Blankets → orders counts each sales order made from a blanket order, for the person who made it; the link is read while the order is open, so orders already closed before 26 Nov 2026 are not in it. Deleted = removed from Wiise without a shipment, receipt or invoice (an order) or without becoming an order (a quote); it is counted for the person who removed it and is not added into the All column. A document counts once per row however many times it was touched. “Changed” leaves out the first hour of the person who raised the order (that is part of raising it) and the fields Wiise fills in by itself on release and posting. Click a figure to list its documents"
            + (levels.length ? "; click a row to open it" + (levels.length > 1 ? ", then a row inside it to go a level deeper." : ".") : ".");
    }

    // ---------- charts ----------
    function paintChart(rows, ms) {
        destroyChart("saChart");
        let oth = null;      // { what, col, rows: [{ label, v, row }], user } — what the grey "Others" part is made of
        const heat = S.chart === "heat";
        $("saChartWrap").style.display = heat ? "none" : ""; $("saHeat").style.display = heat ? "" : "none";
        $("saChartMeasure").style.display = (S.chart === "rank" || heat) ? "none" : "";
        const m = MBY[S.cm], unit = S.show === "amount" ? "value excl. tax" : "documents";
        const colour = i => ENTITY_PALETTE[i % ENTITY_PALETTE.length];
        const tickFmt = v => S.show === "amount" ? fmtMoney(v) : fmtInt(v);
        const dimLabel = (VIEWS.find(v => v[0] === S.view) || [])[1].toLowerCase();
        $("saChartNote").textContent = "";
        if (heat) { paintOthers(null); return paintHeat(rows); }
        // group helper: Map key → node, for one measure or all
        const group = (list, dim) => { const g = new Map(); for (const e of list) { const [k, l] = keyOf(e, dim); if (!g.has(k)) g.set(k, node(k, l)); add(g.get(k), e); } return g; };
        let cfg = null;
        if (S.chart === "trend") {
            const stackDim = S.view === "period" ? "user" : S.view;
            const list = rows.filter(e => e.m === m.k), g = group(list, stackDim);
            const keys = [...g.values()].sort((a, b) => val(b, m.k) - val(a, m.k)), topN = keys.slice(0, 9), rest = keys.slice(9);
            const buckets = []; { let b = bucket(S.from); const end = bucket(S.to); for (let i = 0; i < 400 && b <= end; i++) { buckets.push(b); b = S.grain === "month" ? iso(new Date(Date.UTC(+b.slice(0, 4), +b.slice(5, 7), 1))).slice(0, 7) : addDays(b, S.grain === "week" ? 7 : 1); } }
            const cell = new Map();   // series|bucket → node
            for (const e of list) { const [k] = keyOf(e, stackDim), s = topN.some(n => n.key === k) ? k : "__other", ck = s + "|" + bucket(e.date); if (!cell.has(ck)) cell.set(ck, node(ck, "")); add(cell.get(ck), e); }
            const ds = topN.map((n, i) => ({ label: n.label, stack: "s", backgroundColor: colour(i), borderRadius: 2, data: buckets.map(b => { const c = cell.get(n.key + "|" + b); return c ? val(c, m.k) : 0; }) }));
            if (rest.length) ds.push({ label: "Others (" + rest.length + ")", stack: "s", backgroundColor: "#9e9e9e", borderRadius: 2, data: buckets.map(b => { const c = cell.get("__other|" + b); return c ? val(c, m.k) : 0; }) });
            if (rest.length) oth = { what: rest.length + " more, whole period", col: (VIEWS.find(v => v[0] === stackDim) || [])[1], rows: rest.map(n => ({ label: n.label, v: val(n, m.k), row: [[stackDim, n.key]] })) };
            cfg = { type: "bar", data: { labels: buckets.map(bucketLabel), datasets: ds },
                options: chartBase({ plugins: { legend: { position: "bottom" }, tooltip: { mode: "index", intersect: false, filter: i => i.parsed.y !== 0, callbacks: { label: c => c.dataset.label + ": " + tickFmt(c.parsed.y), footer: it => "Total: " + tickFmt(it.reduce((s, i) => s + i.parsed.y, 0)) } } },
                    scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, ticks: { callback: tickFmt, precision: 0 } } },
                    onClick: (ev, els) => { if (els.length && rest.length && els[0].datasetIndex === topN.length) { S.othOpen = true; paint(); } } }) };
            $("saChartTitle").textContent = m.label + " — " + unit + " per " + (S.grain === "day" ? "day" : S.grain) + ", by " + (VIEWS.find(v => v[0] === stackDim) || [])[1].toLowerCase();
        } else if (S.chart === "rank") {
            const dim = S.view === "period" ? "user" : S.view, g = group(rows, dim);
            const keys = [...g.values()].sort((a, b) => totalOf(b) - totalOf(a)).slice(0, 20);
            cfg = { type: "bar", data: { labels: keys.map(n => truncate(n.label, 28)), datasets: ms.map((mm, i) => ({ label: mm.short, stack: "s", backgroundColor: colour(i), borderRadius: 2, data: keys.map(n => val(n, mm.k)) })) },
                options: chartBase({ indexAxis: "y", plugins: { legend: { position: "bottom" }, tooltip: { mode: "index", axis: "y", intersect: false, filter: i => i.parsed.x !== 0, callbacks: { label: c => c.dataset.label + ": " + tickFmt(c.parsed.x) } } },
                    scales: { x: { stacked: true, beginAtZero: true, ticks: { callback: tickFmt, precision: 0 } }, y: { stacked: true } },
                    onClick: (ev, els) => { if (!els.length) return; const n = keys[els[0].index]; if (dim === "user") S.user = S.user === n.key ? "" : n.key; else if (dim === "branch") S.branch = S.branch === n.key ? "" : n.key; else return; S.pick = null; paint(); } }) };
            $("saChartTitle").textContent = "Every activity by " + (VIEWS.find(v => v[0] === dim) || [])[1].toLowerCase() + " — " + unit + (keys.length === 20 ? " (top 20)" : "");
            if (dim === "user" || dim === "branch") $("saChartNote").textContent = "Click a bar to filter the whole tab to that " + (dim === "user" ? "employee" : "branch") + "; click it again to clear.";
        } else if (S.chart === "share") {
            const dim = S.view === "period" ? "user" : S.view, g = group(rows.filter(e => e.m === m.k), dim);
            const keys = [...g.values()].filter(n => val(n, m.k) > 0).sort((a, b) => val(b, m.k) - val(a, m.k)), topN = keys.slice(0, 11), rest = keys.slice(11);
            const labels = topN.map(n => n.label), data = topN.map(n => val(n, m.k));
            if (rest.length) { labels.push("Others (" + rest.length + ")"); data.push(rest.reduce((s, n) => s + val(n, m.k), 0)); }
            if (rest.length) oth = { what: rest.length + " more", col: (VIEWS.find(v => v[0] === dim) || [])[1], rows: rest.map(n => ({ label: n.label, v: val(n, m.k), row: [[dim, n.key]] })) };
            const sum = data.reduce((s, v) => s + v, 0);
            cfg = { type: "doughnut", data: { labels, datasets: [{ data, backgroundColor: labels.map((_, i) => i === 11 ? "#9e9e9e" : colour(i)), borderWidth: 1 }] },
                options: chartBase({ plugins: { legend: { position: "right" }, tooltip: { callbacks: { label: c => c.label + ": " + tickFmt(c.parsed) + (sum ? " (" + (c.parsed / sum * 100).toFixed(1) + "%)" : "") } } },
                    onClick: (ev, els) => { if (els.length && rest.length && els[0].index === topN.length) { S.othOpen = true; paint(); } } }) };
            $("saChartTitle").textContent = m.label + " — share of " + unit + " by " + dimLabelOf(dim);
        } else if (S.chart === "party") {
            const list = rows.filter(e => e.m === m.k), byUser = group(list, "user"), byParty = group(list, "party");
            const us = [...byUser.values()].sort((a, b) => val(b, m.k) - val(a, m.k)).slice(0, 15);
            const ps = [...byParty.values()].sort((a, b) => val(b, m.k) - val(a, m.k)), topP = ps.slice(0, 9);
            const cell = new Map();
            for (const e of list) { const pk = e.kind + "|" + e.party, s = topP.some(n => n.key === pk) ? pk : "__other", ck = (e.user || "") + "||" + s; if (!cell.has(ck)) cell.set(ck, node(ck, "")); add(cell.get(ck), e); }
            const ds = topP.map((p, i) => ({ label: truncate(p.label, 30), stack: "s", backgroundColor: colour(i), borderRadius: 2, data: us.map(u => { const c = cell.get(u.key + "||" + p.key); return c ? val(c, m.k) : 0; }) }));
            if (ps.length > 9) {
                // everybody's Others, or (after a click on one person's grey part) that person's
                const who = S.othUser && us.some(u => u.key === S.othUser) ? S.othUser : "";
                const pool = who ? [...group(list.filter(e => (e.user || "") === who && !topP.some(n => n.key === e.kind + "|" + e.party)), "party").values()] : ps.slice(9);
                oth = { what: pool.length + (m.side === "purch" ? " more suppliers" : " more customers") + (who ? " for " + userLabel(who) : ", all employees"), col: m.side === "purch" ? "Supplier" : "Customer", user: who,
                    rows: pool.map(n => ({ label: n.label, v: val(n, m.k), row: who ? [["user", who], ["party", n.key]] : [["party", n.key]] })) };
            }
            if (ps.length > 9) ds.push({ label: "Others (" + (ps.length - 9) + ")", stack: "s", backgroundColor: "#9e9e9e", borderRadius: 2, data: us.map(u => { const c = cell.get(u.key + "||__other"); return c ? val(c, m.k) : 0; }) });
            cfg = { type: "bar", data: { labels: us.map(u => truncate(u.label, 26)), datasets: ds },
                options: chartBase({ indexAxis: "y", plugins: { legend: { position: "bottom" }, tooltip: { mode: "index", axis: "y", intersect: false, filter: i => i.parsed.x !== 0, callbacks: { label: c => c.dataset.label + ": " + tickFmt(c.parsed.x) } } },
                    scales: { x: { stacked: true, beginAtZero: true, ticks: { callback: tickFmt, precision: 0 } }, y: { stacked: true } },
                    onClick: (ev, els) => { if (els.length && ps.length > 9 && els[0].datasetIndex === topP.length) { S.othOpen = true; S.othUser = us[els[0].index].key; paint(); } } }) };
            $("saChartTitle").textContent = m.label + " — " + unit + " by employee, split by " + (m.side === "purch" ? "supplier" : "customer") + " (nine largest named)";
        }
        paintOthers(oth);
        if (oth) $("saChartNote").textContent = ($("saChartNote").textContent ? $("saChartNote").textContent + " " : "") + "Click the grey “Others” part of the chart, or the line below, to see what is in it.";
        if (!cfg) return;
        const empty = !cfg.data.datasets.length || cfg.data.datasets.every(d => d.data.every(v => !v));
        if (empty) $("saChartNote").textContent = "Nothing to chart for this period and these filters.";
        state.charts.saChart = new Chart($("saChart").getContext("2d"), cfg);
        function dimLabelOf(d) { return (VIEWS.find(v => v[0] === d) || [])[1].toLowerCase(); }
        void dimLabel;
    }
    // The list behind the grey "Others" part of a chart; a row lists its documents below.
    function paintOthers(o) {
        const box = $("saOthers"); if (!box) return;
        if (!o || !o.rows.length) { box.innerHTML = ""; return; }
        const om = S.othSort.dir === "asc" ? 1 : -1, oc = k => S.othSort.key === k ? "sorted-" + S.othSort.dir : "";
        const rows = o.rows.filter(r => r.v).sort((a, b) => (S.othSort.key === "label" ? a.label.localeCompare(b.label, undefined, { numeric: true }) : a.v - b.v) * om || a.label.localeCompare(b.label)), sum = rows.reduce((t, r) => t + r.v, 0);
        box.innerHTML = `<span class="sa-link" data-sa-oth-toggle="1">${S.othOpen ? "▾ Hide" : "▸ Show"} what is in “Others” — ${esc(o.what)}</span>`
            + (S.othOpen && o.user ? `<span class="sa-link" data-sa-oth-all="1" style="margin-left:14px;">show for all employees</span>` : "")
            + (S.othOpen ? `<div class="table-scroll" style="max-height:340px; margin-top:8px;"><table class="ds-table" id="saOthTable"><thead><tr><th data-sa-osort="label" style="cursor:pointer;" title="Click to sort" class="${oc("label")}">${esc(o.col)}</th><th data-sa-osort="v" style="cursor:pointer;" title="Click to sort" class="num ${oc("v")}">${S.show === "amount" ? "Value (excl. tax)" : "Documents"}</th><th data-sa-osort="v" style="cursor:pointer;" title="Click to sort" class="num">Share of Others</th></tr></thead><tbody>`
                + rows.map(r => `<tr data-sa-oth-row='${escapeAttr(JSON.stringify(r.row))}' data-label="${escapeAttr(r.label)}" style="cursor:pointer;"><td>${esc(r.label)}</td><td class="num">${fmtVal(r.v)}</td><td class="num">${sum ? (r.v / sum * 100).toFixed(1) + "%" : ""}</td></tr>`).join("")
                + `<tr class="ds-total"><td>Others in total</td><td class="num">${fmtVal(sum)}</td><td class="num">100%</td></tr></tbody></table></div><div class="sa-note" style="padding-left:0;">Largest first unless you click a heading to sort. Click a row to list its documents below.</div>` : "");
    }
    function paintHeat(rows) {
        const list = rows.filter(e => e.hour >= 0), grid = new Map(); let max = 0, lo = 7, hi = 17;
        for (const e of list) { const wd = (D(e.date).getUTCDay() + 6) % 7, k = wd + "|" + e.hour, s = grid.get(k) || new Set(); s.add(e.m + e.key); grid.set(k, s); if (s.size > max) max = s.size; if (e.hour < lo) lo = e.hour; if (e.hour > hi) hi = e.hour; }
        const hours = []; for (let h = lo; h <= hi; h++) hours.push(h);
        const hl = h => (h % 12 || 12) + (h < 12 ? "am" : "pm");
        const on = id => S.pick && S.pick.id === "heat|" + id ? "sa-on" : "";
        $("saChartTitle").textContent = "Busiest day and hour — orders and quotes raised or changed";
        $("saHeat").innerHTML = !list.length ? `<div class="sa-empty">Nothing with a time of day in this period.</div>`
            : `<table class="sa-heat"><thead><tr><th>Day</th>${hours.map(h => `<th>${hl(h)}</th>`).join("")}<th>Total</th></tr></thead><tbody>`
            + WD.map((d, wd) => { let t = 0; const tds = hours.map(h => { const n = (grid.get(wd + "|" + h) || new Set()).size; t += n; return `<td${n ? ` data-sa-heat="${wd}|${h}" data-label="${d} ${hl(h)}"` : ""} class="${on(wd + "|" + h)}" style="background:color-mix(in srgb, var(--tree-tint, #3949ab) ${n ? Math.round(12 + 78 * n / max) : 0}%, transparent);${n / max > .55 ? "color:#fff;" : ""}${n ? "cursor:pointer;" : ""}" title="${d} ${hl(h)}: ${n} documents${n ? " — click to list them" : ""}">${n || ""}</td>`; }).join(""); return `<tr><th>${d}</th>${tds}<td${t ? ` data-sa-heat="${wd}|" data-label="${d}, all day" style="cursor:pointer;" title="Click to list them"` : ""} class="${on(wd + "|")}"><strong>${fmtInt(t)}</strong></td></tr>`; }).join("")
            + `</tbody></table>`;
        $("saChartNote").textContent = "Melbourne time. Counts orders and quotes raised or changed in each hour (a document counts once per hour). Posted shipments, invoices and receipts carry a posting date but no time of day in Wiise, so they are not in this grid. Click a figure to list its documents.";
    }

    // ---------- documents ----------
    function pickedRows() {
        let rows = filteredCache;
        if (S.pick) {
            if (S.pick.m) rows = rows.filter(e => e.m === S.pick.m);
            for (const [dim, k] of (S.pick.row || [])) rows = rows.filter(e => keyOf(e, dim)[0] === k);
            if (S.pick.heat) { const [wd, h] = S.pick.heat; rows = rows.filter(e => e.hour >= 0 && (h < 0 || e.hour === h) && (D(e.date).getUTCDay() + 6) % 7 === wd); }
        }
        const q = S.docQ.trim().toLowerCase();
        if (q) rows = rows.filter(e => (e.no + " " + (e.order || "") + " " + userLabel(e.user) + " " + e.user + " " + partyName(e) + " " + e.party + " " + e.sp + " " + MBY[e.m].label).toLowerCase().includes(q));
        return rows;
    }
    function purchLink(page, no) {
        return ' <a href="' + escapeAttr("https://businesscentral.dynamics.com/" + BC_CONFIG.tenantId + "/" + BC_CONFIG.environment + "?page=" + page + "&filter=" + encodeURIComponent("'No.' IS '" + no + "'"))
            + '" target="_blank" rel="noopener" title="Open ' + escapeAttr(no) + ' in Wiise" onclick="event.stopPropagation()" style="text-decoration:none; font-size:11px; opacity:0.7;">↗</a>';
    }
    function docLink(e) {
        if (/^with /.test(e.no)) return "";
        if (e.kind === "S") {
            if (e.doc === "order") return e.open ? wiiseDocLinkHtml("Sales Order", e.no) : "";
            if (e.doc === "quote") return e.open ? wiiseDocLinkHtml("Sales Quote", e.no) : "";
            if (e.doc === "blanket") return "";
            return wiiseDocLinkHtml(e.doc === "shipment" ? "Posted Sales Shipment" : e.doc === "invoice" ? "Posted Sales Invoice" : "Posted Sales CR/Adj Note", e.no);
        }
        if (e.doc === "order") return e.open ? purchLink(PURCH_PAGES.order, e.no) : "";
        return purchLink(PURCH_PAGES[e.doc], e.no);
    }
    function detail(e) {
        const bits = [];
        if (e.m === "soM" || e.m === "poM") bits.push(plural(e.n, "change") + ": " + e.fields.slice(0, 6).join(", ") + (e.fields.length > 6 ? " +" + (e.fields.length - 6) + " more" : ""));
        if (e.m === "sqV") bits.push("from quote " + e.quote + (e.quoteGuess ? " (matched by time and customer)" : ""));
        if (e.m === "sbV") bits.push("from blanket order " + e.blanket);
        if (MBY[e.m].gone) bits.push(e.raised ? "raised " + dateLabel(e.raised[0].slice(0, 10)) + " by " + userLabel(e.raised[1]) : "raised before the Change Log started");
        if ((e.doc === "order" || e.doc === "quote" || e.doc === "blanket") && e.au && e.au !== e.user) bits.push("assigned to " + userLabel(e.au));
        if (e.doc === "blanket" && e.closed) bits.push("removed " + dateLabel(e.closed[0].slice(0, 10)) + " by " + userLabel(e.closed[1]));
        if (e.doc === "order" || e.doc === "quote") bits.push(e.open ? "still open" : e.closed ? (e.doc === "order" && e.dropped ? "deleted " : e.doc === "quote" && MBY[e.m].gone ? "deleted " : "closed ") + dateLabel(e.closed[0].slice(0, 10)) + " by " + userLabel(e.closed[1]) + (e.doc === "order" && e.dropped ? " with no shipment or invoice" : "") : "no longer open");
        if (e.via) bits.push("posted with invoice " + e.via);
        if (e.lines) bits.push(plural(e.lines, "line"));
        if (e.approx) bits.push("value estimated from the GST-inclusive total");
        return bits.join(" · ");
    }
    function paintDocs() {
        if (S.status !== "ready") return;
        const DCOLS = [["date", "Date"], ["time", "Time"], ["m", "Activity"], ["no", "Document"], ["user", "Employee"], ["party", "Customer / supplier"], ["br", "Branch"], ["sp", "Salesperson"],
            ["qty", "Qty", 1], ["amt", "Value (excl. tax)", 1], ["order", "Order"], ["detail", "Detail"]];
        const dv = (e, k) => k === "date" ? e.date + (e.time || "99") : k === "time" ? (e.time || "") : k === "m" ? MBY[e.m].short : k === "user" ? userLabel(e.user).toLowerCase()
            : k === "party" ? partyName(e).toLowerCase() : k === "qty" ? (e.qty == null ? -Infinity : e.qty) : k === "amt" ? (e.amt == null ? -Infinity : e.amt)
            : k === "order" ? (e.doc === "order" ? e.no : e.order || "") : k === "detail" ? detail(e) : String(e[k] || "");
        const dk = S.docSort.key, dm = S.docSort.dir === "asc" ? 1 : -1, newest = (a, b) => (b.date + (b.time || "99")).localeCompare(a.date + (a.time || "99")) || a.no.localeCompare(b.no);
        const rows = pickedRows().slice().sort((a, b) => {
            const x = dv(a, dk), y = dv(b, dk);
            return (typeof x === "number" ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true })) * dm || newest(a, b);
        });
        $("saDocs").querySelector("thead").innerHTML = "<tr>" + DCOLS.map(c => `<th data-sa-dsort="${c[0]}" title="Click to sort" class="${c[2] ? "num " : ""}${dk === c[0] ? "sorted-" + S.docSort.dir : ""}">${c[1]}</th>`).join("") + "</tr>";
        $("saDocsTitle").textContent = "Documents" + (S.pick ? " — " + S.pick.label : "");
        $("saPickClear").style.display = S.pick ? "" : "none";
        const shown = rows.slice(0, S.docLimit);
        let sum = 0; const seen = new Set(); for (const e of rows) if (!seen.has(e.m + e.key)) { seen.add(e.m + e.key); sum += e.amt || 0; }
        $("saDocsMeta").textContent = rows.length ? plural(seen.size, "document") + " · " + money0(sum) : "";
        $("saDocs").querySelector("tbody").innerHTML = shown.length ? shown.map(e => {
            const ord = e.doc === "order" ? e.no : e.order, hasHist = !!ord;
            return `<tr><td>${dateLabel(e.date)}</td><td>${esc(e.time || "")}</td><td>${esc(MBY[e.m].short)}</td>
                <td>${esc(e.no)}${docLink(e)}</td><td title="${escapeAttr(e.user)}">${esc(userLabel(e.user))}</td><td title="${escapeAttr(partyName(e) + " (" + e.party + ")")}">${esc(partyName(e))}</td>
                <td>${esc(e.br || "")}</td><td>${esc(e.sp || "")}</td><td class="num">${e.qty == null ? "" : (Number.isInteger(e.qty) ? fmtInt(e.qty) : fmtQty(e.qty))}</td>
                <td class="num">${e.amt == null ? "—" : fmtMoneyFull(e.amt)}</td>
                <td>${hasHist ? `<span class="sa-link" data-sa-hist="${escapeAttr(ord)}" data-sa-kind="${e.kind}" title="Show who raised, changed, shipped and invoiced this order">${esc(ord)}</span>` : ""}</td>
                <td class="wrap">${esc(detail(e))}</td></tr>`;
        }).join("") + (rows.length > shown.length ? `<tr><td colspan="12" class="empty"><span class="sa-link" data-sa-more="1">Showing the first ${fmtInt(shown.length)} of ${fmtInt(rows.length)} lines — show 500 more</span></td></tr>` : "")
            : `<tr><td colspan="12" class="empty">No documents for this choice.</td></tr>`;
        $("saDocsNote").textContent = "Newest first unless you click a column heading to sort. Click an order number in the Order column for its full story. The ↗ opens the document in Wiise; an order or quote that is no longer open has no link because Wiise removes it once it is fully invoiced.";
    }

    // ---------- one order's story ----------
    function paintHist() {
        const p = $("saHistPanel"); if (!p) return;
        if (!S.hist || S.status !== "ready") { p.style.display = "none"; return; }
        p.style.display = "";
        const { kind, no } = S.hist, d = S.data, x = d.docs[kind + "1|" + no], lines = [];
        const push = (when, user, what, det) => lines.push({ when, user, what, det });
        if (x) {
            push(x.f[0], x.f[1], x.c ? "Raised the order" : "First change on record", x.c ? "" : "The order already existed when the Change Log started, so who raised it is not known.");
            if (x.e) for (const ek in x.e) { const cut = ek.indexOf("|"), f = x.e[ek][1].map(i => d.fields[i]); push(ek.slice(0, cut) + ":00", ek.slice(cut + 1), "Changed the order", x.e[ek][0] + " change" + (x.e[ek][0] === 1 ? "" : "s") + ": " + f.join(", ")); }
            if (x.d) push(x.d[0], x.d[1], "Order closed", "Removed from the open orders (normally by the invoice that completed it).");
        }
        const posted = (bag, what) => { for (const n in bag) { const y = bag[n]; if (y.o === no && y.d) push(y.d, y.u, what + " " + n, (y.a != null ? fmtMoneyFull(y.a) : "") + (y.i ? " · with invoice " + y.i : "")); } };
        if (kind === "S") { posted(d.sh, "Posted shipment"); posted(d.si, "Posted invoice"); }
        else { posted(d.pr, "Posted receipt"); posted(d.pi, "Posted invoice"); }
        const hk = S.histSort.key, hm = S.histSort.dir === "asc" ? 1 : -1;
        const hsort = (list, get) => list.sort((a, b) => String(get(a)).localeCompare(String(get(b)), undefined, { numeric: true }) * hm || a.when.localeCompare(b.when));
        hsort(lines, l => hk === "user" ? userLabel(l.user).toLowerCase() : l[hk] || "");
        $("saHistTable").querySelector("thead").innerHTML = "<tr>" + [["when", "When"], ["user", "Employee"], ["what", "What happened"], ["det", "Detail"]]
            .map(c => `<th data-sa-hsort="${c[0]}" title="Click to sort" class="${hk === c[0] ? "sorted-" + S.histSort.dir : ""}">${c[1]}</th>`).join("") + "</tr>";
        $("saHistTitle").textContent = (kind === "S" ? "Sales order " : "Purchase order ") + no;
        const party = x && x.cu ? ((kind === "S" ? d.cust[x.cu] : d.vend[x.cu]) || x.cu) : "";
        $("saHistMeta").textContent = [party, x && x.au ? "assigned to " + userLabel(x.au) : "", x && x.o ? "open" : ""].filter(Boolean).join(" · ");
        const fmtWhen = w => dateLabel(w.slice(0, 10)) + (w.length > 10 ? " " + w.slice(11, 16) : "");
        const live = Array.isArray(S.hist.live) ? hsort(S.hist.live.slice(), r => hk === "user" ? userLabel(r.user).toLowerCase() : hk === "what" ? r.field : hk === "det" ? r.change : r.when) : S.hist.live;
        $("saHistTable").querySelector("tbody").innerHTML = (lines.length ? lines.map(l => `<tr><td>${fmtWhen(l.when)}</td><td>${esc(userLabel(l.user))}</td><td>${esc(l.what)}</td><td class="wrap">${esc(l.det)}</td></tr>`).join("")
            : `<tr><td colspan="4" class="empty">Nothing on record for this order in the activity file.</td></tr>`)
            + (live === "loading" ? `<tr><td colspan="4" class="empty">Reading the Change Log from Wiise…</td></tr>` : "")
            + (Array.isArray(live) ? `<tr class="ds-total"><td colspan="4">Every change in the Wiise Change Log (${fmtInt(live.length)} lines)</td></tr>` + live.map(r => `<tr><td>${fmtWhen(r.when)}</td><td>${esc(userLabel(r.user))}</td><td>${esc(r.field)}</td><td class="wrap">${esc(r.change)}</td></tr>`).join("") : "")
            + (live && live.error ? `<tr><td colspan="4" class="empty">${esc(live.error)}</td></tr>` : "");
        $("saHistLive").style.display = Array.isArray(live) || live === "loading" ? "none" : "";
        $("saHistNote").textContent = "Shipments and invoices carry a posting date but no time of day. “Read every change from Wiise” fetches the field-by-field Change Log for this one order with your own sign-in.";
    }
    async function histLive() {
        if (!S.hist) return;
        const h = S.hist; h.live = "loading"; paintHist();
        try {
            const co = await bcGetCompanyInternalName();
            const melFmt = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23" });
            const mel = s => { const p = {}; for (const x of melFmt.formatToParts(new Date(s))) p[x.type] = x.value; return p.year + "-" + p.month + "-" + p.day + "T" + (p.hour === "24" ? "00" : p.hour) + ":" + p.minute; };
            const filter = "Table_No eq " + (h.kind === "S" ? 36 : 38) + " and Primary_Key_Field_2_Value eq '" + h.no.replace(/'/g, "''") + "'";
            const rows = await bcFetchAll(BC_ODATA_URL + "/Company('" + encodeURIComponent(co) + "')/ChangeLogEntries?$filter=" + encodeURIComponent(filter)
                + "&$select=Entry_No,Date_and_Time,User_ID,Field_Caption,Type_of_Change,Old_Value,New_Value", "Change log");
            rows.sort((a, b) => a.Entry_No - b.Entry_No);
            let delSeen = false; const out = [];
            for (const r of rows) {
                if (r.Type_of_Change === "Deletion") { if (delSeen) continue; delSeen = true; out.push({ when: mel(r.Date_and_Time), user: String(r.User_ID || "").toUpperCase(), field: "(order removed)", change: "Every field is logged when the order is removed; shown once here." }); continue; }
                out.push({ when: mel(r.Date_and_Time), user: String(r.User_ID || "").toUpperCase(), field: r.Field_Caption || "", change: "“" + (r.Old_Value || "") + "” → “" + (r.New_Value || "") + "”" });
            }
            if (S.hist === h) { h.live = out; paintHist(); }
        } catch (e) { if (S.hist === h) { h.live = { error: "Could not read the Change Log from Wiise: " + (e.message || e) }; paintHist(); } }
        const prog = document.getElementById("bcProgress"); if (prog) prog.style.display = "none";    // the corner progress line bcFetchAll switches on
    }

    // ---------- entry point ----------
    window.renderStaffActivity = function () {
        if (!allowed()) return;
        scaffold();
        if (S.status === "idle") { load(false); return; }
        paint();
    };
    window.staffActivity = { state: S, inject, reload: () => load(true), live,     // inject() is for the headless test rig
        ensure: () => S.status === "ready" ? Promise.resolve() : load(false) };
})();
