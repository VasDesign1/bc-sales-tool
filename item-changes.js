/* ============================================================
   ITEM CHANGES  (Operations → Item Changes; Faik, 2026-11-26)
   Every change made on an item card: who, when, which field, the
   old value and the new value. Also new and deleted items, and the
   lists that hang off the card when Wiise logs them (units of
   measure, item suppliers, kit contents).

   OWNER-ONLY: the nav link carries class restrictedNav and the tab
   is in RESTRICTED_TABS (index.html).

   Data: the same file as Staff Activity (snapshots/activity.bin,
   robot scripts/activity.js, section 1b). It is downloaded once and
   shared: this tab asks staff-activity.js for it (staffActivity.ensure).
     data.ic     [when, user, item, part, type, field, old, new, sub, n]
                 part  I item card · S stock card per branch (stockkeeping unit) · U unit of measure · V supplier · K kit line
                 type  0 new · 1 changed · 2 deleted   (0 / 2: one line per record, n fields)
     data.items  item number → [description, category]
   "Get latest from Wiise" (staffActivity.live) adds what Wiise holds
   beyond the file, on screen only; one press serves both tabs.
   The history starts the day the Item table was switched on in the
   Wiise Change Log Setup; nothing earlier exists.

   Keep help.js in step when tiles, columns or controls change.
   ============================================================ */
(function () {
    "use strict";
    const $ = id => document.getElementById(id);
    const esc = s => escapeHtml(s == null ? "" : s);

    const PARTS = { I: "Item card", S: "Stock card (per branch)", U: "Units of measure", V: "Suppliers", K: "Kit contents" };
    const WHOLE = {
        I: ["New item created", "", "Item deleted"],
        U: ["Unit of measure added", "", "Unit of measure removed"],
        V: ["Supplier added", "", "Supplier removed"],
        K: ["Kit line added", "", "Kit line removed"],
        S: ["Stock card created", "", "Stock card deleted"],
    };
    // Fields Wiise rewrites by itself when a purchase is posted or costs are
    // adjusted. Nobody typed these; they are hidden until asked for.
    const SYSTEM = new Set(["Unit Cost", "Last Direct Cost", "Cost is Adjusted", "Last Unit Cost Calc. Date", "Rolled-up Material Cost", "Rolled-up Capacity Cost",
        "Rolled-up Subcontracted Cost", "Rolled-up Mfg. Ovhd Cost", "Rolled-up Cap. Overhead Cost", "Single-Level Material Cost", "Single-Level Capacity Cost",
        "Single-Level Subcontrd. Cost", "Single-Level Cap. Ovhd Cost", "Single-Level Mfg. Ovhd Cost", "Low-Level Code", "Last Counting Period Update",
        "Next Counting Start Date", "Next Counting End Date", "Last Phys. Invt. Date", "Coupled to Dataverse", "Allow Online Adjustment",
        "Profit %", "Last Landed Cost", "Search Description", "Has Sales Forecast"]);
    const SETUP_MIN = 60;      // fields the creator fills in within this many minutes of creating an item are part of creating it
    const VIEWS = [["user", "Employee"], ["item", "Item"], ["field", "Field"], ["day", "Day"]];

    const S = {
        src: null, ev: [],
        range: "month", from: "", to: "",
        user: "", field: "", part: "", type: "", q: "", sys: false, setup: false,
        view: "user", sort: { key: "n", dir: "desc" }, pick: null,
        sumLimit: 200, limit: 300,
    };
    const allowed = () => typeof currentSignInName === "function" && currentSignInName() === RESTRICTED_NAV_OWNER;
    const SA = () => window.staffActivity;
    const plural = (n, w) => fmtInt(n) + " " + w + (n === 1 ? "" : "s");

    // ---------- dates ----------
    const today = () => (typeof melbourneTodayISO === "function" ? melbourneTodayISO() : new Date().toISOString().slice(0, 10));
    const D = iso => { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
    const iso = d => d.toISOString().slice(0, 10);
    const addDays = (s, n) => { const d = D(s); d.setUTCDate(d.getUTCDate() + n); return iso(d); };
    const monday = s => { const d = D(s); d.setUTCDate(d.getUTCDate() - (d.getUTCDay() + 6) % 7); return iso(d); };
    const WD = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"], MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const dateLabel = s => { if (!s) return "—"; const d = D(s); return String(d.getUTCDate()).padStart(2, "0") + " " + MON[d.getUTCMonth()] + " " + d.getUTCFullYear(); };
    const dayLabel = s => { const d = D(s); return WD[(d.getUTCDay() + 6) % 7] + " " + dateLabel(s); };
    const whenLabel = ts => dateLabel(ts.slice(0, 10)) + " " + ts.slice(11, 16);
    const heldFrom = () => { const d = SA() && SA().state.data; return String((d && d.meta && d.meta.itemFrom) || "").slice(0, 10); };
    function setRange(r) {
        const t = today(); S.range = r;
        if (r === "today") { S.from = S.to = t; }
        else if (r === "yesterday") { S.from = S.to = addDays(t, -1); }
        else if (r === "week") { S.from = monday(t); S.to = t; }
        else if (r === "lastweek") { S.from = addDays(monday(t), -7); S.to = addDays(monday(t), -1); }
        else if (r === "month") { S.from = t.slice(0, 8) + "01"; S.to = t; }
        else if (r === "lastmonth") { const d = D(t.slice(0, 8) + "01"); d.setUTCDate(0); S.to = iso(d); S.from = S.to.slice(0, 8) + "01"; }
        else if (r === "all") { S.from = heldFrom() || t; S.to = t; }
    }

    // ---------- names ----------
    const titleCase = s => String(s).toLowerCase().replace(/(^|[\s._-])([a-z])/g, (m, a, b) => a + b.toUpperCase());
    function userLabel(u) {
        if (!u) return "(no user recorded)";
        const d = SA().state.data, n = d && d.users && d.users[u];
        return n && n.toUpperCase() !== u && !/^user_/i.test(n) ? n : (/^USER_[0-9A-F]{20,}$/.test(u) ? "System account" : titleCase(u));
    }

    // ---------- file → lines ----------
    function build(d) {
        const out = [], items = d.items || {}, fields = d.fields || [], vend = d.vend || {};
        const mins = ts => Date.UTC(+ts.slice(0, 4), +ts.slice(5, 7) - 1, +ts.slice(8, 10), +ts.slice(11, 13), +ts.slice(14, 16)) / 60000;
        const born = new Map();      // item → [minute, user] of its creation
        for (const r of (d.ic || [])) if (r[3] === "I" && r[4] === 0) born.set(r[2], [mins(r[0]), r[1]]);
        for (const r of (d.ic || [])) {
            const ts = r[0], t = r[3], type = r[4], sub = r[8] || "", it = items[r[2]] || ["", ""];
            const subLabel = !sub ? "" : t === "V" ? (vend[sub] ? vend[sub] + " (" + sub + ")" : sub) : t === "K" ? "line " + sub : t === "S" ? "branch " + sub : sub;
            let field, sys = false;
            if (type === 1) { const cap = fields[r[5]] || "(field)"; field = cap + (subLabel ? " — " + subLabel : ""); sys = t === "I" && SYSTEM.has(cap); }
            else field = WHOLE[t][type] + (subLabel ? " — " + subLabel : "");
            out.push({ ts, date: ts.slice(0, 10), user: r[1] || "", item: r[2], name: it[0] || (type !== 1 && t === "I" ? (r[7] || r[6] || "") : ""), cat: it[1] || "",
                t, type, cap: type === 1 ? (fields[r[5]] || "(field)") : WHOLE[t][type], field, oldV: type === 1 ? r[6] : (t === "I" ? "" : ""), newV: type === 1 ? r[7] : "",
                whole: type !== 1, n: r[9] || 0, sys,
                setup: (() => { if (type !== 1) return false; const b = born.get(r[2]); if (!b || b[1] !== r[1]) return false; const m = mins(ts) - b[0]; return m >= 0 && m <= SETUP_MIN; })() });
        }
        return out;
    }
    function events() {
        const d = SA().state.data;
        if (S.src !== d) { S.src = d; S.ev = d ? build(d) : []; }
        return S.ev;
    }
    function filtered(skip) {
        const q = S.q.trim().toLowerCase();
        return events().filter(e => {
            if (e.date < S.from || e.date > S.to) return false;
            if (!S.sys && e.sys) return false;
            if (!S.setup && e.setup) return false;
            if (skip !== "user" && S.user && e.user !== S.user) return false;
            if (skip !== "field" && S.field && e.cap !== S.field) return false;
            if (S.part && e.t !== S.part) return false;
            if (S.type !== "" && e.type !== +S.type) return false;
            if (q && !(e.item + " " + e.name + " " + e.cat + " " + e.field + " " + e.oldV + " " + e.newV + " " + userLabel(e.user)).toLowerCase().includes(q)) return false;
            return true;
        });
    }
    function keyOf(e, dim) {
        switch (dim) {
            case "user":  return [e.user || "", userLabel(e.user)];
            case "item":  return [e.item, e.item + (e.name ? " · " + e.name : "")];
            case "field": return [e.cap, e.cap];
            case "day":   return [e.date, dayLabel(e.date)];
        }
        return ["", ""];
    }

    // ---------- one-off page build ----------
    function scaffold() {
        const root = $("icRoot");
        if (!root || root.dataset.built) return;
        if (!$("icStyle")) {
            const st = document.createElement("style"); st.id = "icStyle";
            st.textContent = `
                .ic-tiles { grid-template-columns: repeat(auto-fit, minmax(165px, 1fr)); }
                .ic-tiles .kpi { padding: 12px 16px; }
                .ic-tiles .kpi .value { font-size: 23px; }
                .ic-bar { display: flex; flex-wrap: wrap; align-items: center; justify-content: flex-start; gap: 8px 10px; }
                .ic-bar .lbl, .ic-ctl .lbl { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .5px; opacity: .7; margin-left: 6px; }
                .ic-bar .lbl:first-child { margin-left: 0; }
                .ic-bar input[type=date], .ic-bar input[type=search], .ic-ctl input[type=search] { font: inherit; font-size: 12px; }
                .ic-bar label.chk { font-size: 12px; display: inline-flex; align-items: center; gap: 5px; cursor: pointer; }
                .ic-ctl { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; margin-left: auto; }
                #icSum tbody tr[data-k] { cursor: pointer; }
                #icSum tbody tr.ic-on td { outline: 2px solid #ffd54f; outline-offset: -2px; }
                #icList td { white-space: nowrap; }
                #icList td.wrap { white-space: normal; min-width: 150px; max-width: 320px; overflow-wrap: anywhere; }
                #icList td.ic-old { opacity: .7; text-decoration: line-through; text-decoration-color: rgba(128,128,128,.6); }
                #icList td.ic-new { font-weight: 600; }
                .ic-link { cursor: pointer; text-decoration: underline dotted; text-underline-offset: 3px; }
                .ic-note { font-size: 11px; opacity: .75; padding: 8px 14px 12px; }
                .ic-chip { display: inline-block; font-size: 10px; padding: 1px 6px; border-radius: 8px; background: rgba(128,128,128,.18); margin-left: 6px; }
            `;
            document.head.appendChild(st);
        }
        const opt = (v, l) => `<option value="${v}">${l}</option>`;
        root.innerHTML = `
            <div class="hint rr-hint"><span id="icHint">Every change made on an item card.</span>
                <span class="rr-tgt"><span class="rr-tgt-lbl">Activity file</span><span id="icAsOf">not loaded</span><button type="button" class="more-btn" id="icReload" title="Read the latest file again">↻</button><button type="button" class="more-btn" id="icLive" title="Read everything newer than the file straight from Wiise, with your own sign-in">Get latest from Wiise</button></span></div>
            <div class="panel"><div class="panel-header ic-bar">
                <span class="lbl">Period</span>
                <select id="icRange" class="mini-select" title="Quick period">
                    ${opt("today", "Today")}${opt("yesterday", "Yesterday")}${opt("week", "This week")}${opt("lastweek", "Last week")}${opt("month", "This month")}${opt("lastmonth", "Last month")}${opt("all", "Everything held")}${opt("custom", "Custom dates")}
                </select>
                <input type="date" id="icFrom" class="mini-select" title="From"><span>to</span><input type="date" id="icTo" class="mini-select" title="To">
                <label class="chk" id="icSysLbl" title="Unit Cost, Last Direct Cost and similar fields are rewritten by Wiise itself whenever a purchase is posted or costs are adjusted"><input type="checkbox" id="icSys"> Include the cost updates Wiise makes by itself</label>
                <label class="chk" id="icSetupLbl" title="The fields somebody fills in during the first hour after creating an item are part of creating it, not changes"><input type="checkbox" id="icSetup"> Include the fields filled in while a new item was set up</label>
            </div><div class="panel-header ic-bar" style="border-bottom:0;">
                <span class="lbl">Filter</span>
                <select id="icUser" class="mini-select" title="Employee"></select>
                <select id="icField" class="mini-select" title="Field that was changed"></select>
                <select id="icPart" class="mini-select" title="Part of the item card">${opt("", "Whole card")}${Object.keys(PARTS).map(k => opt(k, PARTS[k])).join("")}</select>
                <select id="icType" class="mini-select" title="Kind of change">${opt("", "All kinds")}${opt("1", "Changed")}${opt("0", "New / added")}${opt("2", "Deleted / removed")}</select>
                <input type="search" id="icQ" class="mini-select" placeholder="Item number, description, value…" title="Part of an item number, description, field or value">
                <button type="button" class="more-btn" id="icClear" title="Clear the filters (keeps the period)">Clear filters</button>
            </div></div>
            <div class="kpi-row ic-tiles" id="icTiles"></div>
            <div class="panel">
                <div class="panel-header"><h2 id="icSumTitle">Changes by employee</h2><div class="meta" id="icSumMeta"></div>
                    <div class="ic-ctl"><span class="lbl">View by</span>${VIEWS.map(v => `<button type="button" class="gp-filter-btn" data-ic-view="${v[0]}">${v[1]}</button>`).join("")}</div>
                </div>
                <div class="table-scroll"><table id="icSum" class="ds-table"><thead></thead><tbody></tbody></table></div>
                <div class="ic-note" id="icSumNote"></div>
            </div>
            <div class="panel" id="icListPanel">
                <div class="panel-header"><h2 id="icListTitle">Every change</h2><div class="meta" id="icListMeta"></div>
                    <div class="ic-ctl"><button type="button" class="more-btn" id="icPickClear" style="display:none;">Show all changes</button></div>
                </div>
                <div class="table-scroll"><table id="icList" class="ds-table"><thead><tr>
                    <th>When</th><th>Employee</th><th>Item</th><th>Description</th><th>Part of the card</th><th>Field</th><th>Old value</th><th>New value</th>
                </tr></thead><tbody></tbody></table></div>
                <div class="ic-note" id="icListNote"></div>
            </div>`;
        root.dataset.built = "1";

        const reset = () => { S.pick = null; S.limit = 300; S.sumLimit = 200; paint(); };
        $("icReload").addEventListener("click", () => { paint("loading"); SA().reload().then(() => paint()); });
        $("icLive").addEventListener("click", () => { const p = SA().live(); paint(); p.then(() => paint()); });
        $("icRange").addEventListener("change", e => { if (e.target.value !== "custom") setRange(e.target.value); else S.range = "custom"; reset(); });
        const dateChange = () => { const f = $("icFrom").value, t = $("icTo").value; if (!f || !t) return; S.from = f <= t ? f : t; S.to = f <= t ? t : f; S.range = "custom"; reset(); };
        $("icFrom").addEventListener("change", dateChange); $("icTo").addEventListener("change", dateChange);
        $("icSys").addEventListener("change", e => { S.sys = e.target.checked; reset(); });
        $("icSetup").addEventListener("change", e => { S.setup = e.target.checked; reset(); });
        $("icUser").addEventListener("change", e => { S.user = e.target.value; reset(); });
        $("icField").addEventListener("change", e => { S.field = e.target.value; reset(); });
        $("icPart").addEventListener("change", e => { S.part = e.target.value; reset(); });
        $("icType").addEventListener("change", e => { S.type = e.target.value; reset(); });
        $("icQ").addEventListener("input", e => { S.q = e.target.value; reset(); });
        $("icClear").addEventListener("click", () => { S.user = S.field = S.part = S.type = S.q = ""; reset(); });
        $("icPickClear").addEventListener("click", () => { S.pick = null; S.limit = 300; paint(); });
        root.addEventListener("click", e => {
            if (e.target.closest("a")) return;
            let b;
            if ((b = e.target.closest("[data-ic-view]"))) { S.view = b.dataset.icView; S.pick = null; S.sumLimit = 200; S.sort = S.view === "day" ? { key: "key", dir: "desc" } : { key: "n", dir: "desc" }; paint(); return; }
            if ((b = e.target.closest("th[data-ic-sort]"))) {
                const k = b.dataset.icSort;
                if (S.sort.key === k) S.sort.dir = S.sort.dir === "asc" ? "desc" : "asc"; else S.sort = { key: k, dir: k === "label" ? "asc" : "desc" };
                paint(); return;
            }
            if ((b = e.target.closest("[data-ic-more]"))) { if (b.dataset.icMore === "sum") S.sumLimit += 500; else S.limit += 500; paint(); return; }
            if ((b = e.target.closest("[data-ic-item]"))) { S.q = b.dataset.icItem; S.pick = null; S.limit = 300; paint(); return; }
            if ((b = e.target.closest("#icSum tr[data-k]"))) {
                const k = b.dataset.k, same = S.pick && S.pick.dim === S.view && S.pick.k === k;
                S.pick = same ? null : { dim: S.view, k, label: b.dataset.label || k }; S.limit = 300; paint();
                if (!same) $("icListPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
            }
        });
    }

    function itemLink(no) {
        if (typeof BC_CONFIG === "undefined" || !no) return "";
        return ' <a href="' + escapeAttr("https://businesscentral.dynamics.com/" + BC_CONFIG.tenantId + "/" + BC_CONFIG.environment + "?page=30&filter=" + encodeURIComponent("'No.' IS '" + no + "'"))
            + '" target="_blank" rel="noopener" title="Open item ' + escapeAttr(no) + ' in Wiise" style="text-decoration:none; font-size:11px; opacity:0.7;">↗</a>';
    }

    // ---------- painting ----------
    function paint(force) {
        if (!allowed()) return;
        scaffold();
        if (!$("icRoot") || !$("icRoot").dataset.built) return;
        const sa = SA().state, status = force || sa.status, ready = status === "ready", d = sa.data || {};
        const meta = d.meta || {}, built = String(meta.builtAtMelbourne || "");
        const fileAsOf = "as of " + built.slice(11) + " on " + dateLabel(built.slice(0, 10));
        const liveErr = ready && sa.live === "error" ? sa.liveErr : ready && meta.liveAt && meta.icErr ? "The item changes could not be read live (" + meta.icErr + ")." : "";
        $("icAsOf").textContent = status === "loading" || status === "idle" ? "loading…" : !ready ? "not loaded"
            : sa.live === "loading" ? fileAsOf + " · reading Wiise…"
            : liveErr ? fileAsOf + " · live read failed"
            : meta.liveAt ? "LIVE from Wiise at " + meta.liveAt.slice(11) + " (" + meta.liveSecs + " s)" : fileAsOf;
        $("icAsOf").title = liveErr || (meta.liveAt ? "The robot's file " + fileAsOf + ", plus everything Wiise holds since. On screen only; ↻ goes back to the file." : "");
        $("icLive").disabled = !ready || sa.live === "loading";
        $("icLive").textContent = sa.live === "loading" ? "Reading Wiise…" : "Get latest from Wiise";
        const has = ready && (d.ic || []).length > 0;
        const off = "Nothing has been recorded yet. Item changes are collected from the Wiise Change Log once the Item table (27) is switched on in Change Log Setup and people have signed in again; the robot then adds them on its next run.";
        $("icHint").innerHTML = status === "error" ? esc(sa.err) : !ready ? "Loading the activity file…" : !has ? (liveErr ? "<strong>" + esc(liveErr) + "</strong> " : "") + esc(off)
            : (liveErr ? "<strong>" + esc(liveErr) + "</strong> Showing the robot's file. " : "") + "Every change made on an item card: who, when, which field, and the old and new value. From the Wiise Change Log, held from <strong>" + dateLabel(heldFrom()) + "</strong>; nothing earlier exists. The toolbar filters above do not apply here; use the ones below.";
        ["icRange", "icFrom", "icTo", "icSys", "icSetup", "icUser", "icField", "icPart", "icType", "icQ"].forEach(id => { $(id).disabled = !has; });
        if (!has) {
            $("icTiles").innerHTML = ""; $("icSum").querySelector("thead").innerHTML = "";
            $("icSum").querySelector("tbody").innerHTML = `<tr><td class="empty">${esc(status === "error" ? sa.err : ready ? off : "Loading…")}</td></tr>`;
            $("icList").querySelector("tbody").innerHTML = ""; $("icSumMeta").textContent = $("icListMeta").textContent = ""; $("icSumNote").textContent = $("icListNote").textContent = "";
            return;
        }
        if (!S.from) setRange(S.range);
        $("icRange").value = S.range; $("icFrom").value = S.from; $("icTo").value = S.to; $("icSys").checked = S.sys; $("icSetup").checked = S.setup;
        $("icPart").value = S.part; $("icType").value = S.type;
        if ($("icQ").value !== S.q) $("icQ").value = S.q;
        $("icRoot").querySelectorAll("[data-ic-view]").forEach(b => b.classList.toggle("active", b.dataset.icView === S.view));
        const fill = (id, rows, dim, cur, allLabel) => {
            const seen = new Map(); for (const e of rows) { const [k, l] = keyOf(e, dim); seen.set(k, [l, (seen.get(k) || [l, 0])[1] + 1]); }
            if (cur && !seen.has(cur)) seen.set(cur, [dim === "user" ? userLabel(cur) : cur, 0]);
            $(id).innerHTML = `<option value="">${allLabel}</option>` + [...seen].sort((a, b) => a[1][0].localeCompare(b[1][0], undefined, { numeric: true }))
                .map(([v, l]) => `<option value="${escapeAttr(v)}"${v === cur ? " selected" : ""}>${esc(l[0])} (${fmtInt(l[1])})</option>`).join("");
        };
        fill("icUser", filtered("user"), "user", S.user, "All employees");
        fill("icField", filtered("field"), "field", S.field, "All fields");

        const rows = filtered();
        paintTiles(rows); paintSum(rows); paintList(rows);
    }

    function paintTiles(rows) {
        const items = new Set(), people = new Set(), fields = new Set(); let chg = 0, neu = 0, del = 0;
        for (const e of rows) {
            items.add(e.item); people.add(e.user);
            if (e.type === 1) { chg++; fields.add(e.cap); } else if (e.t === "I") { if (e.type === 0) neu++; else del++; } else chg++;
        }
        const hidden = S.sys ? 0 : events().filter(e => e.sys && e.date >= S.from && e.date <= S.to).length;
        const tile = (tint, label, value, sub) => `<div class="kpi ${tint}"><div class="label">${label}</div><div class="value">${value}</div><div class="sub">${sub}</div></div>`;
        $("icTiles").innerHTML =
            tile("k-rev", "Changes", fmtInt(chg), plural(fields.size, "different field"))
            + tile("k-qty", "Items touched", fmtInt(items.size), dateLabel(S.from) + " to " + dateLabel(S.to))
            + tile("k-loc", "New items", fmtInt(neu), "created in the period")
            + tile("k-gpp", "Items deleted", fmtInt(del), "removed in the period")
            + tile("k-avg", "Employees", fmtInt(people.size), "made at least one change")
            + tile("k-inv", "Wiise's own cost updates", S.sys ? "shown" : fmtInt(hidden), S.sys ? "included in every figure" : "hidden; tick the box to include");
    }

    function paintSum(rows) {
        const viewName = (VIEWS.find(v => v[0] === S.view) || [])[1];
        $("icSumTitle").textContent = "Changes by " + viewName.toLowerCase();
        const g = new Map();
        for (const e of rows) {
            const [k, l] = keyOf(e, S.view);
            let n = g.get(k); if (!n) g.set(k, n = { key: k, label: l, n: 0, items: new Set(), people: new Set(), fields: new Set(), neu: 0, del: 0, last: "" });
            n.n++; n.items.add(e.item); n.people.add(e.user); if (e.type === 1) n.fields.add(e.cap);
            if (e.t === "I" && e.type === 0) n.neu++; if (e.t === "I" && e.type === 2) n.del++;
            if (e.ts > n.last) n.last = e.ts;
        }
        const cols = [["label", viewName, false], ["n", "Changes", true],
            S.view === "item" ? ["people", "Employees", true] : ["items", "Items", true],
            S.view === "field" ? ["people", "Employees", true] : ["fields", "Fields", true],
            ["neu", "New items", true], ["del", "Deleted", true], ["last", "Latest change", false]];
        const sv = (n, k) => k === "label" ? n.label.toLowerCase() : k === "key" ? n.key : (n[k] instanceof Set ? n[k].size : n[k]);
        const list = [...g.values()].sort((a, b) => {
            const x = sv(a, S.sort.key), y = sv(b, S.sort.key), m = S.sort.dir === "asc" ? 1 : -1;
            return (typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true })) * m || a.label.localeCompare(b.label);
        });
        $("icSum").querySelector("thead").innerHTML = "<tr>" + cols.map(c => { const k = c[0] === "label" && S.view === "day" ? "key" : c[0];
            return `<th data-ic-sort="${k}" style="cursor:pointer;" class="${c[2] ? "num " : ""}${S.sort.key === k ? "sorted-" + S.sort.dir : ""}">${esc(c[1])}</th>`; }).join("") + "</tr>";
        const cell = (n, c) => { const v = n[c[0]]; return c[0] === "label" ? `<td title="${escapeAttr(n.label)}">${esc(n.label)}</td>` : c[0] === "last" ? `<td>${esc(whenLabel(n.last))}</td>` : `<td class="num">${(v instanceof Set ? v.size : v) ? fmtInt(v instanceof Set ? v.size : v) : "—"}</td>`; };
        const shown = list.slice(0, S.sumLimit);
        $("icSum").querySelector("tbody").innerHTML = shown.length
            ? shown.map(n => `<tr data-k="${escapeAttr(n.key)}" data-label="${escapeAttr(n.label)}" class="${S.pick && S.pick.dim === S.view && S.pick.k === n.key ? "ic-on" : ""}">${cols.map(c => cell(n, c)).join("")}</tr>`).join("")
              + (list.length > shown.length ? `<tr><td colspan="${cols.length}" class="empty"><span class="ic-link" data-ic-more="sum">Showing the first ${fmtInt(shown.length)} of ${fmtInt(list.length)} — show 500 more</span></td></tr>` : "")
            : `<tr><td colspan="${cols.length}" class="empty">No item changes for this period and these filters.</td></tr>`;
        $("icSumMeta").textContent = rows.length ? plural(rows.length, "line") + " · " + dateLabel(S.from) + " to " + dateLabel(S.to) : "";
        $("icSumNote").textContent = "Click a row to list its changes below; click it again to clear. Click a column heading to sort. A new item counts as one line: the fields its creator fills in during the first hour are left out unless the box above is ticked.";
    }

    function paintList(all) {
        let rows = all;
        if (S.pick) rows = rows.filter(e => keyOf(e, S.pick.dim)[0] === S.pick.k);
        rows = rows.slice().sort((a, b) => b.ts.localeCompare(a.ts) || a.item.localeCompare(b.item, undefined, { numeric: true }) || a.field.localeCompare(b.field));
        $("icListTitle").textContent = "Every change" + (S.pick ? " — " + S.pick.label : "");
        $("icPickClear").style.display = S.pick ? "" : "none";
        $("icListMeta").textContent = rows.length ? plural(rows.length, "line") + " · " + plural(new Set(rows.map(e => e.item)).size, "item") : "";
        const shown = rows.slice(0, S.limit);
        $("icList").querySelector("tbody").innerHTML = shown.length ? shown.map(e => {
            const gone = e.t === "I" && e.type === 2;
            const oldCell = e.whole ? `<td class="wrap">${e.type === 2 ? esc(plural(e.n, "field") + " cleared") : ""}</td>` : `<td class="wrap ic-old" title="${escapeAttr(e.oldV)}">${e.oldV === "" ? "(blank)" : esc(e.oldV)}</td>`;
            const newCell = e.whole ? `<td class="wrap">${e.type === 0 ? esc(plural(e.n, "field") + " filled in") : ""}</td>` : `<td class="wrap ic-new" title="${escapeAttr(e.newV)}">${e.newV === "" ? "(blank)" : esc(e.newV)}</td>`;
            return `<tr><td>${esc(whenLabel(e.ts))}</td><td title="${escapeAttr(e.user)}">${esc(userLabel(e.user))}</td>
                <td><span class="ic-link" data-ic-item="${escapeAttr(e.item)}" title="Show only this item">${esc(e.item)}</span>${gone ? "" : itemLink(e.item)}</td>
                <td class="wrap" title="${escapeAttr(e.cat ? "Category " + e.cat : "")}">${esc(e.name)}</td><td>${esc(PARTS[e.t] || "")}</td>
                <td class="wrap">${esc(e.field)}${e.sys ? '<span class="ic-chip" title="Wiise rewrites this field by itself">Wiise</span>' : ""}${e.setup ? '<span class="ic-chip" title="Filled in while the new item was being set up">new item</span>' : ""}</td>${oldCell}${newCell}</tr>`;
        }).join("") + (rows.length > shown.length ? `<tr><td colspan="8" class="empty"><span class="ic-link" data-ic-more="list">Showing the first ${fmtInt(shown.length)} of ${fmtInt(rows.length)} lines — show 500 more</span></td></tr>` : "")
            : `<tr><td colspan="8" class="empty">No item changes for this choice.</td></tr>`;
        $("icListNote").textContent = "Newest first, Melbourne time. Click an item number to keep only that item; the ↗ opens its card in Wiise. A login shared by several people shows as one employee. Values longer than 140 characters are cut short.";
    }

    // ---------- entry point ----------
    window.renderItemChanges = function () {
        if (!allowed() || !SA()) return;
        scaffold();
        if (SA().state.status !== "ready") { paint("loading"); SA().ensure().then(() => paint()); return; }
        paint();
    };
    window.itemChanges = { state: S, repaint: () => paint() };
})();
