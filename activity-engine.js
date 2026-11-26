// ============================================================
// activity-engine.js — the staff activity "top-up", shared by
//   the robot    scripts/activity.js        (GitHub Actions, Node)
//   the browser  staff-activity.js          ("Get latest from Wiise")
// One copy of the rules, so the live figures and the robot's file can
// never disagree. Any change to how raised / changed / shipped /
// invoiced / received or the item changes are worked out goes HERE.
//
// activityTopUp(st, env) reads everything newer than the cursors in
// st.meta.cur from Wiise and adds it to st (the opened activity file).
//   env.odata(set, filter, select, label)  → rows   OData web services
//   env.api(pathAndQuery, label)           → rows   API v2.0, company scope
//   env.users()                            → rows   optional (robot only)
//   env.light                              true in the browser: skips the
//                                          whole-list name reads and looks
//                                          back two days for document headers
//   env.log / env.warn                     counts and field captions only:
//                                          NEVER names or amounts (the
//                                          robot's log is public)
// It changes st in place: hand it a copy if a half-finished run must
// not show (the browser does).
// ============================================================
"use strict";
const ACT_POSTED_FROM = "2026-07-01";             // posted documents are read from the start of FY26-27
const ACT_CREATE_WINDOW_S = 300;                  // a "this was just raised" line must come this soon after the first line
const ACT_BURST_S = 3600;                         // the creator's own lines in the first hour are part of raising it, not edits
const actMelFmt = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23" });
function actMel(d) {
    const p = {}; for (const x of actMelFmt.formatToParts(d)) p[x.type] = x.value;
    return p.year + "-" + p.month + "-" + p.day + "T" + (p.hour === "24" ? "00" : p.hour) + ":" + p.minute;
}
function actIsoAddDays(iso, days) { const d = new Date(iso + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() + days); return d.toISOString().slice(0, 10); }

async function activityTopUp(st, env) {
    const POSTED_FROM = ACT_POSTED_FROM, CREATE_WINDOW_S = ACT_CREATE_WINDOW_S, BURST_S = ACT_BURST_S, mel = actMel, isoAddDays = actIsoAddDays;
    const light = !!env.light, odata = env.odata, api = env.api;
    const log = env.log || (() => {}), warn = env.warn || (() => {});
    const num = (v) => (v == null || v === "") ? 0 : parseFloat(v) || 0;
    const dimBranch = (d) => { const m = /^(\d{2,})/.exec(String(d || "")); return m ? m[1] : ""; };
    const upper = (s) => String(s || "").trim().toUpperCase();
    // Try the tight filter first; if Wiise refuses it, fall back to a looser one
    // (the rows are checked again here either way).
    async function odataEither(set, filters, select, label) {
        let lastErr = null;
        for (let i = 0; i < filters.length; i++) {
            try { return await odata(set, filters[i], select, label); }
            catch (e) { lastErr = e; warn("  [" + label + "] filter " + (i + 1) + " refused (" + e.message.slice(0, 160) + ")" + (i + 1 < filters.length ? " — trying a looser one" : "")); }
        }
        throw lastErr;
    }
    const cur = st.meta.cur, first = !cur.ve;
    const todayMel = mel(new Date()).slice(0, 10);
    const recentFrom = first ? POSTED_FROM : isoAddDays(todayMel, light ? -2 : -14);

    // ---- names (small lists, read whole every run) ----
    const [users, custs, vends, locs] = await Promise.all([
        light || !env.users ? [] : env.users().catch(e => { warn("  [Users] not readable: " + e.message.slice(0, 120)); return []; }),
        light ? [] : api("customers?$select=number,displayName", "Customers"),
        light ? [] : api("vendors?$select=number,displayName", "Vendors"),
        light ? [] : api("locations?$select=code,displayName", "Locations").catch(() => []),
    ]);
    for (const u of users) if (u.userName) st.users[upper(u.userName)] = u.displayName || u.userName;
    for (const c of custs) if (c.number) st.cust[c.number] = c.displayName || "";
    for (const v of vends) if (v.number) st.vend[v.number] = v.displayName || "";
    for (const l of locs) if (l.code) st.locs[l.code] = l.displayName || l.code;
    log("Names: " + users.length + " users · " + custs.length + " customers · " + vends.length + " vendors · " + locs.length + " locations");

    // =====================================================
    // 1. CHANGE LOG → who raised / changed / closed each order
    // =====================================================
    const CL_SEL = "Entry_No,Date_and_Time,User_ID,Table_No,Primary_Key_Field_1_Value,Primary_Key_Field_2_Value,Field_No,Field_Caption,Type_of_Change,Old_Value,New_Value";
    const clBase = "Entry_No gt " + cur.cl + " and (Table_No eq 36 or Table_No eq 38)";
    // A deleted order is logged once per field (80+ lines); the "No." line
    // (field 3) is enough. Wiise refuses an OR across two different fields,
    // so that is two reads; if either is refused, read every line instead.
    let clRows, clSplit = true;
    try {
        // Field 2 of a removed document is its customer / vendor number: the only
        // place the party is written down when it was already on the document
        // as it was created (no "typed in" line exists for that).
        const [live, gone, goneParty] = await Promise.all([
            odata("ChangeLogEntries", clBase + " and Type_of_Change ne 'Deletion'", CL_SEL, "Change log"),
            odata("ChangeLogEntries", clBase + " and Type_of_Change eq 'Deletion' and Field_No eq 3", CL_SEL, "Change log (closed)"),
            odata("ChangeLogEntries", clBase + " and Type_of_Change eq 'Deletion' and (Field_No eq 2 or Field_No eq 151)", CL_SEL, "Change log (closed, party)"),
        ]);
        if (!gone.length && live.length > 3000) warn("  [Change log] no closing lines among " + live.length + " lines — check the Field_No 3 filter");
        clRows = live.concat(gone, goneParty);
    } catch (e) {
        warn("  [Change log] split read refused (" + e.message.slice(0, 160) + ") — reading every line");
        clSplit = false;
        clRows = await odata("ChangeLogEntries", clBase, CL_SEL, "Change log");
    }
    clRows.sort((a, b) => a.Entry_No - b.Entry_No);
    const fieldIdx = new Map(st.fields.map((c, i) => [c, i]));
    const fIdx = (cap) => { if (!fieldIdx.has(cap)) { fieldIdx.set(cap, st.fields.length); st.fields.push(cap); } return fieldIdx.get(cap); };
    const capFreq = { S: new Map(), P: new Map() };
    let clParty = 0;
    let clCreated = 0, clDeleted = 0, clEdits = 0, clSkipped = 0, clMax = cur.cl, logFrom = st.meta.logFrom || "";
    const seenDel = new Set();
    for (const r of clRows) {
        if (r.Entry_No <= cur.cl) { clSkipped++; continue; }
        if (r.Entry_No > clMax) clMax = r.Entry_No;
        if (r.Table_No !== 36 && r.Table_No !== 38) { clSkipped++; continue; }
        const kind = r.Table_No === 36 ? "S" : "P";
        const no = String(r.Primary_Key_Field_2_Value || "");
        if (!no) { clSkipped++; continue; }
        const key = kind + String(r.Primary_Key_Field_1_Value || "") + "|" + no;
        const when = new Date(r.Date_and_Time), sec = Math.floor(when.getTime() / 1000), ts = mel(when), user = upper(r.User_ID);
        if (!logFrom || ts < logFrom) logFrom = ts;
        let d = st.docs[key];
        if (!d) d = st.docs[key] = { f: [ts, user], f0: sec, c: 0, n: 0 };
        d.n++;
        if (r.Type_of_Change === "Deletion") {
            if (r.Field_No === 2 || r.Field_No === 151) {
                const v = String(r.Old_Value == null ? "" : r.Old_Value).trim();
                if (r.Field_No === 2) { if (v && !d.cu) { d.cu = v; clParty++; } }       // the customer / vendor the removed document was for
                else if (v) d.q = v;                                                    // field 151 Quote No.: the quote the order was made from
                if (seenDel.has(key) || clSplit) { d.n--; continue; }
            }
            if (!seenDel.has(key)) { seenDel.add(key); if (!d.d) clDeleted++; d.d = [ts, user]; }
            continue;
        }
        const cap = String(r.Field_Caption || ""), oldV = String(r.Old_Value == null ? "" : r.Old_Value).trim(), newV = String(r.New_Value == null ? "" : r.New_Value).trim();
        // facts the lines carry about the document
        if (/^(Sell-to Customer No\.?|Buy-from Vendor No\.?)$/i.test(cap)) { if (newV) d.cu = newV; }
        else if (/^Location Code$/i.test(cap)) { if (newV) d.br = newV; }
        else if (/^(Salesperson Code|Purchaser Code)$/i.test(cap)) d.sp = newV;
        else if (/^Assigned User ID$/i.test(cap)) d.au = upper(newV);
        else if (/^Quote No\.?$/i.test(cap)) { if (newV) d.q = newV; }
        // "just raised": early on, a field that was blank gets its first value
        if (!d.c && sec - d.f0 <= CREATE_WINDOW_S && oldV === "" && newV !== ""
            && /^(Assigned User ID|Location Code|Sell-to Customer No\.?|Buy-from Vendor No\.?)$/i.test(cap)) { d.c = 1; clCreated++; }
        // the first person's lines in the first hour are part of raising it
        if (user === d.f[1] && sec - d.f0 <= BURST_S) continue;
        clEdits++;
        capFreq[kind].set(cap, (capFreq[kind].get(cap) || 0) + 1);
        const ek = ts.slice(0, 13) + "|" + user;
        if (!d.e) d.e = {};
        const e = d.e[ek] || (d.e[ek] = [0, []]);
        e[0]++;
        const fi = fIdx(cap); if (!e[1].includes(fi)) e[1].push(fi);
    }
    cur.cl = clMax; st.meta.logFrom = logFrom;
    log("Change log: " + clRows.length + " lines read (" + clSkipped + " skipped) · " + clCreated + " documents raised · " + clDeleted + " closed · " + clEdits + " later changes · " + clParty + " closed documents given their customer / vendor · log starts " + (logFrom || "?").slice(0, 10));
    for (const k of ["S", "P"]) {
        const top = [...capFreq[k]].sort((a, b) => b[1] - a[1]).slice(0, 30).map(x => x[0] + " " + x[1]).join(" · ");
        if (top) log("  later changes by field (" + (k === "S" ? "sales" : "purchase") + "): " + top);
    }

    // =====================================================
    // 1b. CHANGE LOG → item card changes (Operations → Item Changes)
    //     Table 27 Item, plus the lists that hang off the card when they
    //     are logged too: 5404 units of measure, 99 item suppliers,
    //     90 kit contents, 5700 stockkeeping units (per-branch stock card). Every changed field is kept with its old and
    //     new value; a new or deleted item is logged once per field by
    //     Wiise and is kept here as one line. Never stops the rest of
    //     the run.
    // =====================================================
    try {
        delete st.meta.icErr;
        if (cur.ic == null) cur.ic = 0;
        if (!st.ic) st.ic = [];
        if (!st.items) st.items = {};
        // 5700 Stockkeeping Unit = the per-branch stock card (reorder point, maximum inventory … by location)
        const IT_TABLES = { 27: "I", 5404: "U", 99: "V", 90: "K", 5700: "S" };
        // Left out altogether: the save stamps, two flags the cost adjustment flips
        // on every posting, and the "... Id" copies of a code field shown beside it.
        const IT_DROP = /^(Last (Date|Time|Date ?Time) Modified|Cost is Adjusted|Allow Online Adjustment|.+ Id)$/i;
        { const before = st.ic.length; st.ic = st.ic.filter(r => r[4] !== 1 || !IT_DROP.test(st.fields[r[5]] || "")); if (before !== st.ic.length) log("Item change log: " + (before - st.ic.length) + " bookkeeping lines removed from the file"); }
        const clip = (s) => s.length > 140 ? s.slice(0, 139) + "…" : s;
        let icRows = await odata("ChangeLogEntries", "Entry_No gt " + cur.ic + " and (Table_No eq 27 or Table_No eq 5404 or Table_No eq 99 or Table_No eq 90 or Table_No eq 5700)", CL_SEL, "Item change log");
        // Table 5700 joined later than the others: its older lines are read once (robot only).
        const skuBackfill = !light && !st.meta.sku && cur.ic > 0;
        if (skuBackfill) {
            const old = await odata("ChangeLogEntries", "Entry_No le " + cur.ic + " and Table_No eq 5700", CL_SEL, "Item change log (stock cards)");
            log("Item change log: " + old.length + " earlier stock card (5700) lines read once");
            icRows = icRows.concat(old);
        }
        if (!light) st.meta.sku = 1;
        icRows.sort((a, b) => a.Entry_No - b.Entry_No);
        const whole = new Map(), icFreq = new Map();
        let icMax = cur.ic, icKept = 0, icNew = 0, icGone = 0, icDropped = 0, itemFrom = st.meta.itemFrom || "";
        for (const r of icRows) {
            if (r.Entry_No <= cur.ic && !(skuBackfill && r.Table_No === 5700)) continue;
            if (r.Entry_No > icMax) icMax = r.Entry_No;
            const t = IT_TABLES[r.Table_No];
            if (!t) continue;
            const pk1 = String(r.Primary_Key_Field_1_Value || ""), pk2 = String(r.Primary_Key_Field_2_Value || "");
            const item = t === "V" || t === "S" ? pk2 : pk1, sub = t === "I" ? "" : t === "V" || t === "S" ? pk1 : pk2;
            if (!item) continue;
            const cap = String(r.Field_Caption || "");
            if (IT_DROP.test(cap)) { icDropped++; continue; }
            const when = new Date(r.Date_and_Time), sec = Math.floor(when.getTime() / 1000), ts = mel(when), user = upper(r.User_ID);
            if (!itemFrom || ts < itemFrom) itemFrom = ts;
            const oldV = String(r.Old_Value == null ? "" : r.Old_Value).trim(), newV = String(r.New_Value == null ? "" : r.New_Value).trim();
            const type = r.Type_of_Change === "Insertion" ? 0 : r.Type_of_Change === "Deletion" ? 2 : 1;
            if (type !== 1) {
                // one line for the whole new / deleted record
                const k = t + "|" + item + "|" + sub + "|" + type + "|" + user;
                let w = whole.get(k);
                if (!w || sec - w.sec > 300) {
                    w = { sec, row: [ts, user, item, t, type, -1, "", "", sub, 0] };
                    whole.set(k, w); st.ic.push(w.row); icKept++;
                    if (type === 0) icNew++; else icGone++;
                }
                w.row[9]++;
                if (/^Description$/i.test(cap)) { if (type === 0) w.row[7] = clip(newV); else w.row[6] = clip(oldV); }
                continue;
            }
            st.ic.push([ts, user, item, t, 1, fIdx(cap), clip(oldV), clip(newV), sub]);
            icKept++;
            icFreq.set(cap, (icFreq.get(cap) || 0) + 1);
        }
        cur.ic = icMax; st.meta.itemFrom = itemFrom;
        // names for the items that have a line (read only when one is missing)
        if (st.ic.some(r => !(r[2] in st.items))) {
            const wanted = new Set(st.ic.map(r => r[2]));
            let items = [];
            if (light) {
                // a live top-up asks for the few missing items by number instead of the whole item list
                const missing = [...wanted].filter(no => !(no in st.items)).slice(0, 200);
                for (let i = 0; i < missing.length; i += 15) {
                    const f = missing.slice(i, i + 15).map(no => "number eq '" + no.replace(/'/g, "''") + "'").join(" or ");
                    items = items.concat(await api("items?$select=number,displayName,itemCategoryCode&$filter=" + encodeURIComponent(f), "Items").catch(() => []));
                }
            } else items = await api("items?$select=number,displayName,itemCategoryCode", "Items").catch(e => { warn("  [Items] not readable: " + e.message.slice(0, 120)); return []; });
            for (const it of items) if (wanted.has(it.number)) st.items[it.number] = [it.displayName || "", it.itemCategoryCode || ""];
            for (const no of wanted) if (!(no in st.items)) st.items[no] = ["", ""];      // deleted since: asked once only
        }
        log("Item change log: " + icRows.length + " lines read · " + icKept + " kept (" + icNew + " new records, " + icGone + " deleted) · " + icDropped + " bookkeeping lines left out · file holds " + st.ic.length + " lines on " + Object.keys(st.items).length + " items" + (itemFrom ? " from " + itemFrom.slice(0, 10) : ""));
        const topIc = [...icFreq].sort((a, b) => b[1] - a[1]).slice(0, 30).map(x => x[0] + " " + x[1]).join(" · ");
        if (topIc) log("  item changes by field: " + topIc);
    } catch (e) {
        st.meta.icErr = String(e.message || e).slice(0, 200);
        warn("  [Item change log] not added this run: " + st.meta.icErr);
    }

    // =====================================================
    // 1c. CHANGE LOG → customer and vendor card changes
    //     (Operations → Customer Changes / Vendor Changes)
    //     18 Customer, 222 Ship-to Address, 287 Customer Bank Account;
    //     23 Vendor, 224 Order Address, 288 Vendor Bank Account.
    //     Same line shape as the item changes. Bank account numbers are
    //     kept by their last three digits only: the file is opened with
    //     the passphrase every user of the tool holds.
    // =====================================================
    {
        const CARD_DROP = /^(Last (Date|Time|Date ?Time) Modified|Last Modified Date Time|.+ Id)$/i;
        const SECRET = /(Bank Account No|Bank Branch No|IBAN|SWIFT|BSB|Credit Card|Tax File)/i;
        const maskV = (v) => v ? "•••" + v.slice(-3) : "";
        const clipV = (s) => s.length > 140 ? s.slice(0, 139) + "…" : s;
        for (const g of [
            { key: "cc", from: "custFrom", label: "Customer", tables: { 18: "C", 222: "A", 287: "B" } },
            { key: "vc", from: "vendFrom", label: "Vendor", tables: { 23: "C", 224: "A", 288: "B" } },
        ]) {
            try {
                delete st.meta[g.key + "Err"];
                if (light && cur[g.key] == null) continue;      // not built by the robot yet: the whole history is not a job for a live top-up
                if (cur[g.key] == null) cur[g.key] = 0;
                if (!st[g.key]) st[g.key] = [];
                const bag = st[g.key], c0 = cur[g.key];
                const rows = await odata("ChangeLogEntries", "Entry_No gt " + c0 + " and (" + Object.keys(g.tables).map(n => "Table_No eq " + n).join(" or ") + ")", CL_SEL, g.label + " change log");
                rows.sort((a, b) => a.Entry_No - b.Entry_No);
                const whole = new Map(), freq = new Map();
                let max = c0, kept = 0, neu = 0, gone = 0, dropped = 0, from = st.meta[g.from] || "";
                for (const r of rows) {
                    if (r.Entry_No <= c0) continue;
                    if (r.Entry_No > max) max = r.Entry_No;
                    const t = g.tables[r.Table_No];
                    if (!t) continue;
                    const no = String(r.Primary_Key_Field_1_Value || ""), sub = t === "C" ? "" : String(r.Primary_Key_Field_2_Value || "");
                    if (!no) continue;
                    const cap = String(r.Field_Caption || "");
                    if (CARD_DROP.test(cap)) { dropped++; continue; }
                    const when = new Date(r.Date_and_Time), sec = Math.floor(when.getTime() / 1000), ts = mel(when), user = upper(r.User_ID);
                    if (!from || ts < from) from = ts;
                    let oldV = String(r.Old_Value == null ? "" : r.Old_Value).trim(), newV = String(r.New_Value == null ? "" : r.New_Value).trim();
                    if (SECRET.test(cap)) { oldV = maskV(oldV); newV = maskV(newV); }
                    const type = r.Type_of_Change === "Insertion" ? 0 : r.Type_of_Change === "Deletion" ? 2 : 1;
                    if (type !== 1) {
                        const k = t + "|" + no + "|" + sub + "|" + type + "|" + user;
                        let w = whole.get(k);
                        if (!w || sec - w.sec > 300) {
                            w = { sec, row: [ts, user, no, t, type, -1, "", "", sub, 0] };
                            whole.set(k, w); bag.push(w.row); kept++;
                            if (type === 0) neu++; else gone++;
                        }
                        w.row[9]++;
                        if (/^Name$/i.test(cap)) { if (type === 0) w.row[7] = clipV(newV); else w.row[6] = clipV(oldV); }
                        continue;
                    }
                    bag.push([ts, user, no, t, 1, fIdx(cap), clipV(oldV), clipV(newV), sub]);
                    kept++;
                    freq.set(cap, (freq.get(cap) || 0) + 1);
                }
                cur[g.key] = max; st.meta[g.from] = from;
                log(g.label + " change log: " + rows.length + " lines read · " + kept + " kept (" + neu + " new records, " + gone + " deleted) · " + dropped + " bookkeeping lines left out · file holds " + bag.length + " lines" + (from ? " from " + from.slice(0, 10) : ""));
                const top = [...freq].sort((a, b) => b[1] - a[1]).slice(0, 30).map(x => x[0] + " " + x[1]).join(" · ");
                if (top) log("  " + g.label.toLowerCase() + " changes by field: " + top);
            } catch (e) {
                st.meta[g.key + "Err"] = String(e.message || e).slice(0, 200);
                warn("  [" + g.label + " change log] not added this run: " + st.meta[g.key + "Err"]);
            }
        }
    }

    // =====================================================
    // 2. VALUE ENTRIES + ITEM LEDGER → who shipped / received
    // =====================================================
    const veScope = "(Item_Ledger_Entry_Type eq 'Sale' or Item_Ledger_Entry_Type eq 'Purchase')";
    const veWhere = first ? "Posting_Date ge " + POSTED_FROM : "Entry_No gt " + cur.ve;
    const VE_SEL = "Entry_No,Posting_Date,Item_Ledger_Entry_Type,Entry_Type,Adjustment,Document_Type,Document_No,Sales_Amount_Expected,Sales_Amount_Actual,Cost_Amount_Expected,Cost_Amount_Actual,Item_Ledger_Entry_Quantity,Location_Code,Salespers_Purch_Code,User_ID,Source_No,Item_Ledger_Entry_No";
    const ileWhere = first ? "Posting_Date ge " + POSTED_FROM : "Entry_No gt " + cur.ile;
    const [veRows, ileRows] = await Promise.all([
        odataEither("ValueEntriesFull", [veWhere + " and " + veScope + " and Adjustment eq false", veWhere + " and " + veScope, veWhere], VE_SEL, "Value entries"),
        odataEither("Item_Ledger_Entries", [ileWhere + " and (Entry_Type eq 'Sale' or Entry_Type eq 'Purchase')", ileWhere], "Entry_No,Document_No,Entry_Type", "Item ledger"),
    ]);
    const ileDoc = new Map();
    let ileMax = cur.ile;
    for (const r of ileRows) { ileDoc.set(r.Entry_No, r.Document_No); if (r.Entry_No > ileMax) ileMax = r.Entry_No; }
    let veMax = cur.ve, veUsed = 0, veOut = 0, noIle = 0;
    const invLoc = new Map();
    veRows.sort((a, b) => a.Entry_No - b.Entry_No);
    for (const v of veRows) {
        if (v.Entry_No > veMax) veMax = v.Entry_No;
        if (!first && v.Entry_No <= cur.ve) { veOut++; continue; }
        const day = String(v.Posting_Date || "").slice(0, 10);
        if (day < POSTED_FROM) { veOut++; continue; }
        if (v.Adjustment === true) continue;
        const type = v.Item_Ledger_Entry_Type, dt = v.Document_Type, docNo = String(v.Document_No || "");
        if (type !== "Sale" && type !== "Purchase") continue;
        if (docNo && v.Location_Code && (dt === "Sales Invoice" || dt === "Sales Credit Memo" || dt === "Purchase Invoice" || dt === "Purchase Credit Memo") && !invLoc.has(dt + "|" + docNo)) invLoc.set(dt + "|" + docNo, v.Location_Code);
        const q = num(v.Item_Ledger_Entry_Quantity);
        if (!q) continue;                                  // an invoice of goods moved earlier: no movement here
        let bag, withInv = "", amt, qty;
        if (type === "Sale" && (dt === "Sales Shipment" || dt === "Sales Invoice")) {
            bag = st.sh; qty = -q;
            amt = dt === "Sales Shipment" ? num(v.Sales_Amount_Expected) || num(v.Sales_Amount_Actual) : num(v.Sales_Amount_Actual) || num(v.Sales_Amount_Expected);
            if (dt === "Sales Invoice") withInv = docNo;
        } else if (type === "Purchase" && (dt === "Purchase Receipt" || dt === "Purchase Invoice")) {
            bag = st.pr; qty = q;
            amt = dt === "Purchase Receipt" ? num(v.Cost_Amount_Expected) || num(v.Cost_Amount_Actual) : num(v.Cost_Amount_Actual) || num(v.Cost_Amount_Expected);
            if (dt === "Purchase Invoice") withInv = docNo;
        } else continue;
        // the goods document: its own number, or (posted together with the invoice) the one on the stock entry
        let no = withInv ? (ileDoc.get(v.Item_Ledger_Entry_No) || "") : docNo;
        if (!no) { noIle++; no = "with " + withInv; }
        const user = upper(v.User_ID);
        let d = bag[no];
        if (!d) d = bag[no] = { d: day, u: user, c: String(v.Source_No || ""), b: String(v.Location_Code || ""), s: String(v.Salespers_Purch_Code || ""), a: 0, q: 0, l: 0 };
        d.a = Math.round((d.a + amt) * 100) / 100; d.q = Math.round((d.q + qty) * 1000) / 1000; d.l++;
        if (withInv) d.i = withInv;
        veUsed++;
    }
    cur.ve = veMax; cur.ile = ileMax;
    log("Value entries: " + veRows.length + " read (" + veOut + " outside the window) · " + veUsed + " stock movements · " + ileRows.length + " item ledger rows · " + noIle + " without a goods document number");
    if (!first && veOut > veRows.length / 2 && veRows.length > 1000) throw new Error("Value entry filter was not applied by Wiise (" + veOut + " of " + veRows.length + " rows outside it) — stopping rather than re-reading everything each run");

    // =====================================================
    // 3. CUSTOMER / VENDOR LEDGER → who posted invoices and credit memos
    // =====================================================
    const docTypes = "(Document_Type eq 'Invoice' or Document_Type eq 'Credit Memo')";
    const cleWhere = first ? "Posting_Date ge " + POSTED_FROM : "Entry_No gt " + cur.cle;
    const vleWhere = first ? "Posting_Date ge " + POSTED_FROM : "Entry_No gt " + cur.vle;
    const [cleRows, vleRows] = await Promise.all([
        odataEither("Customer_Ledger_Entries_Excel", [cleWhere + " and " + docTypes, cleWhere],
            "Entry_No,Posting_Date,Document_Type,Document_No,Customer_No,Salesperson_Code,Global_Dimension_1_Code,Sales_LCY,User_ID", "Customer ledger"),
        odataEither("Vendor_Ledger_Entries_Excel", [vleWhere + " and " + docTypes, vleWhere],
            "Entry_No,Posting_Date,Document_Type,Document_No,Vendor_No,Purchaser_Code,Global_Dimension_1_Code,Amount_LCY,User_ID", "Vendor ledger"),
    ]);
    let cleMax = cur.cle, vleMax = cur.vle, nSi = 0, nSc = 0, nPi = 0, nPc = 0;
    for (const r of cleRows) {
        if (r.Entry_No > cleMax) cleMax = r.Entry_No;
        const day = String(r.Posting_Date || "").slice(0, 10), no = String(r.Document_No || "");
        if (!no || day < POSTED_FROM || (!first && r.Entry_No <= cur.cle)) continue;
        const inv = r.Document_Type === "Invoice", cm = r.Document_Type === "Credit Memo";
        if (!inv && !cm) continue;
        const bag = inv ? st.si : st.sc;
        if (!bag[no]) { if (inv) nSi++; else nSc++; }
        const prev = bag[no] || {};
        bag[no] = Object.assign(prev, { d: day, u: upper(r.User_ID), c: String(r.Customer_No || ""), s: String(r.Salesperson_Code || ""),
            a: Math.round(num(r.Sales_LCY) * 100) / 100,
            b: invLoc.get((inv ? "Sales Invoice|" : "Sales Credit Memo|") + no) || prev.b || dimBranch(r.Global_Dimension_1_Code) });
    }
    for (const r of vleRows) {
        if (r.Entry_No > vleMax) vleMax = r.Entry_No;
        const day = String(r.Posting_Date || "").slice(0, 10), no = String(r.Document_No || "");
        if (!no || day < POSTED_FROM || (!first && r.Entry_No <= cur.vle)) continue;
        const inv = r.Document_Type === "Invoice", cm = r.Document_Type === "Credit Memo";
        if (!inv && !cm) continue;
        const bag = inv ? st.pi : st.pc;
        if (!bag[no]) { if (inv) nPi++; else nPc++; }
        const prev = bag[no] || {};
        // the ledger amount includes GST (and is negative for an invoice); the ex-GST figure is filled in below
        bag[no] = Object.assign(prev, { d: day, u: upper(r.User_ID), c: String(r.Vendor_No || ""), s: String(r.Purchaser_Code || ""),
            g: Math.round(-num(r.Amount_LCY) * 100) / 100,
            b: invLoc.get((inv ? "Purchase Invoice|" : "Purchase Credit Memo|") + no) || prev.b || dimBranch(r.Global_Dimension_1_Code) });
    }
    cur.cle = cleMax; cur.vle = vleMax;
    log("Ledgers: " + cleRows.length + " customer rows → " + nSi + " invoices, " + nSc + " credit memos · " + vleRows.length + " vendor rows → " + nPi + " invoices, " + nPc + " credit memos");

    // =====================================================
    // 4. Document headers → order numbers, ex-GST purchase amounts, open order values
    // =====================================================
    const since = "postingDate ge " + recentFrom;
    const quiet = (p, label) => api(p, label).catch(e => { warn("  [" + label + "] not readable: " + e.message.slice(0, 140)); return []; });
    const [shH, siH, prH, piH, pcH, openSo, openPo, openSq] = await Promise.all([
        quiet("salesShipments?$select=number,orderNumber&$filter=" + encodeURIComponent(since), "Shipment headers"),
        quiet("salesInvoices?$select=number,orderNumber&$filter=" + encodeURIComponent(since), "Invoice headers"),
        quiet("purchaseReceipts?$select=number,orderNumber&$filter=" + encodeURIComponent(since), "Receipt headers"),
        quiet("purchaseInvoices?$select=number,orderNumber,totalAmountExcludingTax&$filter=" + encodeURIComponent(since), "Purchase invoice headers"),
        quiet("purchaseCreditMemos?$select=number,totalAmountExcludingTax&$filter=" + encodeURIComponent(since), "Purchase credit memo headers"),
        quiet("salesOrders?$select=number,totalAmountExcludingTax,customerNumber", "Open sales orders"),
        quiet("purchaseOrders?$select=number,totalAmountExcludingTax,vendorNumber", "Open purchase orders"),
        quiet("salesQuotes?$select=number,totalAmountExcludingTax,customerNumber", "Open sales quotes"),
    ]);
    let linked = 0;
    const link = (rows, bag, amt) => { for (const h of rows) { const d = bag[h.number]; if (!d) continue; if (h.orderNumber) { d.o = h.orderNumber; linked++; } if (amt && h.totalAmountExcludingTax != null) d.a = Math.round(num(h.totalAmountExcludingTax) * 100) / 100; } };
    link(shH, st.sh); link(siH, st.si); link(prH, st.pr); link(piH, st.pi, true); link(pcH, st.pc, true);
    // open documents: today's value, and "still open" for everything else
    const openNow = new Map();
    for (const o of openSo) openNow.set("S1|" + o.number, o);
    for (const o of openPo) openNow.set("P1|" + o.number, o);
    for (const o of openSq) openNow.set("S0|" + o.number, o);
    for (const k of Object.keys(st.docs)) {
        const d = st.docs[k], o = openNow.get(k);
        if (o) { d.a = Math.round(num(o.totalAmountExcludingTax) * 100) / 100; d.o = 1; if (!d.cu) d.cu = String(o.customerNumber || o.vendorNumber || ""); }
        else if (d.o) delete d.o;
    }
    log("Headers: " + linked + " goods/invoice documents tied to their order · open now: " + openSo.length + " sales orders, " + openPo.length + " purchase orders, " + openSq.length + " quotes");

    return st;
}

if (typeof module !== "undefined" && module.exports) {
    module.exports = { activityTopUp, ACT_POSTED_FROM, actMel };
}
