// ============================================================
// activity.js — the "staff activity" robot (Operations → Staff Activity).
// Runs in GitHub Actions (Node 20). Builds ONE small encrypted file that
// says who raised, changed, shipped, invoiced and received what:
//
//   who CREATED / CHANGED an order   ← Wiise Change Log (web service
//                                      "ChangeLogEntries", page 595):
//                                      tables 36 Sales Header and
//                                      38 Purchase Header. The log holds no
//                                      "insert" line; the first line of a
//                                      document (seconds after it is raised)
//                                      names the creator.
//   who POSTED a shipment / receipt  ← ValueEntriesFull.User_ID (the value
//                                      entry written with the stock movement)
//   who POSTED an invoice / credit   ← customer and vendor ledger User_ID
//
// The file is kept on SharePoint beside the snapshots and is added to on
// every run (each list is read from its last entry number on), so the
// history survives even if the Change Log is trimmed in Wiise.
//
// Env (GitHub secrets):
//   BC_CLIENT_ID, BC_TENANT, BC_REFRESH_TOKEN   token exchange
//   SNAPSHOT_PASSPHRASE                         default lock of the file
//   ACTIVITY_PASSPHRASE (optional)              own lock, when set
//   ACTIVITY_REBUILD=true                       start again from nothing
//   ACTIVITY_INCREMENTAL_ONLY=true              do nothing when no file
//                                               exists yet (scheduled runs)
//
// NEVER print user names, customer names or amounts here: the repo and
// its Actions logs are public. Counts and field captions only.
// ============================================================
"use strict";
const path = require("path");
const zlib = require("zlib");
const crypto = require("crypto");

global.state = {};
global.num = (v) => (v == null || v === "") ? 0 : parseFloat(v) || 0;
let lastProgress = 0;
global.updateBCProgress = (label, detail) => {   // one line every few seconds is plenty
    if (Date.now() - lastProgress > 4000) { lastProgress = Date.now(); console.log("  [" + label + "] " + detail); }
};

const TENANT = process.env.BC_TENANT;
const CLIENT_ID = process.env.BC_CLIENT_ID;
const REFRESH_TOKEN = process.env.BC_REFRESH_TOKEN;
const OWN_PASS = (process.env.ACTIVITY_PASSPHRASE || "").trim();
const PASSPHRASE = OWN_PASS || process.env.SNAPSHOT_PASSPHRASE;
const LOCK = OWN_PASS ? "activity" : "snapshot";
const REBUILD = process.env.ACTIVITY_REBUILD === "true";
const INCREMENTAL_ONLY = process.env.ACTIVITY_INCREMENTAL_ONLY === "true";
if (!TENANT || !CLIENT_ID || !REFRESH_TOKEN || !PASSPHRASE) {
    console.error("Missing env: need BC_TENANT, BC_CLIENT_ID, BC_REFRESH_TOKEN, SNAPSHOT_PASSPHRASE");
    process.exit(1);
}

let _tok = null, _tokExp = 0;
global.bcGetToken = async function bcGetToken() {
    if (_tok && Date.now() < _tokExp - 300000) return _tok;
    const resp = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ grant_type: "refresh_token", client_id: CLIENT_ID, refresh_token: REFRESH_TOKEN,
            scope: "https://api.businesscentral.dynamics.com/user_impersonation offline_access" }).toString(),
    });
    const tok = await resp.json();
    if (!tok.access_token) throw new Error("Token exchange failed: " + String(tok.error || "unknown"));
    _tok = tok.access_token;
    _tokExp = Date.now() + (tok.expires_in || 3600) * 1000;
    return _tok;
};

const F = require(path.join(__dirname, "..", "bc-fetchers.js"));
const SP = require(path.join(__dirname, "sp-upload.js"));
const PBKDF2_ITERATIONS = 600000;                 // keep in sync with decryptSnapshot() in index.html
const FILE = F.SP_CONFIG.folder + "/snapshots/activity.bin";
const META = F.SP_CONFIG.folder + "/snapshots/activity.meta.json";
const FORMAT = 1;
const POSTED_FROM = "2026-07-01";                 // posted documents are read from the start of FY26-27
const CREATE_WINDOW_S = 300;                      // a "this was just raised" line must come this soon after the first line
const BURST_S = 3600;                             // the creator's own lines in the first hour are part of raising it, not edits

// ---------- Melbourne wall clock ----------
const melFmt = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23" });
function mel(d) {
    const p = {}; for (const x of melFmt.formatToParts(d)) p[x.type] = x.value;
    return p.year + "-" + p.month + "-" + p.day + "T" + (p.hour === "24" ? "00" : p.hour) + ":" + p.minute;
}
function isoAddDays(iso, days) { const d = new Date(iso + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() + days); return d.toISOString().slice(0, 10); }

// ---------- lock / unlock ----------
function encrypt(obj) {
    const gz = zlib.gzipSync(Buffer.from(JSON.stringify(obj), "utf8"), { level: 9 });
    const salt = crypto.randomBytes(16), iv = crypto.randomBytes(12);
    const key = crypto.pbkdf2Sync(PASSPHRASE, salt, PBKDF2_ITERATIONS, 32, "sha256");
    const c = crypto.createCipheriv("aes-256-gcm", key, iv);
    return Buffer.concat([salt, iv, c.update(gz), c.final(), c.getAuthTag()]);
}
function decrypt(bin) {
    const salt = bin.subarray(0, 16), iv = bin.subarray(16, 28), tag = bin.subarray(bin.length - 16), ct = bin.subarray(28, bin.length - 16);
    const key = crypto.pbkdf2Sync(PASSPHRASE, salt, PBKDF2_ITERATIONS, 32, "sha256");
    const d = crypto.createDecipheriv("aes-256-gcm", key, iv);
    d.setAuthTag(tag);
    return JSON.parse(zlib.gunzipSync(Buffer.concat([d.update(ct), d.final()])).toString("utf8"));
}

function emptyState() {
    return {
        meta: { v: FORMAT, lock: LOCK, postedFrom: POSTED_FROM, logFrom: "", cur: { cl: 0, ve: 0, ile: 0, cle: 0, vle: 0 } },
        users: {}, locs: {}, cust: {}, vend: {}, fields: [],
        docs: {},                 // "S1|SO123" / "P1|PO123" → order, quote … header record (Change Log)
        sh: {}, si: {}, sc: {},   // sales shipments, invoices, credit memos
        pr: {}, pi: {}, pc: {},   // purchase receipts, invoices, credit memos
    };
}

// ---------- reading Wiise ----------
let CO = "", C2 = "";
async function odata(set, filter, select, label) {
    const url = F.BC_ODATA_URL + "/Company('" + encodeURIComponent(CO) + "')/" + set + "?$filter=" + encodeURIComponent(filter) + (select ? "&$select=" + select : "");
    return F.bcFetchAll(url, label);
}
async function api(pathAndQuery, label) { return F.bcFetchAll(C2 + "/" + pathAndQuery, label); }
// Try the tight filter first; if Wiise refuses it, fall back to a looser one
// (the rows are checked again here either way).
async function odataEither(set, filters, select, label) {
    let lastErr = null;
    for (let i = 0; i < filters.length; i++) {
        try { return await odata(set, filters[i], select, label); }
        catch (e) { lastErr = e; console.warn("  [" + label + "] filter " + (i + 1) + " refused (" + e.message.slice(0, 160) + ")" + (i + 1 < filters.length ? " — trying a looser one" : "")); }
    }
    throw lastErr;
}
const dimBranch = (d) => { const m = /^(\d{2,})/.exec(String(d || "")); return m ? m[1] : ""; };
const upper = (s) => String(s || "").trim().toUpperCase();

(async () => {
    const t0 = Date.now();
    const driveId = await SP.resolveDrive(F.SP_CONFIG.host);

    // ---- previous state ----
    let st = null;
    if (!REBUILD) {
        try { st = decrypt(await SP.downloadFile(driveId, FILE)); }
        catch (e) {
            if (/HTTP 404/.test(e.message)) console.log("No activity file on SharePoint yet.");
            else throw new Error("Cannot open the existing activity file (" + e.message + "). Run with rebuild to start again.");
        }
    }
    if (st && (st.meta || {}).v !== FORMAT) { console.log("Activity file is an older format — starting again."); st = null; }
    if (!st) {
        if (INCREMENTAL_ONLY) { console.log("Nothing to add to yet (first build is started by hand). Done."); return; }
        st = emptyState();
    }
    st.meta.lock = LOCK;
    const cur = st.meta.cur, first = !cur.ve;
    console.log((first ? "First build" : "Adding to the existing file") + " · cursors: log " + cur.cl + " · value " + cur.ve + " · item " + cur.ile + " · customer " + cur.cle + " · vendor " + cur.vle);

    await F.bcGetCompanyId();
    CO = await F.bcGetCompanyInternalName();
    C2 = F.BC_API_URL + "/companies(" + (await F.bcGetCompanyId()) + ")";
    const todayMel = mel(new Date()).slice(0, 10);
    const recentFrom = first ? POSTED_FROM : isoAddDays(todayMel, -14);

    // ---- names (small lists, read whole every run) ----
    const [users, custs, vends, locs] = await Promise.all([
        F.bcFetchAll(F.BC_API_BASE + "/api/microsoft/automation/v2.0/companies(" + (await F.bcGetCompanyId()) + ")/users?$select=userName,displayName,state", "Users").catch(e => { console.warn("  [Users] not readable: " + e.message.slice(0, 120)); return []; }),
        api("customers?$select=number,displayName", "Customers"),
        api("vendors?$select=number,displayName", "Vendors"),
        api("locations?$select=code,displayName", "Locations").catch(() => []),
    ]);
    for (const u of users) if (u.userName) st.users[upper(u.userName)] = u.displayName || u.userName;
    for (const c of custs) if (c.number) st.cust[c.number] = c.displayName || "";
    for (const v of vends) if (v.number) st.vend[v.number] = v.displayName || "";
    for (const l of locs) if (l.code) st.locs[l.code] = l.displayName || l.code;
    console.log("Names: " + users.length + " users · " + custs.length + " customers · " + vends.length + " vendors · " + locs.length + " locations");

    // =====================================================
    // 1. CHANGE LOG → who raised / changed / closed each order
    // =====================================================
    const CL_SEL = "Entry_No,Date_and_Time,User_ID,Table_No,Primary_Key_Field_1_Value,Primary_Key_Field_2_Value,Field_No,Field_Caption,Type_of_Change,Old_Value,New_Value";
    const clBase = "Entry_No gt " + cur.cl + " and (Table_No eq 36 or Table_No eq 38)";
    // A deleted order is logged once per field (80+ lines); the "No." line
    // (field 3) is enough. Wiise refuses an OR across two different fields,
    // so that is two reads; if either is refused, read every line instead.
    let clRows;
    try {
        const [live, gone] = await Promise.all([
            odata("ChangeLogEntries", clBase + " and Type_of_Change ne 'Deletion'", CL_SEL, "Change log"),
            odata("ChangeLogEntries", clBase + " and Type_of_Change eq 'Deletion' and Field_No eq 3", CL_SEL, "Change log (closed)"),
        ]);
        if (!gone.length && live.length > 3000) console.warn("  [Change log] no closing lines among " + live.length + " lines — check the Field_No 3 filter");
        clRows = live.concat(gone);
    } catch (e) {
        console.warn("  [Change log] split read refused (" + e.message.slice(0, 160) + ") — reading every line");
        clRows = await odata("ChangeLogEntries", clBase, CL_SEL, "Change log");
    }
    clRows.sort((a, b) => a.Entry_No - b.Entry_No);
    const fieldIdx = new Map(st.fields.map((c, i) => [c, i]));
    const fIdx = (cap) => { if (!fieldIdx.has(cap)) { fieldIdx.set(cap, st.fields.length); st.fields.push(cap); } return fieldIdx.get(cap); };
    const capFreq = { S: new Map(), P: new Map() };
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
    console.log("Change log: " + clRows.length + " lines read (" + clSkipped + " skipped) · " + clCreated + " documents raised · " + clDeleted + " closed · " + clEdits + " later changes · log starts " + (logFrom || "?").slice(0, 10));
    for (const k of ["S", "P"]) {
        const top = [...capFreq[k]].sort((a, b) => b[1] - a[1]).slice(0, 30).map(x => x[0] + " " + x[1]).join(" · ");
        if (top) console.log("  later changes by field (" + (k === "S" ? "sales" : "purchase") + "): " + top);
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
    console.log("Value entries: " + veRows.length + " read (" + veOut + " outside the window) · " + veUsed + " stock movements · " + ileRows.length + " item ledger rows · " + noIle + " without a goods document number");
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
    console.log("Ledgers: " + cleRows.length + " customer rows → " + nSi + " invoices, " + nSc + " credit memos · " + vleRows.length + " vendor rows → " + nPi + " invoices, " + nPc + " credit memos");

    // =====================================================
    // 4. Document headers → order numbers, ex-GST purchase amounts, open order values
    // =====================================================
    const since = "postingDate ge " + recentFrom;
    const quiet = (p, label) => api(p, label).catch(e => { console.warn("  [" + label + "] not readable: " + e.message.slice(0, 140)); return []; });
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
    console.log("Headers: " + linked + " goods/invoice documents tied to their order · open now: " + openSo.length + " sales orders, " + openPo.length + " purchase orders, " + openSq.length + " quotes");

    // ---- totals (counts only) ----
    const count = (o) => Object.keys(o).length;
    const docKinds = {}; let confident = 0, withEdits = 0;
    for (const k of Object.keys(st.docs)) { const p = k.slice(0, 2); docKinds[p] = (docKinds[p] || 0) + 1; if (st.docs[k].c) confident++; if (st.docs[k].e) withEdits++; }
    console.log("File now holds: " + count(st.docs) + " order-type documents " + JSON.stringify(docKinds) + " (" + confident + " seen being raised, " + withEdits + " with later changes) · "
        + count(st.sh) + " shipments · " + count(st.si) + " sales invoices · " + count(st.sc) + " sales credit memos · "
        + count(st.pr) + " receipts · " + count(st.pi) + " purchase invoices · " + count(st.pc) + " purchase credit memos");
    const noUser = (o) => Object.values(o).filter(d => !d.u).length;
    console.log("Without a user: shipments " + noUser(st.sh) + " · invoices " + noUser(st.si) + " · receipts " + noUser(st.pr) + " · purchase invoices " + noUser(st.pi)
        + " · purchase invoices without an ex-GST amount " + Object.values(st.pi).filter(d => d.a == null).length);

    // ---- write ----
    const now = new Date();
    st.meta.builtAtUtc = now.toISOString();
    st.meta.builtAtMelbourne = mel(now).replace("T", " ");
    const bin = encrypt(st);
    await SP.uploadFile(driveId, FILE, bin);
    await SP.uploadFile(driveId, META, Buffer.from(JSON.stringify({ v: FORMAT, lock: LOCK, builtAtUtc: st.meta.builtAtUtc, builtAtMelbourne: st.meta.builtAtMelbourne, bytes: bin.length, logFrom: st.meta.logFrom, postedFrom: POSTED_FROM }, null, 2), "utf8"));
    console.log("Published activity.bin (" + (bin.length / 1048576).toFixed(2) + " MB) in " + ((Date.now() - t0) / 1000).toFixed(0) + "s");
})().catch(e => {
    console.error("ACTIVITY FILE NOT UPDATED:", e.message);
    process.exit(1);
});
