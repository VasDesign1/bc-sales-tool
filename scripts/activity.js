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
// HOW each figure is worked out lives in ../activity-engine.js, shared with
// the tab's "Get latest from Wiise" button. This file signs in, opens the
// existing file, runs the engine and writes the file back.
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
const E = require(path.join(__dirname, "..", "activity-engine.js"));   // the rules, shared with the browser's live top-up
const PBKDF2_ITERATIONS = 600000;                 // keep in sync with decryptSnapshot() in index.html
const FILE = F.SP_CONFIG.folder + "/snapshots/activity.bin";
const META = F.SP_CONFIG.folder + "/snapshots/activity.meta.json";
const FORMAT = 1;
const POSTED_FROM = E.ACT_POSTED_FROM;

// ---------- Melbourne wall clock ----------
const melFmt = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23" });
function mel(d) {
    const p = {}; for (const x of melFmt.formatToParts(d)) p[x.type] = x.value;
    return p.year + "-" + p.month + "-" + p.day + "T" + (p.hour === "24" ? "00" : p.hour) + ":" + p.minute;
}

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

    // ---- everything newer than the cursors (rules live in activity-engine.js, shared with the browser) ----
    await E.activityTopUp(st, {
        odata, api, log: (m) => console.log(m), warn: (m) => console.warn(m),
        users: async () => F.bcFetchAll(F.BC_API_BASE + "/api/microsoft/automation/v2.0/companies(" + (await F.bcGetCompanyId()) + ")/users?$select=userName,displayName,state", "Users"),
    });

    // ---- check on the blanket order figures (counts, field captions and seconds only) ----
    try {
        const docNum = (k) => { const m = /^(.*?)(\d+)$/.exec(k.slice(3)); return m ? [m[1], +m[2]] : null; };
        // Per number series: the 10th lowest number that was seen being raised. Any
        // higher number was certainly raised after the log began, so one of those
        // without "seen being raised" is a document the creator rule missed.
        const floor = {};
        for (const type of ["S0", "S1", "S4", "P1"]) {
            const by = {};
            for (const k of Object.keys(st.docs)) { if (!k.startsWith(type + "|") || !st.docs[k].c) continue; const n = docNum(k); if (n) (by[n[0]] = by[n[0]] || []).push(n[1]); }
            for (const p of Object.keys(by)) { by[p].sort((a, b) => a - b); floor[type + "|" + p] = by[p][Math.min(9, by[p].length - 1)]; }
        }
        const after = (k) => { const n = docNum(k); if (!n) return false; const f = floor[k.slice(0, 2) + "|" + n[0]]; return f != null && n[1] > f; };
        const missed = {};
        for (const type of ["S0", "S1", "S4", "P1"]) {
            let all = 0, c1 = 0, c0After = 0, gone = 0; const who = new Set();
            for (const k of Object.keys(st.docs)) {
                if (!k.startsWith(type + "|")) continue;
                const d = st.docs[k]; all++;
                if (d.c) { c1++; who.add(d.f[1]); } else if (after(k)) { c0After++; (missed[type] = missed[type] || []).push(k); }
                if (d.d) gone++;
            }
            console.log("Check " + type + ": " + all + " in the file · " + c1 + " seen being raised (by " + who.size + " users) · " + c0After + " raised after the log began but NOT seen being raised · " + gone + " removed");
        }
        const bo = Object.keys(st.docs).filter(k => k.startsWith("S1|") && st.docs[k].bo);
        const boMissed = bo.filter(k => !st.docs[k].c && after(k));
        console.log("Check blanket links: " + bo.length + " orders tied to a blanket order · " + bo.filter(k => st.docs[k].c).length + " of them seen being raised · " + boMissed.length + " raised after the log began but not seen being raised · "
            + new Set(bo.map(k => st.docs[k].bo)).size + " different blanket orders · " + bo.filter(k => st.docs["S4|" + st.docs[k].bo]).length + " whose blanket order is in the file");
        // the blanket order list itself, when a service carries it
        try {
            const info = await F.bcDiscoverSalesEntity("blanket");
            if (!info) console.log("Check blanket list: no published service found");
            else {
                let rows = await F.bcFetchAll(F.BC_ODATA_URL + "/Company('" + encodeURIComponent(CO) + "')/" + info.entity, "Blanket order list");
                if (info.fDocType && !info.nameMatchesKind) rows = rows.filter(r => r[info.fDocType] === info.docTypeValue);
                const from = String(st.meta.logFrom || "").slice(0, 10);
                const noF = info.fNo || "No", recent = rows.filter(r => info.fDocDate && String(r[info.fDocDate] || "").slice(0, 10) >= from);
                const inFile = (list) => { let c1 = 0, c0 = 0, none = 0; for (const r of list) { const d = st.docs["S4|" + r[noF]]; if (!d) none++; else if (d.c) c1++; else c0++; } return c1 + " seen being raised, " + c0 + " in the file but not seen being raised, " + none + " not in the file"; };
                console.log("Check blanket list: service " + info.entity + " (date field " + info.fDocDate + ", amount field " + info.fAmount + ") · " + rows.length + " open blanket orders: " + inFile(rows)
                    + " · " + recent.length + " dated since the log began: " + inFile(recent) + " · " + rows.filter(r => num(r[info.fAmount])).length + " carry an amount");
                console.log("  fields: " + (info.fields || []).join(", ").slice(0, 1500));
            }
        } catch (e) { console.warn("Check blanket list: " + String(e.message || e).slice(0, 200)); }
        // what the first lines of a missed document look like
        const sample = boMissed.slice(-3).concat((missed.S4 || []).slice(-3), (missed.S1 || []).filter(k => !st.docs[k].bo).slice(-2), (missed.S0 || []).slice(-2));
        for (const k of sample) {
            const rows = await odata("ChangeLogEntries", "Table_No eq 36 and Primary_Key_Field_2_Value eq '" + k.slice(3).replace(/'/g, "''") + "'", "Entry_No,Date_and_Time,User_ID,Primary_Key_Field_1_Value,Field_Caption,Type_of_Change,Old_Value,New_Value", "Check lines");
            rows.sort((a, b) => a.Entry_No - b.Entry_No);
            const mine = rows.filter(r => String(r.Primary_Key_Field_1_Value || "") === k[1]);
            if (!mine.length) { console.log("  missed " + k.slice(0, 2) + (st.docs[k].bo ? " (from a blanket order)" : "") + ": no lines"); continue; }
            const t0 = new Date(mine[0].Date_and_Time).getTime(), u0 = mine[0].User_ID;
            console.log("  missed " + k.slice(0, 2) + (st.docs[k].bo ? " (from a blanket order)" : "") + ": " + mine.length + " lines · " + mine.slice(0, 14).map(r =>
                "+" + Math.round((new Date(r.Date_and_Time).getTime() - t0) / 1000) + "s " + (r.Type_of_Change || "")[0] + " " + r.Field_Caption + " [" + (String(r.Old_Value || "").trim() ? "set" : "blank") + "→" + (String(r.New_Value || "").trim() ? "set" : "blank") + "]" + (r.User_ID === u0 ? "" : " other user")).join(" | "));
        }
    } catch (e) { console.warn("Check on the blanket figures did not finish: " + String(e.message || e).slice(0, 200)); }

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
    await SP.uploadFile(driveId, META, Buffer.from(JSON.stringify({ v: FORMAT, lock: LOCK, builtAtUtc: st.meta.builtAtUtc, builtAtMelbourne: st.meta.builtAtMelbourne, bytes: bin.length, logFrom: st.meta.logFrom, postedFrom: POSTED_FROM, itemFrom: st.meta.itemFrom || "", itemLines: (st.ic || []).length }, null, 2), "utf8"));
    console.log("Published activity.bin (" + (bin.length / 1048576).toFixed(2) + " MB) in " + ((Date.now() - t0) / 1000).toFixed(0) + "s");
})().catch(e => {
    console.error("ACTIVITY FILE NOT UPDATED:", e.message);
    process.exit(1);
});
