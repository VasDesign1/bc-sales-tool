// ============================================================
// snapshot.js — the "morning snapshot" robot.
// Runs in GitHub Actions (Node 20). Performs the SAME full load the
// tool's Load sales button does — all 16 sources over the last 12
// months — via the shared bc-fetchers.js, then gzips + encrypts the
// result for the Fast lookup menu.
//
// Env (from GitHub secrets):
//   BC_CLIENT_ID, BC_TENANT, BC_REFRESH_TOKEN  — token exchange
//   SNAPSHOT_PASSPHRASE                        — AES key material
//   SLOT                                       — "0500" | "1100" | "1700"
//                                                (empty = fill whichever due
//                                                 slot is missing today)
//
// Output — uploaded to SharePoint (scripts/sp-upload.js → SP_CONFIG in
// bc-fetchers.js); a local copy also lands in ./snapshot-out/ for the
// Actions log:
//   <slot>.bin        salt(16) | iv(12) | AES-256-GCM ciphertext||tag
//                     of gzip(JSON payload) — tag last so browser
//                     WebCrypto can decrypt the ct||tag block directly.
//                     Key = PBKDF2-SHA256(passphrase, salt, 600,000 iters).
//   <slot>.meta.json  { slot, fetchedAtUtc, fetchedAtMelbourne, from,
//                       to, bytes, formatVersion } (plaintext, no data)
// ============================================================
"use strict";
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");
const crypto = require("crypto");

// ---------- Globals contract required by bc-fetchers.js ----------
global.state = {};                    // fetchers read fastTrim + write diagnostics
global.num = (v) => (v == null || v === "") ? 0 : parseFloat(v) || 0;   // keep in sync with index.html
global.updateBCProgress = (label, detail) => console.log("  [" + label + "] " + detail);

const TENANT = process.env.BC_TENANT;
const CLIENT_ID = process.env.BC_CLIENT_ID;
const REFRESH_TOKEN = process.env.BC_REFRESH_TOKEN;
const PASSPHRASE = process.env.SNAPSHOT_PASSPHRASE;
if (!TENANT || !CLIENT_ID || !REFRESH_TOKEN || !PASSPHRASE) {
    console.error("Missing env: need BC_TENANT, BC_CLIENT_ID, BC_REFRESH_TOKEN, SNAPSHOT_PASSPHRASE");
    process.exit(1);
}

let _tok = null, _tokExp = 0;
global.bcGetToken = async function bcGetToken() {
    if (_tok && Date.now() < _tokExp - 300000) return _tok;
    const body = new URLSearchParams({
        grant_type: "refresh_token",
        client_id: CLIENT_ID,
        refresh_token: REFRESH_TOKEN,
        scope: "https://api.businesscentral.dynamics.com/user_impersonation offline_access",
    });
    const resp = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
    });
    const tok = await resp.json();
    if (!tok.access_token) {
        throw new Error("Token exchange failed: " + JSON.stringify(tok).slice(0, 400));
    }
    _tok = tok.access_token;
    _tokExp = Date.now() + (tok.expires_in || 3600) * 1000;
    return _tok;
};

const F = require(path.join(__dirname, "..", "bc-fetchers.js"));
const SP = require(path.join(__dirname, "sp-upload.js"));
const PBKDF2_ITERATIONS = 600000;   // keep in sync with decryptSnapshot() in index.html

// ---------- Melbourne wall clock + slot detection ----------
function melbourneNow() {
    const parts = new Intl.DateTimeFormat("en-AU", {
        timeZone: "Australia/Melbourne",
        year: "numeric", month: "2-digit", day: "2-digit",
        hour: "2-digit", minute: "2-digit", hourCycle: "h23",
    }).formatToParts(new Date());
    const g = (t) => parts.find(p => p.type === t).value;
    return { date: g("year") + "-" + g("month") + "-" + g("day"),
             minutes: parseInt(g("hour"), 10) * 60 + parseInt(g("minute"), 10),
             hhmm: g("hour") + ":" + g("minute") };
}
// ---------- Slots ----------
// Three fixed slots a day on the Melbourne clock. The cron fires every
// 30 minutes; each firing asks "which slot that is already due is still
// missing today's capture?" and fills ONE of them — the current slot
// first, then any earlier slot that missed — or exits in seconds. A slot
// that already holds today's capture is never overwritten by the
// schedule, so a cron GitHub runs hours late can no longer file a capture
// into the wrong bin or leave one sitting on yesterday's data.
//   deadline: once a slot is this late, a failed fill ends the run in
//   failure (→ one email); before that a failure is quiet because the
//   next firing retries anyway.
const SLOTS = [
    { slot: "0500", label: "Morning", due: 5 * 60,  deadline: 10 * 60 + 30 },
    { slot: "1100", label: "Midday",  due: 11 * 60, deadline: 16 * 60 + 30 },
    { slot: "1700", label: "Evening", due: 17 * 60, deadline: 22 * 60 + 30 },
];
// A slot holds today's capture when its meta says so and the capture
// happened after the slot came due (a 04:30 manual fill is not the
// 05:00 capture).
function filledToday(meta, s, mel) {
    if (!meta || !meta.fetchedAtMelbourne) return false;
    const [d, t] = String(meta.fetchedAtMelbourne).split(" ");
    if (d !== mel.date) return false;
    const [h, m] = (t || "0:0").split(":").map(Number);
    return ((h || 0) % 24) * 60 + (m || 0) >= s.due;
}
// Returns { s, reason } or null when there is nothing to do.
async function chooseSlot(mel) {
    const forced = (process.env.SLOT || "").trim();
    if (forced) {
        const s = SLOTS.find(x => x.slot === forced);
        if (!s) throw new Error("Unknown slot '" + forced + "' — use " + SLOTS.map(x => x.slot).join(" / "));
        return { s, reason: "forced by workflow input" };
    }
    const due = SLOTS.filter(s => s.due <= mel.minutes);
    if (!due.length) { console.log("Melbourne " + mel.hhmm + " — before the first slot of the day, nothing to do."); return null; }
    const current = due[due.length - 1];
    // A manual "Run workflow" always refreshes the most recent slot.
    if (process.env.GITHUB_EVENT_NAME === "workflow_dispatch") return { s: current, reason: "manual refresh" };
    let metas;
    try { metas = await SP.readSnapshotMetas(due.map(s => s.slot)); }
    catch (e) {
        console.warn("Cannot read slot state from SharePoint (" + e.message + ") — filling the current slot anyway");
        return { s: current, reason: "slot state unknown" };
    }
    for (const s of due) {
        const m = metas[s.slot];
        console.log("  " + s.label + " (" + s.slot + "): " + (m && m.fetchedAtMelbourne ? "captured " + m.fetchedAtMelbourne : "no snapshot")
            + (filledToday(m, s, mel) ? " ✓ today" : " — missing today"));
    }
    for (let i = due.length - 1; i >= 0; i--) {
        if (!filledToday(metas[due[i].slot], due[i], mel)) return { s: due[i], reason: i === due.length - 1 ? "current slot" : "catch-up" };
    }
    console.log("Melbourne " + mel.hhmm + " — every due slot already holds today's capture, nothing to do.");
    return null;
}

// ---------- ISO date helpers (Melbourne-anchored) ----------
function isoAddDays(iso, days) {
    const d = new Date(iso + "T00:00:00Z");
    d.setUTCDate(d.getUTCDate() + days);
    return d.toISOString().slice(0, 10);
}

let chosen = null;   // visible to the failure handler below
(async () => {
    const mel = melbourneNow();
    chosen = await chooseSlot(mel);
    if (!chosen) return;
    const slot = chosen.s.slot;
    console.log("Filling " + chosen.s.label + " slot (" + slot + ") — " + chosen.reason);

    const to = mel.date;
    const from = isoAddDays(to, -365);
    console.log("Snapshot slot " + slot + " · Melbourne " + mel.date + " " + mel.hhmm + " · range " + from + " → " + to);

    const t0 = Date.now();
    // Same fetches, same order, as index.html handleLoad(). glRecon is the
    // revenue / cost-of-goods-sold ledger slice the Overview's two
    // reconciliation rows need, carried so Fast lookup has them without a
    // second round trip to Business Central.
    const [locs, items, customers, invoices, iles, creditMemos, shipments, returnReceipts,
           quotes, quoteExtras, blanketOrders, valueEntries, salesOrders, salesOrderOutstanding,
           _residentialLookup, _quoteArchive, glRecon, glAccounts] = await Promise.all([
        F.fetchLocations(),
        F.fetchItems(),
        F.fetchCustomers(),
        F.fetchSalesInvoicesWithLines(from, to),
        F.fetchItemLedgerSales(from, to),
        F.fetchSalesCreditMemos(from, to),
        F.fetchSalesShipments(from, to),
        F.fetchSalesReturnReceipts(from, to),
        F.fetchSalesQuotes(from, to),
        F.fetchSalesQuoteExtras(from, to),
        F.fetchBlanketSalesOrders(from, to),
        F.fetchValueEntries(from, to),
        F.fetchSalesOrders(),
        F.fetchSalesOrderOutstandingLines(),
        F.fetchResidentialDocLookup(),
        F.fetchSalesQuoteArchive(from, to),
        F.fetchGLReconSlice(from, to).catch(e => {
            // Never fail a whole snapshot over the ledger slice: the rows
            // that use it fall back to fetching on demand.
            console.warn("  [Ledger] reconciliation slice failed: " + e.message);
            return [];
        }),
        // Account names travel with the entries: the cost row tells an
        // expected-cost clearing account apart by its name.
        F.fetchGLReconAccountNames().catch(e => {
            console.warn("  [Ledger] account names failed: " + e.message);
            return [];
        }),
    ]);
    console.log("Fetched in " + ((Date.now() - t0) / 1000).toFixed(1) + "s: "
        + invoices.length + " invoices · " + (valueEntries || []).length + " VE · "
        + (iles || []).length + " ILE · " + (quotes || []).length + " quotes · "
        + (glRecon || []).length + " ledger rows (" + (glRecon || []).filter(r => Array.isArray(r[7])).length + " cost rows linked) · "
        + (glAccounts || []).length + " accounts");

    // ---- Integrity checks (fail loudly rather than snapshot bad data) ----
    // 1. VE date filter actually applied? (this tenant has form — $select
    //    provably disabled it once already)
    let veMin = "9999", veMax = "0000", veOut = 0;
    for (const v of (valueEntries || [])) {
        const d = String(v.postingDate || "").slice(0, 10);
        if (d < veMin) veMin = d;
        if (d > veMax) veMax = d;
        if (d < from || d > to) veOut++;
    }
    console.log("[VE check] postingDate " + veMin + " … " + veMax + " · outOfRange=" + veOut);
    // 2. Duplicate rows? Entry_No is unique in BC — any repeat means a
    //    pagination/duplication fault and the cost maps would double-count.
    const entryNos = (valueEntries || []).map(v => v.entryNo).filter(n => n != null);
    const distinct = new Set(entryNos).size;
    console.log("[VE check] entryNo present on " + entryNos.length + "/" + (valueEntries || []).length
        + " rows · distinct=" + distinct);
    if (entryNos.length && distinct !== entryNos.length) {
        throw new Error("[VE check] DUPLICATE Value Entry rows detected — aborting snapshot");
    }
    if (veOut > 0) console.warn("[VE check] " + veOut + " rows outside range — server filter ignored?! Snapshot keeps them (aggregation re-filters by date) but investigate.");
    // 3. Truncation? If the newest value entry is more than 3 days older than
    //    the end of the window (capped at today), the fetch stopped early —
    //    a $top-style cap or a broken nextLink. Cost for the latest days
    //    would be silently missing, so fail loudly rather than publish.
    const todayIso = new Date().toISOString().slice(0, 10);
    const expectEnd = to < todayIso ? to : todayIso;
    const staleDays = Math.round((new Date(expectEnd + "T00:00:00") - new Date(veMax + "T00:00:00")) / 86400000);
    if ((valueEntries || []).length && staleDays > 3) {
        throw new Error("[VE check] newest value entry is " + veMax + " but the window runs to " + expectEnd + " (" + staleDays + " days short) — fetch truncated? aborting snapshot");
    }

    const payload = {
        meta: {
            formatVersion: 1,
            slot,
            fetchedAtUtc: new Date().toISOString(),
            fetchedAtMelbourne: mel.date + " " + mel.hhmm,
            from, to,
        },
        data: { locs, items, customers, invoices, iles, creditMemos, shipments, returnReceipts,
                quotes, quoteExtras, blanketOrders, valueEntries, salesOrders, salesOrderOutstanding,
                glRecon, glAccounts, glReconCols: F.GL_RECON_COLS },
        // Side effects the discovery-style fetchers write into `state`,
        // which handleLoad doesn't receive via return values. Maps are
        // serialised as entry arrays.
        sideEffects: {
            docYourReference: [...(state.docYourReference || new Map())],
            docQuoteLink: [...(state.docQuoteLink || new Map())],
            quoteArchive: [...(state.quoteArchive || new Map())],
            veFieldMap: state.veFieldMap || null,
            quoteExtrasFieldMap: state.quoteExtrasFieldMap || null,
            veDiagnostic: state.veDiagnostic || "",
            quoteExtrasDiagnostic: state.quoteExtrasDiagnostic || "",
            quoteArchiveDiagnostic: state.quoteArchiveDiagnostic || "",
            residentialDiagnostic: state.residentialDiagnostic || "",
            blanketOrdersDiagnostic: state.blanketOrdersDiagnostic || "",
        },
    };

    const json = Buffer.from(JSON.stringify(payload), "utf8");
    const gz = zlib.gzipSync(json, { level: 9 });
    const salt = crypto.randomBytes(16);
    const iv = crypto.randomBytes(12);
    const key = crypto.pbkdf2Sync(PASSPHRASE, salt, PBKDF2_ITERATIONS, 32, "sha256");
    const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
    const ct = Buffer.concat([cipher.update(gz), cipher.final(), cipher.getAuthTag()]);
    const bin = Buffer.concat([salt, iv, ct]);

    const outDir = path.join(__dirname, "..", "snapshot-out");
    fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(path.join(outDir, slot + ".bin"), bin);
    const metaJson = JSON.stringify({
        slot,
        fetchedAtUtc: payload.meta.fetchedAtUtc,
        fetchedAtMelbourne: payload.meta.fetchedAtMelbourne,
        from, to,
        bytes: bin.length,
        formatVersion: 1,
    }, null, 2);
    fs.writeFileSync(path.join(outDir, slot + ".meta.json"), metaJson);
    console.log("Wrote " + slot + ".bin (" + (bin.length / 1048576).toFixed(2) + " MB, "
        + (json.length / 1048576).toFixed(1) + " MB raw JSON)");

    // The reconciliation slice on its own, encrypted the same way, so Load
    // sales can take it without downloading the whole snapshot.
    const glPayload = { meta: { slot, fetchedAtUtc: payload.meta.fetchedAtUtc, fetchedAtMelbourne: payload.meta.fetchedAtMelbourne, from, to },
                        glRecon, glAccounts, glReconCols: F.GL_RECON_COLS };
    const glJson = Buffer.from(JSON.stringify(glPayload), "utf8");
    const glGz = zlib.gzipSync(glJson, { level: 9 });
    const glSalt = crypto.randomBytes(16), glIv = crypto.randomBytes(12);
    const glKey = crypto.pbkdf2Sync(PASSPHRASE, glSalt, PBKDF2_ITERATIONS, 32, "sha256");
    const glCipher = crypto.createCipheriv("aes-256-gcm", glKey, glIv);
    const glBin = Buffer.concat([glSalt, glIv, glCipher.update(glGz), glCipher.final(), glCipher.getAuthTag()]);
    fs.writeFileSync(path.join(outDir, slot + ".gl.bin"), glBin);
    console.log("Wrote " + slot + ".gl.bin (" + (glBin.length / 1048576).toFixed(2) + " MB, " + (glJson.length / 1048576).toFixed(1) + " MB raw JSON)");

    console.log("Publishing to SharePoint…");
    await SP.publishSnapshot(slot, bin, metaJson, glBin);
    console.log("Published " + slot + " to SharePoint");
})().catch(e => {
    const mel = melbourneNow();
    const manual = !!(process.env.SLOT || "").trim() || process.env.GITHUB_EVENT_NAME === "workflow_dispatch";
    const late = chosen && mel.minutes >= chosen.s.deadline;
    if (!manual && chosen && !late) {
        // Quiet: the schedule fires again in 30 minutes and will retry this
        // slot. Only a slot that is past its deadline fails the run (→ email).
        console.warn("SNAPSHOT NOT PUBLISHED (" + chosen.s.label + "): " + e.message);
        console.warn("Next firing retries; the run ends in failure only after " + String(Math.floor(chosen.s.deadline / 60)).padStart(2, "0") + ":" + String(chosen.s.deadline % 60).padStart(2, "0") + " Melbourne.");
        return;
    }
    console.error("SNAPSHOT FAILED" + (chosen ? " (" + chosen.s.label + ")" : "") + ":", e.message);
    process.exit(1);
});
