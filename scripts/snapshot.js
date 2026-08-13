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
//   SLOT                                       — "0700" | "1200" | "1630"
//                                                (empty = auto-detect from
//                                                 Melbourne wall clock)
//
// Output (./snapshot-out/):
//   <slot>.bin        salt(16) | iv(12) | AES-256-GCM ciphertext||tag
//                     of gzip(JSON payload) — tag last so browser
//                     WebCrypto can decrypt the ct||tag block directly
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

// ---------- Melbourne wall clock + slot detection ----------
function melbourneNow() {
    const parts = new Intl.DateTimeFormat("en-AU", {
        timeZone: "Australia/Melbourne",
        year: "numeric", month: "2-digit", day: "2-digit",
        hour: "2-digit", minute: "2-digit", hour12: false,
    }).formatToParts(new Date());
    const g = (t) => parts.find(p => p.type === t).value;
    return { date: g("year") + "-" + g("month") + "-" + g("day"),
             minutes: parseInt(g("hour"), 10) * 60 + parseInt(g("minute"), 10),
             hhmm: g("hour") + ":" + g("minute") };
}
const SLOTS = { "0700": 7 * 60, "1200": 12 * 60, "1630": 16 * 60 + 30 };
function detectSlot(mel) {
    let best = null, bestDiff = 1e9;
    for (const [slot, mins] of Object.entries(SLOTS)) {
        const d = Math.abs(mel.minutes - mins);
        if (d < bestDiff) { bestDiff = d; best = slot; }
    }
    return bestDiff <= 45 ? best : null;   // cron drift tolerance; DST-shifted crons fall outside and skip
}

// ---------- ISO date helpers (Melbourne-anchored) ----------
function isoAddDays(iso, days) {
    const d = new Date(iso + "T00:00:00Z");
    d.setUTCDate(d.getUTCDate() + days);
    return d.toISOString().slice(0, 10);
}

(async () => {
    const mel = melbourneNow();
    let slot = (process.env.SLOT || "").trim();
    if (!slot) {
        slot = detectSlot(mel);
        if (!slot) {
            console.log("Melbourne time " + mel.hhmm + " is not near any slot — DST-offset cron firing, skipping.");
            return;
        }
    }
    if (!SLOTS[slot]) { console.error("Unknown slot '" + slot + "'"); process.exit(1); }

    const to = mel.date;
    const from = isoAddDays(to, -365);
    console.log("Snapshot slot " + slot + " · Melbourne " + mel.date + " " + mel.hhmm + " · range " + from + " → " + to);

    const t0 = Date.now();
    // Same sixteen fetches, same order, as index.html handleLoad().
    const [locs, items, customers, invoices, iles, creditMemos, shipments, returnReceipts,
           quotes, quoteExtras, blanketOrders, valueEntries, salesOrders, salesOrderOutstanding] = await Promise.all([
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
    ]);
    console.log("Fetched in " + ((Date.now() - t0) / 1000).toFixed(1) + "s: "
        + invoices.length + " invoices · " + (valueEntries || []).length + " VE · "
        + (iles || []).length + " ILE · " + (quotes || []).length + " quotes");

    const payload = {
        meta: {
            formatVersion: 1,
            slot,
            fetchedAtUtc: new Date().toISOString(),
            fetchedAtMelbourne: mel.date + " " + mel.hhmm,
            from, to,
        },
        data: { locs, items, customers, invoices, iles, creditMemos, shipments, returnReceipts,
                quotes, quoteExtras, blanketOrders, valueEntries, salesOrders, salesOrderOutstanding },
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
    const key = crypto.pbkdf2Sync(PASSPHRASE, salt, 150000, 32, "sha256");
    const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
    const ct = Buffer.concat([cipher.update(gz), cipher.final(), cipher.getAuthTag()]);
    const bin = Buffer.concat([salt, iv, ct]);

    const outDir = path.join(__dirname, "..", "snapshot-out");
    fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(path.join(outDir, slot + ".bin"), bin);
    fs.writeFileSync(path.join(outDir, slot + ".meta.json"), JSON.stringify({
        slot,
        fetchedAtUtc: payload.meta.fetchedAtUtc,
        fetchedAtMelbourne: payload.meta.fetchedAtMelbourne,
        from, to,
        bytes: bin.length,
        formatVersion: 1,
    }, null, 2));
    console.log("Wrote " + slot + ".bin (" + (bin.length / 1048576).toFixed(2) + " MB, "
        + (json.length / 1048576).toFixed(1) + " MB raw JSON)");
})().catch(e => {
    console.error("SNAPSHOT FAILED:", e.message);
    process.exit(1);
});
