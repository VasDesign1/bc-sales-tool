// ============================================================
// decrypt-snapshot.js — open a snapshot (or nav-data.bin) outside the tool.
//
//   node scripts/decrypt-snapshot.js <file.bin> [out.json]
//
// Prompts for the passphrase (not echoed), writes the decrypted JSON next
// to the .bin (or to out.json) and prints a short summary. Works for both
// the sales-tool and forecast-tool files: layout is
//   salt(16) | iv(12) | AES-256-GCM ciphertext || tag(16)   of gzip(JSON)
// with key = PBKDF2-SHA256(passphrase, salt, 600,000 iterations).
// ============================================================
"use strict";
const fs = require("fs"), path = require("path"), crypto = require("crypto"), zlib = require("zlib"), readline = require("readline");

const src = process.argv[2];
if (!src) { console.error("usage: node scripts/decrypt-snapshot.js <file.bin> [out.json]"); process.exit(1); }
const out = process.argv[3] || src.replace(/\.bin$/i, "") + ".json";

function askHidden(q) {
    return new Promise(resolve => {
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true });
        const w = rl._writeToOutput;
        rl.question(q, a => { rl._writeToOutput = w; rl.close(); process.stdout.write("\n"); resolve(a.trim()); });
        rl._writeToOutput = () => {};   // hide typed characters
    });
}

(async () => {
    const bin = fs.readFileSync(src);
    const pass = process.env.SNAPSHOT_PASSPHRASE || await askHidden("Passphrase: ");
    const salt = bin.subarray(0, 16), iv = bin.subarray(16, 28), ct = bin.subarray(28, bin.length - 16), tag = bin.subarray(bin.length - 16);
    let gz;
    for (const iters of [600000, 150000]) {           // 150k = pre-2026-09-22 files
        const key = crypto.pbkdf2Sync(pass, salt, iters, 32, "sha256");
        const d = crypto.createDecipheriv("aes-256-gcm", key, iv); d.setAuthTag(tag);
        try { gz = Buffer.concat([d.update(ct), d.final()]); break; } catch (e) { /* wrong key or iteration count */ }
    }
    if (!gz) { console.error("Could not unlock " + path.basename(src) + " — wrong passphrase?"); process.exit(2); }
    const json = zlib.gunzipSync(gz);
    fs.writeFileSync(out, json);
    const p = JSON.parse(json.toString("utf8"));
    const summary = p.data
        ? Object.entries(p.data).map(([k, v]) => k + "=" + (Array.isArray(v) ? v.length : (v && v.rows ? v.rows.length : "?"))).join(", ")
        : "rows=" + (p.rows || []).length;
    console.log("Decrypted " + path.basename(src) + " → " + out + " (" + (json.length / 1048576).toFixed(1) + " MB)");
    console.log("meta: " + JSON.stringify(p.meta || {}));
    console.log("data: " + summary);
})();
