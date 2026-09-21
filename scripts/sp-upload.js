// ============================================================
// sp-upload.js — SharePoint publisher for the snapshot robot.
//
// Snapshots no longer live on a public GitHub branch: the robot writes
// them into a folder in the Vic Air SharePoint tenant through Microsoft
// Graph, and the tool reads them back with the signed-in user's own
// token (index.html → spFetch). Anyone outside the tenant, or without
// access to the folder, gets 403 from Microsoft before the (still
// AES-256-GCM encrypted) bytes are ever served.
//
// Auth: the same delegated refresh token the BC fetches use. Entra v2
// refresh tokens are multi-resource, so one token exchange with a Graph
// scope is enough — the app registration ("BC Sales Snapshot Bot") has
// Files.ReadWrite.All delegated + admin consent.
//
// Env: BC_TENANT, BC_CLIENT_ID, BC_REFRESH_TOKEN (same as snapshot.js)
//
// Self-test (workflow_dispatch → selftest=true, or `node sp-upload.js --selftest`):
//   uploads a 40-byte marker to the folder, reads it back, deletes it.
// ============================================================
"use strict";
const path = require("path");
const { SP_CONFIG } = require(path.join(__dirname, "..", "bc-fetchers.js"));

const GRAPH = "https://graph.microsoft.com/v1.0";
const SMALL_UPLOAD_LIMIT = 4 * 1024 * 1024;        // Graph: simple PUT ≤ 4 MB
const CHUNK = 10 * 1024 * 1024;                     // upload-session chunks must be multiples of 320 KiB — 10 MiB = 32 × 320 KiB

const TENANT = process.env.BC_TENANT;
const CLIENT_ID = process.env.BC_CLIENT_ID;
const REFRESH_TOKEN = process.env.BC_REFRESH_TOKEN;

let _tok = null, _tokExp = 0;
async function graphToken() {
    if (_tok && Date.now() < _tokExp - 300000) return _tok;
    const body = new URLSearchParams({
        grant_type: "refresh_token",
        client_id: CLIENT_ID,
        refresh_token: REFRESH_TOKEN,
        scope: "https://graph.microsoft.com/Files.ReadWrite.All offline_access",
    });
    const resp = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
    });
    const tok = await resp.json();
    if (!tok.access_token) throw new Error("Graph token exchange failed: " + JSON.stringify(tok).slice(0, 400));
    _tok = tok.access_token;
    _tokExp = Date.now() + (tok.expires_in || 3600) * 1000;
    return _tok;
}

async function graphFetch(url, opts, attempt) {
    attempt = attempt || 0;
    const token = await graphToken();
    const headers = Object.assign({ Authorization: "Bearer " + token }, (opts && opts.headers) || {});
    const resp = await fetch(url, Object.assign({}, opts, { headers }));
    if ((resp.status === 429 || resp.status >= 500) && attempt < 4) {
        const wait = parseInt(resp.headers.get("Retry-After") || "0", 10) * 1000 || (2000 * (attempt + 1));
        console.log("  [graph] HTTP " + resp.status + " — retrying in " + wait + " ms");
        await new Promise(r => setTimeout(r, wait));
        return graphFetch(url, opts, attempt + 1);
    }
    return resp;
}

async function graphJson(url, opts) {
    const resp = await graphFetch(url, opts);
    const text = await resp.text();
    if (!resp.ok) throw new Error("Graph " + (opts && opts.method || "GET") + " " + url.replace(GRAPH, "") + " → HTTP " + resp.status + ": " + text.slice(0, 300));
    return text ? JSON.parse(text) : {};
}

// Each path segment percent-encoded separately so "/" stays a separator
// while spaces and other characters in folder names survive.
function encPath(p) {
    return p.split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

// The site's default document library ("Documents" / SHARED DOCUMENTS).
async function resolveDrive(host) {
    const site = await graphJson(GRAPH + "/sites/" + host + ":/?$select=id,webUrl");
    const drive = await graphJson(GRAPH + "/sites/" + site.id + "/drive?$select=id,name,webUrl");
    console.log("  [graph] site " + site.webUrl + " · drive '" + drive.name + "' " + drive.id.slice(0, 12) + "…");
    return drive.id;
}

function itemUrl(driveId, itemPath) {
    return GRAPH + "/drives/" + driveId + "/root:/" + encPath(itemPath);
}

async function uploadSmall(driveId, itemPath, buf) {
    return graphJson(itemUrl(driveId, itemPath) + ":/content?@microsoft.graph.conflictBehavior=replace", {
        method: "PUT",
        headers: { "Content-Type": "application/octet-stream" },
        body: buf,
    });
}

async function uploadLarge(driveId, itemPath, buf) {
    const session = await graphJson(itemUrl(driveId, itemPath) + ":/createUploadSession", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }),
    });
    const total = buf.length;
    let item = null;
    for (let start = 0; start < total; start += CHUNK) {
        const end = Math.min(start + CHUNK, total);
        const chunk = buf.subarray(start, end);
        let resp;
        for (let attempt = 0; ; attempt++) {
            // Upload-session URLs are pre-authenticated: no Authorization header.
            resp = await fetch(session.uploadUrl, {
                method: "PUT",
                headers: {
                    "Content-Length": String(chunk.length),
                    "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + total,
                },
                body: chunk,
            });
            if (resp.ok || attempt >= 4) break;
            const wait = 2000 * (attempt + 1);
            console.log("  [graph] chunk " + start + "-" + (end - 1) + " HTTP " + resp.status + " — retrying in " + wait + " ms");
            await new Promise(r => setTimeout(r, wait));
        }
        const text = await resp.text();
        if (!resp.ok) throw new Error("Upload chunk " + start + "-" + (end - 1) + " → HTTP " + resp.status + ": " + text.slice(0, 300));
        if (resp.status === 200 || resp.status === 201) item = JSON.parse(text);   // final chunk returns the driveItem
    }
    return item;
}

async function uploadFile(driveId, itemPath, buf) {
    const t0 = Date.now();
    const item = buf.length <= SMALL_UPLOAD_LIMIT
        ? await uploadSmall(driveId, itemPath, buf)
        : await uploadLarge(driveId, itemPath, buf);
    console.log("  [graph] uploaded " + itemPath.split("/").pop() + " (" + (buf.length / 1048576).toFixed(2) + " MB) in "
        + ((Date.now() - t0) / 1000).toFixed(1) + "s → " + (item && item.size) + " bytes on SharePoint");
    if (item && item.size !== buf.length) throw new Error("Size mismatch after upload: sent " + buf.length + ", SharePoint has " + item.size);
    return item;
}

async function downloadFile(driveId, itemPath) {
    const resp = await graphFetch(itemUrl(driveId, itemPath) + ":/content");
    if (!resp.ok) throw new Error("Download " + itemPath + " → HTTP " + resp.status);
    return Buffer.from(await resp.arrayBuffer());
}

async function deleteFile(driveId, itemPath) {
    const resp = await graphFetch(itemUrl(driveId, itemPath), { method: "DELETE" });
    if (!resp.ok && resp.status !== 404) throw new Error("Delete " + itemPath + " → HTTP " + resp.status);
}

// Publish a finished snapshot: .bin first, .meta.json last so the menu
// never advertises a slot whose bytes are still uploading.
async function publishSnapshot(slot, bin, metaJson) {
    const driveId = await resolveDrive(SP_CONFIG.host);
    await uploadFile(driveId, SP_CONFIG.folder + "/" + slot + ".bin", bin);
    await uploadFile(driveId, SP_CONFIG.folder + "/" + slot + ".meta.json", Buffer.from(metaJson, "utf8"));
}

async function selftest() {
    if (!TENANT || !CLIENT_ID || !REFRESH_TOKEN) throw new Error("Missing env: BC_TENANT, BC_CLIENT_ID, BC_REFRESH_TOKEN");
    console.log("SharePoint self-test → " + SP_CONFIG.host + " / " + SP_CONFIG.folder);
    const driveId = await resolveDrive(SP_CONFIG.host);
    const name = SP_CONFIG.folder + "/_selftest.txt";
    const marker = Buffer.from("selftest " + new Date().toISOString() + "\n", "utf8");
    await uploadFile(driveId, name, marker);
    const back = await downloadFile(driveId, name);
    if (!back.equals(marker)) throw new Error("Read-back mismatch");
    console.log("  [graph] read back OK (" + back.length + " bytes)");
    await deleteFile(driveId, name);
    console.log("  [graph] deleted marker");
    console.log("SELF-TEST PASSED");
}

module.exports = { resolveDrive, uploadFile, downloadFile, deleteFile, publishSnapshot, selftest };

if (require.main === module && process.argv.includes("--selftest")) {
    selftest().catch(e => { console.error("SELF-TEST FAILED:", e.message); process.exit(1); });
}
