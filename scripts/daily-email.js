// ============================================================
// daily-email.js — emails the Daily Snapshot tab (Faik, 2026-11-10).
//
// The figures are NOT recalculated here. The robot opens the tool's own
// index.html in headless Chrome, pours the newest snapshot into it the
// same way Fast lookup does, reads the targets workbooks with the page's
// own parser, and asks the page for dsSnapshotFigures() — the exact
// object the Daily Snapshot tab paints. The email is that object laid
// out as an HTML message, so tab and email can never disagree.
//
// Runs in GitHub Actions on Node 22 (needs the built-in WebSocket).
// Env:
//   BC_TENANT, BC_CLIENT_ID, BC_REFRESH_TOKEN   token exchange (Graph:
//                                               Files.ReadWrite.All + Mail.Send, delegated)
//   SNAPSHOT_PASSPHRASE                         opens the snapshot
//   DAILY_EMAIL_TO                              recipients, comma separated
//   DRY_RUN=1                                   build everything, send nothing
//   CHROME_PATH                                 optional, default google-chrome
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

// ---------- collect the inputs into a folder the page can read ----------
async function collectFromSharePoint(dir) {
    const site = await graphJson("/sites/" + SP_CONFIG.host + ":/?$select=id");
    const drive = (await graphJson("/sites/" + site.id + "/drive?$select=id")).id;
    const item = p => "/drives/" + drive + "/root:/" + enc(p);
    // newest snapshot of the three slots (a forced SLOT wins)
    let best = null;
    for (const slot of SLOTS) {
        const r = await graph(item(SP_CONFIG.folder + "/snapshots/" + slot + ".meta.json") + ":/content");
        if (!r.ok) continue;
        const meta = JSON.parse(await r.text());
        if ((process.env.SLOT || "") === slot) { best = meta; break; }
        if (!process.env.SLOT && (!best || String(meta.fetchedAtMelbourne) > String(best.fetchedAtMelbourne))) best = meta;
    }
    if (!best) throw new Error("no snapshot found on SharePoint");
    console.log("Snapshot: slot " + best.slot + " captured " + best.fetchedAtMelbourne + " Melbourne");
    const r = await graph(item(SP_CONFIG.folder + "/snapshots/" + best.slot + ".bin") + ":/content");
    if (!r.ok) throw new Error("snapshot download → HTTP " + r.status);
    const bin = Buffer.from(await r.arrayBuffer());
    const salt = bin.subarray(0, 16), iv = bin.subarray(16, 28), ct = bin.subarray(28, bin.length - 16), tag = bin.subarray(bin.length - 16);
    const key = crypto.pbkdf2Sync(process.env.SNAPSHOT_PASSPHRASE, salt, 600000, 32, "sha256");   // keep in sync with snapshot.js
    const d = crypto.createDecipheriv("aes-256-gcm", key, iv); d.setAuthTag(tag);
    const json = zlib.gunzipSync(Buffer.concat([d.update(ct), d.final()]));
    fs.writeFileSync(path.join(dir, "payload.json"), json);
    console.log("  payload " + (json.length / 1048576).toFixed(1) + " MB");
    // targets workbooks, oldest first (a newer file wins where months overlap — as in the tool)
    const list = await graphJson(item(SP_CONFIG.folder + "/" + TARGETS_FOLDER) + ":/children?$select=name,lastModifiedDateTime,file&$top=200");
    const files = (list.value || []).filter(f => f.file && /\.xls[xm]?$/i.test(f.name) && !/^~\$/.test(f.name))
        .sort((a, b) => (a.lastModifiedDateTime || "").localeCompare(b.lastModifiedDateTime || ""));
    const out = [];
    for (let i = 0; i < files.length; i++) {
        const fr = await graph(item(SP_CONFIG.folder + "/" + TARGETS_FOLDER + "/" + files[i].name) + ":/content");
        if (!fr.ok) { console.warn("  targets file skipped (HTTP " + fr.status + ")"); continue; }
        fs.writeFileSync(path.join(dir, "t" + i + ".xlsx"), Buffer.from(await fr.arrayBuffer()));
        out.push({ file: "t" + i + ".xlsx", name: files[i].name, modified: files[i].lastModifiedDateTime || "" });
    }
    console.log("  targets workbooks: " + out.length);
    fs.writeFileSync(path.join(dir, "targets.json"), JSON.stringify({ files: out }));
}

// ---------- ask the tool's own page for the figures ----------
const PAGE_SCRIPT = `(async () => {
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
    const f = dsSnapshotFigures();
    return JSON.stringify({ f, names: { branch: spBudgetNames("branch"), salesperson: spBudgetNames("salesperson") },
        asOf: state.dataAsOf || "", invoices: state.invoices.length, targetFiles: (spTargets.files || []).length, skipped: spTargets.skipped || [],
        today: localISO(new Date()) });
})()`;
async function figuresFromPage(dataDir) {
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
        "--user-data-dir=" + path.join(dataDir, "chrome"), "--window-size=1600,1200", "about:blank"], { stdio: "ignore" });
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
        await send("Runtime.enable"); await send("Page.enable");
        // The page works on the local clock: make that Melbourne, wherever the robot runs.
        await send("Emulation.setTimezoneOverride", { timezoneId: "Australia/Melbourne" });
        await send("Page.navigate", { url: "http://127.0.0.1:" + port + "/index.html" });
        for (let i = 0; i < 120; i++) {
            const r = await send("Runtime.evaluate", { expression: "typeof dsSnapshotFigures === 'function' && typeof hydrateSnapshot === 'function' && typeof XLSX !== 'undefined' && typeof Chart !== 'undefined'", returnByValue: true });
            if (r.result && r.result.result && r.result.result.value === true) break;
            if (i === 119) throw new Error("the tool page did not finish loading in headless Chrome");
            await sleep(500);
        }
        const r = await send("Runtime.evaluate", { expression: PAGE_SCRIPT, returnByValue: true, awaitPromise: true });
        if (!r.result || !r.result.result || typeof r.result.result.value !== "string") throw new Error("the page could not produce the figures: " + JSON.stringify(r.result || r).slice(0, 600) + " | page errors: " + pageErrors.slice(0, 5).join(" || "));
        if (pageErrors.length) console.warn("  page messages: " + pageErrors.slice(0, 5).join(" || "));
        ws.close();
        return JSON.parse(r.result.result.value);
    } finally { chrome.kill(); srv.close(); }
}

// ---------- the email ----------
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const money = n => (n == null || isNaN(n)) ? "—" : (n < 0 ? "−" : "") + "$" + Math.abs(n).toLocaleString("en-AU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const pct = n => (n == null || isNaN(n)) ? "—" : (n * 100).toFixed(2) + "%";
const INK = "#1f2430", MUTED = "#6b7280", LINE = "#e5e7eb", FONT = "font-family:Segoe UI,Arial,Helvetica,sans-serif;";
const BAND = ["#e53935", "#fb8c00", "#43a047"];
const band = p => (p >= 1 ? 2 : p >= 0.5 ? 1 : 0);
function statusCell(s) {
    const good = /ON TRACK|MET$/.test(s) && !/MISSED/.test(s), none = /NO PLAN|NOT STARTED/.test(s);
    const bg = none ? "#eceff1" : good ? "#c8e6c9" : "#ffcdd2", fg = none ? "#546e7a" : good ? "#1b5e20" : "#b71c1c";
    return '<td style="padding:7px 6px;border-bottom:1px solid ' + LINE + ';"><span style="' + FONT + 'font-size:9.5px;font-weight:700;background:' + bg + ';color:' + fg + ';padding:3px 6px;white-space:nowrap;">' + esc(s === "NO PLAN" ? "No Plan" : s) + "</span></td>";
}
function runRateTable(title, firstCol, data, names, sd) {
    const th = (t, right) => '<th style="' + FONT + 'font-size:9.5px;letter-spacing:.3px;text-transform:uppercase;color:' + MUTED + ';text-align:' + (right ? "right" : "left") + ';padding:8px 6px;border-bottom:2px solid ' + LINE + ';white-space:nowrap;">' + t + "</th>";
    const td = (t, o) => '<td style="' + FONT + "font-size:12px;color:" + ((o && o.color) || INK) + ";text-align:" + (o && o.left ? "left" : "right") + ";padding:7px 6px;border-bottom:1px solid " + LINE + ";white-space:nowrap;" + (o && o.bold ? "font-weight:700;" : "") + '">' + t + "</td>";
    const row = (r, total) => {
        const o = { bold: total };
        const name = !total && names[r.entity] ? ' <span style="font-size:11px;color:' + MUTED + ';">' + esc(names[r.entity]) + "</span>" : "";
        const bal = !r.target ? td("—", o) : td((r.balance >= 0 ? "+" : "−") + money(Math.abs(r.balance)), { bold: total, color: r.balance >= 0 ? "#1b5e20" : "#b71c1c" });
        const ach = r.pctAchieved == null ? td("—", o)
            : '<td style="padding:7px 6px;border-bottom:1px solid ' + LINE + ';text-align:right;white-space:nowrap;"><span style="' + FONT + "font-size:12.5px;font-weight:700;color:" + BAND[band(r.pctAchieved)] + ';">' + pct(r.pctAchieved) + "</span></td>";
        return '<tr' + (total ? ' style="background:#f3f4f6;"' : "") + ">" + td(esc(r.entity) + name, { left: true, bold: total })
            + td(r.pyVal ? money(r.pyVal) : "—", o) + td(r.target ? money(r.target) : "—", o) + td(r.cyVal || total ? money(r.cyVal) : "—", o)
            + td(r.target ? money(r.targetDaily) : "—", o) + td(r.actualDaily ? money(r.actualDaily) : "—", o) + bal
            + td(!r.target ? "—" : !sd.remaining ? "n/a" : money(r.requiredDaily), o) + statusCell(r.status) + ach + "</tr>";
    };
    return '<p style="' + FONT + 'font-size:15px;font-weight:700;color:' + INK + ';margin:26px 0 8px;">' + esc(title) + "</p>"
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">'
        + "<tr>" + th(firstCol) + th("Sales For Month PY", 1) + th("Sales Target", 1) + th("Sales For Month", 1) + th("Target Run Rate", 1) + th("Actual Run Rate", 1) + th("Balance", 1) + th("Required Daily", 1) + th("Status") + th("% Achieved", 1) + "</tr>"
        + data.rows.map(r => row(r, false)).join("") + (data.rows.length ? row(data.total, true) : "") + "</table>";
}
function renderEmail(res) {
    const f = res.f, d = new Date(res.today + "T12:00:00");
    const dateLbl = d.toLocaleDateString("en-AU", { weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "UTC" });
    const asOf = res.asOf ? res.asOf.slice(11) + " on " + new Date(res.asOf.slice(0, 10) + "T12:00:00").toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric", timeZone: "UTC" }) : "";
    const tot = f.branch.total;
    const subject = "Daily Snapshot · " + d.toLocaleDateString("en-AU", { weekday: "short", day: "numeric", month: "short", timeZone: "UTC" })
        + " · " + f.monthLabel + " sales " + money(Math.round(f.sales)).replace(/\.00$/, "") + (tot.pctAchieved != null ? " (" + (tot.pctAchieved * 100).toFixed(0) + "% of target)" : "");
    const lbl = t => '<div style="' + FONT + 'font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:' + MUTED + ';">' + t + "</div>";
    const part = (dot, t, v, sub, italic) => '<td valign="top" style="padding:10px 26px 0 0;">' + '<div style="' + FONT + 'font-size:10px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:' + MUTED + ';"><span style="color:' + dot + ';">&#9632;</span> ' + t + "</div>"
        + '<div style="' + FONT + "font-size:19px;font-weight:" + (italic ? "500;font-style:italic" : "700") + ";color:" + INK + ';padding-top:2px;">' + v + "</div>"
        + '<div style="' + FONT + 'font-size:11px;color:' + MUTED + ';">' + sub + "</div></td>";
    const share = f.total > 0 ? Math.max(1, Math.min(99, Math.round(Math.max(0, f.sales) / f.total * 100))) : 100;
    const pane1 = lbl("Sales + Shipped Not Invoiced · month to date")
        + '<div style="' + FONT + 'font-size:34px;font-weight:700;color:' + INK + ';padding:4px 0 8px;">' + money(f.total) + "</div>"
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td width="' + share + '%" height="7" style="background:#26c6da;font-size:1px;line-height:1px;">&nbsp;</td><td height="7" style="background:#ffb300;font-size:1px;line-height:1px;">&nbsp;</td></tr></table>'
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>'
        + part("#26c6da", "Sales (excl. tax)", money(f.sales), f.invoices.toLocaleString("en-AU") + " invoices")
        + part("#ffb300", "Shipped Not Invoiced", money(f.sni), f.sniOrders + " orders · delivered, awaiting invoice", true) + "</tr></table>";
    const pane2 = lbl("Gross Profit · month to date")
        + '<div style="' + FONT + 'font-size:34px;font-weight:700;color:' + INK + ';padding:4px 0 2px;">' + money(f.gp) + ' <span style="font-size:17px;font-weight:600;color:' + MUTED + ';">' + pct(f.gpPct) + "</span></div>"
        + '<div style="' + FONT + 'font-size:11px;color:' + MUTED + ';">on invoiced sales · cost ' + money(f.cost) + "</div>"
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>'
        + part("#ffb300", "Including Shipped Not Invoiced" + (f.sniGpExact ? "" : " (estimate)"), money(f.gpAll) + ' <span style="font-size:12px;">' + pct(f.gpAllPct) + "</span>",
               f.sniGpExact ? "adds " + money(f.sniGp) + " margin on shipped goods, at order-line cost" : "adds " + money(f.sniGp) + ", assuming shipped goods earn the same margin", true) + "</tr></table>";
    const day = (t, v) => '<td width="25%" valign="top" style="padding:12px 14px;border:1px solid ' + LINE + ';">' + lbl(t) + '<div style="' + FONT + 'font-size:22px;font-weight:700;color:' + INK + ';padding-top:3px;">' + v + "</div></td>";
    const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>'
        + '<body style="margin:0;padding:0;background:#f4f1eb;">'
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f4f1eb;"><tr><td align="center" style="padding:18px 10px;">'
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="1080" style="width:1080px;max-width:100%;background:#ffffff;border:1px solid ' + LINE + ';"><tr><td style="padding:24px 28px 28px;">'
        + '<div style="' + FONT + 'font-size:22px;font-weight:700;color:' + INK + ';">Daily Snapshot</div>'
        + '<div style="' + FONT + 'font-size:13px;color:' + MUTED + ';padding:3px 0 18px;">' + esc(dateLbl) + " · <strong>" + esc(f.monthLabel) + "</strong> so far" + (asOf ? " · figures as of " + esc(asOf) : "") + "</div>"
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>'
        + '<td width="58%" valign="top" style="padding:16px 18px;border:1px solid ' + LINE + ';">' + pane1 + '</td><td width="12" style="font-size:1px;">&nbsp;</td>'
        + '<td valign="top" style="padding:16px 18px;border:1px solid ' + LINE + ';">' + pane2 + "</td></tr></table>"
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top:12px;border-collapse:collapse;"><tr>'
        + day("Total selling days", f.sd.total) + day("Selling days passed", f.sd.passed) + day("Remaining selling days", f.sd.remaining) + day("% month passed", f.sd.total ? pct(f.sd.passed / f.sd.total) : "—") + "</tr></table>"
        + runRateTable("Sales Run Rate · by Branch", "Branch", f.branch, res.names.branch || {}, f.sd)
        + runRateTable("Sales Run Rate · by Salesperson", "Salesperson", f.salesperson, res.names.salesperson || {}, f.sd)
        + '<div style="' + FONT + 'font-size:12px;color:' + MUTED + ';padding-top:22px;">% Achieved: <span style="color:' + BAND[0] + ';font-weight:700;">under 50%</span> · <span style="color:' + BAND[1] + ';font-weight:700;">50–99%</span> · <span style="color:' + BAND[2] + ';font-weight:700;">100%+</span>. Selling days = Mon–Fri minus Vic public holidays. Rental income is excluded.</div>'
        + '<div style="' + FONT + 'font-size:12px;color:' + MUTED + ';padding-top:10px;">Same figures as the Daily Snapshot tab of the Sales Analytics tool: <a href="' + TOOL_URL + '" style="color:#3f51b5;">open the tool</a>. Sent automatically each morning.</div>'
        + "</td></tr></table></td></tr></table></body></html>";
    return { subject, html };
}

async function sendMail(mail, to) {
    const resp = await graph("/me/sendMail", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message: { subject: mail.subject, body: { contentType: "HTML", content: mail.html }, toRecipients: to.map(a => ({ emailAddress: { address: a } })) }, saveToSentItems: true }) });
    if (resp.status !== 202) throw new Error("sendMail → HTTP " + resp.status + ": " + (await resp.text()).slice(0, 400));
}

(async () => {
    const local = process.argv.indexOf("--local");
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "daily-email-"));
    try {
        if (local >= 0) {
            fs.copyFileSync(process.argv[local + 1], path.join(dir, "payload.json"));
            fs.copyFileSync(process.argv[local + 2], path.join(dir, "targets.json"));
        } else {
            for (const k of ["BC_TENANT", "BC_CLIENT_ID", "BC_REFRESH_TOKEN", "SNAPSHOT_PASSPHRASE"]) if (!process.env[k]) throw new Error("Missing env " + k);
            await collectFromSharePoint(dir);
        }
        const res = await figuresFromPage(dir);
        console.log("Figures for " + res.f.monthLabel + ": " + res.invoices + " invoices in the snapshot · " + res.f.branch.rows.length + " branch rows · " + res.f.salesperson.rows.length + " salesperson rows · "
            + res.targetFiles + " targets file(s)" + (res.skipped.length ? " · skipped " + res.skipped.length : "") + " · page date " + res.today);
        const mail = renderEmail(res);
        console.log("Email built: " + (mail.html.length / 1024).toFixed(1) + " KB");
        if (local >= 0) { fs.writeFileSync(process.argv[local + 3], mail.html); console.log("Wrote " + process.argv[local + 3] + " · subject: " + mail.subject); return; }
        const to = (process.env.DAILY_EMAIL_TO || "").split(/[,;\s]+/).filter(Boolean);
        if (process.env.DRY_RUN === "1" || process.env.DRY_RUN === "true") { await graphToken(); console.log("DRY RUN — nothing sent (" + to.length + " recipient(s) configured)."); return; }
        if (!to.length) throw new Error("DAILY_EMAIL_TO is empty — no recipients");
        await sendMail(mail, to);
        console.log("Sent to " + to.length + " recipient(s).");
    } finally { try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) { /* temp dir */ } }
})().catch(e => { console.error("DAILY EMAIL FAILED: " + e.message); process.exit(1); });
