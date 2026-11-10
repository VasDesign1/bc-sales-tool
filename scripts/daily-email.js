// ============================================================
// daily-email.js — emails the Daily Snapshot tab (Faik, 2026-11-10).
//
// The figures are NOT recalculated here. The robot opens the tool's own
// index.html in headless Chrome, pours the newest snapshot into it the
// same way Fast lookup does, reads the targets workbooks with the page's
// own parser, then takes a picture of the Daily Snapshot tab as the tool
// draws it. That picture is the email, so tab and email can never
// disagree in figures or in looks. dsSnapshotFigures() — the object the
// tab paints — is read as well, for the subject line and for the
// optional AI commentary above the picture.
//
// Runs in GitHub Actions on Node 22 (needs the built-in WebSocket).
// Env:
//   BC_TENANT, BC_CLIENT_ID, BC_REFRESH_TOKEN   token exchange (Graph:
//                                               Files.ReadWrite.All + Mail.Send, delegated)
//   SNAPSHOT_PASSPHRASE                         opens the snapshot
//   DAILY_EMAIL_TO                              recipients, comma separated
//   ANTHROPIC_API_KEY                           optional: adds the commentary paragraph
//                                               (needs `npm install @anthropic-ai/sdk`)
//   EMAIL_THEME                                 optional: tool theme for the picture (default linen)
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
const THEME = (process.env.EMAIL_THEME || "linen").toLowerCase().replace(/[^a-z]/g, "") || "linen";
const SHOT_WIDTH = 1560;      // CSS px the tab is laid out at for the picture: wide enough for all ten columns and the two kit panels side by side
const SHOW_WIDTH = 1200;      // px the email shows the picture at (it is taken at 2x, so it stays sharp)
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
    setActiveTab("dailySnapshot");          // paints the tab with the targets now in place
    const f = dsSnapshotFigures();
    return JSON.stringify({ f, names: { branch: spBudgetNames("branch"), salesperson: spBudgetNames("salesperson") },
        asOf: state.dataAsOf || "", invoices: state.invoices.length, targetFiles: (spTargets.files || []).length, skipped: spTargets.skipped || [],
        today: localISO(new Date()) });
})()`;
// Strips the page down to the tab itself for the picture: no sidebar,
// header or filter bar, and none of the tab's own controls that mean
// nothing in an email (targets-file picker, "Full tab" buttons).
const CAPTURE_SCRIPT = `(async () => {
    const st = document.createElement("style");
    st.textContent = ".sidebar, .header, .toolbar, #tabStrip, #filterChips, .filter-chips, #themePicker, .progress, #busyWrap, #exportMenu, #helpTip { display: none !important; }"
        + " .app { display: block !important; } .main { margin: 0 !important; width: auto !important; } .container { padding: 18px 20px !important; max-width: none !important; }"
        + " section[data-tab-content='dailySnapshot'] .rr-tgt, section[data-tab-content='dailySnapshot'] [data-ds-goto], #dsEmpty, #dsKits .panel-header .meta { display: none !important; }";
    document.head.appendChild(st);
    window.scrollTo(0, 0);
    if (document.fonts && document.fonts.ready) await document.fonts.ready;
    await new Promise(r => setTimeout(r, 400));
    const c = document.querySelector(".container").getBoundingClientRect();
    // The page's own background and text colour, so the email around the picture can match the theme.
    let bg = "rgb(255, 255, 255)";
    for (let e = document.querySelector(".container"); e; e = e.parentElement) { const b = getComputedStyle(e).backgroundColor; if (b && b !== "transparent" && b !== "rgba(0, 0, 0, 0)") { bg = b; break; } }
    return JSON.stringify({ clip: { x: 0, y: Math.max(0, Math.floor(c.top + window.scrollY)), width: Math.ceil(document.documentElement.clientWidth), height: Math.ceil(c.height) },
        bg, ink: getComputedStyle(document.querySelector(".container")).color });
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
        "--user-data-dir=" + path.join(dataDir, "chrome"), "--window-size=" + SHOT_WIDTH + ",1200", "--hide-scrollbars", "about:blank"], { stdio: "ignore" });
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
        // Picture settings: the chosen theme from the first paint, and two device pixels per CSS pixel so it stays sharp.
        await send("Page.addScriptToEvaluateOnNewDocument", { source: "try{localStorage.setItem('bcSalesTheme','" + THEME + "');}catch(e){}" });
        await send("Emulation.setDeviceMetricsOverride", { width: SHOT_WIDTH, height: 1200, deviceScaleFactor: 2, mobile: false });
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
        const res = JSON.parse(r.result.result.value);
        const c = await send("Runtime.evaluate", { expression: CAPTURE_SCRIPT, returnByValue: true, awaitPromise: true });
        if (!c.result || !c.result.result || typeof c.result.result.value !== "string") throw new Error("the page could not be prepared for the picture: " + JSON.stringify(c.result || c).slice(0, 400));
        const prep = JSON.parse(c.result.result.value), clip = prep.clip;
        const hex = (rgb, fallback) => { const m = /(\d+)\D+(\d+)\D+(\d+)/.exec(rgb || ""); return m ? "#" + [1, 2, 3].map(i => (+m[i]).toString(16).padStart(2, "0")).join("") : fallback; };
        // Graph takes about 3 MB of attachment in one sendMail call: step the sharpness down if a picture is too heavy.
        let png = null, scale = 2;
        for (const dsf of [2, 1.5, 1]) {
            scale = dsf;
            await send("Emulation.setDeviceMetricsOverride", { width: SHOT_WIDTH, height: 1200, deviceScaleFactor: dsf, mobile: false });
            await sleep(300);
            const shot = await send("Page.captureScreenshot", { format: "png", captureBeyondViewport: true, clip: Object.assign({ scale: 1 }, clip) });
            if (!shot.result || !shot.result.data) throw new Error("screenshot failed: " + JSON.stringify(shot).slice(0, 300));
            png = Buffer.from(shot.result.data, "base64");
            if (png.length <= 2.6 * 1048576) break;
        }
        res.png = png;
        res.shot = { cssWidth: clip.width, pxWidth: Math.round(clip.width * scale), pxHeight: Math.round(clip.height * scale), theme: THEME,
                     bg: hex(prep.bg, "#ffffff"), ink: hex(prep.ink, "#1f2430") };
        ws.close();
        return res;
    } finally { chrome.kill(); srv.close(); }
}

// ---------- AI commentary (optional) ----------
// A few dot points above the picture. The model is handed the figures
// the tab already shows and only describes them; it calculates nothing
// that lands in the email. Skipped quietly when there is no API key, or
// when the call fails or is declined: the email must still go out.
const COMMENTARY_SYSTEM = [
    "You write the short opening dot points of a daily sales email at an Australian HVAC wholesaler. It goes to internal staff every morning, above a picture of the sales dashboard they already know.",
    "You are given the dashboard's own figures as JSON: month-to-date sales, goods shipped but not yet invoiced, gross profit, selling days, each branch and each salesperson against their monthly sales target, and two product quantity targets (AirTouch kits and Omniflow units).",
    "Write 10 to 12 dot points that walk a busy reader through where the month stands and what matters today. Cover, in roughly this order: sales so far against the month's target and against where they should be at this point of the month; goods shipped but not yet invoiced and what they add; gross profit and margin; the daily rate achieved so far and the rate now required; each branch in turn; the salespeople picture; the AirTouch kit target; the Omniflow target; and what would most help close the gap. One point per idea, each a single plain sentence of at most about 25 words, and no two points making the same observation. If the figures do not support ten distinct points, write fewer rather than padding.",
    "Use only the figures provided. Quote them as given, rounded sensibly (for example $401k, 19% of target); do not derive new totals or forecasts of your own. If something needed is missing, leave it out rather than guessing.",
    "This is read by the whole team, so name a salesperson only to recognise a result that is ahead of pace; describe shortfalls at branch or team level, not by person.",
    "Australian English, matter-of-fact and encouraging without hype. Output format: each dot point on its own line starting with a hyphen and a space, and nothing else: no greeting, no intro line, no sign-off, no headings, no bold or other markdown, no emojis.",
].join("\n\n");
function commentaryInput(res) {
    const f = res.f;
    const rows = (t, names) => t.rows.filter(r => r.target > 0).map(r => ({ code: r.entity, name: names[r.entity] || undefined, target: Math.round(r.target), sales_so_far: Math.round(r.cyVal),
        pct_of_target: +(r.pctAchieved * 100).toFixed(1), status: r.status, required_daily: Math.round(r.requiredDaily) }));
    const tot = t => ({ target: Math.round(t.total.target), sales_so_far: Math.round(t.total.cyVal), pct_of_target: t.total.pctAchieved == null ? null : +(t.total.pctAchieved * 100).toFixed(1),
        status: t.total.status, target_per_day: Math.round(t.total.targetDaily), actual_per_day: Math.round(t.total.actualDaily), required_per_day: Math.round(t.total.requiredDaily) });
    return {
        date: res.today, month: f.monthLabel, figures_as_of: res.asOf,
        selling_days: { total: f.sd.total, passed: f.sd.passed, remaining: f.sd.remaining, pct_of_month_passed: f.sd.total ? +(f.sd.passed / f.sd.total * 100).toFixed(1) : null },
        month_to_date: { sales_excl_tax: Math.round(f.sales), shipped_not_invoiced: Math.round(f.sni), sales_plus_shipped_not_invoiced: Math.round(f.total),
                         gross_profit: Math.round(f.gp), gross_profit_pct: +(f.gpPct * 100).toFixed(2), invoices: f.invoices },
        branches: { total: tot(f.branch), rows: rows(f.branch, res.names.branch || {}) },
        salespeople: { total: tot(f.salesperson), rows: rows(f.salesperson, res.names.salesperson || {}) },
        quantity_targets: (f.kits || []).map(k => ({ name: k.title, monthly_target_units: k.target, units_sold: k.qty, pct_of_target: +(k.pctAchieved * 100).toFixed(1), status: k.status,
                                                    required_per_day: +Math.max(0, k.requiredDaily).toFixed(1), by_item: k.rows.map(r => ({ item: r.label, units: r.qty })) })),
    };
}
async function writeCommentary(res) {
    if (!process.env.ANTHROPIC_API_KEY) { console.log("Commentary: no ANTHROPIC_API_KEY — email goes without it."); return ""; }
    try {
        const mod = require("@anthropic-ai/sdk");
        const Anthropic = mod.default || mod;
        const client = new Anthropic();
        const response = await client.beta.messages.create({
            model: "claude-opus-5-5",
            max_tokens: 8000,
            // If a safety classifier declines, the API re-runs the request on its default fallback model.
            betas: ["server-side-fallback-2026-07-01"],
            fallbacks: "default",
            system: COMMENTARY_SYSTEM,
            messages: [{ role: "user", content: "Today's dashboard figures:\n\n" + JSON.stringify(commentaryInput(res), null, 1) }],
        });
        if (response.stop_reason === "refusal") { console.warn("Commentary: request declined — email goes without it."); return ""; }
        const text = response.content.filter(b => b.type === "text").map(b => b.text).join("\n").trim();
        console.log("Commentary: " + text.length + " characters · model " + response.model + " · tokens in " + response.usage.input_tokens + " / out " + response.usage.output_tokens);
        return text;
    } catch (e) {
        console.warn("Commentary: skipped (" + String((e && e.message) || e).slice(0, 200) + ") — email goes without it.");
        return "";
    }
}

// ---------- the email ----------
// The body is a picture of the Daily Snapshot tab exactly as the tool
// draws it (Faik, 2026-11-10: same colours, shape and format as the tool),
// with the date and the optional commentary as real text above it.
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const money0 = n => "$" + Math.round(n).toLocaleString("en-AU");
function renderEmail(res, commentary, imgSrc) {
    const f = res.f, d = new Date(res.today + "T12:00:00Z");
    const tot = f.branch.total;
    const subject = "Daily Snapshot · " + d.toLocaleDateString("en-AU", { weekday: "short", day: "numeric", month: "short", timeZone: "UTC" })
        + " · " + f.monthLabel + " sales " + money0(f.sales) + (tot.pctAchieved != null ? " (" + (tot.pctAchieved * 100).toFixed(0) + "% of target)" : "");
    const alt = "Daily Snapshot for " + f.monthLabel + ": sales plus shipped not invoiced " + money0(f.total) + ", gross profit " + money0(f.gp) + " (" + (f.gpPct * 100).toFixed(2) + "%).";
    // Layout (Faik, 2026-11-10): a bold header, a gap, the commentary as a
    // dot-point list, a gap, then the picture of the tab. Calibri 11pt.
    // The block sits centred in the message, text left-aligned and exactly
    // as wide as the picture. The whole message takes the tab's own page
    // colour (and its text colour), so the picture blends into the email
    // instead of sitting on white.
    const bg = res.shot.bg, ink = res.shot.ink;
    const FONT = "font-family:Calibri,Segoe UI,Arial,sans-serif;font-size:11pt;color:" + ink + ";";
    const header = "DAILY SALES SUMMARY · " + d.toLocaleDateString("en-AU", { weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "UTC" }).toUpperCase();
    const points = commentary ? commentary.split(/\n+/).map(p => p.replace(/^\s*(?:[-•*–]|\d+[.)])\s*/, "").trim()).filter(Boolean) : [];
    const list = points.length ? '<ul style="' + FONT + 'line-height:1.45;margin:0;padding:0 0 0 30px;">'
        + points.map(p => '<li style="' + FONT + 'margin:0 0 5px;">' + esc(p) + "</li>").join("") + "</ul>" : "";
    const gap = h => '<tr><td height="' + h + '" style="height:' + h + 'px;font-size:1px;line-height:1px;">&nbsp;</td></tr>';
    const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>'
        + '<body bgcolor="' + bg + '" style="margin:0;padding:0;background-color:' + bg + ';">'
        + '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="' + bg + '" style="background-color:' + bg + ';"><tr><td align="center" style="padding:18px 14px 24px;">'
        + '<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" width="' + SHOW_WIDTH + '" style="width:' + SHOW_WIDTH + 'px;max-width:100%;margin:0 auto;">'
        + '<tr><td align="left" style="' + FONT + 'font-size:13pt;font-weight:700;letter-spacing:.3px;">' + esc(header) + "</td></tr>"
        + (list ? gap(18) + '<tr><td align="left" style="' + FONT + '">' + list + "</td></tr>" : "")
        + gap(24)
        + '<tr><td align="left"><img src="' + imgSrc + '" width="' + SHOW_WIDTH + '" alt="' + esc(alt) + '" style="display:block;width:100%;max-width:' + SHOW_WIDTH + 'px;height:auto;border:0;margin:0 auto;"></td></tr>'
        + "</table></td></tr></table></body></html>";
    return { subject, html };
}

async function sendMail(mail, png, to) {
    const message = {
        subject: mail.subject,
        body: { contentType: "HTML", content: mail.html },
        toRecipients: to.map(a => ({ emailAddress: { address: a } })),
        attachments: [{ "@odata.type": "#microsoft.graph.fileAttachment", name: "daily-snapshot.png", contentType: "image/png",
                        contentBytes: png.toString("base64"), contentId: "dailysnapshot", isInline: true }],
    };
    const resp = await graph("/me/sendMail", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ message, saveToSentItems: true }) });
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
            + res.targetFiles + " targets file(s)" + (res.skipped.length ? " · skipped " + res.skipped.length : "") + " · page date " + res.today
            + " · picture " + res.shot.pxWidth + "×" + res.shot.pxHeight + " px, " + (res.png.length / 1024).toFixed(0) + " KB (" + res.shot.theme + " theme)");
        const commentary = await writeCommentary(res);
        if (local >= 0) {
            const outHtml = process.argv[local + 3], outPng = outHtml.replace(/\.html?$/i, "") + ".png";
            fs.writeFileSync(outPng, res.png);
            fs.writeFileSync(outHtml, renderEmail(res, commentary, path.basename(outPng)).html);
            console.log("Wrote " + outHtml + " + picture");
            return;
        }
        const mail = renderEmail(res, commentary, "cid:dailysnapshot");
        const to = (process.env.DAILY_EMAIL_TO || "").split(/[,;\s]+/).filter(Boolean);
        if (process.env.DRY_RUN === "1" || process.env.DRY_RUN === "true") { await graphToken(); console.log("DRY RUN — nothing sent (" + to.length + " recipient(s) configured)."); return; }
        if (!to.length) throw new Error("DAILY_EMAIL_TO is empty — no recipients");
        await sendMail(mail, res.png, to);
        console.log("Sent to " + to.length + " recipient(s)" + (commentary ? ", with commentary." : ", without commentary."));
    } finally { try { fs.rmSync(dir, { recursive: true, force: true }); } catch (e) { /* temp dir */ } }
})().catch(e => { console.error("DAILY EMAIL FAILED: " + e.message); process.exit(1); });
