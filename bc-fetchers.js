// ============================================================
// bc-fetchers.js — BUSINESS CENTRAL / WIISE FETCH LAYER
// Shared verbatim between the browser tool (index.html) and the
// snapshot GitHub Action (Node). Keep it environment-neutral:
// no DOM, no MSAL. The host must provide these globals before any
// fetcher is called:
//   bcGetToken()              -> Promise<access token string>
//   updateBCProgress(l, d)    -> progress sink (may be a no-op)
//   state                     -> shared state object (fetchers read
//                                state.fastTrim and write diagnostics)
//   num(v)                    -> numeric coercion util from the tool
//   fetch                     -> browser-native / Node 18+ global
// ============================================================

// ============================================================
// BUSINESS CENTRAL / WIISE CONNECTION LAYER
// ============================================================
const BC_CONFIG = {
    clientId:    "0f3136a8-79cd-4335-9790-7ae3fe5800be",
    tenantId:    "68c88731-a731-4307-bb12-28557affd0ca",
    environment: "Production",
    companyName: "VICAIR Pty Ltd",
};
const BC_TENANT_DOMAIN = "VicAirPtyLtd.onmicrosoft.com";
const BC_API_BASE = "https://api.businesscentral.dynamics.com/v2.0/" + BC_TENANT_DOMAIN + "/" + BC_CONFIG.environment;
const BC_API_URL  = BC_API_BASE + "/api/v2.0";
const BC_ODATA_URL = BC_API_BASE + "/ODataV4";
const BC_SCOPES   = ["https://api.businesscentral.dynamics.com/.default"];

// SharePoint home of the encrypted Fast-lookup snapshots (snapshots/ sub-
// folder of the tool folder). Written by the
// snapshot robot (scripts/sp-upload.js, delegated refresh token) and read
// by the browser through Microsoft Graph with the signed-in user's token —
// Files.Read.All on the "VicAir Forecast Tool" app registration. Knowing
// the path grants nothing: Graph enforces the folder's SharePoint ACL.
const SP_CONFIG = {
    host:   "vicairptyltd.sharepoint.com",
    folder: "Vic Air Shared Docs Folder/03.0 Corporate/3.05 IT and Communications/bc-tools/sales-tool",
};
const GRAPH_SCOPES = ["https://graph.microsoft.com/Files.Read.All"];

let bcCompanyId = null, bcCompanyInternalName = null, bcODataMetadata = null;

// A dropped connection makes fetch throw rather than return a status; retry
// those a few times before giving up, so one network blip doesn't sink a
// whole load or snapshot.
async function bcFetchRetrying(url, opts) {
    for (let attempt = 0; ; attempt++) {
        try { return await fetch(url, opts); }
        catch (e) {
            if (attempt >= 3) throw e;
            console.warn("[BC] connection dropped (" + e.message + ") — retrying in " + (2 * (attempt + 1)) + " s");
            await new Promise(r => setTimeout(r, 2000 * (attempt + 1)));
        }
    }
}
async function bcFetch(url) {
    const token = await bcGetToken();
    let resp = await bcFetchRetrying(url, { headers: { "Authorization": "Bearer " + token, "Accept": "application/json" } });
    if (resp.status === 401) {
        bcAccessToken = null;
        const newToken = await bcGetToken();
        resp = await bcFetchRetrying(url, { headers: { "Authorization": "Bearer " + newToken, "Accept": "application/json" } });
    }
    if (!resp.ok) {
        const errBody = await resp.text().catch(() => "");
        throw new Error("BC API error " + resp.status + ": " + resp.statusText + (errBody ? " — " + errBody.substring(0, 200) : ""));
    }
    return resp.json();
}
async function bcFetchAll(url, progressLabel) {
    let allRows = [], nextUrl = url, page = 0;
    while (nextUrl) {
        page++;
        const data = await bcFetch(nextUrl);
        allRows = allRows.concat(data.value || []);
        updateBCProgress(progressLabel, allRows.length + " rows (page " + page + ")...");
        nextUrl = data["@odata.nextLink"] || null;
    }
    return allRows;
}
async function bcGetCompanyId() {
    if (bcCompanyId) return bcCompanyId;
    const data = await bcFetch(BC_API_URL + "/companies");
    const companies = data.value || [];
    if (!companies.length) throw new Error("No companies found.");
    const target = BC_CONFIG.companyName.toLowerCase();
    let match = companies.find(c => (c.name || "").toLowerCase() === target)
             || companies.find(c => (c.displayName || "").toLowerCase() === target)
             || companies[0];
    bcCompanyId = match.id;
    bcCompanyInternalName = match.name || match.displayName;
    return bcCompanyId;
}
async function bcGetCompanyInternalName() {
    if (bcCompanyInternalName) return bcCompanyInternalName;
    await bcGetCompanyId();
    return bcCompanyInternalName;
}

// ============================================================
// OData $metadata discovery — calendar-viewer pattern.
// Sales Quote / Blanket Sales Order entities are published under
// tenant-specific names. Fetch $metadata once, list every EntitySet,
// find the one matching the desired hints + required fields, and use
// the real published name in the OData URL.
// ============================================================
async function bcGetODataMetadata() {
    if (bcODataMetadata) return bcODataMetadata;
    const token = await bcGetToken();
    const resp = await bcFetchRetrying(BC_ODATA_URL + "/$metadata", { headers: { "Authorization": "Bearer " + token } });
    if (!resp.ok) throw new Error("OData $metadata fetch " + resp.status);
    const xml = await resp.text();
    const entitySets = {};
    const setRe = /<EntitySet\s+Name="([^"]+)"\s+EntityType="[^"]*?\.([^"]+)"/g;
    let m;
    while ((m = setRe.exec(xml)) !== null) entitySets[m[1]] = m[2];
    const entityTypes = {};
    const typeBlockRe = /<EntityType\s+Name="([^"]+)"[^>]*>([\s\S]*?)<\/EntityType>/g;
    while ((m = typeBlockRe.exec(xml)) !== null) {
        const propRe = /<Property\s+Name="([^"]+)"/g;
        const fields = []; let p;
        while ((p = propRe.exec(m[2])) !== null) fields.push(p[1]);
        // BC / Wiise metadata sometimes ships multiple EntityType blocks
        // with the same Name (e.g. one tiny salutation fragment plus the
        // real header). Keep the version with the most fields — that's
        // overwhelmingly the page's actual data shape.
        if (!entityTypes[m[1]] || entityTypes[m[1]].length < fields.length) {
            entityTypes[m[1]] = fields;
        }
    }
    bcODataMetadata = { entitySets, entityTypes };
    return bcODataMetadata;
}
function bcFindField(names, patterns) {
    for (const p of patterns) { const hit = names.find(n => p.test(n)); if (hit) return hit; }
    return null;
}
// Find the matching *lines* entity for a header entity (e.g. for
// "BlanketSalesOrder", look for "BlanketSalesOrderSalesLines" or similar).
// Required fields: Document No, Line No, Item No (or Type+No), Quantity, Amount.
async function bcDiscoverSalesLinesEntity(headerEntity) {
    const md = await bcGetODataMetadata();
    const allSets = Object.keys(md.entitySets);
    const DOC_NO   = [/^document.?no$/i, /documentNumber/i];
    const LINE_NO  = [/^line.?no$/i, /lineNumber/i];
    // v2-style salesDocumentLines uses just "number" for the item/account
    // number. OData/NAV-style uses "No" or "Item_No_". Match all three;
    // do NOT match anything starting with sellTo/billTo/customer/etc.
    const ITEM_NO  = [/^number$/i, /^no$/i, /^no_$/i, /^item.?no_?$/i, /^itemNumber$/i, /^lineObjectNumber$/i];
    const QTY      = [/^quantity$/i, /^qty$/i];
    const AMOUNT   = [/^amount$/i, /^line.?amount$/i, /amountExcludingTax/i, /^amount.*excl/i];
    const UPRICE   = [/^unit.?price$/i, /unitPrice/i];
    const DESC     = [/^description$/i];
    const LOCATION = [/^location.?code$/i, /locationCode/i, /^location$/i];
    const TYPE     = [/^type$/i, /^line.?type$/i, /lineType/i];
    const DISC_ALLOC = [/^inv.*discount.?allocation$/i, /^invoiceDiscountAllocation$/i, /^invDiscount.?Amount$/i, /^invoice.?discount.?amount$/i];
    // Score candidates by name closeness to header + presence of line indicators
    const base = headerEntity.toLowerCase();
    const candidates = allSets.filter(n => {
        const l = n.toLowerCase();
        if (l === base) return false;
        return /line|sales.?line/i.test(l) && l.includes(base.replace(/s$/, "").substring(0, Math.min(8, base.length)));
    });
    // Fallback: any entity starting with header name that has "line" in it
    if (!candidates.length) {
        for (const n of allSets) {
            if (n === headerEntity) continue;
            if (n.toLowerCase().startsWith(base.toLowerCase()) && /line/i.test(n)) candidates.push(n);
        }
    }
    // Final fallback: scan for entities with Document_No + Quantity + Amount fields
    const fallbackPool = candidates.length ? candidates : allSets;
    for (const candidate of fallbackPool) {
        const typeName = md.entitySets[candidate];
        const fields = md.entityTypes[typeName] || [];
        const fDocNo  = bcFindField(fields, DOC_NO);
        const fQty    = bcFindField(fields, QTY);
        const fAmount = bcFindField(fields, AMOUNT);
        if (!fDocNo || !fQty || !fAmount) continue;
        return {
            entity: candidate, fields,
            fDocNo,
            fLineNo:    bcFindField(fields, LINE_NO),
            fItemNo:    bcFindField(fields, ITEM_NO),
            fQty, fAmount,
            fUnitPrice: bcFindField(fields, UPRICE),
            fDesc:      bcFindField(fields, DESC),
            fLocation:  bcFindField(fields, LOCATION),
            fType:      bcFindField(fields, TYPE),
            fDiscAlloc: bcFindField(fields, DISC_ALLOC),
        };
    }
    return null;
}

// Discover the Value Entry entity (BC page 5802). PBI reads cost and
// sales from Value Entries, not from invoice lines. Returns the published
// entity name + a field map for the columns the tool needs.
async function bcDiscoverValueEntryEntity() {
    const md = await bcGetODataMetadata();
    const allSets = Object.keys(md.entitySets);
    const named = allSets.filter(n => /value.?entr/i.test(n));
    const SALES_AMT  = [/^salesAmountActual$/i, /^Sales_Amount_Actual_?$/i, /^sales_amount_actual$/i];
    const COST_AMT   = [/^costAmountActual$/i, /^Cost_Amount_Actual_?$/i, /^cost_amount_actual$/i];
    const COST_NI    = [/^costAmountNonInvtbl$/i, /^Cost_Amount_Non_Invtbl_?$/i, /^cost_amount_non_invtbl$/i];
    const ITEM_NO    = [/^itemNumber$/i, /^Item_No_?$/i, /^itemNo$/i];
    const DOC_NO     = [/^documentNumber$/i, /^Document_No_?$/i, /^documentNo$/i];
    const DOC_TYPE   = [/^documentType$/i, /^Document_Type$/i];
    const SOURCE_NO  = [/^sourceNumber$/i, /^Source_No_?$/i, /^sourceNo$/i];
    const SOURCE_TYP = [/^sourceType$/i, /^Source_Type$/i];
    const DATE       = [/^postingDate$/i, /^Posting_Date$/i, /^posting_date$/i];
    const DOC_DATE   = [/^documentDate$/i, /^Document_Date$/i, /^document_date$/i];
    const QTY_INVD   = [/^invoicedQuantity$/i, /^Invoiced_Quantity$/i, /^invoiced_quantity$/i];
    const ENTRY_TYPE = [/^itemLedgerEntryType$/i, /^Item_Ledger_Entry_Type$/i];
    // Prefer a service that also carries the value-entry type (Direct Cost,
    // Revaluation, Rounding...): same rows, but the drill-down can say what
    // each one is in Wiise's own words.
    const hasType = n => (md.entityTypes[md.entitySets[n]] || []).includes("Entry_Type");
    const candidates = (named.length ? named : allSets).slice().sort((a, b) => hasType(b) - hasType(a));
    for (const candidate of candidates) {
        const typeName = md.entitySets[candidate];
        const fields = md.entityTypes[typeName] || [];
        if (!fields.length) continue;
        const fSales   = bcFindField(fields, SALES_AMT);
        const fCost    = bcFindField(fields, COST_AMT);
        const fItem    = bcFindField(fields, ITEM_NO);
        const fDate    = bcFindField(fields, DATE);
        // Need at least sales OR cost, plus item + date — minimum to be useful
        if (!(fSales || fCost) || !fItem || !fDate) continue;
        return {
            entity: candidate, fields,
            fSales, fCost,
            fCostNI:    bcFindField(fields, COST_NI),
            fItem,      fDate,
            fDocDate:   bcFindField(fields, DOC_DATE),
            fDocNo:     bcFindField(fields, DOC_NO),
            fDocType:   bcFindField(fields, DOC_TYPE),
            fSourceNo:  bcFindField(fields, SOURCE_NO),
            fSourceTyp: bcFindField(fields, SOURCE_TYP),
            fQtyInvd:   bcFindField(fields, QTY_INVD),
            fEntryType: bcFindField(fields, ENTRY_TYPE),
            fEntryNo:   bcFindField(fields, [/^entryNumber$/i, /^Entry_No_?$/i]),
            fCostExp:   bcFindField(fields, [/^costAmountExpected$/i, /^Cost_Amount_Expected_?$/i]),
            fValueType: bcFindField(fields, [/^Entry_Type$/]),
        };
    }
    return null;
}

// Does $select keep the date filter on this value-entry service? Asks for at
// most 200 rows of the last month of the range twice, with and without a
// column list ($top is used deliberately here: a cap is exactly what a probe
// wants). Safe only if both agree and every row is inside the range.
const veSelectOk = {};
async function veProbeSelect(base, info, sel, fromISO, toISO) {
    if (veSelectOk[info.entity] !== undefined) return veSelectOk[info.entity];
    let ok = false;
    try {
        const end = toISO;
        const d = new Date(toISO + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() - 30);
        const start = d.toISOString().slice(0, 10) > fromISO ? d.toISOString().slice(0, 10) : fromISO;
        const filt = "$filter=" + encodeURIComponent(info.fDate + " ge " + start + " and " + info.fDate + " le " + end);
        const a = await bcFetch(base + "?" + filt + "&$select=" + sel.join(",") + "&$top=200");
        const b = await bcFetch(base + "?" + filt + "&$top=200");
        const ra = (a && a.value) || [], rb = (b && b.value) || [];
        const inRange = r => { const k = String(r[info.fDate] || "").slice(0, 10); return k >= start && k <= end; };
        ok = ra.length > 0 && ra.length === rb.length && ra.every(inRange)
          && ra.every(r => Object.keys(r).filter(k => k[0] !== "@").length <= sel.length);
    } catch (e) { ok = false; }
    veSelectOk[info.entity] = ok;
    console.log("[Value Entry] column list on " + info.entity + ": " + (ok ? "honoured — fetching " + sel.length + " columns" : "not safe — fetching all columns"));
    return ok;
}

// Fetch Value Entries for the period, with discovered field names.
// Returns an array of normalised rows {postingDate, itemNumber,
// documentNumber, documentType, sourceNumber, sourceType, salesAmount,
// costAmount, costAmountNonInv, invoicedQty, entryType} — sourceNumber
// is the customer No on sales rows, vendor No on purchase rows.
async function fetchValueEntries(fromISO, toISO) {
    state.veDiagnostic = "";
    let info;
    try { info = await bcDiscoverValueEntryEntity(); }
    catch (e) {
        console.warn("[Value Entry] metadata discovery failed:", e.message);
        state.veDiagnostic = "OData $metadata fetch failed: " + e.message;
        return [];
    }
    if (!info) {
        state.veDiagnostic = "No Value Entry entity published in this tenant. Ask BC admin to publish page 5802 as a Web Service named 'ValueEntries'.";
        console.warn("[Value Entry] no candidate entity found");
        return [];
    }
    const coName = encodeURIComponent(await bcGetCompanyInternalName());
    const filter = info.fDate + " ge " + fromISO + " and " + info.fDate + " le " + toISO;
    // No $top: BC treats it as a TOTAL cap, not a page size. The old
    // $top=200000 silently truncated the table once it passed 200k rows
    // (2026-10-01: snapshot held entries only up to 25 Sep → September
    // cost $360k short vs PBI). The server pages at 20k via nextLink.
    const params = ["$filter=" + encodeURIComponent(filter)];
    // Ask only for the columns the tool reads (15 of ~66 on ValueEntriesFull)
    // — but only if this service still honours the date filter with a
    // column list. The older ValueEntries query object was seen to DROP the
    // $filter when $select was added and return the whole table, so a small
    // probe decides first; if it fails, nothing changes from before.
    const base = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity;
    const sel = [info.fDate, info.fDocDate, info.fItem, info.fDocNo, info.fDocType, info.fSourceNo, info.fSourceTyp,
                 info.fSales, info.fCost, info.fCostNI, info.fQtyInvd, info.fEntryType, info.fEntryNo, info.fCostExp, info.fValueType]
                .filter(Boolean);
    if (await veProbeSelect(base, info, sel, fromISO, toISO)) params.push("$select=" + sel.join(","));
    const url = base + "?" + params.join("&");
    try {
        const t0 = Date.now();
        const rows = await bcFetchAll(url, "Value Entries (" + info.entity + ")");
        console.log("[Value Entry] " + info.entity + " → " + rows.length + " rows in " + ((Date.now() - t0) / 1000).toFixed(1) + "s"
                    + (params.length > 1 ? " (" + sel.length + " columns)" : " (all columns)"));
        if (rows.length) console.log("[Value Entry] sample row keys:", Object.keys(rows[0]).join(", "));
        const normalised = rows.map(r => ({
            postingDate:      info.fDate      ? r[info.fDate]      : "",
            documentDate:     info.fDocDate   ? r[info.fDocDate]   : "",
            itemNumber:       info.fItem      ? r[info.fItem]      : "",
            documentNumber:   info.fDocNo     ? r[info.fDocNo]     : "",
            documentType:     info.fDocType   ? r[info.fDocType]   : "",
            sourceNumber:     info.fSourceNo  ? r[info.fSourceNo]  : "",
            sourceType:       info.fSourceTyp ? r[info.fSourceTyp] : "",
            salesAmount:      info.fSales     ? num(r[info.fSales])  : 0,
            costAmount:       info.fCost      ? num(r[info.fCost])   : 0,
            costAmountNonInv: info.fCostNI    ? num(r[info.fCostNI]) : 0,
            invoicedQty:      info.fQtyInvd   ? num(r[info.fQtyInvd]) : 0,
            entryType:        info.fEntryType ? r[info.fEntryType] : "",
            entryNo:          info.fEntryNo   ? r[info.fEntryNo]   : null,
            costAmountExpected: info.fCostExp ? num(r[info.fCostExp]) : 0,
            valueType:        info.fValueType ? r[info.fValueType] : "",
        }));
        state.veFieldMap = info;
        state.veDiagnostic = "Source: OData " + info.entity + " · " + rows.length + " rows.";
        return normalised;
    } catch (e) {
        console.warn("[Value Entry] OData fetch failed:", e.message);
        state.veDiagnostic = "OData " + info.entity + " fetch failed: " + e.message;
        return [];
    }
}

// kind: "quote" | "blanket"
async function bcDiscoverSalesEntity(kind) {
    const md = await bcGetODataMetadata();
    const allSets = Object.keys(md.entitySets);
    const excluded = /^purchase|^vendor|^item|^job|^journal|^ledger|^chart|^contact|price.?list|^accountant|^company|^segment|^dimension|^stockkeep|template|^bank|^res.?ledger|^fa.?ledger|^cust.?ledger|workflow/i;
    const included = /sales|quote|blanket|document|pipeline|opportunity|dashboard/i;
    const docTypeNeeded = kind === "blanket" ? /blanket/i : /quote/i;
    const docTypeValue = kind === "blanket" ? "Blanket Order" : "Quote";
    const CUST_NO = [/^sell.?to.?cust.*no/i, /^sell.?to.?customer.*no/i, /^customer.?no$/i, /customer.?no/i, /customerNumber/i];
    const CUST_NAME = [/^sell.?to.?cust.*name/i, /^sell.?to.?customer.*name/i, /^customer.?name$/i, /customer.?name/i, /customerName/i];
    const AMOUNT = [/^amount.*excl/i, /^total.*amount.*excl/i, /^total.*excl/i, /excl.?gst/i, /excl.?vat/i, /^amount$/i, /totalAmountExcl/i];
    const CAMPAIGN = [/^campaign.?no/i, /campaign.?no/i, /^campaign$/i, /campaignNumber/i];
    const ASSIGNED = [/^assigned.?user.?id/i, /assigned.?user/i, /assignedUserId/i];
    const DOC_TYPE = [/^document.?type$/i, /^doc.?type$/i, /^documentType$/i];
    const DOC_DATE = [/^document.?date$/i, /documentDate/i, /^order.?date$/i, /orderDate/i];
    const NO_FIELD = [/^no$/i, /^number$/i];
    const candidates = allSets.filter(n => included.test(n) && !excluded.test(n));
    for (const candidate of candidates) {
        const typeName = md.entitySets[candidate];
        const fields = md.entityTypes[typeName] || [];
        if (!fields.length) continue;
        const fCustNo = bcFindField(fields, CUST_NO);
        const fCustName = bcFindField(fields, CUST_NAME);
        const fAmount = bcFindField(fields, AMOUNT);
        const fCampaign = bcFindField(fields, CAMPAIGN);
        const fDocType = bcFindField(fields, DOC_TYPE);
        const fDocDate = bcFindField(fields, DOC_DATE);
        const fAssigned = bcFindField(fields, ASSIGNED);
        const fNo = bcFindField(fields, NO_FIELD);
        if (!fCustNo || !fCustName || !fAmount) continue;
        const nameMatchesKind = docTypeNeeded.test(candidate);
        if (!nameMatchesKind && !fDocType) continue;
        return {
            entity: candidate,
            fields,
            fCustNo, fCustName, fAmount, fCampaign, fAssigned, fDocType, fDocDate, fNo,
            docTypeValue, nameMatchesKind,
        };
    }
    return null;
}
// ============================================================
// FETCHERS
// ============================================================
// Fetch with $select column trimming and automatic fallback: if the
// trimmed request errors (a selected field doesn't exist on this
// tenant's schema), log it and retry the full untrimmed URL so fast
// mode degrades to correct-but-slower instead of silently losing data.
async function bcFetchAllTrimmed(fullUrl, trimmedUrl, label) {
    if (!state.fastTrim || !trimmedUrl) return bcFetchAll(fullUrl, label);
    try {
        return await bcFetchAll(trimmedUrl, label + " (trimmed)");
    } catch (e) {
        console.warn("[" + label + "] trimmed $select failed (" + e.message + ") — retrying untrimmed");
        return bcFetchAll(fullUrl, label);
    }
}

// Whitelist of posted sales invoice numbers, sourced from the OData
// salesInvoiceHeader entity (posted-only) filtered to the same date
// window. Used to strip un-posted drafts that the v2.0 REST endpoint
// returns with status="Open" — that status is ambiguous (unposted
// drafts and posted-unpaid invoices both show it), so status alone
// can't distinguish. Cross-referencing against a posted-only table
// is the only reliable way. Verified with Vic Air: SI0000063 was in
// Wiise's Sales Invoices (draft) list, not Posted Sales Invoices.
async function fetchPostedInvoiceNumbers(fromISO, toISO) {
    try {
        const md = await bcGetODataMetadata();
        if (!md.entitySets["salesInvoiceHeader"]) {
            console.warn("[Posted invoice whitelist] salesInvoiceHeader entity not found in $metadata");
            return null;
        }
        const cname = await bcGetCompanyInternalName();
        const filter = "Posting_Date ge " + fromISO + " and Posting_Date le " + toISO;
        // No $top (it is a TOTAL cap on this tenant): the 12-month window held
        // 9,777 posted invoices on 2026-10-01 against the old 10,000 cap — one
        // more month and real invoices would have been dropped as "unposted".
        // bcFetchAll follows nextLink; the server pages at its default size.
        const url = BC_ODATA_URL + "/Company('" + encodeURIComponent(cname) + "')/salesInvoiceHeader?$select=No&$filter=" + encodeURIComponent(filter);
        const list = await bcFetchAll(url, "Posted invoice numbers");
        const set = new Set((list || []).map(r => r.No));
        console.log("[Posted invoice whitelist] " + set.size + " posted invoice numbers in window");
        return set;
    } catch (e) {
        console.warn("[Posted invoice whitelist] failed:", e.message);
        return null;
    }
}
async function fetchSalesInvoicesWithLines(fromISO, toISO) {
    const compId = await bcGetCompanyId();
    // Filter by posting date only. Status filter is unreliable — BC's
    // v2.0 REST /salesInvoices unifies posted + unposted, and "Open"
    // status appears on both. Posted-only distinction happens below via
    // the salesInvoiceHeader whitelist.
    const filter = "postingDate ge " + fromISO + " and postingDate le " + toISO;
    // No $top — BC v2.0 treats $top as a *total* cap (not a page size hint),
    // and we were silently dropping invoices beyond the first 2000. The
    // server uses its default page size and @odata.nextLink for the rest;
    // bcFetchAll already follows nextLink. Verified diagnostic: 343
    // invoices totalling ~$199K were missing under $top=2000, exactly the
    // gap between Document-mode sales and VE / PBI sales on this tenant.
    const url = BC_API_URL + "/companies(" + compId + ")/salesInvoices?$filter=" + encodeURIComponent(filter) + "&$expand=salesInvoiceLines";
    // Trimmed variant: only the header/line columns the app reads.
    // Verified against this tenant's logged sample row keys — the v2.0
    // salesInvoices header has no documentDate (it's invoiceDate), so
    // that's deliberately absent.
    const INV_HEADER_SEL = "id,number,postingDate,customerId,customerNumber,customerName,salesperson,status,discountAmount,totalAmountExcludingTax,lastModifiedDateTime";
    const INV_LINE_SEL   = "id,sequence,lineType,itemId,lineObjectNumber,description,quantity,unitPrice,amountExcludingTax,invoiceDiscountAllocation,locationId,shipmentDate";
    const trimmedUrl = BC_API_URL + "/companies(" + compId + ")/salesInvoices?$filter=" + encodeURIComponent(filter)
        + "&$select=" + INV_HEADER_SEL + "&$expand=" + encodeURIComponent("salesInvoiceLines($select=" + INV_LINE_SEL + ")");
    const [rowsRaw, postedSet] = await Promise.all([
        bcFetchAllTrimmed(url, trimmedUrl, "Sales invoices + lines"),
        fetchPostedInvoiceNumbers(fromISO, toISO),
    ]);
    let rows = rowsRaw || [];
    if (rows.length) {
        const beforeStatus = {};
        for (const r of rows) beforeStatus[r.status || "(none)"] = (beforeStatus[r.status || "(none)"] || 0) + 1;
        console.log("[Sales Invoices] v2.0 REST returned " + rows.length + " · status distribution: " + JSON.stringify(beforeStatus));
    }
    // Apply posted-only whitelist. If the OData fetch failed (postedSet
    // is null), fall back to the raw list rather than break the tool —
    // the number will be slightly high on that tenant, matching the
    // old behaviour, and the console warning will tell us why.
    if (postedSet) {
        const before = rows.length;
        const kept = [];
        const dropped = [];
        for (const r of rows) {
            if (postedSet.has(r.number)) kept.push(r);
            else dropped.push(r);
        }
        rows = kept;
        console.log("[Sales Invoices] posted-whitelist kept " + rows.length + "/" + before + " · dropped " + dropped.length + " unposted");
        if (dropped.length) {
            for (const r of dropped.slice(0, 5)) {
                console.log("  dropped: " + r.number + " · status=" + r.status + " · total=" + r.totalAmountExcludingTax + " · customer=" + (r.customerName || ""));
            }
        }
    }
    return rows;
}
async function fetchItemLedgerSales(fromISO, toISO) {
    const compId = await bcGetCompanyId();
    // entryType="Sale" covers sales invoices (and credits — caller can filter by documentType if needed).
    // No $top — BC treats it as a TOTAL cap, not a page size, and a
    // 12-month window exceeds 10,000 Sale ILE rows; the old $top=10000
    // silently truncated the cost-fallback data. bcFetchAll pages via
    // @odata.nextLink so the full set arrives regardless of size.
    const filter = "entryType eq 'Sale' and postingDate ge " + fromISO + " and postingDate le " + toISO;
    const url = BC_API_URL + "/companies(" + compId + ")/itemLedgerEntries?$filter=" + encodeURIComponent(filter);
    return bcFetchAll(url, "Item ledger (cost)");
}
async function fetchLocations() {
    const compId = await bcGetCompanyId();
    try {
        const data = await bcFetch(BC_API_URL + "/companies(" + compId + ")/locations"); // no $top anywhere in this file — BC treats it as a total cap
        return data.value || [];
    } catch (e) { console.warn("Could not fetch locations:", e.message); return []; }
}
async function fetchItems() {
    const compId = await bcGetCompanyId();
    try {
        const data = await bcFetchAll(BC_API_URL + "/companies(" + compId + ")/items?$select=id,number,displayName,itemCategoryCode", "Items"); // no $top: 7,672 items on 2026-10-01, cap was 10k
        return data;
    } catch (e) { console.warn("Could not fetch items:", e.message); return []; }
}
async function fetchInvoiceLines(invoiceId) {
    const compId = await bcGetCompanyId();
    return bcFetchAll(BC_API_URL + "/companies(" + compId + ")/salesInvoices(" + invoiceId + ")/salesInvoiceLines", "Invoice lines");
}
async function fetchCustomers() {
    const compId = await bcGetCompanyId();
    try {
        return await bcFetchAll(BC_API_URL + "/companies(" + compId + ")/customers", "Customers"); // no $top (total cap)
    } catch (e) { console.warn("Customers fetch failed:", e.message); return []; }
}
async function fetchSalesCreditMemos(fromISO, toISO) {
    const compId = await bcGetCompanyId();
    const filter = "postingDate ge " + fromISO + " and postingDate le " + toISO;
    const url = BC_API_URL + "/companies(" + compId + ")/salesCreditMemos?$filter=" + encodeURIComponent(filter) + "&$expand=salesCreditMemoLines";
    const CM_HEADER_SEL = "id,number,postingDate,customerNumber,customerName,salesperson,totalAmountExcludingTax,lastModifiedDateTime";
    const CM_LINE_SEL   = "id,sequence,lineType,itemId,lineObjectNumber,description,quantity,unitPrice,amountExcludingTax,invoiceDiscountAllocation,locationId,shipmentDate";
    const trimmedUrl = BC_API_URL + "/companies(" + compId + ")/salesCreditMemos?$filter=" + encodeURIComponent(filter)
        + "&$select=" + CM_HEADER_SEL + "&$expand=" + encodeURIComponent("salesCreditMemoLines($select=" + CM_LINE_SEL + ")");
    try { return await bcFetchAllTrimmed(url, trimmedUrl, "Sales credit memos"); }
    catch (e) { console.warn("Credit memos fetch failed:", e.message); return []; }
}
// Fetches OPEN sales orders (no date filter). Pipeline KPIs need
// current-state visibility — orders that haven't shipped or invoiced
// yet — which is independent of the toolbar date range.
async function fetchSalesOrders() {
    const compId = await bcGetCompanyId();
    const url = BC_API_URL + "/companies(" + compId + ")/salesOrders?$expand=salesOrderLines";
    try { return await bcFetchAll(url, "Sales orders"); }
    catch (e) { console.warn("Sales orders fetch failed:", e.message); return []; }
}

// Fetch sales-order lines via OData v4 so we get BC's server-computed
// outstanding-amount fields (outstandingAmountLcy, shippedNotInvLcyNoVat,
// etc.) — these are the same fields PBI reads from the Sales Line table.
// The v2.0 REST salesOrderLines endpoint exposes only the raw amounts
// and we have to recompute outstanding ourselves, which never matches
// PBI to the dollar because BC applies tax and discount overrides
// per-line that we can't replicate from v2.0 fields alone.
async function fetchSalesOrderOutstandingLines() {
    let headerInfo, linesInfo;
    try { headerInfo = await bcDiscoverSalesEntity("blanket"); }
    catch (e) { console.warn("[SO outstanding] header discovery failed:", e.message); return []; }
    if (!headerInfo) {
        console.warn("[SO outstanding] no OData salesDocuments entity published");
        return [];
    }
    try { linesInfo = await bcDiscoverSalesLinesEntity(headerInfo.entity); }
    catch (e) { console.warn("[SO outstanding] lines discovery failed:", e.message); return []; }
    if (!linesInfo) {
        console.warn("[SO outstanding] no OData salesDocumentLines entity published");
        return [];
    }
    // Field finders for the BC-computed outstanding columns. They live
    // on the Sales Line table and are exposed by the standard
    // SalesDocumentLines OData web service.
    const fields = linesInfo.fields || [];
    const fDocType    = bcFindField(fields, [/^document.?type$/i, /documentType/i]);
    const fDocNo      = linesInfo.fDocNo;
    const fQty        = linesInfo.fQty;
    const fOutQty     = bcFindField(fields, [/^outstanding.?quantity$/i, /outstandingQuantity/i]);
    const fOutAmtLcy  = bcFindField(fields, [/^outstanding.?amount.?lcy$/i, /outstandingAmountLcy/i]);
    const fOutAmt     = bcFindField(fields, [/^outstanding.?amount$/i, /outstandingAmount/i]);
    // shippedNotInvoicedLcy (with VAT) is what PBI's tile reads —
    // confirmed against this tenant: PBI shows $42.92K and that field
    // sums to $42,919.32. The shippedNotInvLcyNoVat variant gives the
    // ex-VAT value (~$39K) which doesn't match PBI. Order the regexes
    // so the with-VAT field wins; fall through to no-VAT only if the
    // with-VAT one isn't published.
    const fSniLcy     = bcFindField(fields, [/^shipped.?not.?invoiced.?lcy$/i, /shippedNotInvoicedLcy/i, /^shipped.?not.?inv.?lcy.?no.?vat$/i, /shippedNotInvLcyNoVat/i]);
    // Ex-GST variant for tenants that publish "shippedNotInvLcyNoVat"
    // alongside the with-VAT field. We keep both around — pipeline UI
    // can pick whichever matches the rest of the tool's tax treatment.
    const fSniLcyNoVat = bcFindField(fields, [/^shipped.?not.?inv.?lcy.?no.?vat$/i, /shippedNotInvLcyNoVat/i]);
    // BC's line "Amount" = ex-VAT line total in document currency. On a
    // single-currency tenant (AUD = LCY) this is the right ex-GST basis
    // for computing outstanding amount as (amount × outstandingQty / qty).
    const fLineAmount  = bcFindField(fields, [/^amount$/i, /^line.?amount$/i]);
    const fSniQty     = bcFindField(fields, [/^qty.?shipped.?not.?invoiced$/i, /qtyShippedNotInvoiced/i]);
    // Per-line shipmentDate — PBI's "Outstanding Sales Orders" tile
    // filters on the line's Shipment Date (each line carries its own
    // expected ship date, independent of the order header). Confirmed
    // against this tenant: L.shipmentDate gives 169 orders ≈ PBI's
    // 171 within data-freshness drift. We also try planned / requested
    // delivery variants in case a tenant uses those instead.
    const fShipDate   = bcFindField(fields, [/^shipment.?date$/i, /shipmentDate/i, /^planned.?shipment.?date$/i, /^planned.?delivery.?date$/i]);
    if (!fDocType || !fDocNo) {
        console.warn("[SO outstanding] missing documentType / documentNumber fields on lines entity");
        return [];
    }
    const coName = encodeURIComponent(await bcGetCompanyInternalName());
    const filter = fDocType + " eq 'Order'";
    // No $top (total cap on this tenant — see fetchValueEntries); nextLink pages the rest.
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + linesInfo.entity +
        "?$filter=" + encodeURIComponent(filter);
    let rows = [];
    try { rows = await bcFetchAll(url, "Sales order lines (OData outstanding)"); }
    catch (e) { console.warn("[SO outstanding] OData fetch failed:", e.message); return []; }
    console.log("[SO outstanding] " + linesInfo.entity + " (documentType=Order) → " + rows.length + " rows · outAmtLcy=" + fOutAmtLcy + " · lineAmount=" + fLineAmount + " · sniLcy=" + fSniLcy + " · sniLcyNoVat=" + fSniLcyNoVat + " · shipDate=" + fShipDate);
    return rows.map(r => ({
        documentNumber: fDocNo ? r[fDocNo] : "",
        quantity:       fQty ? num(r[fQty]) : 0,
        outstandingQuantity:  fOutQty ? num(r[fOutQty]) : 0,
        qtyShippedNotInvoiced: fSniQty ? num(r[fSniQty]) : 0,
        outstandingAmountLcy:  fOutAmtLcy ? num(r[fOutAmtLcy]) : (fOutAmt ? num(r[fOutAmt]) : 0),
        lineAmount:            fLineAmount ? num(r[fLineAmount]) : 0,
        shippedNotInvLcy:      fSniLcy ? num(r[fSniLcy]) : 0,
        shippedNotInvLcyExVat: fSniLcyNoVat ? num(r[fSniLcyNoVat]) : 0,
        shipmentDate:          fShipDate ? (r[fShipDate] || "").toString().slice(0, 10) : "",
    }));
}
async function fetchSalesShipments(fromISO, toISO) {
    const compId = await bcGetCompanyId();
    const filter = "postingDate ge " + fromISO + " and postingDate le " + toISO;
    const url = BC_API_URL + "/companies(" + compId + ")/salesShipments?$filter=" + encodeURIComponent(filter) + "&$expand=salesShipmentLines";
    // Conservative line list — shipment lines carry no amount fields the
    // app needs (cost bridging only uses item + qty; drill leaves show
    // description; sales contribution is zeroed for shipments anyway).
    const SH_HEADER_SEL = "id,number,postingDate,customerNumber,customerName,salesperson,lastModifiedDateTime";
    const SH_LINE_SEL   = "id,sequence,lineType,itemId,lineObjectNumber,description,quantity,locationId";
    const trimmedUrl = BC_API_URL + "/companies(" + compId + ")/salesShipments?$filter=" + encodeURIComponent(filter)
        + "&$select=" + SH_HEADER_SEL + "&$expand=" + encodeURIComponent("salesShipmentLines($select=" + SH_LINE_SEL + ")");
    try { return await bcFetchAllTrimmed(url, trimmedUrl, "Sales shipments"); }
    catch (e) { console.warn("Shipments fetch failed:", e.message); return []; }
}
async function fetchSalesReturnReceipts(fromISO, toISO) {
    const compId = await bcGetCompanyId();
    const filter = "postingDate ge " + fromISO + " and postingDate le " + toISO;
    const tries = [
        "/salesReturnReceipts?$filter=" + encodeURIComponent(filter) + "&$expand=salesReturnReceiptLines",
        "/salesReturnReceipts?$filter=" + encodeURIComponent(filter),
    ];
    for (const path of tries) {
        try { return await bcFetchAll(BC_API_URL + "/companies(" + compId + ")" + path, "Sales return receipts"); }
        catch (e) { console.warn("Return receipts try failed:", e.message); }
    }
    return [];
}
async function fetchSalesQuotes(fromISO, toISO) {
    const compId = await bcGetCompanyId();
    // Quotes use documentDate (not postingDate)
    const filter = "documentDate ge " + fromISO + " and documentDate le " + toISO;
    const url = BC_API_URL + "/companies(" + compId + ")/salesQuotes?$filter=" + encodeURIComponent(filter) + "&$expand=salesQuoteLines";
    try { return await bcFetchAll(url, "Sales quotes"); }
    catch (e) { console.warn("Quotes fetch failed:", e.message); return []; }
}

// Blanket Sales Orders aren't a documented v2.0 REST endpoint. Discover
// the actual published OData entity via $metadata (same pattern as the
// calendar-viewer tool) and pull from there.
async function fetchBlanketSalesOrders(fromISO, toISO) {
    state.blanketOrdersDiagnostic = "";
    state.blanketOrdersSource = "";
    let info;
    try { info = await bcDiscoverSalesEntity("blanket"); }
    catch (e) {
        console.warn("[Blanket SO] metadata discovery failed:", e.message);
        state.blanketOrdersDiagnostic = "OData $metadata fetch failed: " + e.message;
        return [];
    }
    if (!info) {
        state.blanketOrdersDiagnostic = "No published OData entity found for Blanket Sales Orders. Ask BC admin to publish page 9305 (BlanketSalesOrder) as a Web Service.";
        console.warn("[Blanket SO] no candidate entity found in $metadata");
        return [];
    }
    const coName = encodeURIComponent(await bcGetCompanyInternalName());
    const filters = [];
    if (info.fDocDate) filters.push(info.fDocDate + " ge " + fromISO + " and " + info.fDocDate + " le " + toISO);
    if (!info.nameMatchesKind && info.fDocType) filters.push(info.fDocType + " eq '" + info.docTypeValue + "'");
    const params = []; // no $top (total cap); bcFetchAll pages via nextLink
    if (filters.length) params.push("$filter=" + encodeURIComponent(filters.join(" and ")));
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity + (params.length ? "?" + params.join("&") : "");
    try {
        const rows = await bcFetchAll(url, "Blanket sales orders (" + info.entity + ")");
        console.log("[Blanket SO] " + info.entity + " → " + rows.length + " rows");
        if (rows.length) console.log("[Blanket SO] sample row keys:", Object.keys(rows[0]).join(", "));
        // Extra field hints for the normalize step
        const fields = info.fields || [];
        const fStatus = bcFindField(fields, [/^status$/i]);
        const fSalesperson = bcFindField(fields, [/^salesperson.?code/i, /salespersonCode/i, /^salesperson$/i]);
        const fOrderDate = bcFindField(fields, [/^order.?date$/i, /orderDate/i]);
        // Map OData → v2-shape so the renderer code doesn't need to change
        const normalized = rows.map(r => ({
            number:         info.fNo ? r[info.fNo] : (r.No || r.No_ || r.Number || ""),
            customerNumber: info.fCustNo ? r[info.fCustNo] : "",
            customerName:   info.fCustName ? r[info.fCustName] : "",
            totalAmountExcludingTax: info.fAmount ? r[info.fAmount] : 0,
            campaignNumber: info.fCampaign ? r[info.fCampaign] : "",
            assignedUserID: info.fAssigned ? r[info.fAssigned] : "",
            status:         fStatus ? r[fStatus] : "",
            salesperson:    fSalesperson ? r[fSalesperson] : "",
            documentDate:   info.fDocDate ? r[info.fDocDate] : "",
            orderDate:      fOrderDate ? r[fOrderDate] : "",
            blanketSalesOrderLines: [],
            _raw: r,
        }));
        state.blanketOrdersSource = "OData: " + info.entity;
        state.blanketOrdersFieldMap = info;

        // ----- Lines: discover entity, fetch all, group by Document_No -----
        let linesNote = "Lines not loaded.";
        try {
            const linesInfo = await bcDiscoverSalesLinesEntity(info.entity);
            if (linesInfo) {
                state.blanketOrdersLinesFieldMap = linesInfo;
                const docNos = normalized.map(o => o.number).filter(Boolean);
                // Prefer a constant-size documentType filter. The old
                // approach OR-ed every header's document number into one
                // URL, which worked at 62 orders but returned HTTP 414
                // (URI Too Long) at 219 — silently leaving every order
                // line-less. The join below drops any extra doc numbers,
                // so the wider type filter is safe.
                const fLineDocType = bcFindField(linesInfo.fields || [], [/^documentType$/i, /^Document_Type$/i]);
                let lineRows;
                if (fLineDocType && info.docTypeValue) {
                    const linesUrl = BC_ODATA_URL + "/Company('" + coName + "')/" + linesInfo.entity
                        + "?$filter=" + encodeURIComponent(fLineDocType + " eq '" + info.docTypeValue + "'");
                    lineRows = await bcFetchAll(linesUrl, "Blanket order lines (" + linesInfo.entity + ")");
                } else {
                    // Fallback: chunked OR-filters, 40 document numbers per
                    // request, so the URL stays far below server limits.
                    lineRows = [];
                    for (let i = 0; i < docNos.length; i += 40) {
                        const chunk = docNos.slice(i, i + 40);
                        const inClause = chunk.map(n => linesInfo.fDocNo + " eq '" + String(n).replace(/'/g, "''") + "'").join(" or ");
                        const linesUrl = BC_ODATA_URL + "/Company('" + coName + "')/" + linesInfo.entity
                            + "?$filter=" + encodeURIComponent(inClause);
                        lineRows = lineRows.concat(await bcFetchAll(linesUrl, "Blanket order lines " + Math.min(i + 40, docNos.length) + "/" + docNos.length));
                    }
                }
                console.log("[Blanket SO lines] " + linesInfo.entity + " → " + lineRows.length + " line rows");
                if (lineRows.length) {
                    console.log("[Blanket SO lines] sample line keys:", Object.keys(lineRows[0]).join(", "));
                    console.log("[Blanket SO lines] sample line row:", lineRows[0]);
                    state.blanketOrdersLinesSampleKeys = Object.keys(lineRows[0]);
                    state.blanketOrdersLinesSample = lineRows[0];
                }

                // ----- Sample-row fallback: item-No / description discovery
                // sometimes picks a field that's null for every line. Re-pick
                // by scanning the first 50 rows for a column whose values
                // actually look like item codes / descriptions. -----
                const sample = lineRows.slice(0, 50);
                const fieldHasValues = (f) => sample.some(r => r[f] != null && String(r[f]).trim() !== "");
                // If discovery picked a wrong fItemNo (e.g. sellToCustomerNumber)
                // OR didn't pick one, re-pick from the actual line keys. Prefer
                // exact "number" / "No" matches first (BC v2 / NAV conventions),
                // then fall back to a value-based scan. Exclude any customer /
                // shipping / billing field so we don't get the customer number.
                const exclude = /^(?:sellTo|billTo|payTo|shipTo|customer|vendor|bill|ship|salesperson|location|currency|invoice|shipping|bin|variant|posting|gen|vat|tax|prepayment|prepmt|return|reserved|original|special|appl|item.?reference|item.?category|nonstock|whse|deferral|attached|attach|job|work|shortcut|dimension|deprec|fa|ic|exit|area|entry|response|requested|promised|planned)/i;
                const looksLikeItemNo = (k) =>
                    /^(?:number|no|no_|item.?no_?|itemNumber|lineObjectNumber)$/i.test(k) && !exclude.test(k);
                if (lineRows.length && (!linesInfo.fItemNo || !looksLikeItemNo(linesInfo.fItemNo) || !fieldHasValues(linesInfo.fItemNo))) {
                    const allKeys = Object.keys(sample[0] || {});
                    // First priority: exact name match
                    let pick = allKeys.find(k => looksLikeItemNo(k) && fieldHasValues(k));
                    // Second priority: contains "item" + "no/number" but not excluded
                    if (!pick) pick = allKeys.find(k =>
                        /item.*(?:no|number)/i.test(k) && !exclude.test(k) && fieldHasValues(k));
                    if (pick) {
                        const before = linesInfo.fItemNo;
                        linesInfo.fItemNo = pick;
                        console.log("[Blanket SO lines] fItemNo fallback: '" + (before || "—") + "' → '" + pick + "'");
                    }
                }
                if (lineRows.length && (!linesInfo.fDesc || !fieldHasValues(linesInfo.fDesc))) {
                    const allKeys = Object.keys(sample[0] || {});
                    const descCandidate = allKeys.find(k => /desc/i.test(k) && fieldHasValues(k));
                    if (descCandidate) {
                        const before = linesInfo.fDesc;
                        linesInfo.fDesc = descCandidate;
                        console.log("[Blanket SO lines] fDesc fallback: '" + (before || "—") + "' → '" + descCandidate + "'");
                    }
                }
                // Group by document number
                const byDoc = new Map();
                for (const lr of lineRows) {
                    const docNo = String(lr[linesInfo.fDocNo] || "");
                    if (!docNo) continue;
                    const qty = num(lr[linesInfo.fQty]);
                    const amt = num(lr[linesInfo.fAmount]);
                    const disc = linesInfo.fDiscAlloc ? num(lr[linesInfo.fDiscAlloc]) : 0;
                    const typeVal = linesInfo.fType ? lr[linesInfo.fType] : "Item";
                    const itemNo = linesInfo.fItemNo ? lr[linesInfo.fItemNo] : "";
                    const locCode = linesInfo.fLocation ? lr[linesInfo.fLocation] : "";
                    const v2Line = {
                        lineObjectNumber: itemNo || "",
                        itemId:           itemNo || "",
                        quantity:         qty,
                        unitPrice:        linesInfo.fUnitPrice ? num(lr[linesInfo.fUnitPrice]) : 0,
                        amountExcludingTax:        amt,
                        invoiceDiscountAllocation: disc,
                        description:      linesInfo.fDesc ? (lr[linesInfo.fDesc] || "") : "",
                        locationId:       locCode || "",   // code, not GUID — locationCode() will pass through unknown codes
                        lineType:         (typeVal && /^item$/i.test(typeVal)) ? "Item" : (typeVal || "Item"),
                    };
                    if (!byDoc.has(docNo)) byDoc.set(docNo, []);
                    byDoc.get(docNo).push(v2Line);
                }
                let attached = 0;
                for (const o of normalized) {
                    const lines = byDoc.get(String(o.number));
                    if (lines && lines.length) { o.blanketSalesOrderLines = lines; attached++; }
                }
                linesNote = lineRows.length + " line rows · attached to " + attached + "/" + normalized.length + " orders.";
            } else {
                linesNote = "No matching lines entity found in $metadata (looked for entities matching '" + info.entity + "' + 'line').";
                console.warn("[Blanket SO lines] no candidate lines entity");
            }
        } catch (le) {
            console.warn("[Blanket SO lines] fetch failed:", le.message);
            linesNote = "Lines fetch failed: " + le.message;
        }
        state.blanketOrdersDiagnostic = "Source: OData " + info.entity + " · " + rows.length + " orders. " + linesNote;
        return normalized;
    } catch (e) {
        console.warn("[Blanket SO] OData fetch failed:", e.message);
        state.blanketOrdersDiagnostic = "OData " + info.entity + " fetch failed: " + e.message;
        return [];
    }
}

// BC v2.0 REST doesn't expose `Your Reference` or `Quote No.` on
// salesOrders / salesInvoices. Discover the OData v4 sales-header
// entities via $metadata, pull (No, Your_Reference, Quote_No) for every
// document, and build two lookup maps consumed by getYourReference /
// getOriginatingQuoteNumber.
async function fetchResidentialDocLookup() {
    state.docYourReference = new Map();
    state.docQuoteLink     = new Map();
    state.residentialDiagnostic = "";
    let md;
    try { md = await bcGetODataMetadata(); }
    catch (e) {
        state.residentialDiagnostic = "OData $metadata fetch failed: " + e.message;
        console.warn("[Residential lookup] metadata fetch failed:", e.message);
        return;
    }
    // Field-name patterns. BC uses underscores in OData v4 entity property
    // names ("Your_Reference", "Quote_No") but some tenants flatten them.
    const YOUR_REF = [/^your.?reference$/i, /yourReference/i];
    const QUOTE_NO = [/^quote.?no$/i, /quoteNo$/i, /quoteNumber$/i];
    const DOC_NO   = [/^no$/i, /^number$/i, /^document.?no$/i];
    const candidates = [];
    // Track which (typeName, field-shape) combinations we've already
    // accepted so we don't fetch the same dataset twice (Wiise tenants
    // commonly publish `salesDocuments` + `workflowSalesDocuments`,
    // `SalesOrder` + `Sales_Order_Excel` etc. with identical content).
    const seenShapes = new Set();
    for (const [setName, typeName] of Object.entries(md.entitySets)) {
        // Sales-only — explicitly drop purchase entities even when their
        // name contains "document".
        if (!/sales|invoice|order|quote|document/i.test(setName)) continue;
        if (/purchase|vendor|payable|requisition/i.test(setName)) continue;
        if (/line|footer|comment|child|attachment|prepayment|archive/i.test(setName)) continue;
        const fields = md.entityTypes[typeName] || [];
        if (!fields.length) continue;
        const fYourRef = bcFindField(fields, YOUR_REF);
        const fNo      = bcFindField(fields, DOC_NO);
        if (!fYourRef || !fNo) continue;
        const fQuote   = bcFindField(fields, QUOTE_NO);
        // Same EntityType (i.e. same underlying BC page/table) →
        // skip subsequent EntitySets pointing at it.
        const shapeKey = typeName + "|" + fNo + "|" + fYourRef + "|" + (fQuote || "");
        if (seenShapes.has(shapeKey)) {
            console.log("[Residential lookup] skipping duplicate-shape entity " + setName + " (already covered by " + typeName + ")");
            continue;
        }
        seenShapes.add(shapeKey);
        candidates.push({ entity: setName, fields, fYourRef, fNo, fQuote });
    }
    if (!candidates.length) {
        state.residentialDiagnostic =
            "No OData entity exposes 'Your Reference'. Ask BC admin to publish Sales Header (table 36) and Posted Sales Invoice Header (table 112) as Web Services.";
        console.warn("[Residential lookup] no candidates in $metadata");
        return;
    }
    // Dump the full entity-set list grouped by intent so we can spot any
    // Wiise-custom posted-invoice entity that our regex above missed.
    // Helpful only on first inspection — costs nothing to print.
    const allSets = Object.keys(md.entitySets);
    const postedLike = allSets.filter(n => /post|sinv|invoiceheader|crmemo|cr_memo|^pInv|pstd/i.test(n));
    const yourRefBearing = [];
    for (const [setName, typeName] of Object.entries(md.entitySets)) {
        const f = md.entityTypes[typeName] || [];
        if (f.some(x => /your.?reference/i.test(x))) yourRefBearing.push(setName);
    }
    // Print as joined strings — Chrome collapses array values to "Array(N)"
    // in the console which hides exactly the names we need to inspect.
    console.log("[Residential lookup] candidate entities (after dedupe):\n  " + candidates.map(c => c.entity).join("\n  "));
    console.log("[Residential lookup] ALL entities with a Your_Reference column:\n  " + yourRefBearing.join("\n  "));
    console.log("[Residential lookup] entity names containing post/sinv/invoiceheader/crmemo/pstd:\n  " + postedLike.join("\n  "));
    // Also dump every entity name that mentions "archive" or "history" or
    // "quote" — quote-archive support is the Plan-B we discussed.
    const archiveLike = allSets.filter(n => /archive|history|quoteHeader|salesQuote/i.test(n));
    console.log("[Residential lookup] entity names containing archive/history/quoteHeader/salesQuote:\n  " + archiveLike.join("\n  "));
    console.log("[Residential lookup] total entity sets in $metadata:", allSets.length);
    const coName = encodeURIComponent(await bcGetCompanyInternalName());
    let totalHits = 0;
    for (const c of candidates) {
        const selectCols = [c.fNo, c.fYourRef, c.fQuote].filter(Boolean).join(",");
        const url = BC_ODATA_URL + "/Company('" + coName + "')/" + c.entity +
            "?$select=" + selectCols;
        try {
            const rows = await bcFetchAll(url, "Residential lookup (" + c.entity + ")");
            let entityHits = 0;
            for (const r of rows) {
                const docNo  = r[c.fNo];
                const yourRef = c.fYourRef ? (r[c.fYourRef] || "") : "";
                const quoteNo = c.fQuote ? (r[c.fQuote] || "") : "";
                if (!docNo) continue;
                if (yourRef) {
                    state.docYourReference.set(docNo.toString(), yourRef.toString());
                    entityHits++;
                }
                if (quoteNo) {
                    state.docQuoteLink.set(docNo.toString(), quoteNo.toString());
                }
            }
            totalHits += entityHits;
            console.log("[Residential lookup] " + c.entity + " → " + rows.length + " rows · " + entityHits + " with Your_Reference");
        } catch (e) {
            console.warn("[Residential lookup] " + c.entity + " fetch failed:", e.message);
        }
    }
    if (!totalHits) {
        state.residentialDiagnostic =
            "Found " + candidates.length + " entit" + (candidates.length === 1 ? "y" : "ies") + " with 'Your Reference' but no rows returned values. " +
            "Ask BC admin to confirm the field is being populated.";
    }
    console.log("[Residential lookup] total docs with Your_Reference: " + state.docYourReference.size +
        " · Quote_No links: " + state.docQuoteLink.size);
}

// Sales Quote Archive (BC table 5107, page 5152). Once published as a
// Web Service, the OData v4 metadata exposes an entity carrying every
// quote that's been converted — preserving Quote No., Your Reference,
// Assigned User ID, original quoted amount, and customer info. This is
// the cleanest source for "this quote was Won" because BC writes the
// archive row at the moment of conversion and never modifies it again.
async function fetchSalesQuoteArchive(fromISO, toISO) {
    state.quoteArchive = new Map();   // quoteNumber -> normalised row
    state.quoteArchiveDiagnostic = "";
    let md;
    try { md = await bcGetODataMetadata(); }
    catch (e) {
        state.quoteArchiveDiagnostic = "OData $metadata fetch failed: " + e.message;
        return;
    }
    // Match common publishing names: salesQuoteArchive, SalesQuoteArchive,
    // archivedSalesQuotes, salesHeaderArchive, etc.
    const ARCHIVE_NAME = /quote.{0,3}archive|archive.{0,3}quote|salesheader.{0,3}archive/i;
    const NO_FIELD     = [/^no$/i, /^number$/i];
    const YOUR_REF     = [/^your.?reference$/i, /yourReference/i];
    const ASSIGNED     = [/^assigned.?user.?id/i, /assigned.?user/i, /assignedUserId/i];
    const SP_CODE      = [/^salesperson.?code$/i, /^salesperson$/i, /salespersonCode/i];
    const CUST_NO      = [/^sell.?to.?customer.?no/i, /^customer.?no/i, /customerNumber/i, /sellToCustomerNumber/i];
    const CUST_NAME    = [/sell.?to.?customer.?name/i, /^customer.?name/i, /customerName/i, /sellToCustomerName/i];
    const DOC_DATE     = [/^document.?date/i, /documentDate/i, /^order.?date/i, /orderDate/i];
    const AMOUNT       = [/^amount.*excl/i, /^total.*amount.*excl/i, /^total.*excl/i, /totalAmountExcl/i, /^amount$/i];
    const CAMPAIGN     = [/^campaign.?no/i, /campaignNumber/i, /campaign.?no/i];
    const DOC_TYPE     = [/^document.?type/i, /documentType/i];
    let info = null;
    const archiveSkipReasons = [];
    for (const [setName, typeName] of Object.entries(md.entitySets)) {
        if (!ARCHIVE_NAME.test(setName)) continue;
        const fields = md.entityTypes[typeName] || [];
        // Log every candidate so we can see why one was skipped — typeName
        // mismatches and missing No fields are the common culprits.
        console.log("[Quote archive] candidate '" + setName + "' (typeName=" + typeName + ", " + fields.length + " fields)");
        if (fields.length < 20) console.log("[Quote archive]   fields:", fields.join(", "));
        if (!fields.length) { archiveSkipReasons.push(setName + ": typeName '" + typeName + "' had no fields in $metadata"); continue; }
        const fNo = bcFindField(fields, NO_FIELD);
        if (!fNo) { archiveSkipReasons.push(setName + ": no field matched No/Number — saw: " + fields.slice(0, 30).join(", ")); continue; }
        info = {
            entity:    setName,
            fNo,
            fYourRef:  bcFindField(fields, YOUR_REF),
            fAssigned: bcFindField(fields, ASSIGNED),
            fSp:       bcFindField(fields, SP_CODE),
            fCustNo:   bcFindField(fields, CUST_NO),
            fCustName: bcFindField(fields, CUST_NAME),
            fDocDate:  bcFindField(fields, DOC_DATE),
            fAmount:   bcFindField(fields, AMOUNT),
            fCampaign: bcFindField(fields, CAMPAIGN),
            fDocType:  bcFindField(fields, DOC_TYPE),
        };
        break;
    }
    if (!info) {
        if (archiveSkipReasons.length) {
            state.quoteArchiveDiagnostic =
                "Sales Quote Archive entity found in metadata but skipped: " + archiveSkipReasons.join(" · ");
            console.warn("[Quote archive] entity skipped:", archiveSkipReasons);
        } else {
            state.quoteArchiveDiagnostic =
                "Sales Quote Archive entity not found in OData metadata. Publish Page 5152 as a Web Service in BC.";
            console.warn("[Quote archive] no entity matched ARCHIVE_NAME regex");
        }
        return;
    }
    const coName = encodeURIComponent(await bcGetCompanyInternalName());
    // Filter by document date so we only pull archived quotes from the
    // current window — keeps payload sane on tenants with years of history.
    let filter = "";
    if (info.fDocDate && fromISO && toISO) {
        filter = info.fDocDate + " ge " + fromISO + " and " + info.fDocDate + " le " + toISO;
    }
    // The archive holds rows for every documentType (Quote, Order, Invoice, ...);
    // Quote-only filter cuts the noise.
    if (info.fDocType) {
        filter = (filter ? "(" + filter + ") and " : "") + info.fDocType + " eq 'Quote'";
    }
    const params = []; // no $top (total cap)
    if (filter) params.push("$filter=" + encodeURIComponent(filter));
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity + (params.length ? "?" + params.join("&") : "");
    let rows = [];
    try { rows = await bcFetchAll(url, "Sales Quote Archive (" + info.entity + ")"); }
    catch (e) {
        state.quoteArchiveDiagnostic = "Archive fetch failed: " + e.message;
        console.warn("[Quote archive] fetch failed:", e.message);
        return;
    }
    if (rows.length) {
        console.log("[Quote archive] sample row keys:", Object.keys(rows[0]).join(", "));
    }
    for (const r of rows) {
        const quoteNo = r[info.fNo]; if (!quoteNo) continue;
        const yourRef = info.fYourRef  ? (r[info.fYourRef] || "")  : "";
        state.quoteArchive.set(quoteNo.toString(), {
            quoteNumber:    quoteNo.toString(),
            yourReference:  yourRef.toString(),
            assignedUserId: info.fAssigned ? (r[info.fAssigned] || "") : "",
            salesperson:    info.fSp       ? (r[info.fSp]       || "") : "",
            customerNumber: info.fCustNo   ? (r[info.fCustNo]   || "") : "",
            customerName:   info.fCustName ? (r[info.fCustName] || "") : "",
            documentDate:   info.fDocDate  ? (r[info.fDocDate]  || "") : "",
            amount:         info.fAmount   ? num(r[info.fAmount])      : 0,
            campaignNumber: info.fCampaign ? (r[info.fCampaign] || "") : "",
        });
    }
    console.log("[Quote archive] " + info.entity + " → " + rows.length + " rows · " + state.quoteArchive.size + " unique quote numbers");
}

// BC v2.0 REST API doesn't expose Campaign No or Assigned User ID on
// salesQuotes. Discover the actual published OData entity via $metadata
// and pull Campaign No / Assigned User ID from there.
async function fetchSalesQuoteExtras(fromISO, toISO) {
    state.quoteExtrasDiagnostic = "";
    let info;
    try { info = await bcDiscoverSalesEntity("quote"); }
    catch (e) {
        console.warn("[Quote extras] metadata discovery failed:", e.message);
        state.quoteExtrasDiagnostic = "OData $metadata fetch failed: " + e.message;
        return [];
    }
    if (!info) {
        state.quoteExtrasDiagnostic = "No published OData entity found for Sales Quotes. Ask BC admin to publish page 9300 (SalesQuotes) as a Web Service.";
        console.warn("[Quote extras] no candidate entity found in $metadata");
        return [];
    }
    const coName = encodeURIComponent(await bcGetCompanyInternalName());
    const filters = [];
    if (info.fDocDate) filters.push(info.fDocDate + " ge " + fromISO + " and " + info.fDocDate + " le " + toISO);
    if (!info.nameMatchesKind && info.fDocType) filters.push(info.fDocType + " eq '" + info.docTypeValue + "'");
    const selectFields = [info.fNo, info.fCustNo, info.fCustName, info.fAmount, info.fCampaign, info.fAssigned, info.fDocDate].filter(Boolean);
    const params = []; // no $top (total cap)
    if (selectFields.length) params.push("$select=" + selectFields.join(","));
    if (filters.length) params.push("$filter=" + encodeURIComponent(filters.join(" and ")));
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity + (params.length ? "?" + params.join("&") : "");
    try {
        const rows = await bcFetchAll(url, "Sales quote extras (" + info.entity + ")");
        console.log("[Quote extras] " + info.entity + " → " + rows.length + " rows");
        if (rows.length) console.log("[Quote extras] sample row keys:", Object.keys(rows[0]).join(", "));
        state.quoteExtrasSource = "OData: " + info.entity;
        state.quoteExtrasDiagnostic = "Source: OData " + info.entity + " (" + rows.length + " rows). Campaign field: " + (info.fCampaign || "—") + ", Assigned User: " + (info.fAssigned || "—") + ".";
        state.quoteExtrasFieldMap = info;
        return rows;
    } catch (e) {
        console.warn("[Quote extras] OData fetch failed:", e.message);
        state.quoteExtrasDiagnostic = "OData " + info.entity + " fetch failed: " + e.message;
        return [];
    }
}


// ============================================================
// P&L MAPPING TABLE — the single source of truth
// ============================================================
// Shared by the browser tool (P&L tab, Overview reconciliation rows) and
// the snapshot robot. Add or move an account here and every consumer picks
// it up; there is no second copy to keep in step.
// ---- The mapping file (VAS_Account_Mapping_Table, 2026-09-24) ----
// P&L Bucket per G/L code, and BS Grouping for balance-sheet codes.
// Synthetic FS-/NEW-/SID- rows in the file do not exist in Wiise and are
// left out. Update here when accounts are added in Wiise.
const PNL_MAP_BUCKETS = {
    "Total Revenue (net)": ["4010", "4015", "4020", "4030", "4050", "4055", "4060", "4065"],
    "Cost of Goods Sold": ["5010", "5020", "5030", "5125", "5126", "5130", "5131", "5135", "5140", "5145", "5150", "5155", "5160", "5165", "5170", "5175", "5180", "5195", "5196"],
    "Direct Labour": ["5205", "5206", "5210", "5220", "5230", "5235", "5240", "5245", "5250", "5255", "5256", "5260"],
    "Marketing & Promotion": ["6010", "6015", "6020", "6025"],
    "ICT": ["6055", "6060", "6065", "6070"],
    "Vehicle Expenses": ["6105", "6110", "6115", "6120", "6125"],
    "Staff - Other Expenses": ["6155", "6160", "6165", "6170", "6175"],
    "Amortisation & Depreciation": ["6210"],
    "Insurance Expenses": ["6260"],
    "Financing & Leasing": ["6285", "6315", "6319"],
    "Bank & Merchant Fees": ["6305", "6309", "6310"],
    "Property Expenses": ["6335", "6340", "6345", "6350", "6355", "6360", "6365", "6370", "6375", "6380"],
    "Other Expenses": ["6405", "6410", "6415", "6420", "6425", "6430", "6435", "6440", "6445", "6450", "6460", "6465", "6470", "6475", "6477", "6480", "6488"],
    "Other Income": ["7005", "7010", "7025", "7030", "9999"],
    "Non-Operating Costs": ["8005", "8015", "8020", "8026"]
};
const PNL_MAP_BS = {
    "Cash and Cash Equivalents": ["1111", "1116", "1155", "1160"],
    "Trade and Other Receivables": ["1115", "1205", "1220", "1474", "2190"],
    "Trade and Other Payables": ["1140", "1150", "2110", "2120", "2146", "2147", "2160", "2165", "2166", "2170", "2175", "2180", "2195", "2212", "2215", "2220", "2225", "2226", "2255", "2280", "2285", "23999"],
    "Inventory": ["1300", "1310", "1320", "1330", "1340", "1350", "1360", "1475"],
    "Prepayments": ["1450"],
    "Intangibles": ["1460"],
    "Low Value Pool": ["1560"],
    "Less Accumulated Depreciation on Low Value Pool": ["1565"],
    "In-house Software Pool": ["1567", "1680", "1685"],
    "Plant and Equipment at Cost": ["1570", "1580", "1590", "1670"],
    "Accumulated Depreciation of Plant and Equipment": ["1575", "1585", "1595", "1650"],
    "Vehicles at Cost": ["1600"],
    "Accumulated Depreciation of Vehicles": ["1605"],
    "Buildings at Cost": ["1610", "1613"],
    "Accumulated Depreciation of Buildings": ["1615"],
    "Other Assets": ["1720", "2430", "2435"],
    "Financial Liabilities": ["2140", "2145", "2230", "2231", "2236", "2410", "2420", "2422", "3160"],
    "Equity": ["3100", "3200"]
};

// ============================================================
// GENERAL LEDGER — reconciliation slice
// ============================================================
// The Overview's two reconciliation rows bridge the sales tool's figures to
// the P&L. They look at the mapping's revenue and cost-of-goods-sold
// accounts, whatever accounts those are, so the slice is defined by the
// mapping table above and nothing else.
const GL_RECON_REVENUE  = PNL_MAP_BUCKETS["Total Revenue (net)"];
const GL_RECON_COGS     = PNL_MAP_BUCKETS["Cost of Goods Sold"];
const GL_RECON_ACCOUNTS = GL_RECON_REVENUE.concat(GL_RECON_COGS);
// Rows travel as compact arrays: a 12-month snapshot holds hundreds of
// thousands of them and key names would dwarf the data. valueEntries is the
// list of value-entry numbers that created a cost entry (0 when none, and
// always 0 on revenue rows).
const GL_RECON_COLS = ["postingDate", "documentNumber", "documentType", "accountNumber", "debitAmount", "creditAmount", "entryNumber", "valueEntries"];
async function fetchGLReconAccountNames() {
    const compId = await bcGetCompanyId();
    const rows = await bcFetchAll(BC_API_URL + "/companies(" + compId + ")/accounts?$select=number,displayName", "Chart of accounts (names)");
    return (rows || []).map(a => [String(a.number || ""), a.displayName || ""]);
}
function glReconDecode(s) { return (s == null ? "" : String(s)).replace(/_x([0-9a-fA-F]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16))).trim(); }
async function fetchGLReconEntries(fromISO, toISO) {
    const compId = await bcGetCompanyId();
    const dateFilter = "postingDate ge " + fromISO + " and postingDate le " + toISO;
    const acctFilter = "(" + GL_RECON_ACCOUNTS.map(a => "accountNumber eq '" + a + "'").join(" or ") + ")";
    const base = BC_API_URL + "/companies(" + compId + ")/generalLedgerEntries"
               + "?$select=entryNumber,postingDate,documentNumber,documentType,accountNumber,debitAmount,creditAmount"
               + "&$orderby=entryNumber&$filter=";
    // No $top anywhere in this file — BC treats it as a total cap.
    const rows = await bcFetchAll(base + encodeURIComponent(dateFilter + " and " + acctFilter), "Ledger (reconciliation accounts)");
    return (rows || []).map(e => [
        (e.postingDate || "").toString().slice(0, 10),
        e.documentNumber || "",
        glReconDecode(e.documentType),
        String(e.accountNumber || ""),
        num(e.debitAmount),
        num(e.creditAmount),
        Number(e.entryNumber) || 0,
        0,
    ]);
}

// ---- the ledger-entry <-> value-entry link (page 5823 "G/L - Item Ledger Relation")
// Business Central records which value entry created each ledger entry. The
// cost row is built on that record, so it needs no account lists and no
// guessing from document numbers. Published in Wiise as a web service and
// found by its fields, whatever it is named.
async function bcDiscoverGLItemRelation() {
    const md = await bcGetODataMetadata();
    for (const set of Object.keys(md.entitySets)) {
        const fields = md.entityTypes[md.entitySets[set]] || [];
        if (fields.includes("G_L_Entry_No") && fields.includes("Value_Entry_No")) return { entity: set, fields };
    }
    return null;
}
// $select has been seen to drop $filter on some of this tenant's OData
// objects, which would turn a range request into a whole-table download.
// Probe once with a tiny range and use $select only if it behaves.
let glRelationSelectOk = null;
async function glRelationProbeSelect(url, lo) {
    if (glRelationSelectOk !== null) return glRelationSelectOk;
    const filter = "G_L_Entry_No ge " + lo + " and G_L_Entry_No le " + (lo + 50);
    try {
        const withSel = await bcFetchAll(url + "?$filter=" + encodeURIComponent(filter) + "&$select=G_L_Entry_No,Value_Entry_No", "Ledger links (probe)");
        const plain   = await bcFetchAll(url + "?$filter=" + encodeURIComponent(filter), "Ledger links (probe)");
        const inRange = (withSel || []).every(r => Number(r.G_L_Entry_No) >= lo && Number(r.G_L_Entry_No) <= lo + 50);
        glRelationSelectOk = inRange && (withSel || []).length === (plain || []).length;
    } catch (e) { glRelationSelectOk = false; }
    return glRelationSelectOk;
}
// For every cost entry in the rows, attach the value-entry numbers that
// created it (row[7]). Fetched in entry-number windows so memory stays flat
// on a 12-month pull, skipping stretches that hold no cost entries.
async function attachGLReconLinks(rows) {
    const cogs = new Set(GL_RECON_COGS);
    const nos = rows.filter(r => cogs.has(r[3])).map(r => r[6]).filter(n => n > 0).sort((a, b) => a - b);
    if (!nos.length) return rows;
    const rel = await bcDiscoverGLItemRelation();
    if (!rel) throw new Error("The ledger-entry link (page 5823 G/L - Item Ledger Relation) is not published as a web service");
    const coName = encodeURIComponent(await bcGetCompanyInternalName());
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + rel.entity;
    const sel = (await glRelationProbeSelect(url, nos[0])) ? "&$select=G_L_Entry_No,Value_Entry_No" : "";
    const want = new Set(nos);
    const links = new Map();
    const WINDOW = 50000;
    let i = 0;
    while (i < nos.length) {
        const lo = nos[i], hi = lo + WINDOW - 1;
        const filter = "G_L_Entry_No ge " + lo + " and G_L_Entry_No le " + hi;
        const got = await bcFetchAll(url + "?$filter=" + encodeURIComponent(filter) + sel, "Ledger links");
        for (const r of (got || [])) {
            const g = Number(r.G_L_Entry_No);
            if (!want.has(g)) continue;
            const arr = links.get(g) || [];
            arr.push(Number(r.Value_Entry_No));
            links.set(g, arr);
        }
        while (i < nos.length && nos[i] <= hi) i++;
    }
    for (const r of rows) if (cogs.has(r[3])) r[7] = links.get(r[6]) || 0;
    return rows;
}
// The whole slice the reconciliation rows need: entries plus their links.
async function fetchGLReconSlice(fromISO, toISO) {
    return attachGLReconLinks(await fetchGLReconEntries(fromISO, toISO));
}
// Top-up for a slice taken from a snapshot: every reconciliation-account
// entry posted after the snapshot's last entry number, with its links.
// Ledger entries are never changed once posted, so the snapshot's rows stay
// right; going by entry number (not date) also catches back-dated postings.
async function fetchGLReconSince(afterEntryNo, fromISO, toISO) {
    const compId = await bcGetCompanyId();
    const filter = "entryNumber gt " + afterEntryNo + " and postingDate ge " + fromISO + " and postingDate le " + toISO
                 + " and (" + GL_RECON_ACCOUNTS.map(a => "accountNumber eq '" + a + "'").join(" or ") + ")";
    const rows = await bcFetchAll(BC_API_URL + "/companies(" + compId + ")/generalLedgerEntries"
        + "?$select=entryNumber,postingDate,documentNumber,documentType,accountNumber,debitAmount,creditAmount"
        + "&$orderby=entryNumber&$filter=" + encodeURIComponent(filter), "Ledger (since the snapshot)");
    const packed = (rows || []).map(e => [
        (e.postingDate || "").toString().slice(0, 10), e.documentNumber || "", glReconDecode(e.documentType),
        String(e.accountNumber || ""), num(e.debitAmount), num(e.creditAmount), Number(e.entryNumber) || 0, 0,
    ]);
    return attachGLReconLinks(packed);
}

// Node (snapshot Action) entry point. Classic-script browsers skip this.
if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        BC_CONFIG, BC_TENANT_DOMAIN, BC_API_BASE, BC_API_URL, BC_ODATA_URL, BC_SCOPES, SP_CONFIG, GRAPH_SCOPES,
        bcFetch, bcFetchAll, bcFetchAllTrimmed,
        bcGetCompanyId, bcGetCompanyInternalName, bcGetODataMetadata, bcFindField,
        bcDiscoverSalesLinesEntity, bcDiscoverValueEntryEntity, bcDiscoverSalesEntity,
        fetchValueEntries, fetchPostedInvoiceNumbers, fetchSalesInvoicesWithLines,
        fetchItemLedgerSales, fetchLocations, fetchItems, fetchInvoiceLines,
        fetchCustomers, fetchSalesCreditMemos, fetchSalesOrders,
        fetchSalesOrderOutstandingLines, fetchSalesShipments, fetchSalesReturnReceipts,
        fetchSalesQuotes, fetchBlanketSalesOrders, fetchResidentialDocLookup,
        fetchSalesQuoteArchive, fetchSalesQuoteExtras,
        fetchGLReconEntries, fetchGLReconAccountNames, fetchGLReconSlice, attachGLReconLinks, bcDiscoverGLItemRelation,
        fetchGLReconSince,
        PNL_MAP_BUCKETS, PNL_MAP_BS,
        GL_RECON_REVENUE, GL_RECON_COGS, GL_RECON_ACCOUNTS, GL_RECON_COLS,
    };
}
