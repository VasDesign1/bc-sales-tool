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

let bcCompanyId = null, bcCompanyInternalName = null, bcODataMetadata = null;

async function bcFetch(url) {
    const token = await bcGetToken();
    let resp = await fetch(url, { headers: { "Authorization": "Bearer " + token, "Accept": "application/json" } });
    if (resp.status === 401) {
        bcAccessToken = null;
        const newToken = await bcGetToken();
        resp = await fetch(url, { headers: { "Authorization": "Bearer " + newToken, "Accept": "application/json" } });
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
    const resp = await fetch(BC_ODATA_URL + "/$metadata", { headers: { "Authorization": "Bearer " + token } });
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
    const candidates = named.length ? named : allSets;
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
        };
    }
    return null;
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
    const params = ["$top=200000", "$filter=" + encodeURIComponent(filter)];
    // Deliberately NO $select here — verified on this tenant that adding
    // $select to the ValueEntries query object silently DROPS the date
    // $filter and returns the whole table (102k rows for a 3-month ask).
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity + "?" + params.join("&");
    try {
        const rows = await bcFetchAll(url, "Value Entries (" + info.entity + ")");
        console.log("[Value Entry] " + info.entity + " → " + rows.length + " rows");
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
        const token = await bcGetToken();
        const cname = await bcGetCompanyInternalName();
        const filter = "Posting_Date ge " + fromISO + " and Posting_Date le " + toISO;
        const url = BC_ODATA_URL + "/Company('" + encodeURIComponent(cname) + "')/salesInvoiceHeader?$select=No&$filter=" + encodeURIComponent(filter) + "&$top=10000";
        const resp = await fetch(url, { headers: { "Authorization": "Bearer " + token } });
        if (!resp.ok) {
            console.warn("[Posted invoice whitelist] HTTP " + resp.status);
            return null;
        }
        const j = await resp.json();
        const set = new Set((j.value || []).map(r => r.No));
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
        const data = await bcFetch(BC_API_URL + "/companies(" + compId + ")/locations?$top=1000");
        return data.value || [];
    } catch (e) { console.warn("Could not fetch locations:", e.message); return []; }
}
async function fetchItems() {
    const compId = await bcGetCompanyId();
    try {
        const data = await bcFetchAll(BC_API_URL + "/companies(" + compId + ")/items?$top=10000&$select=id,number,displayName,itemCategoryCode", "Items");
        return data;
    } catch (e) { console.warn("Could not fetch items:", e.message); return []; }
}
async function fetchInvoiceLines(invoiceId) {
    const compId = await bcGetCompanyId();
    return bcFetchAll(BC_API_URL + "/companies(" + compId + ")/salesInvoices(" + invoiceId + ")/salesInvoiceLines?$top=10000", "Invoice lines");
}
async function fetchCustomers() {
    const compId = await bcGetCompanyId();
    try {
        return await bcFetchAll(BC_API_URL + "/companies(" + compId + ")/customers?$top=10000", "Customers");
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
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + linesInfo.entity +
        "?$top=200000&$filter=" + encodeURIComponent(filter);
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
    const params = ["$top=10000"];
    if (filters.length) params.push("$filter=" + encodeURIComponent(filters.join(" and ")));
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity + "?" + params.join("&");
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
                const lineParams = ["$top=50000"];
                // Filter to just our headers' document numbers if we have a manageable count
                const docNos = normalized.map(o => o.number).filter(Boolean);
                if (docNos.length && docNos.length <= 400) {
                    const inClause = docNos.map(n => linesInfo.fDocNo + " eq '" + String(n).replace(/'/g, "''") + "'").join(" or ");
                    lineParams.push("$filter=" + encodeURIComponent(inClause));
                }
                const linesUrl = BC_ODATA_URL + "/Company('" + coName + "')/" + linesInfo.entity + "?" + lineParams.join("&");
                const lineRows = await bcFetchAll(linesUrl, "Blanket order lines (" + linesInfo.entity + ")");
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
            "?$select=" + selectCols + "&$top=20000";
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
    const params = ["$top=20000"];
    if (filter) params.push("$filter=" + encodeURIComponent(filter));
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity + "?" + params.join("&");
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
    const params = ["$top=10000"];
    if (selectFields.length) params.push("$select=" + selectFields.join(","));
    if (filters.length) params.push("$filter=" + encodeURIComponent(filters.join(" and ")));
    const url = BC_ODATA_URL + "/Company('" + coName + "')/" + info.entity + "?" + params.join("&");
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


// Node (snapshot Action) entry point. Classic-script browsers skip this.
if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        BC_CONFIG, BC_TENANT_DOMAIN, BC_API_BASE, BC_API_URL, BC_ODATA_URL, BC_SCOPES,
        bcFetch, bcFetchAll, bcFetchAllTrimmed,
        bcGetCompanyId, bcGetCompanyInternalName, bcGetODataMetadata, bcFindField,
        bcDiscoverSalesLinesEntity, bcDiscoverValueEntryEntity, bcDiscoverSalesEntity,
        fetchValueEntries, fetchPostedInvoiceNumbers, fetchSalesInvoicesWithLines,
        fetchItemLedgerSales, fetchLocations, fetchItems, fetchInvoiceLines,
        fetchCustomers, fetchSalesCreditMemos, fetchSalesOrders,
        fetchSalesOrderOutstandingLines, fetchSalesShipments, fetchSalesReturnReceipts,
        fetchSalesQuotes, fetchBlanketSalesOrders, fetchResidentialDocLookup,
        fetchSalesQuoteArchive, fetchSalesQuoteExtras,
    };
}
