// ============================================================
// cost-engine.js — LINE COST ENGINE
// Shared verbatim between the browser tool (index.html) and the
// snapshot GitHub Action (scripts/snapshot.js, Node). Keep it
// environment-neutral: no DOM. The host must provide these globals
// before any function here is called:
//   state   -> shared state object (invoices, creditMemos, shipments,
//              returnReceipts are read; the cost maps are written)
//   num(v)  -> numeric coercion util from the tool
//
// What lives here: the item-ledger / value-entry cost maps
// (buildCostMaps, ingestValueEntries) and the per-line cost lookup
// (costFromLine, costNonInvFromLine, totalCostFromLine) with its
// tiered fallbacks. Moved out of index.html unchanged so the snapshot
// robot can pre-compute the same per-line cost the browser would.
// ============================================================

function monthKey(iso) {
    const d = new Date(iso);
    if (isNaN(d)) return iso;
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
}

// ============================================================
// COST / GP COMPUTATION (from item ledger entries)
// ============================================================
function buildCostMaps(iles) {
    // Hybrid cost lookup. Build BOTH:
    //   (a) per-document × per-item maps   — gives exact line cost for the
    //       Ship+Invoice workflow where ILE.documentNumber matches the
    //       invoice/credit-memo number. This is the PBI-equivalent path.
    //   (b) the existing avg-unit-cost maps (item, item+month,
    //       cust+item+month) — fallback when the document-key lookup
    //       misses (Ship-first-Invoice-later workflows, etc.).
    invalidateLineCostMemo();
    state.itemAvgCost = new Map();
    state.itemAvgCostNonInv = new Map();
    state.itemMonthAvgCost = new Map();         // key "itemNo|YYYY-MM" -> avg unit cost
    state.itemMonthAvgCostNonInv = new Map();
    state.custItemMonthAvgCost = new Map();     // key "custNo|itemNo|YYYY-MM" -> avg unit cost
    state.custItemMonthAvgCostNonInv = new Map();
    state.costByItemDocKey      = new Map();    // key "docNo|itemNo" -> {cost, qty}
    state.costNonInvByItemDocKey = new Map();   // key "docNo|itemNo" -> {cost, qty}
    state.custItemDateAvgCost = new Map();      // key "cust|item|YYYY-MM-DD" -> avg unit cost
    state.custItemDateAvgCostNonInv = new Map();
    state.costStats = { rows: 0, items: 0, monthBuckets: 0, custBuckets: 0, docKeys: 0, totalCost: 0, totalQty: 0, sampleLines: 0, matched: 0, exactDocHits: 0, fields: null };

    // Build documentNumber → customerNumber map from all loaded posted documents.
    // ILE entries store the shipment / return-receipt number (or invoice number
    // in Ship+Invoice workflows); use this map to find the customer per ILE row.
    const docToCust = new Map();
    for (const inv of state.invoices)      { if (inv.number) docToCust.set(inv.number, inv.customerNumber || ""); }
    for (const cm  of state.creditMemos)   { if (cm.number)  docToCust.set(cm.number,  cm.customerNumber  || ""); }
    for (const s   of state.shipments)     { if (s.number)   docToCust.set(s.number,   s.customerNumber   || ""); }
    for (const r   of state.returnReceipts){ if (r.number)   docToCust.set(r.number,   r.customerNumber   || ""); }
    if (!iles.length) { console.warn("[ILE] no rows fetched — cost will be unavailable"); return; }
    const pick = (row, candidates) => { for (const c of candidates) if (c in row) return c; return null; };
    const s = iles[0];
    const fItem   = pick(s, ["itemNumber", "item_Number", "Item_No", "itemNo"]) || "itemNumber";
    const fCost   = pick(s, ["costAmountActual", "Cost_Amount_Actual_", "cost_amount_actual"]) || "costAmountActual";
    const fCostNI = pick(s, ["costAmountNonInvtbl", "Cost_Amount_Non_Invtbl", "cost_amount_non_invtbl"]) || "costAmountNonInvtbl";
    const fQty    = pick(s, ["quantity", "Quantity"]) || "quantity";
    const fDate   = pick(s, ["postingDate", "posting_Date", "Posting_Date"]) || "postingDate";
    // ILE document number field varies: v2 REST = documentNumber, OData v4
    // = Document_No, older NAV = Document_No_. Pick whichever the tenant
    // returns; without it the per-doc cost lookup builds an empty map.
    const fDoc    = pick(s, ["documentNumber", "Document_No", "Document_No_", "documentNo", "document_Number"]);
    state.costStats.fields = { fItem, fCost, fCostNI, fQty, fDate, fDoc };
    console.log("[ILE] using fields:", state.costStats.fields);
    console.log("[ILE] sample row keys:", Object.keys(s).slice(0, 30).join(", "));

    const byItem = new Map();
    const byItemMonth = new Map();
    const byCustItemMonth = new Map();
    const byCustItemDate  = new Map();   // tighter than byCustItemMonth — same-day match
    let custLinked = 0;
    for (const e of iles) {
        state.costStats.rows++;
        const itemNo = e[fItem] || "";
        if (!itemNo) continue;
        const qty    = Math.abs(num(e[fQty]));
        if (qty === 0) continue;
        const cost   = Math.abs(num(e[fCost]));
        const costNI = Math.abs(num(e[fCostNI]));
        state.costStats.totalCost += cost;
        state.costStats.totalQty  += qty;
        const cur = byItem.get(itemNo) || { qty: 0, cost: 0, costNI: 0 };
        cur.qty += qty; cur.cost += cost; cur.costNI += costNI;
        byItem.set(itemNo, cur);
        // (a) exact per-(documentNumber, item) cost — used by the Ship+Invoice
        // workflow where ILE.documentNumber == the invoice / credit-memo No.
        const docNo = (fDoc ? e[fDoc] : e.documentNumber) || "";
        if (docNo) {
            const dk = docNo + "|" + itemNo;
            const curD = state.costByItemDocKey.get(dk) || { qty: 0, cost: 0 };
            curD.qty += qty; curD.cost += cost;
            state.costByItemDocKey.set(dk, curD);
            const curDN = state.costNonInvByItemDocKey.get(dk) || { qty: 0, cost: 0 };
            curDN.qty += qty; curDN.cost += costNI;
            state.costNonInvByItemDocKey.set(dk, curDN);
        }
        const dateRaw = e[fDate];
        const mk = dateRaw ? monthKey(dateRaw) : null;
        const dk = dateRaw ? String(dateRaw).substring(0, 10) : null;   // YYYY-MM-DD
        if (mk) {
            const k = itemNo + "|" + mk;
            const curM = byItemMonth.get(k) || { qty: 0, cost: 0, costNI: 0 };
            curM.qty += qty; curM.cost += cost; curM.costNI += costNI;
            byItemMonth.set(k, curM);
            const cust = docToCust.get(docNo);
            if (cust) {
                custLinked++;
                const ck = cust + "|" + itemNo + "|" + mk;
                const curC = byCustItemMonth.get(ck) || { qty: 0, cost: 0, costNI: 0 };
                curC.qty += qty; curC.cost += cost; curC.costNI += costNI;
                byCustItemMonth.set(ck, curC);
                // Daily granularity — catches same-day Ship+Invoice flow
                // without averaging across the month.
                if (dk) {
                    const cdk = cust + "|" + itemNo + "|" + dk;
                    const curCD = byCustItemDate.get(cdk) || { qty: 0, cost: 0, costNI: 0 };
                    curCD.qty += qty; curCD.cost += cost; curCD.costNI += costNI;
                    byCustItemDate.set(cdk, curCD);
                }
            }
        }
    }
    state.costStats.docKeys = state.costByItemDocKey.size;
    for (const [k, v] of byItem) {
        state.itemAvgCost.set(k, v.qty ? v.cost / v.qty : 0);
        state.itemAvgCostNonInv.set(k, v.qty ? v.costNI / v.qty : 0);
    }
    for (const [k, v] of byItemMonth) {
        state.itemMonthAvgCost.set(k, v.qty ? v.cost / v.qty : 0);
        state.itemMonthAvgCostNonInv.set(k, v.qty ? v.costNI / v.qty : 0);
    }
    for (const [k, v] of byCustItemMonth) {
        state.custItemMonthAvgCost.set(k, v.qty ? v.cost / v.qty : 0);
        state.custItemMonthAvgCostNonInv.set(k, v.qty ? v.costNI / v.qty : 0);
    }
    for (const [k, v] of byCustItemDate) {
        state.custItemDateAvgCost.set(k, v.qty ? v.cost / v.qty : 0);
        state.custItemDateAvgCostNonInv.set(k, v.qty ? v.costNI / v.qty : 0);
    }
    state.costStats.items = state.itemAvgCost.size;
    state.costStats.monthBuckets = state.itemMonthAvgCost.size;
    state.costStats.custBuckets  = state.custItemMonthAvgCost.size;
    console.log("[ILE] cost layers: " + state.costStats.items + " items · " +
                state.costStats.monthBuckets + " item-month · " +
                state.costStats.custBuckets + " cust-item-month (" + custLinked + " ILE rows linked to a customer)");

    console.log("[ILE] " + state.costStats.docKeys + " (docNumber|item) keys built for exact-cost lookup");

    // ---- Shipment-line bridge ----
    // On Ship-first-Invoice-later tenants the ILE document numbers are
    // shipment numbers, never invoice numbers, so the docKey tier above is
    // 0% useful. salesShipments is already fetched and its lines carry
    // (customer, item, qty, shipmentNumber, date). Build a per-
    // (customer|item) sorted list of shipment lines so costFromLine can
    // walk it for invoices that miss same-day, average-weighted cost
    // across the shipments inside a ~30 day lookback window, and pull the
    // actual cost from state.costByItemDocKey keyed by shipmentNumber.
    state.shipmentLinesByCustItem = new Map();
    let shipLines = 0;
    for (const sh of (state.shipments || [])) {
        const cust = sh.customerNumber || "";
        const sNo  = sh.number || "";
        const date = sh.postingDate ? String(sh.postingDate).substring(0, 10) : "";
        if (!cust || !sNo || !date) continue;
        for (const ln of (sh.salesShipmentLines || [])) {
            const itemNo = ln.lineObjectNumber || ln.itemId || "";
            if (!itemNo) continue;
            const q = Math.abs(num(ln.quantity));
            if (q === 0) continue;
            const k = cust + "|" + itemNo;
            let arr = state.shipmentLinesByCustItem.get(k);
            if (!arr) { arr = []; state.shipmentLinesByCustItem.set(k, arr); }
            arr.push({ date, shipmentNumber: sNo, qty: q });
            shipLines++;
        }
    }
    // Newest first — the lookup walks until it crosses the 30-day cutoff
    for (const arr of state.shipmentLinesByCustItem.values()) {
        arr.sort((a, b) => a.date < b.date ? 1 : a.date > b.date ? -1 : 0);
    }
    state.costStats.shipKeys = state.shipmentLinesByCustItem.size;
    console.log("[Shipments] indexed " + shipLines + " shipment lines into " +
                state.costStats.shipKeys + " (cust|item) buckets for the cost-bridge tier");

    // Spot-check: how many invoice lines have a matching avg-cost AND
    // how many would hit the exact-document cost path (Ship+Invoice).
    if (state.invoices && state.invoices.length) {
        let lineCount = 0, matched = 0, exact = 0, sameDay = 0, bridged = 0;
        for (const inv of state.invoices.slice(0, 100)) {
            const invDate = inv.postingDate ? String(inv.postingDate).substring(0, 10) : "";
            for (const ln of (inv.salesInvoiceLines || [])) {
                lineCount++;
                const itemNo = ln.lineObjectNumber || ln.itemId || "";
                if (itemNo && state.itemAvgCost.has(itemNo) && state.itemAvgCost.get(itemNo) > 0) matched++;
                if (inv.number && itemNo && state.costByItemDocKey.has(inv.number + "|" + itemNo)) exact++;
                const hitSameDay = inv.customerNumber && itemNo && invDate &&
                    state.custItemDateAvgCost.has(inv.customerNumber + "|" + itemNo + "|" + invDate);
                if (hitSameDay) sameDay++;
                // Bridge counts a line if same-day MISSED but a shipment exists for
                // this (cust|item) within the 30-day lookback before invDate.
                if (!hitSameDay && inv.customerNumber && itemNo && invDate &&
                    shipmentBridgeUnitCost(inv.customerNumber, itemNo, inv.postingDate, "inv", 30) > 0) bridged++;
            }
        }
        state.costStats.sampleLines = lineCount;
        state.costStats.matched = matched;
        const pct = lineCount ? Math.round(matched / lineCount * 100) : 0;
        const exactPct = lineCount ? Math.round(exact / lineCount * 100) : 0;
        const sameDayPct = lineCount ? Math.round(sameDay / lineCount * 100) : 0;
        const bridgedPct = lineCount ? Math.round(bridged / lineCount * 100) : 0;
        const combinedPct = lineCount ? Math.round((sameDay + bridged) / lineCount * 100) : 0;
        console.log("[ILE] exact-doc coverage: " + exact + "/" + lineCount + " (" + exactPct + "%) sample lines match by invoice No.");
        console.log("[ILE] same-day coverage: " + sameDay + "/" + lineCount + " (" + sameDayPct + "%) sample lines match by (cust+item+date).");
        console.log("[ILE] shipment-bridge coverage: " + bridged + "/" + lineCount + " (" + bridgedPct + "%) extra lines bridged via shipments; combined precise = " + combinedPct + "%");
        console.log("[ILE] line-item coverage: " + matched + "/" + lineCount + " (" + pct + "%) sample lines have an avg-cost match");
    }
}

// Item-line check used by cost-from-line (cost only flows for stock items).
function isItemLine(line) {
    if (line == null) return false;
    if (line.lineType) return line.lineType === "Item";
    return !!(line.lineObjectNumber || line.itemId);
}

// Walk shipment lines for this (customer, item), starting from the
// invoice posting date and looking back up to `windowDays`. For each
// shipment found, accumulate the ILE cost via the per-doc cost map.
// Returns avg unit cost (totalCost / totalQty) or 0 when no match.
// `which`: "inv" → state.costByItemDocKey, "noninv" → state.costNonInvByItemDocKey.
function shipmentBridgeUnitCost(customerNumber, itemNo, postingDate, which, windowDays) {
    if (!customerNumber || !itemNo || !postingDate) return 0;
    const arr = state.shipmentLinesByCustItem && state.shipmentLinesByCustItem.get(customerNumber + "|" + itemNo);
    if (!arr || !arr.length) return 0;
    const invISO = String(postingDate).substring(0, 10);
    const cutoffMs = new Date(invISO).getTime() - (windowDays || 30) * 86400000;
    const docMap = (which === "noninv") ? state.costNonInvByItemDocKey : state.costByItemDocKey;
    if (!docMap) return 0;
    let totalCost = 0, totalQty = 0;
    for (const c of arr) {                                    // newest first
        if (c.date > invISO) continue;                        // future of invoice — skip
        if (new Date(c.date).getTime() < cutoffMs) break;     // sorted desc; everything older crosses cutoff
        const ile = docMap.get(c.shipmentNumber + "|" + itemNo);
        if (ile && ile.qty > 0) { totalCost += ile.cost; totalQty += ile.qty; }
    }
    return totalQty > 0 ? totalCost / totalQty : 0;
}

function costFromLineRaw(line, sign, postingDate, customerNumber, documentNumber) {
    if (!isItemLine(line)) return 0;
    sign = (sign == null) ? 1 : sign;
    const itemNo = line.lineObjectNumber || line.itemId || "";
    const qty = Math.abs(num(line.quantity));
    // Tier 0a: Value Entry per-(invoice|item) — authoritative for direct-
    // invoice items where BC posts cost on the invoice VE row itself
    // (no preceding shipment, no ILE bridge). Skip when VE cost is zero
    // (Ship+Invoice flows post cost on the shipment side, not invoice).
    if (documentNumber && itemNo && state.veCostByDocItem) {
        const v = state.veCostByDocItem.get(documentNumber + "|" + itemNo);
        if (v && v.qty !== 0 && Math.abs(v.cost) > 0) {
            return (Math.abs(v.cost) / Math.abs(v.qty)) * qty * sign;
        }
    }
    // Tier 0: exact per-document cost (Ship+Invoice workflow — ILE doc
    // number matches the invoice / credit-memo number). Gives PBI-exact
    // line cost when present.
    if (documentNumber && itemNo) {
        const d = state.costByItemDocKey && state.costByItemDocKey.get(documentNumber + "|" + itemNo);
        if (d && d.qty > 0) {
            if (state.costStats) state.costStats.exactDocHits++;
            return (d.cost / d.qty) * qty * sign;
        }
    }
    // Tiered fallback, tightest-match first:
    //   1. cust + item + date     (same-day ILE entry — Ship+Invoice exact)
    //   2. shipment-line bridge   (Ship-first-Invoice-later: find earlier
    //                              shipments to this customer for this item
    //                              within 30 days, average their actual cost)
    //   3. cust + item + month    (any shipment of this item to this customer)
    //   4. item + month           (any shipment of this item that month)
    //   5. global item average    (last resort)
    let uc = 0;
    if (postingDate && customerNumber) {
        const dk = String(postingDate).substring(0, 10);
        const v = state.custItemDateAvgCost.get(customerNumber + "|" + itemNo + "|" + dk);
        if (v != null && v > 0) uc = v;
    }
    if (uc === 0) uc = shipmentBridgeUnitCost(customerNumber, itemNo, postingDate, "inv", 30);
    if (uc === 0 && postingDate) {
        const mk = monthKey(postingDate);
        if (customerNumber) {
            const v = state.custItemMonthAvgCost.get(customerNumber + "|" + itemNo + "|" + mk);
            if (v != null && v > 0) uc = v;
        }
        if (uc === 0) {
            const v = state.itemMonthAvgCost.get(itemNo + "|" + mk);
            if (v != null && v > 0) uc = v;
        }
    }
    if (uc === 0) uc = state.itemAvgCost.get(itemNo) || 0;
    return uc * qty * sign;
}
// Power BI's "Adjusted Cost (LCY)" measure sums BOTH cost components
// from Value Entries: Cost Amount (Actual) + Cost Amount (Non-Invtbl).
// All headline aggregations and KPI / table cost columns use this
// combined helper so the headline number matches PBI. The drill-down
// trees keep the two parts separate via costFromLine + costNonInvFromLine
// because they expose distinct "Adj Cost" and "Cost Non-Inv" columns.
// ---- Per-line memo ----
// A line's cost depends only on the line itself, its document (posting
// date / customer / number) and the cost maps built at load. None of
// those move when the user changes the toolbar dates, filters or tab,
// so the answer is kept against the line object (WeakMap: it goes
// away with the line) and reused. buildCostMaps / ingestValueEntries
// reset the memo so a fresh load never reads a stale value. Lines that
// arrive with preCost / preCostNonInv already on them (snapshot robot,
// formatVersion 2) skip the lookup altogether.
let _lineCostMemo = new WeakMap();
let _lineCostGen = 0;
function invalidateLineCostMemo() { _lineCostGen++; _lineCostMemo = new WeakMap(); }
function lineCostPair(line, postingDate, customerNumber, documentNumber) {
    if (line == null || typeof line !== "object") return { inv: 0, ni: 0 };
    if (line.preCost != null) return { inv: num(line.preCost), ni: num(line.preCostNonInv) };
    let m = _lineCostMemo.get(line);
    if (m && m.gen === _lineCostGen && m.doc === documentNumber) return m;
    m = { gen: _lineCostGen, doc: documentNumber,
          inv: costFromLineRaw(line, 1, postingDate, customerNumber, documentNumber),
          ni:  costNonInvFromLineRaw(line, 1, postingDate, customerNumber, documentNumber) };
    _lineCostMemo.set(line, m);
    return m;
}
function costFromLine(line, sign, postingDate, customerNumber, documentNumber) {
    return lineCostPair(line, postingDate, customerNumber, documentNumber).inv * ((sign == null) ? 1 : sign);
}
function costNonInvFromLine(line, sign, postingDate, customerNumber, documentNumber) {
    return lineCostPair(line, postingDate, customerNumber, documentNumber).ni * ((sign == null) ? 1 : sign);
}
function totalCostFromLine(line, sign, postingDate, customerNumber, documentNumber) {
    const p = lineCostPair(line, postingDate, customerNumber, documentNumber);
    return (p.inv + p.ni) * ((sign == null) ? 1 : sign);
}
// Snapshot robot: stamp every line of `docs` with the cost the browser
// would compute, so the tool reads it instead of running the lookup.
function precostLines(docs, linesKey) {
    let lines = 0, costed = 0;
    for (const d of (docs || [])) {
        for (const ln of (d[linesKey] || [])) {
            ln.preCost       = costFromLineRaw(ln, 1, d.postingDate, d.customerNumber, d.number);
            ln.preCostNonInv = costNonInvFromLineRaw(ln, 1, d.postingDate, d.customerNumber, d.number);
            lines++;
            if (ln.preCost || ln.preCostNonInv) costed++;
        }
    }
    return { lines, costed };
}
function costNonInvFromLineRaw(line, sign, postingDate, customerNumber, documentNumber) {
    if (!isItemLine(line)) return 0;
    sign = (sign == null) ? 1 : sign;
    const itemNo = line.lineObjectNumber || line.itemId || "";
    const qty = Math.abs(num(line.quantity));
    // Tier 0a: Value Entry per-(invoice|item) non-invtbl cost (same
    // reasoning as costFromLine's tier 0a — catches direct-invoice items
    // missing from ILE bridge).
    if (documentNumber && itemNo && state.veCostNonInvByDocItem) {
        const v = state.veCostNonInvByDocItem.get(documentNumber + "|" + itemNo);
        if (v && v.qty !== 0 && Math.abs(v.cost) > 0) {
            return (Math.abs(v.cost) / Math.abs(v.qty)) * qty * sign;
        }
    }
    // Tier 0: exact per-document non-invtbl cost (Ship+Invoice workflow).
    if (documentNumber && itemNo) {
        const d = state.costNonInvByItemDocKey && state.costNonInvByItemDocKey.get(documentNumber + "|" + itemNo);
        if (d && d.qty > 0) return (d.cost / d.qty) * qty * sign;
    }
    let uc = 0;
    if (postingDate && customerNumber) {
        const dk = String(postingDate).substring(0, 10);
        const v = state.custItemDateAvgCostNonInv.get(customerNumber + "|" + itemNo + "|" + dk);
        if (v != null && v > 0) uc = v;
    }
    if (uc === 0) uc = shipmentBridgeUnitCost(customerNumber, itemNo, postingDate, "noninv", 30);
    if (uc === 0 && postingDate) {
        const mk = monthKey(postingDate);
        if (customerNumber) {
            const v = state.custItemMonthAvgCostNonInv.get(customerNumber + "|" + itemNo + "|" + mk);
            if (v != null && v > 0) uc = v;
        }
        if (uc === 0) {
            const v = state.itemMonthAvgCostNonInv.get(itemNo + "|" + mk);
            if (v != null && v > 0) uc = v;
        }
    }
    if (uc === 0) uc = state.itemAvgCostNonInv.get(itemNo) || 0;
    return uc * qty * sign;
}

function isSalesValueEntry(v) {
    const et = (v.entryType || "").toString();
    return /^sale$/i.test(et);
}

// Shared between full and fast load: stash VE rows and build the
// per-(doc|item) / per-invoice cost maps every cost consumer reads.
function ingestValueEntries(valueEntries) {
    invalidateLineCostMemo();
    state.valueEntries = valueEntries || [];
    // Build per-(docNumber|itemNo) cost map from VE rows so direct-
    // invoice items (no shipment / no ILE entry) still get accurate
    // cost. Sum signed values; consumers take abs at lookup time.
    state.veCostByDocItem = new Map();
    state.veCostNonInvByDocItem = new Map();
    // Per-invoice total cost map. Sums signed cost+costNI across all
    // sales-type VE rows for each invoice. invoiceCost() reads from
    // this so Branch / Salesperson / Sales-by-Item / EOM / MTD use
    // the same cost number the Overview tab does (PBI-aligned).
    state.veCostByInvoice = new Map();
    for (const v of state.valueEntries) {
        if (!isSalesValueEntry(v)) continue;
        const d = v.documentNumber;
        if (d) {
            const prev = state.veCostByInvoice.get(d) || 0;
            state.veCostByInvoice.set(d, prev + (v.costAmount || 0) + (v.costAmountNonInv || 0));
        }
        if (!v.documentNumber || !v.itemNumber) continue;
        const k = v.documentNumber + "|" + v.itemNumber;
        const c = state.veCostByDocItem.get(k) || { cost: 0, qty: 0 };
        c.cost += v.costAmount; c.qty += v.invoicedQty;
        state.veCostByDocItem.set(k, c);
        const cn = state.veCostNonInvByDocItem.get(k) || { cost: 0, qty: 0 };
        cn.cost += v.costAmountNonInv; cn.qty += v.invoicedQty;
        state.veCostNonInvByDocItem.set(k, cn);
    }
    if (state.veCostByInvoice.size) {
        console.log("[Value Entry] built veCostByInvoice map · " + state.veCostByInvoice.size + " distinct documents");
    }
    if (state.veCostByDocItem.size) {
        console.log("[Value Entry] built veCostByDocItem map · " + state.veCostByDocItem.size + " (doc, item) keys");
    }
}

// Node (snapshot Action) entry point. Classic-script browsers skip this.
if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        monthKey, isItemLine, isSalesValueEntry,
        buildCostMaps, ingestValueEntries,
        shipmentBridgeUnitCost, costFromLine, costNonInvFromLine, totalCostFromLine,
        costFromLineRaw, costNonInvFromLineRaw, invalidateLineCostMemo, precostLines,
    };
}
