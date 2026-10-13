// ============================================================
// help.js — Help mode ("?" button in the header).
// OFF by default: no hover explanations. ON: hover anything — a sidebar
// tab, a toolbar filter, a button, a KPI tile, a chart card, a column
// header, a table cell, a status badge — and a floating box says what it
// is and how it is worked out. Clicks keep working while help is on.
// Wording lives here, grouped by what it matches. When something has no
// entry, the element's own title text (if any) is shown instead.
// ============================================================
(function () {
    'use strict';
    let helpOn = false;
    const B = (title, body, note) => '<strong>' + title + '</strong>' + body + (note ? '<em>' + note + '</em>' : '');

    // ---- Sidebar groups ----
    const GROUPS = {
        'Dashboards': B('Dashboards', 'The day-to-day views: overall figures, branch and salesperson results, items, customers, quotes, blanket orders and returns. All of them follow the toolbar period and filters.'),
        'Time Analytics': B('Time Analytics', 'Views that look at how sales move over time: month to date, rolling 12 months, month on month, daily pattern, customer retention, forecasting, Pareto and the decomposition tree.'),
        'Targets & Run Rate': B('Targets & Run Rate', 'Actual sales and margin against the budget. Targets come from the Excel in the Sales Budgets-Targets folder on SharePoint, read with your own sign-in; nothing is typed into the tool.'),
        'Lists': B('Lists', 'Working lists still being built.'),
        'Finance': B('Finance', 'The general ledger view. Visible to the finance group only; the figures come straight from the Wiise ledger, not from the sales documents.'),
    };

    // ---- Sidebar tabs (also used by the tab strip in Tab View) ----
    const TABS = {
        overview: B('Overview', 'The headline figures for the loaded period: sales, gross profit, quantity, the adjusted set (with returns taken off), the open pipeline, and the main charts. Everything follows the Period and the filters in the toolbar.'),
        deepInsights: B('Deep Insights', 'Nineteen visual analyses built from the same loaded invoices: sales flow, sunburst, salesperson radar, funnel, treemap, heatmaps, RFM segments, cohorts, geography and more. Good for spotting patterns rather than reading exact numbers.'),
        branch: B('Branch Sales', 'Sales, gross profit and volume by branch. The branch is the Wiise location on each invoice line, so one invoice can be split across branches. Monthly trend, top branches, GP, volume, a fiscal-year by month table and a drill-down tree.'),
        salesperson: B('Salesperson Sales', 'Sales, gross profit and volume by salesperson. The salesperson is the code on the invoice header. Monthly trend, top salespeople, GP, volume, a fiscal-year by month table and a drill-down tree.'),
        byItem: B('Sales by Item', 'What sold: the top 20 items, the split by item category, an item Pareto, and a searchable table of every item in the period. Click an item to see which customers bought it.'),
        customerItem: B('Customer × Item', 'Who buys what. Builds a universe of customers against items, then shows buyers and non-buyers, how much of the universe each salesperson covers, and the items and customers in scope.'),
        eom: B('EOM Report', 'The end-of-month report at invoice-line level: every line with its sell price, cost, GP and GP %. Flags for unusually low or high margins, a drill-down tree and an export to Excel.'),
        quotes: B('Sales Quote Overview', 'Sales quotes from Wiise, open and historic: value, count, win rate, quoting trend, and the split by salesperson, item category, assigned user and campaign. Expand a customer to see its quotes and lines.'),
        blanketOrders: B('Blanket Sales Order Overview', 'Blanket (standing) sales orders from Wiise: value, count, status, win rate, trend, and the split by salesperson, item category, assigned user and campaign.'),
        returns: B('Sales Returns Overview', 'Posted credit memos in the period: the return rate against sales, credit by customer and by item, the GP impact, and the memo list.'),
        mtd: B('Month-to-Date', 'This month so far, day by day, building up to today. It uses the current calendar month regardless of the Period in the toolbar; the other filters still apply.'),
        mat: B('Moving Annual Total', 'A rolling 12-month total of sales for each month, compared with the same rolling total a year earlier. Smooths out seasonality so the underlying direction is visible.'),
        pop: B('Period-over-Period', 'Each month against the month before it: sales, growth in dollars and percent, and the GP for the month.'),
        dailySales: B('Daily Sales Matrix', 'A calendar grid: one row per week, one column per weekday, sales in each cell. Shows which days of the week carry the sales.'),
        demographics: B('Sales Demographics', 'Sales split three ways: by item category, by the customer’s city, and by the customer posting group from Wiise.'),
        retention: B('Customer Retention', 'For each month, how many customers bought for the first time in the period and how many were returning. Click a month to see the customers.'),
        retentionHistory: B('Customer Retention History', 'One row per customer: first and last purchase, days since, the longest gap, invoices and sales, and a status of Active, New, Returning or Lost with the reason.'),
        forecast: B('Sales Forecasting', 'Monthly sales projected forward. Pick a method and how many months ahead, then Update forecast. Actual months are drawn solid, forecast months dashed.'),
        pareto: B('Sales Pareto (80/20)', 'How concentrated the sales are: the few customers, products, salespeople or days that make up most of the total, with the cumulative percentage line.'),
        decomposition: B('Sales Decomposition', 'A tree that splits the total sales through up to five levels you choose, for example Branch, then Salesperson, then Customer. Each level shows its share of the total.'),
        branchTarget: B('Branch Sales & Targets', 'Every loaded month by branch: Sales $, GP $, GP %, the Sales $ target for that branch and month from the budget file, the difference, and the percent achieved. The month row carries the total of all branch targets.'),
        branchSalesRate: B('Branch Sales Run Rate', 'For the chosen month: each branch’s sales against its target, as a daily run rate across the selling days of the month. Shows whether the branch is on track, what it still needs per day, and comparison charts underneath.'),
        branchMarginRate: B('Branch Margin Run Rate', 'Gross margin dollars by branch for the current month against a target, as a daily run rate. Targets on this tab are typed in and saved on this computer; there is no branch margin budget in the file yet.'),
        salespersonSalesRate: B('Salesperson Sales Run Rate', 'For the chosen month: each salesperson’s sales against their budget target, as a daily run rate across the selling days of the month. Shows whether they are on track, what they still need per day, and comparison charts underneath.'),
        salespersonMarginRate: B('Salesperson Margin Run Rate', 'For the chosen month: each salesperson’s gross margin dollars against their GP budget target, as a daily run rate. Same layout as the sales run rate.'),
        pnl: B('Profit & Loss (Ledger)', 'The P&L read straight from the Wiise general ledger for the toolbar period: income, cost of goods sold, gross profit, expenses and net profit, month by month, by account, with the account map. Independent of the sales documents.'),
        keyCustomers: B('Key Customer Sales', 'Not built yet.'),
        ordersPastDel: B('Orders past Delivery', 'Not built yet.'),
    };

    // ---- Header, toolbar and other things found by element id ----
    const IDS = {
        headerTitle: B('Current tab', 'The name of the tab you are on.'),
        loadBtn: B('Load sales', 'Loads live data from Wiise for the Period in the toolbar: posted invoices and their lines, credit memos, shipments, value entries for cost, locations and customers. Takes longer for long periods; for a quick look use Fast lookup.'),
        loadLookupBtn: B('Fast lookup', 'Opens a ready-made snapshot of the last 12 months, captured automatically three times a day (morning, midday, evening, Melbourne time). Everything on every tab then works on that snapshot in a few seconds. Needs your Microsoft sign-in and the snapshot passphrase, asked once per device.'),
        loadLookupMenu: B('Snapshot slots', 'Pick the morning, midday or evening capture. The time shown is when that snapshot was taken.'),
        userBadge: B('Signed-in user', 'The Microsoft account the tool is using. Wiise, SharePoint and the targets file are all read with this account’s own permissions.'),
        periodBadge: B('Loaded period', 'The date range of the data currently loaded. The toolbar can narrow it, not widen it; widen the Period and Load sales again for more.'),
        costQualityBadge: B('Cost quality', 'How the cost figures were obtained. Value entries from Wiise give the exact item cost; when they are missing the tool falls back to the unit cost on the line and says so here.'),
        snapshotBadge: B('Figures as of', 'Shown when a Fast lookup snapshot is in use: the time the snapshot was captured. Live Load sales clears it.'),
        themeBtn: B('Theme', 'Changes the colours and fonts of the whole tool. Your choice is kept on this computer.'),
        signOutBtn: B('Sign out', 'Signs the tool out of your Microsoft account.'),
        helpBtn: B('Help (?)', 'Turns help on or off. While it is on, hover over anything to see what it means.'),
        tabViewWrap: B('Tab View', 'ON: the sidebar hides and every tab appears as a strip of buttons under the filters, so the page uses the full width. OFF: the usual sidebar layout. Kept on this computer.'),
        tabViewSwitch: B('Tab View', 'ON: the sidebar hides and every tab appears as a strip of buttons under the filters, so the page uses the full width. OFF: the usual sidebar layout. Kept on this computer.'),
        tabStrip: B('Tabs', 'The same tabs as the sidebar, in the same groups. Click one to open it.'),
        fromDate: B('From', 'Start of the Period. Only invoices posted on or after this date count. Typing a date clears Quick range and Month.'),
        toDate: B('To', 'End of the Period. Only invoices posted on or before this date count. Typing a date clears Quick range and Month.'),
        quickRange: B('Quick range', 'Sets From and To in one go: last 30 or 90 days, 6 or 12 months, year to date, or the current fiscal year (July to June).'),
        monthPicker: B('Month', 'Sets the Period to one calendar month. The list runs from Wiise’s first month to now. Picking a month clears Quick range.'),
        monthPickerField: B('Month', 'Sets the Period to one calendar month. The list runs from Wiise’s first month to now. Picking a month clears Quick range.'),
        locationFilter: B('Branch filter', 'Keeps only invoice lines whose Wiise location is this branch. Line-based, so an invoice with lines from two branches is split. “(no location)” is lines with no branch at all, such as freight and other charges.'),
        salespersonFilter: B('Salesperson filter', 'Keeps only invoices whose header carries this salesperson code. The code is the person’s initials in Wiise.'),
        customerFilter: B('Customer filter', 'Keeps only invoices for this customer. The list holds the customers billed in the loaded period.'),
        itemFilterBox: B('Item filter', 'Type part of an item number or description and pick one. Keeps only invoice lines for that item.'),
        itemCombo: B('Item filter', 'Type part of an item number or description and pick one. Keeps only invoice lines for that item.'),
        includeRentalsToggle: B('Rentals', 'Rental income (G/L line 7005, for example the monthly Hunter Leisure rental) is left out of every figure by default. Switch it on to include it, which is what Wiise and Power BI totals do.'),
        includeRentalsKnob: B('Rentals', 'Rental income (G/L line 7005) is left out of every figure by default. Switch it on to include it, which is what Wiise and Power BI totals do.'),
        includeRentalsState: B('Rentals', 'Off: rental income excluded. On: included.'),
        filterChips: B('Active filters', 'The filters currently applied, as chips. Click the × on a chip to clear that filter.'),
        // Run-rate tabs
        spsrMonthSel: B('Month', 'Which budgeted month to look at. The current month shows a live run rate; a past month shows the finished result; a future month shows the target only.'),
        spmrMonthSel: B('Month', 'Which budgeted month to look at. The current month shows a live run rate; a past month shows the finished result; a future month shows the target only.'),
        brsrMonthSel: B('Month', 'Which budgeted month to look at. The current month shows a live run rate; a past month shows the finished result; a future month shows the target only.'),
        spsrCmpView: B('Comparison view', 'Which chart to draw underneath: target vs actual, percent achieved, ahead or behind, monthly targets across the year, one person across the year, year-to-date cumulative, or share of total.'),
        spmrCmpView: B('Comparison view', 'Which chart to draw underneath: target vs actual, percent achieved, ahead or behind, monthly targets across the year, one person across the year, year-to-date cumulative, or share of total.'),
        brsrCmpView: B('Comparison view', 'Which chart to draw underneath: target vs actual, percent achieved, ahead or behind, monthly targets across the year, one branch across the year, year-to-date cumulative, or share of total.'),
        spsrCmpRep: B('Salesperson', 'The salesperson drawn in the “one salesperson, whole FY” chart.'),
        spmrCmpRep: B('Salesperson', 'The salesperson drawn in the “one salesperson, whole FY” chart.'),
        brsrCmpRep: B('Branch', 'The branch drawn in the “one branch, whole FY” chart.'),
        // Forecast / decomposition / pareto controls
        fcMethod: B('Forecast method', 'Linear trend fits a straight line through the monthly history. 3-month moving average carries the average of the last three months forward. 3-month MA + trend does the same and adds the recent slope.'),
        fcHorizon: B('Horizon', 'How many months ahead to forecast.'),
        decompL1: B('Level 1', 'The first way the total is split. Each further level splits the one above it.'),
        decompL2: B('Level 2', 'Splits each level-1 group further. Leave blank to stop here.'),
        decompL3: B('Level 3', 'Splits each level-2 group further. Leave blank to stop here.'),
        decompL4: B('Level 4', 'Splits each level-3 group further. Leave blank to stop here.'),
        decompL5: B('Level 5', 'Splits each level-4 group further. Leave blank to stop here.'),
        rhStatusFilter: B('Status', 'Show only customers with this status: New is a first purchase in the last 90 days, Returning came back after a gap of more than 180 days, Active buys regularly, Lost has not bought for more than 180 days.'),
        cxiView: B('View by', 'Customer → items lists the items for one customer; Item → customers lists the customers for one item.'),
        cxiCustomer: B('Customer', 'The customer in scope for the Customer → items view.'),
        cxiItem: B('Item', 'The item in scope for the Item → customers view.'),
        cxiSalesperson: B('Salesperson', 'Narrows the universe to customers who belong to this salesperson on their customer card.'),
        cxiStatus: B('Show', 'Buyers bought the item in the period; non-buyers are in the universe but did not.'),
        diRadarPeriod: B('Radar period', 'Which slice of the loaded data the salesperson radar uses.'),
        diRfmTopN: B('Top N', 'How many customers to include in the RFM segmentation.'),
        diHeatmapTopN: B('Top N', 'How many customers to show on the profit density heatmap.'),
        diSparkTopN: B('Top N', 'How many customers to show in the sparkline grid.'),
        pnlMonthlyType: B('Chart type', 'How to draw the month-by-month P&L.'),
        pnlBreakGroup: B('Breakdown group', 'Which part of the statement the breakdown chart draws: income, COGS, expense accounts, or the four totals.'),
        pnlBreakType: B('Chart type', 'How to draw the breakdown by account.'),
        pnlBreakTop: B('Top N', 'How many accounts to show before the rest are grouped as Other.'),
    };

    // ---- KPI tiles, matched on the tile label ----
    const KPIS = [
        [/^sales \(excl\. tax\)$/i, B('Sales (excl. tax)', 'Posted sales invoice lines in the period, net of invoice discount, before GST. Credit memos are not taken off here; see Adjusted Sales for that.')],
        [/^gross profit \$$/i, B('Gross Profit $', 'Sales less the item cost of what was sold. Cost comes from the Wiise value entries, so it is the real posted cost, not the unit cost on the line.')],
        [/^gross profit %$/i, B('Gross Profit %', 'Gross Profit $ divided by Sales.')],
        [/^quantity sold$/i, B('Quantity sold', 'Units on the invoice lines in the period.')],
        [/^adjusted sales$/i, B('Adjusted Sales', 'Sales less posted credit memos (returns and credits) in the period. The figure to compare with the ledger.')],
        [/^adjusted cost$/i, B('Adjusted Cost', 'Item cost of what was sold, with the cost of returned stock taken back out.')],
        [/^adjusted profit$/i, B('Adjusted Profit', 'Adjusted Sales less Adjusted Cost.')],
        [/^adjusted profit margin$/i, B('Adjusted Profit Margin', 'Adjusted Profit divided by Adjusted Sales.')],
        [/^outstanding sales orders$/i, B('Outstanding Sales Orders', 'Open sales orders in Wiise right now, not yet shipped. Current state, not tied to the Period.')],
        [/^shipped not invoiced$/i, B('Shipped Not Invoiced', 'Goods shipped on sales orders that have not been invoiced yet. Current state, not tied to the Period.')],
        [/^outstanding invoices$/i, B('Outstanding Invoices', 'Posted invoices not yet paid. Current state, not tied to the Period.')],
        [/^revenue$/i, B('Revenue', 'Sales for this customer in the period, net of invoice discount, before GST.')],
        [/^gp\$$/i, B('GP $', 'Sales less item cost for this customer in the period.')],
        [/^invoices$/i, B('Invoices', 'Number of posted invoices in the period.')],
        [/^total selling days$/i, B('Total Selling Days', 'Working days in the month: Monday to Friday, less Victorian public holidays.')],
        [/^selling days passed$/i, B('Selling Days Passed', 'Working days from the 1st up to and including today. For a past month it equals the total.')],
        [/^remaining selling days$/i, B('Remaining Selling Days', 'Working days left after today in the month.')],
        [/^% month passed$/i, B('% Month Passed', 'Selling days passed divided by total selling days.')],
        [/^universe size$/i, B('Universe size', 'How many customer-and-item pairs are in scope for the chosen view and filters.')],
        [/^buyers$/i, B('Buyers', 'Pairs where the customer did buy the item in the period.')],
        [/^non-buyers$/i, B('Non-buyers', 'Pairs where the customer did not buy the item in the period. Prospects for a follow-up.')],
        [/^sales captured$/i, B('Sales captured', 'Sales in the period from the buyers in scope.')],
        [/^sales amount$/i, B('Sales Amount', 'Sales on the lines in the report, net of invoice discount, before GST.')],
        [/^sales quantity$/i, B('Sales Quantity', 'Units on the lines in the report.')],
        [/^gross profit$/i, B('Gross Profit', 'Sales less item cost on the lines in the report.')],
        [/^gross profit margin$/i, B('Gross Profit Margin', 'Gross Profit divided by Sales.')],
        [/^total quote value$/i, B('Total Quote Value', 'Value of every quote in the period, open and closed.')],
        [/^open quote value$/i, B('Open Quote Value', 'Value of quotes still open.')],
        [/^no\. of quotes$/i, B('No. of Quotes', 'How many quotes were created in the period.')],
        [/^quote quantity$/i, B('Quote Quantity', 'Units across all quote lines.')],
        [/^avg quote value$/i, B('Avg Quote Value', 'Total quote value divided by the number of quotes.')],
        [/^won$/i, B('Won', 'Quotes or orders that turned into a sale.')],
        [/^lost$/i, (tab) => tab === 'retention' || tab === 'retentionHistory' ? B('Lost', 'Customers who have not bought for more than 180 days.') : B('Lost', 'Quotes or orders marked lost.')],
        [/^archived \(historic\)$/i, B('Archived (historic)', 'Quotes that have been archived in Wiise, kept for history.')],
        [/^open \(in progress\)$/i, B('Open (in progress)', 'Quotes still being worked on.')],
        [/^win rate$/i, B('Win rate', 'Won divided by won plus lost.')],
        [/^total blanket order value$/i, B('Total Blanket Order Value', 'Value of every blanket order in the period.')],
        [/^no\. of blanket orders$/i, B('No. of Blanket Orders', 'How many blanket orders were created in the period.')],
        [/^order quantity$/i, B('Order Quantity', 'Units across all blanket order lines.')],
        [/^avg order value$/i, B('Avg Order Value', 'Total value divided by the number of orders.')],
        [/^follow-up$/i, B('Follow-up', 'Blanket orders marked for follow-up.')],
        [/^uncategorised$/i, B('Uncategorised', 'Blanket orders with no status set.')],
        [/^return rate \(amount\)$/i, B('Return Rate (Amount)', 'Credit memo value divided by sales in the period.')],
        [/^total credit \(amount\)$/i, B('Total Credit (Amount)', 'Value of posted credit memos in the period, shown as a negative.')],
        [/^total credit \(qty\)$/i, B('Total Credit (Qty)', 'Units returned on credit memos.')],
        [/^credit adj profit$/i, B('Credit Adj Profit', 'The GP lost through credits: credit value less the cost that came back into stock.')],
        [/^total unique customers$/i, B('Total unique customers', 'Distinct customers invoiced in the period.')],
        [/^one-time customers$/i, B('One-time customers', 'Customers invoiced in only one month of the period.')],
        [/^repeat customers$/i, B('Repeat customers', 'Customers invoiced in more than one month of the period.')],
        [/^total$/i, (tab) => tab === 'retentionHistory' ? B('Total', 'All customers with at least one invoice in the loaded data.') : null],
        [/^active$/i, B('Active', 'Customers buying regularly, with no long gap.')],
        [/^new \(last 90 days\)$/i, B('New (last 90 days)', 'Customers whose first ever purchase was in the last 90 days.')],
        [/^returning$/i, B('Returning', 'Customers who came back after a gap of more than 180 days.')],
        [/^sales$/i, (tab) => tab === 'branchTarget' ? B('Sales', 'Sales in the loaded period across the branches shown, net of invoice discount, before GST.') : null],
        [/^income$/i, B('Income', 'Every ledger entry in the period on accounts Wiise classes as Income, netted credits less debits. Straight from the ledger, no sales-tool adjustments.')],
        [/^cost of goods sold$/i, B('Cost of Goods Sold', 'Ledger entries on accounts Wiise classes as Cost of Goods Sold, netted debits less credits. Raw ledger COGS, including purchase and accrual postings, not just item cost.')],
        [/^gross margin$/i, B('Gross Margin', 'Gross Profit divided by Income.')],
        [/^expenses$/i, B('Expenses', 'Ledger entries on accounts Wiise classes as Expense: wages, rent, vehicles, bank fees and so on. Accounts with no category are listed as Unclassified instead.')],
        [/^net profit$/i, B('Net Profit', 'Gross Profit less Expenses. Balance-sheet accounts never enter this figure.')],
    ];

    // ---- Column headers, matched on the header text (sort arrows stripped) ----
    const COLS = [
        [/^fiscal year$/i, B('Fiscal Year', 'July to June. FY 2027 runs from July 2026 to June 2027.')],
        [/^month$/i, B('Month', 'The posting month of the invoices in the row.')],
        [/^branch$/i, B('Branch', 'The Wiise location code on the invoice line, with the branch name: 10 Keilor, 20 Shepparton, 30 Dandenong. “(no location)” is lines with no branch, such as freight and charges.')],
        [/^salesperson$/i, B('Salesperson', 'The salesperson code on the invoice header, the person’s initials in Wiise.')],
        [/^sales \$$/i, B('Sales $', 'Invoice lines net of invoice discount, before GST.')],
        [/^gp \$$/i, B('GP $', 'Sales less the item cost from the Wiise value entries.')],
        [/^gp %$/i, B('GP %', 'GP $ divided by Sales $.')],
        [/^qty$/i, B('Qty', 'Units on the invoice lines.')],
        [/^invoices$/i, B('Invoices', 'Number of posted invoices.')],
        [/^customers$/i, B('Customers', 'Number of distinct customers.')],
        [/^customer no$/i, B('Customer No', 'The Wiise customer number. Click the row to open the customer.')],
        [/^name$/i, B('Name', 'The customer or account name.')],
        [/^last invoice$/i, B('Last invoice', 'Posting date of the most recent invoice in the loaded data.')],
        [/^item no$/i, B('Item No', 'The Wiise item number. Click the row to see who bought it.')],
        [/^description$/i, B('Description', 'The item or account description from Wiise.')],
        [/^category$/i, B('Category', 'The item category from the Wiise item card.')],
        [/^item category$/i, B('Item Category', 'The item category from the Wiise item card.')],
        [/^lines$/i, B('Lines', 'How many invoice lines carried this item.')],
        [/^sales \(lcy\)$/i, B('Sales (LCY)', 'Sales in local currency, net of invoice discount, before GST.')],
        [/^adj cost$/i, B('Adj Cost', 'Item cost from the value entries, with the cost of returned stock taken back out.')],
        [/^cost non-inv$/i, B('Cost Non-Inv', 'Cost not tied to stock: item charges, freight and similar postings.')],
        [/^adj profit$/i, B('Adj Profit', 'Sales less Adj Cost and Cost Non-Inv.')],
        [/^margin %$/i, B('Margin %', 'Adj Profit divided by Sales.')],
        [/^% of gt$/i, B('% of GT', 'Share of the grand total at the top of the tree.')],
        [/^customer$/i, B('Customer', 'The customer on the document.')],
        [/^invoice no$/i, B('Invoice No', 'The posted invoice number. The arrow opens it in Wiise.')],
        [/^unit sell$/i, B('Unit Sell', 'Sell price per unit on the line, after line discount.')],
        [/^unit cost$/i, B('Unit Cost', 'Cost per unit on the line.')],
        [/^fy$/i, B('FY', 'Fiscal year, July to June.')],
        [/^quotes$/i, B('Quotes', 'Number of quotes.')],
        [/^status$/i, (tab) => /Rate$/.test(tab || '') ? B('Status', 'ON TRACK when today’s actual run rate is at or above the target run rate; BEHIND when it is under. For a finished month: TARGET MET or TARGET MISSED. No Plan means there is no target for this row.') : B('Status', 'The document status in Wiise.')],
        [/^assigned user$/i, B('Assigned User', 'The Wiise user the quote or order is assigned to.')],
        [/^campaign$/i, B('Campaign', 'The campaign code on the quote or order.')],
        [/^orders$/i, B('Orders', 'Number of blanket orders.')],
        [/^value \(lcy\)$/i, B('Value (LCY)', 'Order value in local currency.')],
        [/^order date$/i, B('Order date', 'The order date on the blanket order.')],
        [/^cr no$/i, B('CR No', 'The posted credit memo number. The arrow opens it in Wiise.')],
        [/^posting date$/i, B('Posting Date', 'The date the document was posted in Wiise.')],
        [/^credit amount$/i, B('Credit Amount', 'Value credited to the customer, shown as a negative.')],
        [/^gp impact$/i, B('GP Impact', 'GP lost through the credit: credit value less the cost that came back into stock.')],
        [/^date$/i, B('Date', 'Calendar day.')],
        [/^sales$/i, (tab) => tab === 'mat' ? B('Sales', 'Sales posted in that month.') : B('Sales', 'Sales for the row, net of invoice discount, before GST.')],
        [/^mtd sales$/i, B('MTD Sales', 'Sales added up from the 1st of the month to this day.')],
        [/^mtd gp$/i, B('MTD GP', 'Gross profit added up from the 1st of the month to this day.')],
        [/^mat$/i, B('MAT', 'Moving annual total: the 12 months ending this month, added up.')],
        [/^py-mat$/i, B('PY-MAT', 'The same 12-month total one year earlier.')],
        [/^growth \$$/i, B('Growth $', 'This figure less the comparison figure.')],
        [/^growth %$/i, B('Growth %', 'Growth $ divided by the comparison figure.')],
        [/^prior month$/i, B('Prior Month', 'Sales in the month before.')],
        [/^period$/i, B('Period', 'The week, by its first day.')],
        [/^(mon|tue|wed|thu|fri|sat|sun)$/i, B('Weekday', 'Sales posted on that day of the week.')],
        [/^total$/i, (tab) => tab === 'dailySales' ? B('Total', 'Sales for the whole week.') : B('Total', 'The column total.')],
        [/^customer posting group$/i, B('Customer Posting Group', 'The posting group on the customer card in Wiise.')],
        [/^% of total$/i, B('% of total', 'Share of the total for the table.')],
        [/^total customers$/i, B('Total Customers', 'Distinct customers invoiced in the month.')],
        [/^first-in-period$/i, B('First-in-period', 'Customers whose first invoice in the loaded period was in this month.')],
        [/^returning$/i, B('Returning', 'Customers invoiced in this month who had already bought earlier in the period.')],
        [/^actual$/i, B('Actual', 'Sales posted in the month.')],
        [/^forecast$/i, B('Forecast', 'Sales projected for the month by the chosen method.')],
        [/^hierarchy$/i, B('Hierarchy', 'The tree: each level splits the one above it. Click a row to expand or collapse it.')],
        [/^first purchase$/i, B('First purchase', 'The customer’s first invoice date in the loaded data.')],
        [/^last purchase$/i, B('Last purchase', 'The customer’s most recent invoice date.')],
        [/^days since$/i, B('Days since', 'Days from the last purchase to today.')],
        [/^max gap \(days\)$/i, B('Max gap (days)', 'The longest stretch between two purchases.')],
        [/^why$/i, B('Why', 'The rule that gave the customer its status.')],
        [/^sales \$ target$/i, B('Sales $ Target', 'The budget target for this branch and month, read from the Branch Targets sheet.')],
        [/^sales vs target$/i, B('Sales vs Target', 'Sales $ less the target: plus when ahead, minus when behind.')],
        [/^% achieved$/i, B('% Achieved', 'Actual divided by target. Red under 50%, orange 50 to 99%, green at 100% and over.')],
        [/^sales for month py$/i, B('Sales For Month PY', 'Sales in the same month one year earlier, if that month is in the loaded data.')],
        [/^gm \$ for month py$/i, B('GM $ For Month PY', 'Gross margin in the same month one year earlier, if that month is in the loaded data.')],
        [/^sales target$/i, B('Sales Target', 'The budget target for the month, read from the targets file.')],
        [/^gm \$ target$/i, B('GM $ Target', 'The gross margin target for the month.')],
        [/^sales for month$/i, B('Sales For Month', 'Sales posted so far in the chosen month.')],
        [/^gm \$ for month$/i, B('GM $ For Month', 'Gross margin dollars posted so far in the chosen month: sales less item cost.')],
        [/^target run rate$/i, B('Target Run Rate', 'Target divided by the total selling days in the month: what is needed per working day.')],
        [/^actual run rate$/i, B('Actual Run Rate', 'Actual so far divided by the selling days passed: what has been done per working day.')],
        [/^balance$/i, B('Balance', 'Actual less target: plus when ahead of target, minus when behind.')],
        [/^required daily$/i, B('Required Daily', 'What is still needed per remaining selling day to reach the target. n/a once the month is over.')],
        [/^account$/i, B('Account', 'The G/L account number. The arrow opens its entries in Wiise.')],
        [/^entries$/i, B('Entries', 'Number of ledger entries on the account in the period.')],
        [/^debit$/i, B('Debit', 'Total debits posted to the account in the period.')],
        [/^credit$/i, B('Credit', 'Total credits posted to the account in the period.')],
        [/^net$/i, B('Net', 'Debits less credits, or the reverse for income accounts, so that income reads positive.')],
        [/^% of income$/i, B('% of income', 'The account’s net as a share of total income.')],
        [/^income$/i, B('Income', 'Income accounts for the month, netted the P&L way.')],
        [/^cogs$/i, B('COGS', 'Cost of goods sold accounts for the month.')],
        [/^gross profit$/i, B('Gross Profit', 'Income less COGS.')],
        [/^gm %$/i, B('GM %', 'Gross Profit divided by Income.')],
        [/^expenses$/i, B('Expenses', 'Expense accounts for the month.')],
        [/^net profit$/i, B('Net Profit', 'Gross Profit less Expenses.')],
        [/^unclassified$/i, B('Unclassified', 'Entries on accounts that have no category in Wiise. They are not in any P&L line until the account is categorised.')],
        [/^ledger check$/i, B('Ledger check', 'Debits should equal credits across every account including balance sheet. Shows the difference if they do not.')],
        [/^wiise category$/i, B('Wiise category', 'The Account Category on the Wiise chart of accounts.')],
        [/^treated as$/i, B('Treated as', 'Where the tool places the account: Income, COGS, Expense, Balance sheet or Unclassified.')],
        [/^net \(p&l sign\)$/i, B('Net (P&L sign)', 'The account’s movement in the period with the sign the P&L uses.')],
        [/^balance now$/i, B('Balance now', 'The account’s current balance in Wiise today, regardless of the period.')],
        [/\/ document|\/ item|\/ customer|\/ doc/i, B('Drill-down', 'A tree: click a row to expand it to the next level. Each row shows its share of the grand total.')],
    ];

    // ---- Chart cards and panel headings, matched on the heading text ----
    const HEADS = [
        [/^revenue over time$/i, B('Revenue over time', 'Sales by day, week or month across the period. Split by branch or salesperson to see who drove it. Click a point to filter the tab to that period.')],
        [/^top 10 customers$/i, B('Top 10 customers', 'The ten customers with the most sales in the period. Click a bar to open the customer; Show more lists them all.')],
        [/^top 10 items$/i, B('Top 10 items', 'The ten items with the most sales in the period. Click a bar to see who bought it.')],
        [/^sales by branch$/i, B('Sales by branch', 'Sales split by the Wiise location on each invoice line. Click a bar to filter by that branch.')],
        [/^sales by salesperson$/i, B('Sales by salesperson', 'Sales split by the salesperson code on each invoice. Click a bar to filter by that person.')],
        [/^invoice status$/i, B('Invoice status', 'Posted invoices by their status in Wiise, as a count or as dollars.')],
        [/^posting pattern by day of week$/i, B('Posting pattern by day of week', 'How much is invoiced on each weekday across the period.')],
        [/^all customers$/i, B('All customers', 'Every customer invoiced in the period with their sales and GP. Click a row to open the customer.')],
        [/^sales by month$/i, B('Sales by month', 'Monthly sales, stacked by branch or salesperson. Click a segment to filter.')],
        [/^top branches by sales$/i, B('Top branches by sales', 'Branches ranked by sales. Click a bar to filter.')],
        [/^top salespeople by sales$/i, B('Top salespeople by sales', 'Salespeople ranked by sales. Click a bar to filter.')],
        [/^gp \$ vs gp % by/i, B('GP $ vs GP %', 'Gross profit dollars as bars and the margin percent as a line, side by side.')],
        [/^invoices · qty · customers/i, B('Volume', 'Number of invoices, units sold and distinct customers, side by side.')],
        [/fiscal year × month$/i, B('Fiscal year by month', 'One row per branch or salesperson per month, with sales, GP, quantity and volume.')],
        [/^drill-down/i, B('Drill-down tree', 'Click a row to expand it to the next level, down to the item line. Each row shows sales, adjusted cost and profit and its share of the grand total.')],
        [/^top 20 items by sales$/i, B('Top 20 items by sales', 'The twenty items with the most sales. Click a bar to filter the table to that item.')],
        [/^sales by item category$/i, B('Sales by item category', 'Sales split by the item category on the item card. Click a slice to filter.')],
        [/^item pareto/i, B('Item Pareto', 'Items ranked by sales with the cumulative share as a line: how many items make up 80% of sales.')],
        [/^items$/i, B('Items', 'Every item sold in the period. Search, sort, and click a row to see who bought it.')],
        [/^buyer coverage by salesperson$/i, B('Buyer coverage by salesperson', 'For each salesperson, how many of their customers in the universe bought and how many did not.')],
        [/^monthly sales captured$/i, B('Monthly sales captured', 'Sales from the buyers in scope, month by month.')],
        [/^top items in scope$/i, B('Top items in scope', 'The items with the most sales within the current universe.')],
        [/^top customers in scope$/i, B('Top customers in scope', 'The customers with the most sales within the current universe.')],
        [/^eom report/i, B('EOM Report', 'Every invoice line in the period with sell, cost and GP. The GP flags highlight lines under 20% or over 90% margin for checking.')],
        [/^quote status$/i, B('Quote status', 'Quotes by status, as a count or as dollars.')],
        [/^quoting trend$/i, B('Quoting trend', 'Quotes created over time, by count or value.')],
        [/^top 10 customers quoted$/i, B('Top 10 customers quoted', 'Customers with the most quote value.')],
        [/^top 10 items quoted$/i, B('Top 10 items quoted', 'Items with the most quote value.')],
        [/^order status$/i, B('Order status', 'Blanket orders by status.')],
        [/^blanket order trend$/i, B('Blanket order trend', 'Blanket orders created over time, by count or value.')],
        [/^top 10 customers \(blanket orders\)$/i, B('Top 10 customers', 'Customers with the most blanket order value.')],
        [/^top 10 items \(blanket orders\)$/i, B('Top 10 items', 'Items with the most blanket order value.')],
        [/^by salesperson$/i, B('By salesperson', 'Split by the salesperson code on the document.')],
        [/^by item category$/i, B('By item category', 'Split by the item category of the lines.')],
        [/^by assigned user$/i, B('By assigned user', 'Split by the Wiise user the document is assigned to.')],
        [/^by campaign$/i, B('By campaign', 'Split by the campaign code on the document.')],
        [/by customer · click/i, B('By customer', 'One row per customer. Click the arrow to expand to the documents and then the lines.')],
        [/^return rate \(amount\) by customer$/i, B('Return rate by customer', 'The customers with the most credit value, with their return rate against sales. Click a bar to open the customer.')],
        [/^total credit \(amount\) by item$/i, B('Credit by item', 'The items with the most credit value. Click a bar to filter.')],
        [/^credit memo detail$/i, B('Credit memo detail', 'Every posted credit memo in the period.')],
        [/^month-to-date sales performance/i, B('Month-to-date', 'Sales building up day by day from the 1st of this month.')],
        [/^mtd details/i, B('MTD details', 'One row per day of the month with the day’s sales and the running totals.')],
        [/^monthly comparison: mat/i, B('MAT vs prior year', 'The rolling 12-month total each month against the same total a year earlier.')],
        [/^moving annual total details$/i, B('MAT details', 'The rolling totals month by month with the growth.')],
        [/^monthly sales — current vs prior month$/i, B('Month vs prior month', 'Each month’s sales beside the month before.')],
        [/^period-over-period growth details$/i, B('Growth details', 'Month-on-month growth in dollars and percent, with the GP for the month.')],
        [/^daily sales by date$/i, B('Daily sales', 'One row per week, one column per weekday.')],
        [/^sales by customer city$/i, B('Sales by customer city', 'Sales split by the city on the customer card.')],
        [/^sales by customer posting group$/i, B('Sales by customer posting group', 'Sales split by the posting group on the customer card.')],
        [/^first-appearance vs returning/i, B('New vs returning', 'Each month: customers buying for the first time in the period against those who had bought before.')],
        [/^monthly retention detail/i, B('Monthly retention detail', 'The same figures as a table. Click a row to see the customers.')],
        [/^actual \+ forecast/i, B('Actual and forecast', 'Monthly sales to date, then the projection. Actual months solid, forecast months dashed.')],
        [/^forecast detail$/i, B('Forecast detail', 'The month-by-month figures behind the chart.')],
        [/^decomposition tree$/i, B('Decomposition tree', 'The total split through the chosen levels. Click a row to expand or collapse it.')],
        [/^customer retention detail/i, B('Customer retention detail', 'One row per customer with purchase dates, gaps and status. Click a row to open the customer.')],
        [/sales vs targets$/i, B('Sales vs targets', 'Every loaded month by branch against the budget target. Fiscal Year and Month headers flip the month order; the other headers sort the branches within each month.')],
        [/^sales run rate · by/i, B('Sales run rate', 'For the chosen month: target, actual so far, both as a per-day rate, the gap, what is still needed per day, and the percent achieved. Click a header to sort.')],
        [/^margin run rate · by/i, B('Margin run rate', 'For the chosen month: gross margin target, actual so far, both as a per-day rate, the gap, what is still needed per day, and the percent achieved.')],
        [/^comparisons ·/i, B('Comparisons', 'Charts of actual against budget. Pick a view in the dropdown: this month by person or branch, percent achieved, ahead or behind, monthly targets across the year, one person or branch across the year, year-to-date cumulative, or share of total.')],
        [/^ledger$/i, B('Ledger', 'What is loaded from the general ledger and the controls: Wiise categories or the mapping table, opening balances, all months.')],
        [/^month by month$/i, B('Month by month', 'Income, COGS, gross profit, expenses and net profit for each month in the period.')],
        [/^breakdown by account$/i, B('Breakdown by account', 'The chosen section of the statement drawn account by account.')],
        [/^profit & loss statement$/i, B('Profit & Loss statement', 'Income less COGS gives gross profit; less expenses gives net profit. Click a category to open its accounts, and an account to see its entries.')],
        [/^monthly summary$/i, B('Monthly summary', 'The statement totals for each month. Click a month to set the period to it.')],
        [/^account map/i, B('Account map', 'Every G/L account with its Wiise category, where the tool places it, and its balance now.')],
        [/^sales flow/i, B('Sales flow', 'Sales flowing from each salesperson into the item categories they sold. Thicker bands are more sales.')],
        [/^sales hierarchy/i, B('Sales hierarchy', 'Rings from the centre out: branch, salesperson, category. Larger arcs are more sales.')],
        [/^salesperson performance radar$/i, B('Salesperson radar', 'Each salesperson on several measures at once: sales, GP, margin, customers, invoices, average sale.')],
        [/^sales pipeline funnel$/i, B('Pipeline funnel', 'From quotes through orders to invoices, showing where value drops off.')],
        [/^monthly revenue ridges/i, B('Revenue ridges', 'One ridge per category: monthly sales stacked so seasonality is easy to see.')],
        [/^sales treemap/i, B('Treemap', 'Categories as blocks sized by sales, split into their items.')],
        [/^customer concentration$/i, B('Customer concentration', 'Sales ranked by customer with the cumulative share: how much depends on the top few.')],
        [/^item bubble matrix$/i, B('Item bubble matrix', 'Each item as a bubble: quantity across, GP% up, size for sales.')],
        [/^period-over-period waterfall$/i, B('Waterfall', 'How sales moved from one period to the next, customer by customer or category by category.')],
        [/^top customers · slope chart$/i, B('Slope chart', 'Top customers in the earlier period against the later one: lines going up grew, down shrank.')],
        [/^seasonality polar chart$/i, B('Seasonality', 'Sales by calendar month around the clock face.')],
        [/^invoice value distribution$/i, B('Invoice value distribution', 'How many invoices fall into each value band.')],
        [/^rfm customer segmentation$/i, B('RFM segmentation', 'Customers scored on Recency, Frequency and Monetary value, then grouped into segments such as champions and at-risk.')],
        [/^profit density heatmap/i, B('Profit density heatmap', 'Customers across, categories down, colour for GP.')],
        [/^customer lifecycle stacked area$/i, B('Customer lifecycle', 'Each month’s sales split into new, returning and lapsed-then-returned customers.')],
        [/^geographic sales/i, B('Geographic sales', 'Sales by state from the customer address.')],
        [/^top customers · sparkline grid$/i, B('Sparkline grid', 'A small monthly trend line for each top customer.')],
        [/^daily sales calendar heatmap$/i, B('Calendar heatmap', 'Every day of the period coloured by its sales.')],
        [/^customer cohort retention$/i, B('Cohort retention', 'Customers grouped by their first month, and how many of each group kept buying in the months after.')],
        [/^sales over time$/i, B('Sales over time', 'This customer’s or item’s sales by day, week or month.')],
    ];

    // ---- Buttons, matched on text or id ----
    const BTNS = [
        [/^load sales$/i, IDS.loadBtn],
        [/^fast lookup/i, IDS.loadLookupBtn],
        [/^theme$/i, IDS.themeBtn],
        [/^sign out$/i, IDS.signOutBtn],
        [/^day$/i, B('Day', 'Group the chart by day.')],
        [/^week$/i, B('Week', 'Group the chart by week.')],
        [/^month$/i, B('Month', 'Group the chart by month.')],
        [/^gross$/i, B('Gross', 'Sales as invoiced, before credit memos.')],
        [/^adj$/i, B('Adj', 'Adjusted: sales less posted credit memos (returns).')],
        [/^show more…$/i, B('Show more', 'Opens the full list behind the chart, searchable and sortable.')],
        [/^hide$/i, B('Hide', 'Closes the list.')],
        [/^count$/i, B('Count', 'Show the number of documents.')],
        [/^\$ amount$/i, B('$ Amount', 'Show the dollar value.')],
        [/^close$/i, B('Close', 'Closes this view and goes back.')],
        [/^load ledger$/i, B('Load ledger', 'Reads the general ledger entries for the toolbar period from Wiise. Needs G/L read rights on your Wiise login.')],
        [/^gp < 20%$/i, B('GP < 20%', 'Show only lines with a gross margin under 20%: possible pricing or cost errors.')],
        [/^gp > 90%$/i, B('GP > 90%', 'Show only lines with a gross margin over 90%: possible missing cost.')],
        [/^export to excel$/i, B('Export to Excel', 'Downloads the table as an Excel file.')],
        [/^export csv$/i, B('Export CSV', 'Downloads the table as a CSV file.')],
        [/^update forecast$/i, B('Update forecast', 'Recalculates the projection with the chosen method and horizon.')],
        [/^build tree$/i, B('Build tree', 'Builds the decomposition tree with the levels chosen above.')],
        [/^wiise$/i, B('Wiise categories', 'Lay the statement out by the Account Category on the Wiise chart of accounts.')],
        [/^mapping$/i, B('Mapping table', 'Lay the statement out by the finance mapping table: revenue, COGS, direct labour, operating expense buckets, other income, non-operating.')],
        [/^← back$/i, B('Back', 'Returns to the period you had before clicking a month.')],
        [/^↻$/i, B('Reload targets', 'Reads the Sales Budgets-Targets folder on SharePoint again and rebuilds the targets.')],
        [/^apply$/i, B('Apply', 'Uses the ticked files and redraws the tab.')],
        [/^reset$/i, B('Reset', 'Clears the filters.')],
    ];

    const BADGES = [
        [/^on track$/i, B('ON TRACK', 'Today’s actual run rate is at or above the target run rate.')],
        [/^behind/i, B('BEHIND', 'The actual run rate is under the target run rate. Required Daily shows what is needed per remaining day to catch up.')],
        [/^target met$/i, B('TARGET MET', 'The month is over and the actual reached the target.')],
        [/^target missed$/i, B('TARGET MISSED', 'The month is over and the actual fell short of the target.')],
        [/^no plan$/i, B('No Plan', 'There is no target for this row in the targets file.')],
        [/^not started$/i, B('Not started', 'The chosen month is in the future.')],
        [/^paid$/i, B('Paid', 'The invoice has been paid.')],
        [/^open$/i, B('Open', 'The invoice is posted and unpaid.')],
        [/^draft$/i, B('Draft', 'Not posted yet.')],
        [/^wip$/i, B('WIP', 'Work in progress: this tab is not built yet.')],
    ];

    // ---- Lookup ----
    const text = el => (el.textContent || '').replace(/[▲▼⇅]/g, '').replace(/\s+/g, ' ').trim();
    const matchList = (list, t, tab) => {
        for (const [re, h] of list) if (re.test(t)) { const v = typeof h === 'function' ? h(tab) : h; if (v) return v; }
        return null;
    };
    function tabOf(el) {
        const sec = el.closest && el.closest('section[data-tab-content]');
        if (sec) return sec.dataset.tabContent;
        return (typeof state !== 'undefined' && state && state.tab) || '';
    }
    function helpFor(target) {
        if (!target || !target.closest) return null;
        const tab = tabOf(target);
        // Sidebar / tab strip
        const navA = target.closest('.sidebar a[data-tab], #tabStrip [data-ts-tab]');
        if (navA) { const k = navA.dataset.tab || navA.dataset.tsTab; if (TABS[k]) return TABS[k]; }
        const grp = target.closest('.sidebar .group-label, #tabStrip .ts-group');
        if (grp) return GROUPS[text(grp)] || B(text(grp), 'A group of tabs.');
        if (target.closest('.sidebar .brand')) return B('Sales Analytics', 'Sales reporting read live from Wiise (Business Central), or from the automatic snapshots.');
        // Badges and run-rate bits
        const badge = target.closest('.badge');
        if (badge) { const h = matchList(BADGES, text(badge), tab); if (h) return h; }
        if (target.closest('.rr-bar')) return matchList(COLS, '% Achieved', tab);
        if (target.closest('.rr-legend')) return B('Colour legend', 'The % Achieved bar is red under 50%, orange from 50 to 99%, green at 100% and over.');
        if (target.closest('.rr-tgt-pop')) return B('Targets files', 'The Excel files found in the Sales Budgets-Targets folder. Tick the ones to use, then Apply. Where two files cover the same month, the newer file wins.');
        if (target.closest('.rr-tgt')) return B('Targets', 'Which targets file is in use and when it was last changed. Click the name to choose files; ↻ reads the folder again. Every file in the folder is read and merged; the newest wins where months overlap.');
        if (target.closest('.rr-name')) return B('Name', 'The person or branch behind the code, from the targets file.');
        if (target.closest('.chip')) return B('Filter chip', 'A filter in force. Click × to clear it.');
        // Elements with a known id (walk up a few levels)
        let el = target;
        for (let i = 0; i < 6 && el && el !== document.body; i++, el = el.parentElement) {
            if (el.id && IDS[el.id]) return IDS[el.id];
            if (el.id && /MonthSel$/.test(el.id)) return IDS.spsrMonthSel;
            if (el.id && /CmpView$/.test(el.id)) return IDS.spsrCmpView;
            if (el.id && /TgtStatus$/.test(el.id)) return B('Targets file', 'The file in use and when it was last changed. Click to choose which files to use.');
        }
        // Tiles
        const kpi = target.closest('.kpi');
        if (kpi) {
            const lab = kpi.querySelector('.label');
            const h = lab && matchList(KPIS, text(lab), tab);
            if (h) return h;
            if (kpi.title) return B(text(lab) || 'Tile', kpi.title);
            const row = kpi.closest('.kpi-row'); if (row && row.title) return B(text(lab) || 'Tile', row.title);
        }
        // Column header, or a cell under one
        const th = target.closest('th');
        if (th) { const h = matchList(COLS, text(th), tab); if (h) return h; if (th.title) return B(text(th), th.title); }
        const td = target.closest('td');
        if (td) {
            const tr = td.parentElement, table = td.closest('table');
            const ths = table ? table.querySelectorAll('thead th') : [];
            const idx = tr ? Array.prototype.indexOf.call(tr.children, td) : -1;
            const hth = idx >= 0 && ths[idx];
            if (hth) { const h = matchList(COLS, text(hth), tab); if (h) return h; if (hth.title) return B(text(hth), hth.title); }
        }
        // Buttons, selects
        const btn = target.closest('button, .toggle button, .chart-view-toggle button');
        if (btn) {
            const h = matchList(BTNS, text(btn), tab); if (h) return h;
            if (btn.title) return B(text(btn) || 'Button', btn.title);
        }
        const sel = target.closest('select');
        if (sel) {
            if (/ChartType$|Type$/.test(sel.id)) return B('Chart type', 'How to draw this chart. The figures do not change.');
            if (/Split$/.test(sel.id)) return B('Split by', 'Stack the chart by branch or salesperson, or show it as one series.');
            if (/Metric$/.test(sel.id)) return B('Measure', 'Which figure the chart shows.');
            if (sel.title) return B('Option', sel.title);
            const lab = sel.id && document.querySelector('label[for="' + sel.id + '"]');
            if (lab && IDS[sel.id]) return IDS[sel.id];
        }
        const lab = target.closest('label[for]');
        if (lab && IDS[lab.htmlFor]) return IDS[lab.htmlFor];
        // Cards and panels: heading text
        const card = target.closest('.card, .panel');
        if (card) {
            const hd = card.querySelector('.head h3, .panel-header h2, h3, h2');
            const h = hd && matchList(HEADS, text(hd), tab);
            if (h) return h;
        }
        // Anything with its own title text
        const titled = target.closest('[title]');
        if (titled && titled.title) return B(text(titled).slice(0, 60) || 'Info', titled.title);
        return null;
    }

    // ---- Floating box ----
    let box = document.getElementById('helpTip');
    if (!box) { box = document.createElement('div'); box.id = 'helpTip'; document.body.appendChild(box); }
    let lastTarget = null, lastHtml = null, mx = 0, my = 0, raf = 0;
    function place() {
        raf = 0;
        if (box.style.display !== 'block') return;
        const w = box.offsetWidth, h = box.offsetHeight;
        let left = mx + 16, top = my + 18;
        if (left + w > window.innerWidth - 8) left = Math.max(8, mx - w - 12);
        if (top + h > window.innerHeight - 8) top = Math.max(8, my - h - 12);
        box.style.left = left + 'px'; box.style.top = top + 'px';
    }
    function hide() { box.style.display = 'none'; lastTarget = null; lastHtml = null; }
    document.addEventListener('mousemove', function (e) {
        if (!helpOn) return;
        mx = e.clientX; my = e.clientY;
        if (e.target !== lastTarget) {
            lastTarget = e.target;
            const h = helpFor(e.target);
            if (!h) { box.style.display = 'none'; lastHtml = null; }
            else if (h !== lastHtml) { box.innerHTML = h; lastHtml = h; box.style.display = 'block'; }
            else box.style.display = 'block';
        }
        if (!raf) raf = requestAnimationFrame(place);
    }, true);
    document.addEventListener('mouseleave', hide);
    window.addEventListener('blur', hide);

    window.toggleHelpMode = function (force) {
        helpOn = typeof force === 'boolean' ? force : !helpOn;
        document.body.classList.toggle('help-mode', helpOn);
        const b = document.getElementById('helpBtn');
        if (b) b.classList.toggle('on', helpOn);
        if (!helpOn) hide();
        else {
            const r = b ? b.getBoundingClientRect() : null;
            if (r) { mx = r.left; my = r.bottom; }
            box.innerHTML = B('Help is on', 'Hover over anything — tabs, filters, buttons, tiles, charts, column headers, numbers — to see what it means. Click ? again to turn help off.');
            box.style.display = 'block'; lastHtml = null; lastTarget = null; place();
        }
    };
    window.isHelpMode = () => helpOn;
})();
