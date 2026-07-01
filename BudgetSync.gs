/**
 * BudgetSync.gs — DCR Personal Budget Automation
 * ================================================
 * Architecture : CSV-only (Monzo + Amex BA Black)
 * Trigger      : Time-based (daily) + manual run
 * Alerts       : Email at 80% category spend
 * Monthly brief: Scheduled 1st of each month at 08:00
 *
 * SETUP INSTRUCTIONS (one-time):
 * 1. Open your Google Sheet → Extensions → Apps Script → paste this file
 * 2. Run setupTriggers() once to register all time-based triggers
 * 3. Create a folder in My Drive named "BudgetCSVs"
 * 4. Each month drop your CSVs there (naming convention below)
 *
 * CSV NAMING CONVENTION:
 *   Monzo   : monzo_YYYY-MM.csv   (e.g. monzo_2026-04.csv)
 *   Amex    : amex_YYYY-MM.csv    (e.g. amex_2026-04.csv)
 *
 * MONZO CSV COLUMNS  (standard export, no header changes needed):
 *   Transaction ID, Date, Time, Type, Name, Emoji, Category,
 *   Amount, Currency, Local amount, Local currency, Notes and #tags,
 *   Address, Receipt, Description, Category split, Money Out, Money In
 *
 * AMEX CSV COLUMNS (standard UK Amex export):
 *   Date, Description, Amount
 *   (Amount is negative for charges, positive for credits/refunds)
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────────

const CONFIG = {
  SHEET_ID      : SpreadsheetApp.getActiveSpreadsheet().getId(),
  DRIVE_FOLDER  : "BudgetCSVs",          // folder name in My Drive
  IFTTT_SHEET_NAME : "Monzo Transactions", // name of the spreadsheet created by IFTTT
  ALERT_EMAIL   : "daniel.cruz.rosso@gmail.com",
  ALERT_THRESHOLD: 0.80,                 // 80%
  CURRENCY      : "GBP",

  // Sheet tab names
  TABS: {
    TRANSACTIONS : "Transactions",
    BUDGET       : "Monthly Budget",
    CONFIG       : "Config",
    DASHBOARD    : "Dashboard",
  },

  // Transactions sheet column positions (1-indexed)
  TX_COLS: {
    DATE        : 1,  // A
    MERCHANT    : 2,  // B
    AMOUNT      : 3,  // C
    CATEGORY    : 4,  // D
    ACCOUNT     : 5,  // E
    MONTH       : 6,  // F
    YEAR        : 7,  // G
    TX_ID       : 8,  // H
  },

  // Valid categories (must match Config sheet dropdown)
  CATEGORIES: [
    "Dining", "Groceries", "Health", "Transport",
    "Shopping", "Entertainment", "Subscriptions",
    "Personal care", "Travel", "Misc", "Rent",
  ],

  // Merchant → category keyword map (extend as needed)
  MERCHANT_MAP: {
    // Rent
    "savills"         : "Rent",
    // Dining
    "deliveroo"       : "Dining",
    "uber eats"       : "Dining",
    "just eat"        : "Dining",
    "restaurant"      : "Dining",
    "cafe"            : "Dining",
    "coffee"          : "Dining",
    "starbucks"       : "Dining",
    "costa"           : "Dining",
    "pret"            : "Dining",
    "mcdonalds"       : "Dining",
    "nando"           : "Dining",
    "wagamama"        : "Dining",
    "five guys"       : "Dining",
    "barrafina"       : "Dining",
    "wetherspoon"     : "Dining",
    "marlborough"     : "Dining",
    "pelican"         : "Dining",
    "pub"             : "Dining",
    "arms"            : "Dining",
    "bar"             : "Dining",
    "inn"             : "Dining",
    "tavern"          : "Dining",
    "casa"            : "Dining",
    // Groceries
    "tesco"           : "Groceries",
    "sainsbury"       : "Groceries",
    "waitrose"        : "Groceries",
    "marks & spencer" : "Groceries",
    "m&s food"        : "Groceries",
    "ocado"           : "Groceries",
    "lidl"            : "Groceries",
    "aldi"            : "Groceries",
    "whole foods"     : "Groceries",
    // Health
    "gym"             : "Health",
    "fitness"         : "Health",
    "pharmacy"        : "Health",
    "boots"           : "Health",
    "dentist"         : "Health",
    "doctor"          : "Health",
    "physiotherapy"   : "Health",
    "nuffield"        : "Health",
    "barry"           : "Health",  // Barry's Bootcamp
    "equinox"         : "Health",
    // Transport (commuting, public transit, taxis)
    "tfl"             : "Transport",
    "uber"            : "Transport",
    "bolt"            : "Transport",
    "addison lee"     : "Transport",
    "national rail"   : "Transport",
    "trainline"       : "Transport",
    "heathrow"        : "Transport",
    "gatwick"         : "Transport",
    "hertz"           : "Transport",
    "avis"            : "Transport",
    // Travel & Accommodation (flights, hotels, airbnb)
    "airbnb"          : "Travel",
    "british airways" : "Travel",
    "ryanair"         : "Travel",
    "easyjet"         : "Travel",
    "turkish airlines": "Travel",
    "jetblue"         : "Travel",
    "airlines"        : "Travel",
    "airways"         : "Travel",
    // Shopping
    "amazon"          : "Shopping",
    "asos"            : "Shopping",
    "zara"            : "Shopping",
    "h&m"             : "Shopping",
    "selfridges"      : "Shopping",
    "john lewis"      : "Shopping",
    "apple store"     : "Shopping",
    "nike"            : "Shopping",
    "adidas"          : "Shopping",
    // Entertainment
    "cinema"          : "Entertainment",
    "odeon"           : "Entertainment",
    "vue"             : "Entertainment",
    "ticketmaster"    : "Entertainment",
    "eventbrite"      : "Entertainment",
    "theatre"         : "Entertainment",
    "bowling"         : "Entertainment",
    // Subscriptions & Bills
    "netflix"         : "Subscriptions",
    "spotify"         : "Subscriptions",
    "apple"           : "Subscriptions",
    "google one"      : "Subscriptions",
    "dropbox"         : "Subscriptions",
    "notion"          : "Subscriptions",
    "claude"          : "Subscriptions",
    "chatgpt"         : "Subscriptions",
    "linkedin"        : "Subscriptions",
    "nytimes"         : "Subscriptions",
    "economist"       : "Subscriptions",
    "ft.com"          : "Subscriptions",
    "islington council": "Subscriptions",
    "do energy"       : "Subscriptions",
    // Personal care
    "hairdresser"     : "Personal care",
    "barber"          : "Personal care",
    "salon"           : "Personal care",
    "spa"             : "Personal care",
    // Rent
    "rent"            : "Rent",
  },
};

// ─── ENTRY POINTS ────────────────────────────────────────────────────────────

/**
 * Run daily — picks up any new CSVs dropped in BudgetCSVs folder
 */
function syncCSVs() {
  const folder   = getDriveFolder_(CONFIG.DRIVE_FOLDER);
  const existing = getExistingTxIds_();
  let   imported = 0;

  // 1. Sync any manually dropped CSVs
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().toLowerCase();
    if (!name.endsWith(".csv")) continue;

    let rows = [];
    if (name.startsWith("monzo_")) {
      rows = parseMonzoCSV_(file);
    } else if (name.startsWith("amex_")) {
      rows = parseAmexCSV_(file);
    } else {
      Logger.log("Skipping unrecognised file: " + file.getName());
      continue;
    }

    // Deduplicate and append
    const newRows = rows.filter(r => !existing.has(r[CONFIG.TX_COLS.TX_ID - 1]));
    if (newRows.length > 0) {
      appendTransactions_(newRows);
      newRows.forEach(r => existing.add(r[CONFIG.TX_COLS.TX_ID - 1]));
      imported += newRows.length;
    }
  }

  // 2. Sync from IFTTT automated Monzo sheet
  try {
    const iftttImported = syncMonzoFromIFTTT_();
    imported += iftttImported;
  } catch (e) {
    Logger.log("Error running IFTTT sync: " + e.message);
  }

  if (imported > 0) {
    Logger.log("Sync completed. Total new transactions: " + imported);
  } else {
    Logger.log("No new transactions found.");
  }
}

/**
 * Run on 1st of each month at 08:00 — sends the CFO brief
 */
function sendMonthlyBrief() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const cfg    = ss.getSheetByName(CONFIG.TABS.CONFIG);
  const budget = ss.getSheetByName(CONFIG.TABS.BUDGET);
  const tx     = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);

  // Determine last month
  const now       = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const month     = lastMonth.getMonth() + 1;
  const year      = lastMonth.getFullYear();
  const monthName = Utilities.formatDate(lastMonth, "GMT", "MMMM yyyy");

  const analysis = buildMonthlyAnalysis_(month, year);
  const html     = buildBriefHTML_(analysis, monthName);

  GmailApp.sendEmail(
    CONFIG.ALERT_EMAIL,
    "💼 Budget Brief: " + monthName + " — CFO Analysis",
    "Please view this email in HTML format.",
    {
      htmlBody : html,
      name     : "DCR Budget System",
    }
  );

  Logger.log("Monthly brief sent for " + monthName);
}

/**
 * Manual trigger — check alerts right now
 */
function runAlertCheck() {
  checkBudgetAlerts_();
}

// ─── CSV PARSERS ─────────────────────────────────────────────────────────────

/**
 * Parse Monzo CSV export
 * Returns array of row arrays matching Transactions sheet columns
 */
function parseMonzoCSV_(file) {
  const raw  = file.getBlob().getDataAsString("UTF-8");
  const rows = Utilities.parseCsv(raw);
  if (rows.length < 2) return [];

  const header = rows[0].map(h => h.trim().toLowerCase());
  const idxId       = findCol_(header, ["transaction id", "id"]);
  const idxDate     = findCol_(header, ["date"]);
  const idxName     = findCol_(header, ["name", "merchant name"]);
  const idxAmtOut   = findCol_(header, ["money out"]);
  const idxAmtIn    = findCol_(header, ["money in"]);
  const idxAmount   = findCol_(header, ["amount"]);
  const idxCategory = findCol_(header, ["category"]);

  const results = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => !c.trim())) continue;

    const txId    = idxId >= 0 ? row[idxId].trim() : "monzo_" + i;
    const rawDate = idxDate >= 0 ? row[idxDate].trim() : "";
    const date    = parseDate_(rawDate);
    if (!date) continue;

    const merchant = idxName >= 0 ? row[idxName].trim() : "Unknown";

    // Amount: prefer Money Out column, fall back to Amount (negative = spend)
    let amount = 0;
    if (idxAmtOut >= 0 && row[idxAmtOut].trim()) {
      amount = Math.abs(parseFloat(row[idxAmtOut].replace(/[£,]/g, "")) || 0);
    } else if (idxAmount >= 0 && row[idxAmount].trim()) {
      const raw = parseFloat(row[idxAmount].replace(/[£,]/g, "")) || 0;
      // Monzo: negative = money out (spend)
      if (raw < 0) amount = Math.abs(raw);
    }

    if (amount === 0) continue; // skip zero/credit transactions

    const monzoCategory = idxCategory >= 0 ? row[idxCategory].trim() : "";
    const category      = mapCategory_(merchant, monzoCategory);

    results.push(buildTxRow_(date, merchant, amount, category, "Monzo", txId));
  }

  return results;
}

/**
 * Parse Amex UK CSV export
 * Columns: Date, Description, Amount
 * Amount negative = charge (spend), positive = credit/refund
 */
function parseAmexCSV_(file) {
  const raw  = file.getBlob().getDataAsString("UTF-8");
  const rows = Utilities.parseCsv(raw);
  if (rows.length < 2) return [];

  // Amex UK export sometimes has no header row — detect by checking row 0
  let startRow = 0;
  const firstCell = rows[0][0] ? rows[0][0].toLowerCase().trim() : "";
  if (firstCell === "date" || firstCell.includes("date")) {
    startRow = 1; // has header
  }

  const results = [];

  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length < 3 || row.every(c => !c.trim())) continue;

    const rawDate    = row[0].trim();
    const date       = parseDate_(rawDate);
    if (!date) continue;

    const merchant   = row[1].trim();
    const rawAmount  = parseFloat((row[2] || "0").replace(/[£,\s]/g, "")) || 0;

    // Amex: negative = charge (spend), skip positives (credits/payments)
    if (rawAmount >= 0) continue;
    const amount = Math.abs(rawAmount);

    const category = mapCategory_(merchant, "");
    // Unique ID: date + merchant slug + amount
    const txId = "amex_" + Utilities.formatDate(date, "GMT", "yyyyMMdd")
                         + "_" + merchant.replace(/\s+/g, "").substring(0, 10).toLowerCase()
                         + "_" + amount;

    results.push(buildTxRow_(date, merchant, amount, category, "Amex BA Black", txId));
  }

  return results;
}

// ─── CATEGORY MAPPING ────────────────────────────────────────────────────────

/**
 * Map merchant name + Monzo native category → our budget category
 */
function mapCategory_(merchant, monzoCategory) {
  const m = merchant.toLowerCase();

  // Check merchant keyword map first
  for (const [keyword, cat] of Object.entries(CONFIG.MERCHANT_MAP)) {
    if (m.includes(keyword)) return cat;
  }

  // Fall back to Monzo's own category
  if (monzoCategory) {
    const mc = monzoCategory.toLowerCase();
    if (mc === "eating out" || mc === "restaurants")  return "Dining";
    if (mc === "groceries" || mc === "supermarkets") return "Groceries";
    if (mc === "transport" || mc === "travel")        return "Transport";
    if (mc === "shopping" || mc === "clothing")       return "Shopping";
    if (mc === "health"   || mc === "medical")        return "Health";
    if (mc === "entertainment" || mc === "hobbies")   return "Entertainment";
    if (mc === "bills" || mc === "subscriptions")     return "Subscriptions";
    if (mc === "personal care")                       return "Personal care";
    if (mc === "rent"    || mc === "housing")         return "Rent";
  }

  return "Misc";
}

// ─── BUDGET ALERTS ───────────────────────────────────────────────────────────

function checkBudgetAlerts_() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const cfg    = ss.getSheetByName(CONFIG.TABS.CONFIG);
  const budgetSheet = ss.getSheetByName(CONFIG.TABS.BUDGET);
  const txSheet     = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);

  const month = cfg.getRange("C32").getValue();
  const year  = cfg.getRange("C33").getValue();
  const now   = new Date();

  // Read all budget rows from Monthly Budget sheet (rows 5–14 = 10 categories)
  const budgetData = budgetSheet.getRange(5, 1, 10, 7).getValues();
  const alerts     = [];

  budgetData.forEach(row => {
    const cat    = row[1]; // column B
    const budget = row[2]; // column C — budget target
    const actual = row[3]; // column D — SUMIFS actual
    if (!cat || !budget || budget === 0) return;

    const pct = actual / budget;
    if (pct >= CONFIG.ALERT_THRESHOLD && pct < 1.0) {
      alerts.push({ cat, actual, budget, pct });
    } else if (pct >= 1.0) {
      alerts.push({ cat, actual, budget, pct, over: true });
    }
  });

  if (alerts.length === 0) return;

  // Check if we already sent an alert today (use PropertiesService to avoid spam)
  const props  = PropertiesService.getScriptProperties();
  const sentKey = "alert_sent_" + year + "_" + month;
  const sent    = JSON.parse(props.getProperty(sentKey) || "[]");

  const newAlerts = alerts.filter(a => !sent.includes(a.cat));
  if (newAlerts.length === 0) return;

  // Send email
  const subject = newAlerts.some(a => a.over)
    ? "🔴 Budget Alert: Category Overspend — " + formatMonthYear_(month, year)
    : "⚠️ Budget Alert: Approaching Limit — " + formatMonthYear_(month, year);

  const html = buildAlertHTML_(newAlerts, month, year);
  GmailApp.sendEmail(CONFIG.ALERT_EMAIL, subject, "", { htmlBody: html, name: "DCR Budget System" });

  // Record sent
  newAlerts.forEach(a => sent.push(a.cat));
  props.setProperty(sentKey, JSON.stringify(sent));
  Logger.log("Budget alert sent for: " + newAlerts.map(a => a.cat).join(", "));
}

// ─── MONTHLY ANALYSIS ENGINE ─────────────────────────────────────────────────

/**
 * Aggregate transactions for a given month/year and build analysis object
 */
function buildMonthlyAnalysis_(month, year) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const tx   = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);
  const cfg  = ss.getSheetByName(CONFIG.TABS.CONFIG);

  // Read all transactions
  const lastRow = tx.getLastRow();
  if (lastRow < 3) return null;
  const data = tx.getRange(3, 1, lastRow - 2, 8).getValues();

  // Budget targets from Config
  const budgetMap = {};
  const budgetRangeData = cfg.getRange(4, 2, 11, 2).getValues(); // B4:C14 (includes new Travel row)
  budgetRangeData.forEach(r => { if (r[0]) budgetMap[r[0]] = r[1] || 0; });

  // Aggregate actuals
  const actuals = {};
  CONFIG.CATEGORIES.forEach(c => actuals[c] = 0);

  const prevMonth = month === 1 ? 12 : month - 1;
  const prevYear  = month === 1 ? year - 1 : year;

  const prevActuals = {};
  CONFIG.CATEGORIES.forEach(c => prevActuals[c] = 0);

  data.forEach(row => {
    const txDate = row[0];
    if (!txDate) return;
    const d = new Date(txDate);
    const m = d.getMonth() + 1;
    const y = d.getFullYear();
    const cat = row[3];
    const amt = parseFloat(row[2]) || 0;

    if (m === month && y === year) {
      if (actuals[cat] !== undefined) actuals[cat] += amt;
    }
    if (m === prevMonth && y === prevYear) {
      if (prevActuals[cat] !== undefined) prevActuals[cat] += amt;
    }
  });

  // Total lifestyle (excl. Rent)
  const lifestyleCats = CONFIG.CATEGORIES.filter(c => c !== "Rent");
  const totalActual   = lifestyleCats.reduce((s, c) => s + actuals[c], 0);
  const totalBudget   = lifestyleCats.reduce((s, c) => s + (budgetMap[c] || 0), 0);
  const totalPrev     = lifestyleCats.reduce((s, c) => s + prevActuals[c], 0);

  // Flags and recommendations
  const flags = generateFlags_(actuals, prevActuals, budgetMap, lifestyleCats);

  return {
    month, year, actuals, prevActuals, budgetMap,
    totalActual, totalBudget, totalPrev,
    flags,
  };
}

/**
 * Generate analyst flags — over budget, MoM spikes, quick wins
 */
function generateFlags_(actuals, prevActuals, budgetMap, cats) {
  const flags = [];

  cats.forEach(cat => {
    const actual = actuals[cat] || 0;
    const budget = budgetMap[cat] || 0;
    const prev   = prevActuals[cat] || 0;
    if (!budget) return;

    const pct    = actual / budget;
    const momChg = prev > 0 ? (actual - prev) / prev : null;

    if (pct > 1.0) {
      flags.push({
        type   : "over",
        cat,
        actual,
        budget,
        pct,
        momChg,
        message: `${cat} over budget by £${fmt_(actual - budget)} (+${Math.round((pct-1)*100)}%)`,
        rec    : `Review ${cat.toLowerCase()} transactions — consider reducing by £${fmt_(actual - budget)} next month.`,
      });
    } else if (momChg !== null && momChg > 0.30) {
      flags.push({
        type   : "spike",
        cat,
        actual,
        budget,
        pct,
        momChg,
        message: `${cat} up ${Math.round(momChg*100)}% vs last month (£${fmt_(prev)} → £${fmt_(actual)})`,
        rec    : `Spike in ${cat.toLowerCase()} — verify if one-off or new trend. If recurring, revise budget up by £${fmt_(actual - prev)}.`,
      });
    } else if (pct < 0.60 && actual > 0) {
      flags.push({
        type   : "underspend",
        cat,
        actual,
        budget,
        pct,
        momChg,
        message: `${cat} only ${Math.round(pct*100)}% used — £${fmt_(budget - actual)} unspent`,
        rec    : `Consider reallocating £${fmt_(Math.round((budget - actual) * 0.5))} from ${cat.toLowerCase()} to savings or another category.`,
      });
    }
  });

  // Sort: over budget first, then spikes, then underspend
  const order = { over: 0, spike: 1, underspend: 2 };
  flags.sort((a, b) => order[a.type] - order[b.type]);

  return flags.slice(0, 5); // max 5 flags per brief
}

// ─── EMAIL HTML BUILDERS ─────────────────────────────────────────────────────

function buildBriefHTML_(analysis, monthName) {
  if (!analysis) return "<p>No data available for this period.</p>";

  const { actuals, prevActuals, budgetMap, totalActual, totalBudget, totalPrev, flags } = analysis;
  const lifestyleCats = CONFIG.CATEGORIES.filter(c => c !== "Rent");
  const varAmt  = totalActual - totalBudget;
  const varSign = varAmt >= 0 ? "+" : "";
  const momChg  = totalPrev > 0 ? ((totalActual - totalPrev) / totalPrev * 100).toFixed(1) : "N/A";

  const tableRows = lifestyleCats.map(cat => {
    const actual = actuals[cat] || 0;
    const budget = budgetMap[cat] || 0;
    const pct    = budget > 0 ? actual / budget : 0;
    const bar    = Math.min(Math.round(pct * 100), 100);
    const color  = pct > 1.0 ? "#ef4444" : pct > 0.8 ? "#f59e0b" : "#22c55e";
    const status = pct > 1.0 ? "🔴 OVER" : pct > 0.8 ? "⚠️ WARN" : "✅ OK";
    return `
      <tr style="border-bottom:1px solid #1e3a5f;">
        <td style="padding:8px 12px;color:#94a3b8;font-size:13px;">${cat}</td>
        <td style="padding:8px 12px;text-align:right;color:#e2e8f0;font-family:monospace;">£${fmt_(actual)}</td>
        <td style="padding:8px 12px;text-align:right;color:#64748b;font-family:monospace;">£${fmt_(budget)}</td>
        <td style="padding:8px 12px;min-width:120px;">
          <div style="background:#1e3a5f;border-radius:4px;height:8px;width:100%;">
            <div style="background:${color};border-radius:4px;height:8px;width:${bar}%;"></div>
          </div>
        </td>
        <td style="padding:8px 12px;text-align:right;color:#94a3b8;font-size:12px;">${Math.round(pct*100)}%</td>
        <td style="padding:8px 12px;text-align:center;font-size:12px;">${status}</td>
      </tr>`;
  }).join("");

  const flagRows = flags.map((f, i) => {
    const icon  = f.type === "over" ? "🔴" : f.type === "spike" ? "⚠️" : "💡";
    return `
      <tr style="border-bottom:1px solid #1e3a5f;">
        <td style="padding:10px 12px;font-size:13px;">${icon}</td>
        <td style="padding:10px 12px;color:#e2e8f0;font-size:13px;">${f.message}</td>
        <td style="padding:10px 12px;color:#94a3b8;font-size:12px;">${f.rec}</td>
      </tr>`;
  }).join("");

  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#0a1628;font-family:'Segoe UI',Arial,sans-serif;">
  <div style="max-width:700px;margin:0 auto;padding:24px;">

    <!-- Header -->
    <div style="background:linear-gradient(135deg,#0f2952,#1a4480);border-radius:12px;padding:28px 32px;margin-bottom:20px;">
      <div style="font-size:11px;color:#64748b;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;">PRIVATE & CONFIDENTIAL</div>
      <h1 style="margin:0 0 4px;font-size:24px;font-weight:700;color:#f8fafc;">Budget Brief</h1>
      <div style="font-size:15px;color:#94a3b8;">${monthName} — CFO Analysis</div>
    </div>

    <!-- KPI Cards -->
    <div style="display:flex;gap:12px;margin-bottom:20px;">
      <div style="flex:1;background:#0f2952;border:1px solid #1e3a5f;border-radius:10px;padding:18px 20px;">
        <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Total Spend</div>
        <div style="font-size:26px;font-weight:700;color:#f8fafc;font-family:monospace;">£${fmt_(totalActual)}</div>
        <div style="font-size:12px;color:#94a3b8;margin-top:4px;">vs £${fmt_(totalBudget)} budget</div>
      </div>
      <div style="flex:1;background:#0f2952;border:1px solid #1e3a5f;border-radius:10px;padding:18px 20px;">
        <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Variance</div>
        <div style="font-size:26px;font-weight:700;color:${varAmt <= 0 ? "#22c55e" : "#ef4444"};font-family:monospace;">${varSign}£${fmt_(Math.abs(varAmt))}</div>
        <div style="font-size:12px;color:#94a3b8;margin-top:4px;">${varAmt <= 0 ? "Under budget" : "Over budget"}</div>
      </div>
      <div style="flex:1;background:#0f2952;border:1px solid #1e3a5f;border-radius:10px;padding:18px 20px;">
        <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">vs Prior Month</div>
        <div style="font-size:26px;font-weight:700;color:#f8fafc;font-family:monospace;">${momChg}%</div>
        <div style="font-size:12px;color:#94a3b8;margin-top:4px;">MoM change</div>
      </div>
    </div>

    <!-- Category Breakdown -->
    <div style="background:#0f2952;border:1px solid #1e3a5f;border-radius:10px;margin-bottom:20px;overflow:hidden;">
      <div style="padding:16px 20px;border-bottom:1px solid #1e3a5f;">
        <h2 style="margin:0;font-size:14px;font-weight:600;color:#f8fafc;letter-spacing:1px;text-transform:uppercase;">Category Breakdown</h2>
      </div>
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#0a1e3d;">
            <th style="padding:8px 12px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Category</th>
            <th style="padding:8px 12px;text-align:right;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Actual</th>
            <th style="padding:8px 12px;text-align:right;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Budget</th>
            <th style="padding:8px 12px;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Progress</th>
            <th style="padding:8px 12px;text-align:right;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Used</th>
            <th style="padding:8px 12px;text-align:center;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Status</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>

    <!-- Analyst Flags & Recommendations -->
    <div style="background:#0f2952;border:1px solid #1e3a5f;border-radius:10px;margin-bottom:20px;overflow:hidden;">
      <div style="padding:16px 20px;border-bottom:1px solid #1e3a5f;">
        <h2 style="margin:0;font-size:14px;font-weight:600;color:#f8fafc;letter-spacing:1px;text-transform:uppercase;">Analyst Flags &amp; Recommendations</h2>
      </div>
      ${flags.length > 0 ? `
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#0a1e3d;">
            <th style="padding:8px 12px;width:30px;"></th>
            <th style="padding:8px 12px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Finding</th>
            <th style="padding:8px 12px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Recommendation</th>
          </tr>
        </thead>
        <tbody>${flagRows}</tbody>
      </table>` : `<p style="padding:20px;color:#64748b;font-size:13px;">No significant flags this month. Strong budget discipline maintained.</p>`}
    </div>

    <!-- Footer -->
    <div style="text-align:center;padding:16px;color:#334155;font-size:11px;">
      Generated by DCR Budget System · ${new Date().toLocaleDateString("en-GB", { weekday:"long",year:"numeric",month:"long",day:"numeric" })}
      <br>To update budgets or review transactions, open DCR_Budget_Tracker.xlsx
    </div>
  </div>
</body>
</html>`;
}

function buildAlertHTML_(alerts, month, year) {
  const monthName = formatMonthYear_(month, year);
  const rows = alerts.map(a => {
    const icon  = a.over ? "🔴" : "⚠️";
    const label = a.over ? "OVER BUDGET" : "APPROACHING LIMIT";
    const color = a.over ? "#ef4444" : "#f59e0b";
    const pctStr = Math.round(a.pct * 100) + "%";
    return `
      <tr style="border-bottom:1px solid #1e3a5f;">
        <td style="padding:10px 14px;font-size:14px;">${icon}</td>
        <td style="padding:10px 14px;color:#e2e8f0;font-weight:600;">${a.cat}</td>
        <td style="padding:10px 14px;text-align:right;font-family:monospace;color:#e2e8f0;">£${fmt_(a.actual)}</td>
        <td style="padding:10px 14px;text-align:right;font-family:monospace;color:#64748b;">£${fmt_(a.budget)}</td>
        <td style="padding:10px 14px;text-align:right;font-family:monospace;color:${color};font-weight:700;">${pctStr}</td>
        <td style="padding:10px 14px;"><span style="background:${color};color:#fff;border-radius:4px;padding:2px 8px;font-size:11px;font-weight:700;">${label}</span></td>
      </tr>`;
  }).join("");

  return `<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#0a1628;font-family:'Segoe UI',Arial,sans-serif;">
  <div style="max-width:640px;margin:0 auto;padding:24px;">
    <div style="background:linear-gradient(135deg,#0f2952,#1a4480);border-radius:12px;padding:24px 28px;margin-bottom:18px;">
      <h1 style="margin:0 0 4px;font-size:22px;color:#f8fafc;">Budget Alert</h1>
      <div style="font-size:14px;color:#94a3b8;">${monthName}</div>
    </div>
    <div style="background:#0f2952;border:1px solid #1e3a5f;border-radius:10px;overflow:hidden;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#0a1e3d;">
            <th style="padding:8px 14px;width:30px;"></th>
            <th style="padding:8px 14px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Category</th>
            <th style="padding:8px 14px;text-align:right;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Actual</th>
            <th style="padding:8px 14px;text-align:right;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Budget</th>
            <th style="padding:8px 14px;text-align:right;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Used</th>
            <th style="padding:8px 14px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Status</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <div style="text-align:center;padding:14px;color:#334155;font-size:11px;">
      DCR Budget System · Check DCR_Budget_Tracker.xlsx for full detail
    </div>
  </div>
</body>
</html>`;
}

// ─── SHEET HELPERS ────────────────────────────────────────────────────────────

function appendTransactions_(rows) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);

  // Apply date format to column A for new rows
  sheet.getRange(lastRow + 1, 1, rows.length, 1)
       .setNumberFormat("dd/mm/yyyy");
}

function buildTxRow_(date, merchant, amount, category, account, txId) {
  return [
    date,
    merchant,
    amount,
    category,
    account,
    date.getMonth() + 1,
    date.getFullYear(),
    txId,
  ];
}

function getExistingTxIds_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return new Set();
  const ids = sheet.getRange(3, CONFIG.TX_COLS.TX_ID, lastRow - 2, 1).getValues();
  return new Set(ids.map(r => r[0]).filter(Boolean));
}

// ─── UTILITY FUNCTIONS ────────────────────────────────────────────────────────

function getDriveFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  // Auto-create if missing
  Logger.log("Creating Drive folder: " + name);
  return DriveApp.createFolder(name);
}

/**
 * Parse a date string in any of the formats Monzo or Amex might export:
 *   DD/MM/YYYY  |  YYYY-MM-DD  |  DD MMM YYYY  |  MM/DD/YYYY (US Amex)
 */
function parseDate_(str) {
  if (!str) return null;
  str = str.trim();

  // YYYY-MM-DD (ISO)
  let m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

  // DD/MM/YYYY
  m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);

  // DD MMM YYYY  (e.g. "14 Apr 2026")
  m = str.match(/^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})/);
  if (m) {
    const months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
    const mo = months[m[2].toLowerCase()];
    if (mo !== undefined) return new Date(+m[3], mo, +m[1]);
  }

  return null;
}

function findCol_(header, candidates) {
  for (const cand of candidates) {
    const idx = header.indexOf(cand);
    if (idx >= 0) return idx;
  }
  return -1;
}

function fmt_(n) {
  return Number(n).toLocaleString("en-GB", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function formatMonthYear_(month, year) {
  const d = new Date(year, month - 1, 1);
  return Utilities.formatDate(d, "GMT", "MMMM yyyy");
}

/**
 * Automatically syncs Monzo transactions from the Google Sheet populated by IFTTT
 */
function syncMonzoFromIFTTT_() {
  const existing = getExistingTxIds_();
  let imported = 0;

  // Search for the IFTTT spreadsheet in Google Drive by name
  const files = DriveApp.getFilesByName(CONFIG.IFTTT_SHEET_NAME || "Monzo Transactions");
  if (!files.hasNext()) {
    Logger.log("IFTTT spreadsheet '" + (CONFIG.IFTTT_SHEET_NAME || "Monzo Transactions") + "' not found in Google Drive. Skipping IFTTT sync.");
    return 0;
  }

  const iftttFile = files.next();
  let iftttSs;
  try {
    iftttSs = SpreadsheetApp.open(iftttFile);
  } catch (e) {
    Logger.log("Failed to open IFTTT spreadsheet: " + e.message);
    return 0;
  }

  // IFTTT appends rows to the first sheet
  const iftttSheet = iftttSs.getSheets()[0];
  const lastRow = iftttSheet.getLastRow();
  if (lastRow < 1) {
    Logger.log("IFTTT sheet is empty.");
    return 0;
  }

  // Read all rows
  // Columns:
  // A: CreatedAt (e.g. "June 16, 2026 at 04:00PM")
  // B: Amount (e.g. "-15.50" or "£-15.50")
  // C: Currency (e.g. "GBP")
  // D: MerchantName (e.g. "Tesco")
  // E: MerchantCategory (e.g. "groceries")
  // F: Description (e.g. "TESCO STORES 1234")
  // G: Category (Monzo native category, e.g. "Groceries")
  const data = iftttSheet.getRange(1, 1, lastRow, 7).getValues();
  const results = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0] || !row[1]) continue;

    const rawDate = String(row[0]).trim();
    const date = parseIFTTTDate_(rawDate);
    if (!date) continue;

    // Parse amount: negative = spend, positive = refund/income (skip)
    let rawAmt = String(row[1]).replace(/[£,\s]/g, "");
    let amount = parseFloat(rawAmt) || 0;
    
    // We only track spends (negative in Monzo/IFTTT)
    if (amount >= 0) continue; 
    amount = Math.abs(amount);

    if (amount === 0) continue;

    const merchant = String(row[3]).trim() || String(row[5]).trim() || "Unknown";
    const monzoCategory = String(row[6]).trim() || String(row[4]).trim() || "";
    
    const category = mapCategory_(merchant, monzoCategory);
    if (category === null) continue; // skip categories flagged as None
    
    // Create unique ID from date, merchant, and amount to prevent double imports
    const dateSlug = Utilities.formatDate(date, "GMT", "yyyyMMdd");
    const merchantSlug = merchant.replace(/\s+/g, "").substring(0, 10).toLowerCase();
    const txId = "ifttt_" + dateSlug + "_" + merchantSlug + "_" + amount;

    if (!existing.has(txId)) {
      results.push(buildTxRow_(date, merchant, amount, category, "Monzo Current", txId));
      existing.add(txId);
    }
  }

  if (results.length > 0) {
    appendTransactions_(results);
    imported = results.length;
    Logger.log("Imported " + imported + " new transactions from IFTTT.");
  }
  
  return imported;
}

/**
 * Parses IFTTT date format like "June 16, 2026 at 04:00PM" or standard JS date strings
 */
function parseIFTTTDate_(str) {
  if (!str) return null;
  
  // Try standard JS parsing first
  let d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  
  // Custom parse for "Month DD, YYYY at HH:MMAM/PM" (IFTTT format)
  // e.g. "June 16, 2026 at 04:00PM"
  const match = str.match(/^([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})\s+at\s+(\d{1,2}):(\d{2})(AM|PM|am|pm)?/);
  if (match) {
    const monthNames = {
      january:0, february:1, march:2, april:3, may:4, june:5,
      july:6, august:7, september:8, october:9, november:10, december:11,
      jan:0, feb:1, mar:2, apr:3, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
    };
    const month = monthNames[match[1].toLowerCase()];
    const day = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);
    let hour = parseInt(match[4], 10);
    const min = parseInt(match[5], 10);
    const ampm = match[6];
    
    if (ampm) {
      if (ampm.toLowerCase() === "pm" && hour < 12) hour += 12;
      if (ampm.toLowerCase() === "am" && hour === 12) hour = 0;
    }
    
    return new Date(year, month, day, hour, min);
  }
  
  // Fall back to parseDate_
  return parseDate_(str);
}

// ─── TRIGGER SETUP ───────────────────────────────────────────────────────────

/**
 * Run this ONCE from the Apps Script editor to register all triggers.
 * Extensions → Apps Script → Run → setupTriggers
 */
function setupTriggers() {
  // Delete existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Daily CSV sync at 07:00 (silent)
  ScriptApp.newTrigger("syncCSVs")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  // Hourly Gmail alerts sync (silent)
  ScriptApp.newTrigger("syncGmailAlerts")
    .timeBased()
    .everyHours(1)
    .create();

  // Weekly Brief email on Mondays at 08:00
  ScriptApp.newTrigger("sendWeeklyBrief")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  // Weekly Budget alerts check on Mondays at 08:15
  ScriptApp.newTrigger("runAlertCheck")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .nearMinute(15)
    .create();

  Logger.log("✅ Triggers registered: syncCSVs (daily 07:00), syncGmailAlerts (hourly), sendWeeklyBrief (weekly Monday 08:00), and runAlertCheck (weekly Monday 08:15)");
}

/**
 * Scans emails labeled "Ledger-Alerts", parses transactions for BofA and Monzo,
 * and appends them to the Transactions sheet.
 */
function syncGmailAlerts() {
  const labelName = "Ledger-Alerts";
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    Logger.log("⚠️ Gmail label '" + labelName + "' does not exist. Please create it.");
    return;
  }

  const threads = label.getThreads();
  if (threads.length === 0) {
    Logger.log("No new emails under label '" + labelName + "'.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);
  const existingIds = getExistingTxIds_(txSheet);

  let newRows = [];

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const subject = msg.getSubject();
      const body = msg.getPlainBody();
      const date = msg.getDate();

      let parsedTx = null;

      // 1. Monzo Alert (IFTTT Email Alert)
      // Subject format: Monzo Alert: MerchantName | Amount | Currency
      if (subject.indexOf("Monzo Alert:") !== -1) {
        try {
          const parts = subject.replace("Monzo Alert:", "").split("|");
          if (parts.length >= 3) {
            const merchant = parts[0].trim();
            const amountStr = parts[1].replace(/[^\d.-]/g, "").trim();
            const currency = parts[2].trim();
            const rawAmount = parseFloat(amountStr);

            if (!isNaN(rawAmount)) {
              // Standardize spend as positive, refund as negative
              let amount = Math.abs(rawAmount);
              let isRefund = rawAmount > 0; // standard IFTTT spend is negative, refund is positive
              if (isRefund) {
                amount = -Math.abs(rawAmount);
              }

              const cleanName = cleanMerchantName_(merchant);
              const category = mapCategory_(cleanName);
              const txId = "monzo_mail_" + Utilities.formatDate(date, "GMT", "yyyyMMdd_HHmmss") + "_" + Math.floor(Math.abs(amount) * 100);

              parsedTx = {
                date: date,
                merchant: cleanName,
                amount: amount,
                category: category,
                account: "Monzo Current",
                txId: txId
              };
            }
          }
        } catch (err) {
          Logger.log("Error parsing Monzo alert: " + err.message);
        }
      }
      
      // 2. Bank of America Transaction Alert
      else if (msg.getFrom().indexOf("bankofamerica.com") !== -1 || subject.indexOf("Bank of America Alert") !== -1 || subject.indexOf("Bank of America Transaction") !== -1) {
        try {
          // BofA alert body has: "A transaction of $XX.XX occurred" or "withdrawal of $XX.XX"
          const amtMatch = body.match(/\$([\d,]+\.\d{2})/);
          if (amtMatch) {
            const rawAmount = parseFloat(amtMatch[1].replace(/,/g, ""));
            
            // Extract description: usually "at [Merchant]" or "to [Merchant]"
            let merchant = "Bank of America Withdrawal";
            const descMatch = body.match(/(?:at|to|with)\s+([A-Za-z0-9\s#&*'-]+?)(?:\.\s|on\s|\nat\s|\n)/);
            if (descMatch) {
              merchant = descMatch[1].trim();
            }

            const cleanName = cleanMerchantName_(merchant);
            let category = mapCategory_(cleanName);
            
            const txId = "bofa_mail_" + Utilities.formatDate(date, "GMT", "yyyyMMdd_HHmmss") + "_" + Math.floor(rawAmount * 100);

            parsedTx = {
              date: date,
              merchant: cleanName,
              amount: rawAmount,
              category: category,
              account: "Bank of America Checking",
              txId: txId
            };
          }
        } catch (err) {
          Logger.log("Error parsing BofA alert: " + err.message);
        }
      }

      if (parsedTx && !existingIds.has(parsedTx.txId)) {
        const row = buildTxRow_(parsedTx.date, parsedTx.merchant, parsedTx.amount, parsedTx.category, parsedTx.account, parsedTx.txId);
        newRows.push(row);
        existingIds.add(parsedTx.txId);
      }
    }

    // Remove the label so it is not scanned again
    thread.removeLabel(label);
  }

  if (newRows.length > 0) {
    appendTransactions_(newRows);
    Logger.log("✅ syncGmailAlerts finished. Added " + newRows.length + " new transactions from Gmail alerts.");
  } else {
    Logger.log("No new transaction alerts found.");
  }
}

/**
 * Helper: Clean merchant name text formatting.
 */
function cleanMerchantName_(desc) {
  let cleaned = desc.replace(/\s{2,}.*/, '').trim();
  return cleaned.split(' ').map(w => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase()).join(' ');
}

/**
 * Serves the transaction database as JSON for PWA integration.
 * Deploy as a Web App: Extensions -> Apps Script -> Deploy -> New Deployment -> Web App (Anyone has access).
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    
    // Read config
    const cfg = ss.getSheetByName(CONFIG.TABS.CONFIG);
    const startDateVal = cfg.getRange("C35").getValue();
    const endDateVal = cfg.getRange("C36").getValue();
    const startingCapital = parseFloat(cfg.getRange("C37").getValue());
    
    // Handle date formatting
    let startDate = "2026-07-01";
    let endDate = "2026-08-31";
    if (startDateVal instanceof Date) {
      startDate = Utilities.formatDate(startDateVal, "GMT", "yyyy-MM-dd");
    } else if (startDateVal) {
      startDate = startDateVal.toString();
    }
    if (endDateVal instanceof Date) {
      endDate = Utilities.formatDate(endDateVal, "GMT", "yyyy-MM-dd");
    } else if (endDateVal) {
      endDate = endDateVal.toString();
    }
    
    // Read transactions
    const sheet = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);
    const lastRow = sheet.getLastRow();
    let txs = [];
    
    if (lastRow >= 3) {
      const data = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue;
        
        let dateStr = "";
        if (row[0] instanceof Date) {
          dateStr = Utilities.formatDate(row[0], "GMT", "yyyy-MM-dd");
        } else {
          dateStr = row[0].toString();
        }
        
        const merchant = row[1];
        const amount = parseFloat(row[2]);
        const category = row[3];
        const account = row[4];
        const txId = row[7];
        
        // Deduce currency based on account name
        const currency = account.indexOf("America") !== -1 ? "USD" : "GBP";
        
        txs.push({
          date: dateStr,
          merchant: merchant,
          amount: amount,
          category: category,
          account: account,
          currency: currency,
          tx_id: txId
        });
      }
    }
    
    const payload = {
      start_date: startDate,
      end_date: endDate,
      starting_capital: startingCapital,
      transactions: txs
    };
    
    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Run weekly on Monday at 08:00 — sends the Weekly CFO Brief
 */
function sendWeeklyBrief() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const txSheet = ss.getSheetByName(CONFIG.TABS.TRANSACTIONS);
  const now = new Date();
  
  // Past 7 days
  const msPerDay = 1000 * 60 * 60 * 24;
  const oneWeekAgo = new Date(now.getTime() - (7 * msPerDay));
  
  // Read all transactions
  const lastRow = txSheet.getLastRow();
  let weekTxs = [];
  let totalSpent = 0;
  let categorySpends = {};
  
  if (lastRow >= 3) {
    const data = txSheet.getRange(3, 1, lastRow - 2, 8).getValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const txDate = new Date(row[0]);
      if (txDate >= oneWeekAgo && txDate <= now) {
        const amount = parseFloat(row[2]);
        const merchant = row[1];
        const category = row[3];
        const account = row[4];
        
        weekTxs.push({ date: txDate, merchant: merchant, amount: amount, category: category, account: account });
        
        // Exclude rent from weekly averages
        if (category !== 'Rent') {
          totalSpent += amount;
          categorySpends[category] = (categorySpends[category] || 0) + amount;
        }
      }
    }
  }
  
  // Build HTML email brief
  let catHtml = "";
  Object.keys(categorySpends).sort((a,b) => categorySpends[b] - categorySpends[a]).forEach(cat => {
    catHtml += `<tr>
      <td style="padding: 8px 0; border-bottom: 1px solid #EDE8E0; font-weight: 600;">${cat}</td>
      <td style="padding: 8px 0; border-bottom: 1px solid #EDE8E0; text-align: right; font-family: monospace;">£${categorySpends[cat].toFixed(2)}</td>
    </tr>`;
  });
  
  const formattedDate = Utilities.formatDate(now, "GMT", "dd MMM yyyy");
  
  const html = `
    <div style="font-family: sans-serif; background-color: #FAF7F2; padding: 30px; color: #2A2421; max-width: 600px; margin: 0 auto; border: 1px solid #EDE8E0; border-radius: 16px;">
      <div style="text-align: center; border-bottom: 2px solid #C9A84C; padding-bottom: 15px; margin-bottom: 20px;">
        <h1 style="font-family: Georgia, serif; font-size: 24px; font-weight: 700; color: #2A2421; margin: 0;">DCR LEDGER</h1>
        <p style="font-size: 11px; text-transform: uppercase; letter-spacing: 0.1em; color: #6B6259; margin: 4px 0 0 0;">Weekly CFO Brief — ${formattedDate}</p>
      </div>
      
      <p style="font-size: 14px; line-height: 1.5; color: #6B6259;">Here is your weekly financial summary for the last 7 days of observation:</p>
      
      <div style="background-color: #FFFFFF; border: 1px solid #EDE8E0; border-radius: 12px; padding: 20px; text-align: center; margin-bottom: 20px;">
        <span style="font-size: 10px; text-transform: uppercase; color: #6B6259; letter-spacing: 0.05em; display: block; margin-bottom: 4px;">Total Variable Outflow</span>
        <span style="font-family: Georgia, serif; font-size: 36px; font-weight: 500; color: #2A2421;">£${totalSpent.toFixed(2)}</span>
      </div>
      
      <h3 style="font-family: Georgia, serif; font-size: 16px; color: #2A2421; margin-bottom: 10px; border-bottom: 1px solid #C9A84C; padding-bottom: 4px;">Outflows by Category</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <tbody>
          ${catHtml || '<tr><td colspan="2" style="text-align: center; color: #6B6259; padding: 20px;">No variable transactions logged this week.</td></tr>'}
        </tbody>
      </table>
      
      <div style="margin-top: 30px; text-align: center; font-size: 11px; color: #6B6259; border-top: 1px solid #EDE8E0; padding-top: 15px;">
        <p>This email is sent automatically by your DCR Ledger App.</p>
        <p><a href="https://DCRcode-dev.github.io/dcr-budget-tracker" style="color: #C9A84C; text-decoration: none; font-weight: 600;">Open Dashboard →</a></p>
      </div>
    </div>
  `;
  
  GmailApp.sendEmail(
    CONFIG.ALERT_EMAIL,
    "📊 Weekly Budget Brief — DCR Ledger",
    "Please view this email in HTML format.",
    {
      htmlBody: html,
      name: "DCR Budget System"
    }
  );
  
  Logger.log("Weekly brief email sent successfully.");
}
