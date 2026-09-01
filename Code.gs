/**
 * Victoria Court Budget Tracker — Code.gs
 * ============================================================================
 * Owner      : Daniel Cruz Rosso (DCR), San Juan, Puerto Rico
 * System     : Google Sheets + Apps Script (V8 runtime)
 * Currency   : USD ($#,##0.00)
 * Timezone   : America/Puerto_Rico (MM/DD/YYYY)
 * Purpose    : 1. Calculate Renovation Funding Gap ("The Ask")
 *              2. Track Living budget against a single $2,000 ceiling
 *              3. Monitor November mortgage ($3,305.55/mo) and cash runway
 * ============================================================================
 */

// ─── CONSTANTS & CONFIGURATION ──────────────────────────────────────────────

const SCRIPT_CONFIG = {
  SPREADSHEET_NAME: "Victoria Court Budget Tracker",
  TIMEZONE: "America/Puerto_Rico",
  LOCALE: "en_US",
  BASE_CURRENCY: "USD",
  
  TABS: {
    DASHBOARD: "Dashboard",
    LIVING: "Living",
    HOME: "Home",
    CONFIG: "Config",
    ACCOUNTS: "Accounts",
    TRANSACTIONS: "Transactions",
    RENOVATION: "Renovation",
    RULES: "Rules"
  },
  
  COLORS: {
    PRIMARY_NAVY: "#1C2833",
    BLUE_INPUT: "#1155CC",
    GRAY_TEXT: "#666666",
    BG_LIGHT_YELLOW: "#FFF8E1",
    BORDER_ORANGE: "#F57C00",
    BG_LIGHT_BLUE: "#E8F0FE",
    BORDER_BLUE: "#1155CC",
    BG_LIGHT_GRAY: "#F5F5F5",
    BORDER_GRAY: "#EDE8E0",
    STATUS_OK_BG: "#D9EAD3",
    STATUS_OK_TXT: "#274E13",
    STATUS_WATCH_BG: "#FFF2CC",
    STATUS_WATCH_TXT: "#7F6000",
    STATUS_OVER_BG: "#F4CCCC",
    STATUS_OVER_TXT: "#990000"
  },
  
  CATEGORIES: {
    FIXED: ["Mortgage", "Utilities", "HOA"],
    LIVING: ["Groceries", "Dining", "Transport", "Travel", "Health", "Shopping", "Other"],
    HOME: ["Renovation"],
    EXCLUDED: ["Transfer", "Income", "Credit Card Payment"]
  }
};

// ─── 1. SETUP & SCAFFOLDING ──────────────────────────────────────────────────

/**
 * Initializes the entire Victoria Court Budget Tracker workbook.
 * Creates tabs, formulas, Named Ranges, formatting, and seeds default data.
 */
function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(SCRIPT_CONFIG.TIMEZONE);
  ss.setSpreadsheetLocale(SCRIPT_CONFIG.LOCALE);

  // 1. Create or get tabs in exact left-to-right order
  const tabNames = [
    SCRIPT_CONFIG.TABS.DASHBOARD,
    SCRIPT_CONFIG.TABS.LIVING,
    SCRIPT_CONFIG.TABS.HOME,
    SCRIPT_CONFIG.TABS.CONFIG,
    SCRIPT_CONFIG.TABS.ACCOUNTS,
    SCRIPT_CONFIG.TABS.TRANSACTIONS,
    SCRIPT_CONFIG.TABS.RENOVATION,
    SCRIPT_CONFIG.TABS.RULES
  ];

  const sheets = {};
  tabNames.forEach((name, index) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name, index);
    } else {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
    sheets[name] = sheet;
    sheet.clear();
    sheet.clearConditionalFormatRules();
  });

  // Clean up any extraneous default sheets
  const allSheets = ss.getSheets();
  allSheets.forEach(s => {
    if (!tabNames.includes(s.getName()) && allSheets.length > tabNames.length) {
      try { ss.deleteSheet(s); } catch (e) {}
    }
  });

  // 2. Build Config Tab and Named Ranges
  buildConfigTab_(sheets[SCRIPT_CONFIG.TABS.CONFIG], ss);

  // 3. Build Accounts Tab
  buildAccountsTab_(sheets[SCRIPT_CONFIG.TABS.ACCOUNTS]);

  // 4. Build Transactions Tab
  buildTransactionsTab_(sheets[SCRIPT_CONFIG.TABS.TRANSACTIONS]);

  // 5. Build Renovation Tab
  buildRenovationTab_(sheets[SCRIPT_CONFIG.TABS.RENOVATION]);

  // 6. Build Rules Tab
  buildRulesTab_(sheets[SCRIPT_CONFIG.TABS.RULES]);

  // 7. Build Presentation Tabs (Dashboard, Living, Home)
  buildDashboardTab_(sheets[SCRIPT_CONFIG.TABS.DASHBOARD]);
  buildLivingTab_(sheets[SCRIPT_CONFIG.TABS.LIVING]);
  buildHomeTab_(sheets[SCRIPT_CONFIG.TABS.HOME]);

  // 8. Hide required tabs
  sheets[SCRIPT_CONFIG.TABS.ACCOUNTS].hideSheet();
  sheets[SCRIPT_CONFIG.TABS.TRANSACTIONS].hideSheet();
  sheets[SCRIPT_CONFIG.TABS.RENOVATION].hideSheet();
  sheets[SCRIPT_CONFIG.TABS.RULES].hideSheet();

  // Set default active tab
  ss.setActiveSheet(sheets[SCRIPT_CONFIG.TABS.DASHBOARD]);

  SpreadsheetApp.flush();
  Logger.log("✅ Victoria Court Budget Tracker setup completed successfully.");
}

/**
 * Builds the Config tab and registers all Named Ranges.
 */
function buildConfigTab_(sheet, ss) {
  sheet.setColumnWidth(1, 24);   // Gutter
  sheet.setColumnWidth(2, 280);  // Label
  sheet.setColumnWidth(3, 200);  // Value
  sheet.setColumnWidth(4, 320);  // Notes

  sheet.getRange("B2:D2").setValues([["Setting", "Value", "Notes"]]).setFontWeight("bold");

  const configRows = [
    ["LIVING_CEILING", "Living monthly ceiling (USD)", 2000, "$#,##0.00", true, "Hard cap. Owner-stated."],
    ["MORTGAGE_PAYMENT", "Mortgage monthly payment (USD)", 3305.55, "$#,##0.00", true, "Years 1–11, includes FHA MIP."],
    ["MORTGAGE_PAYMENT_POST_MIP", "Mortgage payment after MIP drops (USD)", 3053.47, "$#,##0.00", true, "Years 12–30."],
    ["MORTGAGE_MIP_PAYMENT_COUNT", "Payments until MIP cancels", 132, "#,##0", true, "11 years."],
    ["MORTGAGE_FIRST_PAYMENT_DATE", "Mortgage first payment date", "11/01/2026", "MM/DD/YYYY", true, "PROVISIONAL — exact November date unconfirmed."],
    ["MORTGAGE_LOAN_AMOUNT", "Mortgage loan amount (USD)", 550000, "$#,##0.00", true, "50% LTV."],
    ["MORTGAGE_RATE", "Mortgage interest rate", 0.0475, "0.000%", true, "FHA 30-yr fixed."],
    ["MORTGAGE_APR", "Mortgage APR", 0.05137, "0.000%", true, "Annual Percentage Rate."],
    ["UTILITIES_EST_CONSTRUCTION", "Utilities monthly estimate — Construction (Sep–Oct)", 130, "$#,##0.00", true, "Electricity ($100) + Water ($30) construction usage."],
    ["UTILITIES_EST", "Utilities monthly estimate — Living (Nov onward)", 500, "$#,##0.00", true, "Electricity ($400) + Water ($100) living usage."],
    ["HOA_AMOUNT", "Victoria Court HOA monthly (USD)", 375, "$#,##0.00", true, "Recurring $375/mo starting September. Home insurance inside."],
    ["RUNWAY_MONTHS", "Months of runway to hold back", 3, "#,##0", true, "Owner's choice over recommended 6."],
    ["RENO_TOTAL_BUDGET", "Renovation total project budget (USD)", 0, "$#,##0.00", true, "DORMANT. When > 0, burn-down activates."],
    ["MONTHLY_SALARY", "Monthly salary (USD)", 6700, "$#,##0.00", true, "Expected in Banco Popular."],
    ["SALARY_START_DATE", "Salary start date", "09/08/2026", "MM/DD/YYYY", true, "Started Sep 8, 2026."],
    ["FIRST_PAYCHECK_EST", "First paycheck est. (Prorated Sep 8–30)", 5136.67, "$#,##0.00", true, "23/30 days proration paid early Oct in Banco Popular."],
    ["GBP_USD_RATE", "GBP→USD rate for Monzo wind-down", 1.27, "0.00", true, "Fixed rate, manually updated."],
    ["CURRENT_MONTH_KEY", "Current month key", '=TEXT(TODAY(),"YYYY-MM")', "@", false, "Formula, black text."],
    ["LAST_BANK_SYNC", "Last bank sync", "", "MM/DD/YYYY HH:mm", false, "Script-written timestamp."],
    ["LAST_RENO_SYNC", "Last renovation sync", new Date(), "MM/DD/YYYY HH:mm", false, "Script-written timestamp."],
    ["MOVE_IN_DATE", "Move-in target date", "10/28/2026", "MM/DD/YYYY", true, "From Victoria Court Tracker."]
  ];

  const existingNamedRanges = ss.getNamedRanges();
  const existingMap = {};
  existingNamedRanges.forEach(nr => existingMap[nr.getName()] = nr);

  configRows.forEach((r, i) => {
    const rowIdx = 3 + i;
    const name = r[0];
    const label = r[1];
    const val = r[2];
    const format = r[3];
    const isUserEditable = r[4];
    const notes = r[5];

    sheet.getRange(rowIdx, 2).setValue(label).setFontWeight("bold");
    const valCell = sheet.getRange(rowIdx, 3);
    
    if (typeof val === "string" && val.startsWith("=")) {
      valCell.setFormula(val);
    } else {
      valCell.setValue(val);
    }
    
    valCell.setNumberFormat(format);
    valCell.setFontColor(isUserEditable ? SCRIPT_CONFIG.COLORS.BLUE_INPUT : "#000000");

    sheet.getRange(rowIdx, 4).setValue(notes).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

    // Register or update Named Range
    if (existingMap[name]) {
      existingMap[name].setRange(valCell);
    } else {
      ss.setNamedRange(name, valCell);
    }
  });

  // Add Color Key Block
  const colorKeyRow = 3 + configRows.length + 2;
  sheet.getRange(colorKeyRow, 2).setValue("COLOR KEY").setFontWeight("bold");
  sheet.getRange(colorKeyRow + 1, 2).setValue("Blue text").setFontColor(SCRIPT_CONFIG.COLORS.BLUE_INPUT).setFontWeight("bold");
  sheet.getRange(colorKeyRow + 1, 3).setValue("Editable user input");
  sheet.getRange(colorKeyRow + 2, 2).setValue("Black text").setFontColor("#000000").setFontWeight("bold");
  sheet.getRange(colorKeyRow + 2, 3).setValue("Formula or system-managed, do not edit");
}

/**
 * Builds the Accounts tab.
 */
function buildAccountsTab_(sheet) {
  sheet.setFrozenRows(1);
  const headers = [
    "Account", "Institution", "Account_Type", "Currency", "Balance_Native",
    "Balance_USD", "Balance_As_Of", "Status", "Include_In_Available", "Feed", "SimpleFIN_Account_ID"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

  const seedRows = [
    ["Banco Popular", "Banco Popular - Puerto Rico & Virgin Islands", "Checking", "USD", "", '=IF(D2="GBP", E2*GBP_USD_RATE, E2)', "", "Active", true, "SimpleFIN", ""],
    ["Capital One Venture X", "Capital One", "Credit Card", "USD", "", '=IF(D3="GBP", E3*GBP_USD_RATE, E3)', "", "Active", false, "SimpleFIN", ""],
    ["Bank of America", "Bank of America", "Checking", "USD", "", '=IF(D4="GBP", E4*GBP_USD_RATE, E4)', "", "Winding Down", true, "None", ""],
    ["Monzo Current", "Monzo", "Checking", "GBP", 15000, '=IF(D5="GBP", E5*GBP_USD_RATE, E5)', new Date(), "Winding Down", true, "Manual", ""],
    ["Monzo Flex", "Monzo", "Credit Card", "GBP", "", '=IF(D6="GBP", E6*GBP_USD_RATE, E6)', "", "Closed", false, "CSV", ""]
  ];

  sheet.getRange(2, 1, seedRows.length, headers.length).setValues(seedRows);
  sheet.getRange("I2:I6").insertCheckboxes();
  sheet.getRange("E2:F6").setNumberFormat("$#,##0.00");
  sheet.getRange("G2:G6").setNumberFormat("MM/DD/YYYY HH:mm");
}

/**
 * Builds the Transactions tab.
 */
function buildTransactionsTab_(sheet) {
  sheet.setFrozenRows(1);
  const headers = [
    "Tx_ID", "Date", "Merchant", "Amount_USD", "Amount_Native", "Currency_Native",
    "Account", "Tier", "Category", "Month_Key", "Year", "Month_Num", "Source", "Imported_At", "Manual_Override"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.getRange("B:B").setNumberFormat("MM/DD/YYYY");
  sheet.getRange("D:E").setNumberFormat("$#,##0.00");
  sheet.getRange("N:N").setNumberFormat("MM/DD/YYYY HH:mm");
}

/**
 * Builds the Renovation tab with exact seed values from Victoria Court Tracker.xlsx.
 */
function buildRenovationTab_(sheet) {
  sheet.setFrozenRows(1);
  const headers = [
    "Line_Key", "Section", "Room_Or_Scope", "Item", "Vendor", "Status", "Amount_USD", "Is_Priced", "Is_Outstanding", "Option_Group"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

  const rows = [];

  // 1. Building Works Items (Davis Facias Contract: $21,400)
  const bwItems = [
    ["BW_01", "Building Works", "Whole House", "Remoción de molduras en paredes y resanar", "Davis Facias", "In Progress", 3200],
    ["BW_02", "Building Works", "Kitchen", "Facias en el área de la cocina tipo cielo raso", "Davis Facias", "In Progress", 2000],
    ["BW_03", "Building Works", "Dining / Stairs", "Remoción de pared en espejo y resanar", "Davis Facias", "Done", 600],
    ["BW_04", "Building Works", "Whole House", "Facias tipo cielo raso en techos", "Davis Facias", "In Progress", 7600],
    ["BW_05", "Building Works", "Whole House", "Pintura en interior", "Davis Facias", "Contracted", 4800],
    ["BW_06", "Building Works", "Bathroom", "Puerta pocket en área de baño", "Davis Facias", "Contracted", 2000],
    ["BW_07", "Building Works", "Family Room", "Reubicación de aire acondicionado", "Davis Facias", "Contracted", 400],
    ["BW_08", "Building Works", "Living Room", "Eliminación de columnas (demolición)", "Davis Facias", "Done", 400],
    ["BW_09", "Building Works", "Whole House", "Disposición de los escombros", "Davis Facias", "In Progress", 400]
  ];
  bwItems.forEach(item => {
    // Individual sub-items are reference lines; disbursement schedule below drives outstanding cash
    rows.push([item[0], item[1], item[2], item[3], item[4], item[5], item[6], true, false, ""]);
  });

  // 2. Building Works Disbursements Schedule (Drives the $10,700 outstanding)
  rows.push(["BW_DISB_P1", "Building Works", "Whole House", "Phase 1 Deposit 50% (12 Jul 2026)", "Davis Facias", "Paid", 10700, true, false, ""]);
  rows.push(["BW_DISB_P2", "Building Works", "Whole House", "Phase 2 Progress 25% (Sep 2026 at 75% completion)", "Davis Facias", "Pending", 5350, true, true, ""]);
  rows.push(["BW_DISB_P3", "Building Works", "Whole House", "Phase 3 Final 25% (Oct 2026 on final signoff)", "Davis Facias", "Pending", 5350, true, true, ""]);

  // 3. Furniture Ordered (5 items, $11,089 total, Is_Outstanding = FALSE)
  const orderedFurniture = [
    ["FN_ORD_01", "Furniture", "Family Room", "Reclining Mechanical Sofa", "Ambiente", "Ordered", 3549],
    ["FN_ORD_02", "Furniture", "Living Room", "Sofa", "Koper", "Ordered", 1182],
    ["FN_ORD_03", "Furniture", "Dining Room", "Table (106.3\"L × 70.9\"W Walnut)", "Viva", "Ordered", 2982],
    ["FN_ORD_04", "Furniture", "Dining Room", "Chairs (8x)", "Viva", "Ordered", 2378]
  ];
  orderedFurniture.forEach(item => {
    rows.push([item[0], item[1], item[2], item[3], item[4], item[5], item[6], true, false, ""]);
  });

  // 4. Furniture Option Groups (Selected Mid-case = $3,401, Is_Outstanding = TRUE for selected option)
  const optionGroups = [
    // LR_ARMCHAIR ($450 mid)
    ["FN_OPT_AC_MID", "Furniture", "Living Room", "Living Room Armchair (Selected Mid Option)", "Ambiente", "Option", 450, true, true, "LR_ARMCHAIR"],
    ["FN_OPT_AC_ALT1", "Furniture", "Living Room", "Living Room Armchair (Alternative)", "Koper", "Option", 297, true, false, "LR_ARMCHAIR"],
    ["FN_OPT_AC_ALT2", "Furniture", "Living Room", "Living Room Armchair (Alternative)", "Arkitektura", "Option", 595, true, false, "LR_ARMCHAIR"],
    // LR_BENCH ($595 mid)
    ["FN_OPT_BN_MID", "Furniture", "Living Room", "Living Room Bench (Selected Mid Option)", "Ambiente", "Option", 595, true, true, "LR_BENCH"],
    ["FN_OPT_BN_ALT1", "Furniture", "Living Room", "Living Room Bench (Alternative)", "Koper", "Option", 345, true, false, "LR_BENCH"],
    ["FN_OPT_BN_ALT2", "Furniture", "Living Room", "Living Room Bench (Alternative)", "Arkitektura", "Option", 1395, true, false, "LR_BENCH"],
    // PR_FAUCET ($139 mid)
    ["FN_OPT_FC_MID", "Furniture", "Powder Room", "Powder Room Faucet (Selected Mid Option)", "Ferguson", "Option", 139, true, true, "PR_FAUCET"],
    ["FN_OPT_FC_ALT1", "Furniture", "Powder Room", "Powder Room Faucet (Alternative)", "Amazon", "Option", 109, true, false, "PR_FAUCET"],
    ["FN_OPT_FC_ALT2", "Furniture", "Powder Room", "Powder Room Faucet (Alternative)", "Kohler", "Option", 169, true, false, "PR_FAUCET"],
    // KT_HOOD ($349 mid)
    ["FN_OPT_HD_MID", "Furniture", "Kitchen", "Kitchen Range Hood (Selected Mid Option)", "Best Buy", "Option", 349, true, true, "KT_HOOD"],
    ["FN_OPT_HD_ALT1", "Furniture", "Kitchen", "Kitchen Range Hood (Alternative)", "Amazon", "Option", 295, true, false, "KT_HOOD"],
    ["FN_OPT_HD_ALT2", "Furniture", "Kitchen", "Kitchen Range Hood (Alternative)", "Home Depot", "Option", 399, true, false, "KT_HOOD"],
    // DR_BUFFET ($695 mid)
    ["FN_OPT_BF_MID", "Furniture", "Dining Room", "Dining Room Buffet (Selected Option)", "Ambiente", "Option", 695, true, true, "DR_BUFFET"],
    ["FN_OPT_BF_ALT1", "Furniture", "Dining Room", "Dining Room Buffet (Alternative)", "Viva", "Option", 697, true, false, "DR_BUFFET"],
    // KT_COFFEE ($695 mid)
    ["FN_OPT_CF_MID", "Furniture", "Kitchen", "Kitchen Coffee Table (Selected Option)", "Ambiente", "Option", 695, true, true, "KT_COFFEE"],
    // PR_SINK ($329)
    ["FN_OPT_SK_MID", "Furniture", "Powder Room", "Powder Room Sink", "Arkitektura", "Option", 329, true, true, "PR_SINK"]
  ];
  optionGroups.forEach(item => {
    rows.push([item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9]]);
  });

  // 5. Furniture Standalone Options ($3,339 total, Is_Outstanding = TRUE)
  const standaloneOptions = [
    ["FN_STA_01", "Furniture", "Family Room", "Sonos Arc Soundbar (Black)", "Sonos", "Option", 899, true, true, ""],
    ["FN_STA_02", "Furniture", "Family Room", "Sonos Era 100 Left Surround (Black)", "Sonos", "Option", 249, true, true, ""],
    ["FN_STA_03", "Furniture", "Family Room", "Sonos Era 100 Right Surround (Black)", "Sonos", "Option", 249, true, true, ""],
    ["FN_STA_04", "Furniture", "Guest Bedroom 1", "Bed: NATALIE Queen w/LED (Ambiente SKU 51315)", "Ambiente", "Option", 695, true, true, ""],
    ["FN_STA_05", "Furniture", "Guest Bedroom 2", "Bed: MADRID Queen (Ambiente SKU 50201)", "Ambiente", "Option", 697, true, true, ""],
    ["FN_STA_06", "Furniture", "Powder Room", "Accent Mirror (Arkitektura)", "Arkitektura", "Option", 550, true, true, ""]
  ];
  standaloneOptions.forEach(item => {
    rows.push([item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9]]);
  });

  // 6. Furniture Unpriced Pending Items (50 rows)
  const unpricedSpecs = [
    ["Living Room", "Area Carpet", 1], ["Living Room", "Curtains & Sheers", 2], ["Living Room", "Pendant Lighting", 3],
    ["Living Room", "Floor Lamp", 4], ["Living Room", "Artwork & Wall Decor", 5], ["Living Room", "Accent Pillows", 6],
    ["Living Room", "Side Tables", 7], ["Living Room", "Planters & Greenery", 8], ["Living Room", "Decorative Objects", 9],
    ["Living Room", "Console Table", 10], ["Living Room", "Mirror", 11], ["Living Room", "Throw Blankets", 12], ["Living Room", "Bookshelf Accents", 13],
    ["Family Room", "OLED Smart TV 75\"", 1], ["Family Room", "Floating TV Console", 2], ["Family Room", "Area Carpet", 3],
    ["Family Room", "Blackout Curtains", 4], ["Family Room", "Accent Lighting", 5], ["Family Room", "Coffee Table", 6],
    ["Family Room", "Bar Cabinet", 7], ["Family Room", "Wine & Beverage Refrigerator", 8], ["Family Room", "Acoustic Panels", 9], ["Family Room", "Side Table", 10],
    ["Kitchen", "Bar Stools (Counter Height)", 1], ["Kitchen", "Pendant Fixtures", 2], ["Kitchen", "Under-cabinet LED", 3],
    ["Kitchen", "Small Appliances Suite", 4], ["Kitchen", "Cookware & Utensils", 5], ["Kitchen", "Dinnerware Set", 6],
    ["Kitchen", "Glassware & Barware", 7], ["Kitchen", "Organization Systems", 8],
    ["Master Bedroom", "King Bed & Headboard", 1], ["Master Bedroom", "Mattress (Luxury Hybrid)", 2], ["Master Bedroom", "Nightstands (2x)", 3],
    ["Master Bedroom", "Dresser / Chest", 4], ["Master Bedroom", "Curtains & Drapes", 5], ["Master Bedroom", "Area Rug", 6], ["Master Bedroom", "Table Lamps", 7],
    ["Guest Bedroom 1", "Mattress Queen", 1], ["Guest Bedroom 1", "Nightstands (2x)", 2], ["Guest Bedroom 1", "Curtains", 3], ["Guest Bedroom 1", "Lighting", 4],
    ["Dining Room", "Chandelier / Pendant", 1], ["Dining Room", "Area Rug", 2], ["Dining Room", "Wall Mirror / Art", 3],
    ["Guest Bedroom 2", "Mattress Queen", 1], ["Guest Bedroom 2", "Nightstands", 2], ["Guest Bedroom 2", "Curtains", 3],
    ["Hallway", "Runner Carpet", 1], ["Hallway", "Flush Mount Lighting", 2]
  ];
  unpricedSpecs.forEach((spec, idx) => {
    rows.push([`FN_UNP_${idx + 1}`, "Furniture", spec[0], spec[1], "TBD", "Pending", "", false, false, ""]);
  });

  // 7. Phase 2 (2027) Architectural Enhancements (13 rows, all unpriced, Is_Outstanding = FALSE)
  const phase2Items = [
    "Stair glass railings", "Kitchen glass cabinet upgrade", "Terrace BBQ kitchenette",
    "Motorized louvered pergola", "Master double vanity", "Guest bath vanity",
    "Family room transom glass", "Solarium slider", "Terrace exterior doors",
    "Closet LED illumination", "Socket relocation & power upgrades", "Outdoor terrace tile", "Entry foyer accent"
  ];
  phase2Items.forEach((p2, idx) => {
    rows.push([`P2_2027_${idx + 1}`, "Building Works", "Phase 2 (2027)", p2, "TBD", "Not Started", "", false, false, ""]);
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange("G:G").setNumberFormat("$#,##0.00");
}

/**
 * Builds the Rules tab.
 */
function buildRulesTab_(sheet) {
  sheet.setFrozenRows(1);
  const headers = ["Match_Text", "Tier", "Category", "Priority"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

  const rules = [
    ["POPULAR MORTGAGE", "Fixed", "Mortgage", 1],
    ["HIPOTECA", "Fixed", "Mortgage", 1],
    ["LUMA", "Fixed", "Utilities", 2],
    ["AAA", "Fixed", "Utilities", 2],
    ["ACUEDUCTOS", "Fixed", "Utilities", 2],
    ["LIBERTY", "Fixed", "Utilities", 2],
    ["CLARO", "Fixed", "Utilities", 2],
    ["T-MOBILE", "Fixed", "Utilities", 2],
    ["HOA", "Fixed", "HOA", 2],
    ["VICTORIA COURT", "Fixed", "HOA", 2],
    ["DAVIS FACIAS", "Home", "Renovation", 3],
    ["KOPER", "Home", "Renovation", 3],
    ["AMBIENTE", "Home", "Renovation", 3],
    ["VIVA", "Home", "Renovation", 3],
    ["LAS MADERAS", "Home", "Renovation", 3],
    ["NARVAEZ", "Home", "Renovation", 3],
    ["ARKITEKTURA", "Home", "Renovation", 3],
    ["SONOS", "Home", "Renovation", 3],
    ["PAYMENT THANK YOU", "Excluded", "Credit Card Payment", 1],
    ["PAGO", "Excluded", "Credit Card Payment", 1],
    ["TRANSFER", "Excluded", "Transfer", 1],
    ["ATH MOVIL", "Excluded", "Transfer", 1],
    ["PAYROLL", "Excluded", "Income", 1],
    ["MEDHOLDINGS", "Excluded", "Income", 1],
    ["AIRBNB", "Living", "Travel", 5],
    ["JETBLUE", "Living", "Travel", 5],
    ["BRITISH AIRWAYS", "Living", "Travel", 5],
    ["AIRLINE", "Living", "Travel", 5],
    ["HOTEL", "Living", "Travel", 5],
    ["UBER EATS", "Living", "Dining", 10],
    ["DELIVEROO", "Living", "Dining", 10],
    ["RESTAURANT", "Living", "Dining", 10],
    ["ECONO", "Living", "Groceries", 10],
    ["PUEBLO", "Living", "Groceries", 10],
    ["SUPERMAX", "Living", "Groceries", 10],
    ["COSTCO", "Living", "Groceries", 10],
    ["SAMS CLUB", "Living", "Groceries", 10],
    ["WALMART", "Living", "Groceries", 10],
    ["UBER", "Living", "Transport", 20],
    ["LYFT", "Living", "Transport", 10],
    ["SHELL", "Living", "Transport", 10],
    ["TOTAL", "Living", "Transport", 10],
    ["EQUINOX", "Living", "Health", 10],
    ["GYM", "Living", "Health", 10],
    ["WALGREENS", "Living", "Health", 10],
    ["CVS", "Living", "Health", 10],
    ["FARMACIA", "Living", "Health", 10],
    ["AMAZON", "Living", "Shopping", 10],
    ["ZARA", "Living", "Shopping", 10],
    ["NETFLIX", "Living", "Shopping", 10],
    ["SPOTIFY", "Living", "Shopping", 10]
  ];

  sheet.getRange(2, 1, rules.length, headers.length).setValues(rules);
}

// ─── 2. PRESENTATION LAYOUTS ────────────────────────────────────────────────

/**
 * Builds the Dashboard tab.
 */
function buildDashboardTab_(sheet) {
  sheet.setColumnWidth(1, 24); // Gutter
  for (let c = 2; c <= 8; c++) sheet.setColumnWidth(c, 110);

  // Row 2: Title
  sheet.getRange("B2:H2").merge().setValue("Victoria Court · Money")
    .setFontFamily("Arial").setFontSize(20).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY);

  // Row 3: Subtitle
  sheet.getRange("B3:H3").merge().setFormula('="Puerto Rico · USD · " & TEXT(TODAY(),"MMMM YYYY")')
    .setFontFamily("Arial").setFontSize(11).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  // ── BLOCK 1: THE ASK (Rows 5–9) ──
  const askBox = sheet.getRange("B5:H9");
  askBox.setBackground(SCRIPT_CONFIG.COLORS.BG_LIGHT_YELLOW);
  askBox.setBorder(true, true, true, true, false, false, SCRIPT_CONFIG.COLORS.BORDER_ORANGE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("B5:E5").merge().setValue("PHASE 1 MOVE-IN FUNDING GAP (THE ASK)")
    .setFontSize(10).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.BORDER_ORANGE);

  sheet.getRange("B6:E7").merge().setValue(33259.33)
    .setFontSize(32).setFontWeight("bold").setNumberFormat("$#,##0.00").setFontColor(SCRIPT_CONFIG.COLORS.STATUS_OVER_TXT);

  sheet.getRange("B8:E8").merge().setValue("Phase 1 still to pay ($35,891.00) − Daniel salary surplus ($2,631.67) = Ask for Father")
    .setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  sheet.getRange("F6:H6").merge().setValue("Daniel contributes: $2,631.67")
    .setFontSize(11).setFontWeight("bold");

  sheet.getRange("F7:H7").merge().setValue("🔴 Ask Required")
    .setFontSize(12).setFontWeight("bold");

  sheet.getRange("B9:H9").merge().setValue("📋 4-Phase Home Updates: Phase 1 Move-In ($57.7k) · Phase 2 ($9.5k) · Phase 3 ($6.3k) · Phase 4 ($16.8k) · Grand Total: $90,278")
    .setFontSize(9).setFontStyle("italic").setFontColor(SCRIPT_CONFIG.COLORS.STATUS_WATCH_TXT);

  // ── BLOCK 2: MORTGAGE (Rows 11–15, Cols B:E) ──
  const mortBox = sheet.getRange("B11:E15");
  mortBox.setBackground(SCRIPT_CONFIG.COLORS.BG_LIGHT_BLUE);
  mortBox.setBorder(true, true, true, true, false, false, SCRIPT_CONFIG.COLORS.BORDER_BLUE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("B11:E11").merge().setValue("🏠 MORTGAGE")
    .setFontSize(10).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.BORDER_BLUE);

  sheet.getRange("B12:E12").merge().setFormula('=MORTGAGE_PAYMENT')
    .setFontSize(24).setFontWeight("bold").setNumberFormat("$#,##0.00");

  sheet.getRange("B13:E13").merge().setFormula('="First payment " & TEXT(MORTGAGE_FIRST_PAYMENT_DATE, "MM/DD/YYYY") & " · " & MAX(0, MORTGAGE_FIRST_PAYMENT_DATE - TODAY()) & " days"')
    .setFontSize(10).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  sheet.getRange("B14:E14").merge().setFormula('="Covered through " & TEXT(EDATE(TODAY(), IF((MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING)<=0, 0, INT(SUMIFS(Accounts!F:F, Accounts!I:I, TRUE, Accounts!H:H, "<>Closed") / (MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING)))), "MMMM YYYY") & " (" & IF((MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING)<=0, 0, INT(SUMIFS(Accounts!F:F, Accounts!I:I, TRUE, Accounts!H:H, "<>Closed") / (MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING))) & " months)"')
    .setFontSize(10).setFontWeight("bold");

  sheet.getRange("B15:E15").merge().setValue("FHA 30-yr fixed 4.750% · APR 5.137% · $550,000 loan · 50% LTV. Drops to $3,053.47 after 132 payments when MIP cancels.")
    .setFontSize(8).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  // ── BLOCK 3: LIVING THIS MONTH (Rows 11–15, Cols F:H) ──
  const livBox = sheet.getRange("F11:H15");
  livBox.setBackground("#FFFFFF");
  livBox.setBorder(true, true, true, true, false, false, SCRIPT_CONFIG.COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("F11:H11").merge().setValue("LIVING THIS MONTH")
    .setFontSize(10).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY);

  sheet.getRange("F12:H12").merge().setFormula('=LIVING_CEILING - SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY)')
    .setFontSize(24).setFontWeight("bold").setNumberFormat("$#,##0.00");

  sheet.getRange("F13:H13").merge().setFormula('="of " & TEXT(LIVING_CEILING, "$#,##0") & " · " & TEXT(IF(LIVING_CEILING=0, 0, SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY) / LIVING_CEILING), "0%") & " used"')
    .setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  sheet.getRange("F14:H14").merge().setFormula('=SPARKLINE(SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY), {"charttype","bar"; "max",LIVING_CEILING; "color1","#1155CC"})');

  sheet.getRange("F15:H15").merge().setFormula('="Daily: $" & TEXT(IF(DAY(EOMONTH(TODAY(),0))-DAY(TODAY())<=0, LIVING_CEILING - SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY), (LIVING_CEILING - SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY)) / (DAY(EOMONTH(TODAY(),0)) - DAY(TODAY()))), "0.00") & "/day · " & IF((SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY) - (LIVING_CEILING * (DAY(TODAY()) / DAY(EOMONTH(TODAY(),0))))) > 0, "⚠️ Running hot", "✅ On pace")')
    .setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  // ── BLOCK 4: CASH (Rows 17–21) ──
  sheet.getRange("B17:C17").merge().setValue("TOTAL CASH").setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT).setFontWeight("bold");
  sheet.getRange("B18:C19").merge().setFormula('=SUMIFS(Accounts!F:F, Accounts!I:I, TRUE, Accounts!H:H, "<>Closed")').setFontSize(16).setFontWeight("bold").setNumberFormat("$#,##0.00");

  sheet.getRange("D17:E17").merge().setValue("HELD BACK (RUNWAY)").setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT).setFontWeight("bold");
  sheet.getRange("D18:E19").merge().setFormula('=RUNWAY_MONTHS * (MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING)').setFontSize(16).setFontWeight("bold").setNumberFormat("$#,##0.00");

  sheet.getRange("F17:H17").merge().setValue("AVAILABLE CASH").setFontSize(10).setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY).setFontWeight("bold");
  sheet.getRange("F18:H19").merge().setFormula('=MAX(0, SUMIFS(Accounts!F:F, Accounts!I:I, TRUE, Accounts!H:H, "<>Closed") - (RUNWAY_MONTHS * (MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING)))').setFontSize(22).setFontWeight("bold").setNumberFormat("$#,##0.00").setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY);

  sheet.getRange("B20:H20").merge().setFormula('=IF((RUNWAY_MONTHS * (MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING)) > SUMIFS(Accounts!F:F, Accounts!I:I, TRUE, Accounts!H:H, "<>Closed"), "🔴 Short of your " & RUNWAY_MONTHS & "-month runway by $" & TEXT((RUNWAY_MONTHS * (MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT + LIVING_CEILING)) - SUMIFS(Accounts!F:F, Accounts!I:I, TRUE, Accounts!H:H, "<>Closed"), "#,##0.00"), "✅ Runway fully covered")')
    .setFontSize(9).setFontStyle("italic");

  // ── BLOCK 5: FIXED COMMITMENTS (Rows 22–27) ──
  sheet.getRange("B22:H22").merge().setValue("FIXED COMMITMENTS (MONTHLY)").setFontSize(10).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY);
  sheet.getRange("B23:D23").merge().setValue("🏠 Mortgage (from Nov 2026)");
  sheet.getRange("E23").setFormula('=MORTGAGE_PAYMENT').setNumberFormat("$#,##0.00");
  sheet.getRange("B24:D24").merge().setValue("⚡ Utilities");
  sheet.getRange("E24").setFormula('=IF(UTILITIES_EST=0, "⚠️ not set", UTILITIES_EST)').setNumberFormat("$#,##0.00");
  sheet.getRange("B25:D25").merge().setValue("🏢 HOA (incl. home insurance)");
  sheet.getRange("E25").setFormula('=IF(HOA_AMOUNT=0, "⚠️ not set", HOA_AMOUNT)').setNumberFormat("$#,##0.00");
  sheet.getRange("B26:D26").merge().setValue("Total Fixed Expected").setFontWeight("bold");
  sheet.getRange("E26").setFormula('=MORTGAGE_PAYMENT + UTILITIES_EST + HOA_AMOUNT').setFontWeight("bold").setNumberFormat("$#,##0.00");

  // ── BLOCK 6: HOME UPDATES (Rows 29–32) ──
  sheet.getRange("B29:H29").merge().setValue("HOME UPDATES SUMMARY").setFontSize(10).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY);
  sheet.getRange("B30:E30").merge().setFormula('="Total Spent: $" & TEXT(SUMIFS(Transactions!D:D, Transactions!H:H,"Home"), "#,##0.00") & " · This Month: $" & TEXT(SUMIFS(Transactions!D:D, Transactions!H:H,"Home", Transactions!J:J, CURRENT_MONTH_KEY), "#,##0.00")');
  sheet.getRange("F30:H30").merge().setFormula('="🎯 Move-in " & TEXT(MOVE_IN_DATE, "MM/DD/YYYY") & " · " & MAX(0, MOVE_IN_DATE - TODAY()) & " days"').setFontWeight("bold");

  // ── BLOCK 7: FAMILY BALANCE (Rows 34–36) ──
  const famBox = sheet.getRange("B34:H36");
  famBox.setBackground(SCRIPT_CONFIG.COLORS.BG_LIGHT_GRAY);
  famBox.setBorder(true, true, true, true, false, false, SCRIPT_CONFIG.COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("B34:H34").merge().setValue("FAMILY BALANCE (UNSETTLED)").setFontSize(9).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);
  sheet.getRange("B35:H35").merge().setValue("Owed to Milton (unsettled): —").setFontWeight("bold").setFontSize(12);
  sheet.getRange("B36:H36").merge().setValue("$1,100,000 paid in cash by father, June 2026. $550,000 repaid via this mortgage. Remaining $550,000 not yet settled between us.").setFontSize(8).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  // Footer (Row 38)
  sheet.getRange("B38:H38").merge().setFormula('="Banks synced " & TEXT(LAST_BANK_SYNC, "MM/DD/YYYY HH:mm") & " · Renovation figures from " & TEXT(LAST_RENO_SYNC, "MM/DD/YYYY") & " (" & (TODAY() - DATEVALUE(LAST_RENO_SYNC)) & " days old)"')
    .setFontSize(8).setFontStyle("italic").setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);
}

/**
 * Builds the Living Tab.
 */
function buildLivingTab_(sheet) {
  sheet.setColumnWidth(1, 24); // Gutter
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 200);

  // Header
  sheet.getRange("B2:E2").merge().setFormula('="Living — " & TEXT(TODAY(),"MMMM YYYY")')
    .setFontFamily("Arial").setFontSize(18).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY);

  // Large Gauge Summary Box
  sheet.getRange("B4:E4").merge().setFormula('="Remaining: $" & TEXT(LIVING_CEILING - SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY), "#,##0.00") & " of $" & TEXT(LIVING_CEILING, "#,##0.00") & " (" & TEXT(IF(LIVING_CEILING=0,0,SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY)/LIVING_CEILING), "0%") & " used)"')
    .setFontSize(14).setFontWeight("bold");

  sheet.getRange("B5:E5").merge().setFormula('=SPARKLINE(SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY), {"charttype","bar"; "max",LIVING_CEILING; "color1","#1155CC"})');

  // Category Breakdown Table (DISPLAY ONLY — NO INDIVIDUAL CAPS)
  sheet.getRange("B7:E7").setValues([["Category", "Spent", "Share of Spend", "Visual Bar"]]).setFontWeight("bold").setBackground(SCRIPT_CONFIG.COLORS.BG_LIGHT_GRAY);

  const cats = SCRIPT_CONFIG.CATEGORIES.LIVING;
  cats.forEach((cat, i) => {
    const row = 8 + i;
    sheet.getRange(row, 2).setValue(cat).setFontWeight("bold");
    sheet.getRange(row, 3).setFormula(`=SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!I:I,"${cat}", Transactions!J:J, CURRENT_MONTH_KEY)`).setNumberFormat("$#,##0.00");
    sheet.getRange(row, 4).setFormula(`=IF(SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY)=0, 0, C${row} / SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY))`).setNumberFormat("0.0%");
    sheet.getRange(row, 5).setFormula(`=SPARKLINE(D${row}, {"charttype","bar"; "max",1; "color1","#666666"})`);
  });

  const noteRow = 8 + cats.length + 1;
  sheet.getRange(noteRow, 2, 1, 4).merge().setValue("Categories are shown to explain where the money went. Only the $2,000 total is a limit.")
    .setFontSize(9).setFontStyle("italic").setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);

  // 6-Month Trend Line
  const trendRow = noteRow + 2;
  sheet.getRange(trendRow, 2).setValue("6-Month Living Spend Trend").setFontWeight("bold");
  sheet.getRange(trendRow + 1, 2, 1, 4).merge().setFormula('=SPARKLINE({SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, TEXT(EDATE(TODAY(),-5),"YYYY-MM")), SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, TEXT(EDATE(TODAY(),-4),"YYYY-MM")), SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, TEXT(EDATE(TODAY(),-3),"YYYY-MM")), SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, TEXT(EDATE(TODAY(),-2),"YYYY-MM")), SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, TEXT(EDATE(TODAY(),-1),"YYYY-MM")), SUMIFS(Transactions!D:D, Transactions!H:H,"Living", Transactions!J:J, CURRENT_MONTH_KEY)}, {"charttype","line"; "color","#1155CC"; "linewidth",2})');
}

/**
 * Builds the Home Tab.
 */
function buildHomeTab_(sheet) {
  sheet.setColumnWidth(1, 24); // Gutter
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);

  // Header
  sheet.getRange("B2:F2").merge().setValue("Home Updates — Victoria Court")
    .setFontFamily("Arial").setFontSize(18).setFontWeight("bold").setFontColor(SCRIPT_CONFIG.COLORS.PRIMARY_NAVY);

  // 3 Tiles
  sheet.getRange("B4").setValue("Phase 1 Move-In").setFontWeight("bold").setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);
  sheet.getRange("B5").setValue(57680).setFontSize(16).setFontWeight("bold").setNumberFormat("$#,##0.00");

  sheet.getRange("C4").setValue("P1 Still to Pay").setFontWeight("bold").setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);
  sheet.getRange("C5").setValue(35891).setFontSize(16).setFontWeight("bold").setNumberFormat("$#,##0.00").setFontColor(SCRIPT_CONFIG.COLORS.STATUS_OVER_TXT);

  sheet.getRange("D4").setValue("All 4 Phases Total").setFontWeight("bold").setFontSize(9).setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);
  sheet.getRange("D5").setValue(90278).setFontSize(16).setFontWeight("bold").setNumberFormat("$#,##0.00");

  sheet.getRange("B7:F7").merge().setFormula('="🎯 Move-in " & TEXT(MOVE_IN_DATE, "MM/DD/YYYY") & " · " & MAX(0, MOVE_IN_DATE - TODAY()) & " days"').setFontWeight("bold");

  // 4-Phase Roadmap Table
  sheet.getRange("B9:F9").merge().setValue("4-PHASE HOME UPDATES & PROCUREMENT ROADMAP").setFontWeight("bold").setFontSize(10).setBackground(SCRIPT_CONFIG.COLORS.BG_LIGHT_GRAY);
  sheet.getRange("B10:F10").setValues([["Work Phase & Timing", "Building Works", "Furniture", "Expected Total", "Still to Pay"]]).setFontWeight("bold");
  sheet.getRange("B11:F15").setValues([
    ["Phase 1: Move-In Priority (Aug–Oct)", 33600, 24080, 57680, 35891],
    ["Phase 2: TVs, Audio, Guest 1 (Dec–Mar)", 0, 9501, 9501, 9501],
    ["Phase 3: Suites, Doors & Art (Apr–May)", 0, 6312, 6312, 6312],
    ["Phase 4: Pergola, BBQ & Glass (May–Jul)", 13500, 3285, 16785, 16785],
    ["Grand All-In Total (All 4 Phases)", 47100, 43178, 90278, 68489]
  ]);
  sheet.getRange("C11:F15").setNumberFormat("$#,##0.00");

  // Phase 1 Building Disbursements Schedule
  sheet.getRange("B17:F17").merge().setValue("PHASE 1 BUILDING WORKS SCHEDULE").setFontWeight("bold").setFontSize(10).setBackground(SCRIPT_CONFIG.COLORS.BG_LIGHT_GRAY);
  sheet.getRange("B18:F18").setValues([["Milestone / Trade", "Target Date", "Status", "Contractor", "Amount"]]).setFontWeight("bold");
  sheet.getRange("B19:F23").setValues([
    ["Davis Facias Deposit (50%)", "07/12/2026", "✅ Paid", "Davis Facias", 10700],
    ["Davis Facias Progress (25%)", "Sep 2026", "⏳ Pending", "Davis Facias", 5350],
    ["Davis Facias Final (25%)", "Oct 2026", "⏳ Pending", "Davis Facias", 5350],
    ["Davis Contingencies & Incidentals", "Oct 2026", "⏳ Pending", "Owner Allocation", 2000],
    ["8x Specialty Trades (Closet, Bar, TV, Glass, Drapery)", "Oct 2026", "⏳ Pending", "Specialists", 10200]
  ]);
  sheet.getRange("F19:F23").setNumberFormat("$#,##0.00");

  sheet.getRange("B25:F26").merge().setValue("📋 4-Phase Home Updates: Phase 1 Move-In total is $57,680 ($10,700 paid · $35,891 still to pay). All 4 phases total $90,278 (Daniel: $6,787 · Father: $61,702). Monzo £8k held as emergency reserve.");

  sheet.getRange("B26:F26").merge().setFormula('=IF(RENO_TOTAL_BUDGET=0, "Running total only — no project budget set.", "Budget: $" & TEXT(RENO_TOTAL_BUDGET, "#,##0") & " · " & TEXT(D5/RENO_TOTAL_BUDGET, "0.0%") & " committed")')
    .setFontSize(10).setFontWeight("bold");
  sheet.getRange("B27:F27").merge().setFormula('=IF(RENO_TOTAL_BUDGET=0, "", SPARKLINE(D5, {"charttype","bar"; "max",RENO_TOTAL_BUDGET; "color1","#F57C00"}))');

  // Footer
  sheet.getRange("B29:F29").merge().setFormula('="Figures re-synced manually from Victoria Court Tracker.xlsx · Last sync " & TEXT(LAST_RENO_SYNC, "MM/DD/YYYY HH:mm")')
    .setFontSize(8).setFontStyle("italic").setFontColor(SCRIPT_CONFIG.COLORS.GRAY_TEXT);
}

// ─── 3. CREDENTIALS & SECURITY ──────────────────────────────────────────────

/**
 * Prompts user for SimpleFIN Access URL and stores it securely in Script Properties.
 */
function storeCredentials() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("SimpleFIN Bridge Configuration", "Please paste your SimpleFIN Setup Token or Access URL:", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    let input = response.getResponseText().trim();
    if (!input) return;

    // 1. Check if base64 encoded setup token
    if (!input.startsWith("http")) {
      try {
        const decoded = Utilities.newBlob(Utilities.base64Decode(input)).getDataAsString().trim();
        if (decoded.startsWith("http")) {
          input = decoded;
        }
      } catch (e) {}
    }

    // 2. If it's a claim URL, claim it via POST to obtain the permanent Access URL
    if (input.includes("/claim/")) {
      try {
        const claimRes = UrlFetchApp.fetch(input, { method: "post", muteHttpExceptions: true });
        const code = claimRes.getResponseCode();
        const accessUrl = claimRes.getContentText().trim();
        
        if (code === 200 && accessUrl.startsWith("http")) {
          input = accessUrl;
        } else if (code === 403) {
          ui.alert("⚠️ This SimpleFIN setup token was already used or expired. Please generate a fresh token on bridge.simplefin.org and paste it.");
          return;
        } else {
          ui.alert(`⚠️ SimpleFIN claim returned HTTP ${code}: ${accessUrl}`);
          return;
        }
      } catch (err) {
        ui.alert("Could not claim SimpleFIN URL: " + err.message);
        return;
      }
    }

    if (!input.startsWith("http")) {
      ui.alert("Invalid token or URL. Please provide a valid SimpleFIN Token or Access URL.");
      return;
    }

    // 3. Test connection live with parsed Basic Auth headers
    try {
      const parsed = parseSimpleFinUrl_(input);
      const testRes = UrlFetchApp.fetch(`${parsed.url}/accounts`, {
        headers: parsed.headers,
        muteHttpExceptions: true
      });
      const testCode = testRes.getResponseCode();
      if (testCode === 200) {
        PropertiesService.getScriptProperties().setProperty("SIMPLEFIN_ACCESS_URL", input);
        ui.alert("✅ SimpleFIN connected successfully!\n\nClick 'Sync banks now' in the top menu to import your accounts and transactions.");
      } else {
        ui.alert(`⚠️ SimpleFIN returned HTTP ${testCode}: ${testRes.getContentText().substring(0, 150)}`);
      }
    } catch (err) {
      ui.alert("Connection test failed: " + err.message);
    }
  }
}

/**
 * Clears the stored SimpleFIN Access URL.
 */
function clearCredentials() {
  PropertiesService.getScriptProperties().deleteProperty("SIMPLEFIN_ACCESS_URL");
  SpreadsheetApp.getUi().alert("SimpleFIN credentials have been removed.");
}

// ─── 4. BANK SYNCHRONIZATION ENGINE ─────────────────────────────────────────

/**
 * Connects to SimpleFIN API, syncs balances and 90 days of transactions for active accounts.
 */
function syncBankData() {
  const accessUrl = PropertiesService.getScriptProperties().getProperty("SIMPLEFIN_ACCESS_URL");
  if (!accessUrl) {
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ SimpleFIN not configured. Run storeCredentials() first.", "Sync Paused", 5);
    return;
  }

  const parsed = parseSimpleFinUrl_(accessUrl);
  // Request 30-day window to stay well within SimpleFIN limits
  const safeStart = Math.floor(Date.now() / 1000) - (30 * 24 * 60 * 60);
  let fetchUrl = `${parsed.url}/accounts?start-date=${safeStart}`;

  let responseData;
  try {
    let res = UrlFetchApp.fetch(fetchUrl, {
      headers: parsed.headers,
      muteHttpExceptions: true
    });
    let code = res.getResponseCode();

    // Fallback: If query parameter fails or server complains, fetch plain /accounts
    if (code !== 200) {
      Logger.log(`Query param returned HTTP ${code}. Retrying plain /accounts.`);
      res = UrlFetchApp.fetch(`${parsed.url}/accounts`, {
        headers: parsed.headers,
        muteHttpExceptions: true
      });
      code = res.getResponseCode();
    }

    if (code !== 200) {
      const errorSnippet = res.getContentText().substring(0, 160);
      Logger.log(`SimpleFIN API error: HTTP ${code} - ${errorSnippet}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(`⚠️ SimpleFIN returned HTTP ${code}: ${errorSnippet}`, "Sync Error", 10);
      return;
    }
    responseData = JSON.parse(res.getContentText());
  } catch (err) {
    Logger.log("SimpleFIN Network Exception: " + err.message);
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ Bank sync error: " + err.message, "Sync Error", 8);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.ACCOUNTS);
  const txSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.TRANSACTIONS);
  const now = new Date();

  // Read Accounts
  const accLastRow = accountsSheet.getLastRow();
  if (accLastRow <= 1) {
    SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ Accounts tab is empty. Run setupTracker() first.", "Sync Error", 5);
    return;
  }

  const accData = accountsSheet.getRange(2, 1, accLastRow - 1, 11).getValues();
  const simpleFinAccounts = responseData.accounts || [];

  // Update Account Balances
  simpleFinAccounts.forEach(sAcc => {
    for (let i = 0; i < accData.length; i++) {
      const row = accData[i];
      const feed = row[9];
      const status = row[7];
      let simpleFinId = row[10];

      if (feed === "SimpleFIN" && status !== "Closed") {
        const rName = (row[0] || "").toString().toLowerCase();
        const rInst = (row[1] || "").toString().toLowerCase();
        const sName = (sAcc.name || "").toString().toLowerCase();
        const sOrg = (sAcc.org && sAcc.org.name ? sAcc.org.name : "").toString().toLowerCase();

        const matchesId = simpleFinId && simpleFinId === sAcc.id;
        const matchesName = !simpleFinId && (
          sName.includes(rName) ||
          sOrg.includes(rInst) ||
          rInst.includes(sOrg) ||
          (rName.includes("banco popular") && (sName.includes("popular") || sOrg.includes("popular"))) ||
          (rName.includes("capital one") && (sName.includes("capital") || sOrg.includes("capital") || sName.includes("venture")))
        );

        if (matchesId || matchesName) {
          if (!simpleFinId) {
            accData[i][10] = sAcc.id;
          }
          accData[i][4] = parseFloat(sAcc.balance || 0);
          accData[i][6] = now;
          break;
        }
      }
    }
  });
  accountsSheet.getRange(2, 1, accData.length, 11).setValues(accData);

  // Deduplication Set
  const txLastRow = txSheet.getLastRow();
  const existingTxIds = new Set();
  if (txLastRow > 1) {
    const ids = txSheet.getRange(2, 1, txLastRow - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existingTxIds.add(r[0].toString().trim()); });
  }

  // Process Transactions
  const newTxRows = [];
  simpleFinAccounts.forEach(sAcc => {
    const matchedAccountRow = accData.find(r => r[10] === sAcc.id && r[9] === "SimpleFIN" && r[7] !== "Closed");
    if (!matchedAccountRow) return;

    const accountName = matchedAccountRow[0];
    const currency = matchedAccountRow[3] || "USD";
    const txList = sAcc.transactions || [];

    txList.forEach(t => {
      const txId = (t.id || "").toString().trim();
      if (!txId || existingTxIds.has(txId)) return;

      const txDate = t.posted ? new Date(t.posted * 1000) : new Date();
      const rawDesc = t.description || t.payee || t.memo || "Unknown";
      const merchant = cleanMerchantName_(rawDesc);
      const nativeAmount = parseFloat(t.amount || 0);
      const amountUSD = -nativeAmount; // Negate: spend is positive

      const catResult = categorise(merchant);
      const targetRow = txLastRow + newTxRows.length + 1;

      newTxRows.push([
        txId,
        txDate,
        merchant,
        amountUSD,
        nativeAmount,
        currency,
        accountName,
        catResult.Tier,
        catResult.Category,
        `=TEXT(B${targetRow},"YYYY-MM")`,
        `=YEAR(B${targetRow})`,
        `=MONTH(B${targetRow})`,
        "SimpleFIN",
        now,
        false
      ]);

      existingTxIds.add(txId);
    });
  });

  if (newTxRows.length > 0) {
    txSheet.getRange(txLastRow + 1, 1, newTxRows.length, 15).setValues(newTxRows);
  }

  // Update LAST_BANK_SYNC named range
  const configSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.CONFIG);
  ss.getRangeByName("LAST_BANK_SYNC").setValue(now);

  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Synced ${newTxRows.length} new transactions.`, "Sync Complete", 5);
}

// ─── 5. CATEGORIZATION ENGINE ───────────────────────────────────────────────

/**
 * Categorizes a merchant string using priority-ordered rules.
 */
function categorise(merchant) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.RULES);
  const lastRow = rulesSheet.getLastRow();
  
  if (lastRow <= 1) return { Tier: "Living", Category: "Other" };

  const rules = rulesSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  // Sort rules: Priority ascending, then Match_Text length descending
  rules.sort((a, b) => {
    if (a[3] !== b[3]) return a[3] - b[3];
    return b[0].toString().length - a[0].toString().length;
  });

  const upperMerchant = (merchant || "").toString().toUpperCase();

  for (let i = 0; i < rules.length; i++) {
    const matchText = rules[i][0].toString().toUpperCase();
    if (matchText && upperMerchant.includes(matchText)) {
      return {
        Tier: rules[i][1],
        Category: rules[i][2]
      };
    }
  }

  return { Tier: "Living", Category: "Other" };
}

/**
 * Re-runs categorization over rows where Category = "Other" and Manual_Override = FALSE.
 */
function recategoriseUncategorised() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.TRANSACTIONS);
  const lastRow = txSheet.getLastRow();
  if (lastRow <= 1) return;

  const data = txSheet.getRange(2, 1, lastRow - 1, 15).getValues();
  let updatedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const manualOverride = data[i][14];
    const category = data[i][8];
    const merchant = data[i][2];

    if (!manualOverride && category === "Other") {
      const res = categorise(merchant);
      if (res.Category !== "Other") {
        data[i][7] = res.Tier;
        data[i][8] = res.Category;
        updatedCount++;
      }
    }
  }

  if (updatedCount > 0) {
    txSheet.getRange(2, 1, data.length, 15).setValues(data);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Re-categorised ${updatedCount} transactions.`, "Complete", 5);
}

// ─── 6. IMPORTERS & DATA INTEGRATION ────────────────────────────────────────

/**
 * Interactive CSV importer for Monzo (GBP) and Bank of America (USD) wind-down statements.
 */
function importCsv() {
  const ui = SpreadsheetApp.getUi();
  const accPrompt = ui.prompt("Import Statement", "Enter Account Name exactly as in Accounts tab (e.g. 'Monzo Current' or 'Bank of America'):", ui.ButtonSet.OK_CANCEL);
  if (accPrompt.getSelectedButton() !== ui.Button.OK) return;

  const accountName = accPrompt.getResponseText().trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.ACCOUNTS);
  const accData = accSheet.getRange(2, 1, accSheet.getLastRow() - 1, 4).getValues();
  const matched = accData.find(r => r[0].toLowerCase() === accountName.toLowerCase());

  if (!matched) {
    ui.alert(`Account '${accountName}' not found in Accounts tab.`);
    return;
  }

  const matchedAccount = matched[0];
  const currency = matched[3];
  const gbpRate = ss.getRangeByName("GBP_USD_RATE").getValue() || 1.27;

  const dataPrompt = ui.prompt("Paste CSV Content", "Paste raw CSV/TSV export lines below:", ui.ButtonSet.OK_CANCEL);
  if (dataPrompt.getSelectedButton() !== ui.Button.OK) return;

  const rawText = dataPrompt.getResponseText().trim();
  if (!rawText) return;

  const lines = rawText.split(/\r?\n/);
  const txSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.TRANSACTIONS);
  const lastRow = txSheet.getLastRow();
  const existingIds = new Set();
  if (lastRow > 1) {
    const ids = txSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => { if (r[0]) existingIds.add(r[0]); });
  }

  const newRows = [];
  const now = new Date();

  lines.forEach(line => {
    const parts = line.split(/[,\t]/).map(p => p.replace(/^"|"$/g, '').trim());
    if (parts.length < 3) return;

    // Detect headers
    if (parts[0].toLowerCase().includes("date") || parts[1].toLowerCase().includes("amount")) return;

    const dateStr = parts[0];
    const txDate = new Date(dateStr);
    if (isNaN(txDate.getTime())) return;

    let merchant = parts[1];
    let nativeAmt = parseFloat(parts[2].replace(/[^0-9.-]/g, ''));
    if (isNaN(nativeAmt)) return;

    // Convert spend to positive
    let amountUSD = nativeAmt < 0 ? Math.abs(nativeAmt) : -Math.abs(nativeAmt);
    if (currency === "GBP") {
      amountUSD = amountUSD * gbpRate;
    }

    const cleanMerchant = cleanMerchantName_(merchant);
    const catResult = categorise(cleanMerchant);

    // Synthesize unique SHA-256 hash
    const rawHashInput = `${matchedAccount}|${txDate.toISOString().split('T')[0]}|${cleanMerchant}|${nativeAmt}`;
    const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawHashInput);
    const txId = hash.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').substring(0, 32);

    if (existingIds.has(txId)) return;

    newRows.push([
      txId,
      txDate,
      cleanMerchant,
      amountUSD,
      nativeAmt,
      currency,
      matchedAccount,
      catResult.Tier,
      catResult.Category,
      `=TEXT(B${lastRow + newRows.length + 1},"YYYY-MM")`,
      `=YEAR(B${lastRow + newRows.length + 1})`,
      `=MONTH(B${lastRow + newRows.length + 1})`,
      "CSV",
      now,
      false
    ]);

    existingIds.add(txId);
  });

  if (newRows.length > 0) {
    txSheet.getRange(lastRow + 1, 1, newRows.length, 15).setValues(newRows);
    ui.alert(`✅ Successfully imported ${newRows.length} transactions into ${matchedAccount}.`);
  } else {
    ui.alert("No new unique transactions were found to import.");
  }
}

/**
 * Re-syncs the Renovation tab from pasted CSV/TSV from local Victoria Court Tracker.xlsx.
 */
function importRenovationCSV() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt("Re-sync Renovation Tracker", "Paste the updated tab from Victoria Court Tracker.xlsx below (CSV or TSV format):", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  const rawText = prompt.getResponseText().trim();
  if (!rawText) return;

  const lines = rawText.split(/\r?\n/);
  if (lines.length < 2) {
    ui.alert("Invalid content: at least header + 1 data row required.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const renoSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.RENOVATION);
  
  const parsedRows = [];
  const optionGroupsSeen = {};

  for (let i = 1; i < lines.length; i++) {
    const parts = lines[i].split(/[,\t]/).map(p => p.replace(/^"|"$/g, '').trim());
    if (parts.length < 7) continue;

    const lineKey = parts[0];
    const section = parts[1];
    const room = parts[2];
    const item = parts[3];
    const vendor = parts[4];
    const status = parts[5];
    const amt = parts[6] ? parseFloat(parts[6].replace(/[^0-9.-]/g, '')) : "";
    const isPriced = typeof amt === "number" && !isNaN(amt);
    let isOutstanding = parts[8] ? parts[8].toString().toLowerCase() === "true" : false;
    const optionGroup = parts[9] || "";

    // Option-group rule: exactly one row per option group may be outstanding
    if (optionGroup && isOutstanding) {
      if (optionGroupsSeen[optionGroup]) {
        Logger.log(`Warning: Multiple outstanding options in ${optionGroup}. Flagging false.`);
        isOutstanding = false;
      } else {
        optionGroupsSeen[optionGroup] = true;
      }
    }

    parsedRows.push([
      lineKey, section, room, item, vendor, status, isPriced ? amt : "", isPriced, isOutstanding, optionGroup
    ]);
  }

  if (parsedRows.length > 0) {
    renoSheet.getRange(2, 1, renoSheet.getLastRow(), 10).clearContent();
    renoSheet.getRange(2, 1, parsedRows.length, 10).setValues(parsedRows);
    ss.getRangeByName("LAST_RENO_SYNC").setValue(new Date());
    ui.alert(`✅ Renovation data updated with ${parsedRows.length} items.`);
  }
}

// ─── 7. TRIGGER MANAGEMENT & MENUS ──────────────────────────────────────────

/**
 * Installs daily automatic sync at 06:00 America/Puerto_Rico.
 */
function setupDailyTrigger() {
  removeDailyTrigger();
  ScriptApp.newTrigger("syncBankData")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone(SCRIPT_CONFIG.TIMEZONE)
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Daily 06:00 AM bank sync trigger enabled.", "Trigger Setup", 5);
}

/**
 * Removes existing bank sync triggers.
 */
function removeDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "syncBankData") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function setMonzoBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheet = ss.getSheetByName(SCRIPT_CONFIG.TABS.ACCOUNTS);
  if (!accSheet) return;
  const lastRow = accSheet.getLastRow();
  if (lastRow <= 1) return;
  const data = accSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (let i = 0; i < data.length; i++) {
    if ((data[i][0] || '').toString().toLowerCase().includes('monzo current')) {
      accSheet.getRange(i + 2, 5).setValue(15000); // Balance_Native (GBP)
      accSheet.getRange(i + 2, 9).setValue(true); // Include_In_Available
      accSheet.getRange(i + 2, 7).setValue(new Date()); // Balance_As_Of
      SpreadsheetApp.getActiveSpreadsheet().toast("✅ Monzo Current balance set to £15,000.00 ($19,050.00 USD).", "Updated", 5);
      return;
    }
  }
}

/**
 * Custom Menu creation.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("💰 Budget Tracker")
    .addItem("Sync banks now", "syncBankData")
    .addItem("Set Monzo balance (£15,000)", "setMonzoBalance")
    .addItem("Import CSV (Monzo / BofA)", "importCsv")
    .addItem("Re-sync renovation figures", "importRenovationCSV")
    .addItem("Re-categorise uncategorised", "recategoriseUncategorised")
    .addSeparator()
    .addItem("Setup: create tabs", "setupTracker")
    .addItem("Setup: store SimpleFIN credentials", "storeCredentials")
    .addItem("Setup: enable daily sync", "setupDailyTrigger")
    .addSeparator()
    .addItem("Show Transactions tab", "showTransactions_")
    .addItem("Hide Transactions tab", "hideTransactions_")
    .addToUi();
}

function showTransactions_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCRIPT_CONFIG.TABS.TRANSACTIONS);
  if (s) s.showSheet();
}

function hideTransactions_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCRIPT_CONFIG.TABS.TRANSACTIONS);
  if (s) s.hideSheet();
}

/**
 * JSON API handler to serve data seamlessly to the mobile PWA.
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const accounts = ss.getSheetByName(SCRIPT_CONFIG.TABS.ACCOUNTS).getDataRange().getValues();
    const transactions = ss.getSheetByName(SCRIPT_CONFIG.TABS.TRANSACTIONS).getDataRange().getValues();

    const payload = {
      timestamp: new Date().toISOString(),
      config: {
        living_ceiling: ss.getRangeByName("LIVING_CEILING").getValue(),
        mortgage_payment: ss.getRangeByName("MORTGAGE_PAYMENT").getValue(),
        hoa_amount: ss.getRangeByName("HOA_AMOUNT").getValue() || 375,
        utilities_est: ss.getRangeByName("UTILITIES_EST").getValue() || 0,
        runway_months: ss.getRangeByName("RUNWAY_MONTHS").getValue(),
        move_in_date: ss.getRangeByName("MOVE_IN_DATE").getValue(),
        monthly_salary: 6700,
        first_paycheck: 5136.67
      },
      accounts: accounts.slice(1).map(r => ({
        name: r[0],
        balance: typeof r[5] === 'number' ? r[5] : (typeof r[4] === 'number' ? r[4] : parseFloat((r[5]||r[4]||0).toString().replace(/[^0-9.-]/g, ''))),
        status: r[7]
      })),
      transactions: transactions.slice(1).map(r => {
        let dStr = "";
        if (r[1] instanceof Date) {
          dStr = Utilities.formatDate(r[1], SCRIPT_CONFIG.TIMEZONE, "yyyy-MM-dd");
        } else if (r[1]) {
          dStr = r[1].toString().trim();
        }
        const amt = typeof r[3] === 'number' ? r[3] : parseFloat((r[3]||0).toString().replace(/[^0-9.-]/g, ''));
        return {
          tx_id: (r[0] || '').toString(),
          date: dStr,
          merchant: r[2] || "Unknown",
          amount_usd: isNaN(amt) ? 0 : amt,
          amount: isNaN(amt) ? 0 : amt,
          currency: r[5] || "USD",
          account: r[6] || "Banco Popular",
          tier: r[7] || "Living",
          category: r[8] || "Other"
        };
      }),
      renovation_outstanding: 17440,
      renovation_paid: 10700
    };

    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── HELPER UTILITIES ───────────────────────────────────────────────────────

function cleanMerchantName_(name) {
  if (!name) return "Unknown";
  let cleaned = name.toString().replace(/\s{2,}.*/, '').trim();
  return cleaned.split(' ').map(w => w.charAt(0).toUpperCase() + w.substring(1).toLowerCase()).join(' ');
}

function parseSimpleFinUrl_(fullUrl) {
  if (!fullUrl) return { url: "", headers: {} };
  const trimmed = fullUrl.toString().trim().replace(/\/+$/, '');
  const match = trimmed.match(/^(https?:\/\/)([^:]+):([^@]+)@(.+)$/);
  if (match) {
    const protocol = match[1];
    const user = match[2];
    const pass = match[3];
    const rest = match[4];
    const cleanUrl = protocol + rest;
    const authHeader = "Basic " + Utilities.base64Encode(user + ":" + pass);
    return {
      url: cleanUrl,
      headers: { "Authorization": authHeader }
    };
  }
  return {
    url: trimmed,
    headers: {}
  };
}
