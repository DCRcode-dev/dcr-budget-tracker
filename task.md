# Victoria Court Budget Tracker — Implementation Checklist

- [x] 1. Create `task.md` and initialize tracking
- [x] 2. Implement complete `Code.gs` (Google Apps Script project)
  - [x] 2.1 Tab scaffolding & Named Ranges initialization
  - [x] 2.2 Seed data generation (`Accounts`, `Rules`, `Renovation`, `Config`)
  - [x] 2.3 Visual layout builders (`Dashboard`, `Living`, `Home`)
  - [x] 2.4 Formula implementation & conditional formatting
  - [x] 2.5 Security, protection & SimpleFIN credentials management
  - [x] 2.6 Core sync engine (`syncBankData`, `categorise`, `importCsv`, `importRenovationCSV`)
  - [x] 2.7 Trigger management & custom menus (`onOpen`, `doGet`)
- [x] 3. Create non-developer `README.md` with verbatim `KNOWN_OPEN_ITEMS`
- [x] 4. Overhaul Mobile Web App (`index.html`, `manifest.json`, `sw.js`)
  - [x] 4.1 Update branding, typography, palette & layout shell
  - [x] 4.2 Build Dashboard screen with The Ask, Mortgage, Living, and Cash cards
  - [x] 4.3 Build Living screen with $2,000 gauge, explanatory categories & sparklines
  - [x] 4.4 Build Home screen with Renovation disbursements, furniture options & callouts
  - [x] 4.5 Build Settings & Bank Sync screen with CSV uploader & Google Sheets API sync
  - [x] 4.6 Update `manifest.json` and `sw.js` for PWA installation & live updates
- [x] 5. Verify and compile walkthrough documentation
