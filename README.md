# Budget Classifier

A personal finance tool that imports Capital One and Venmo transactions, auto-categorizes them, checks for duplicates, and writes them to a Google Sheet — all from a single HTML file backed by a Google Apps Script web app.

## How It Works

1. Open `index.html` in your browser
2. Paste your Apps Script web app URL into the Setup field
3. Drop in a Capital One CSV and/or a Venmo statement CSV
4. Review and adjust categories in the table; click the **✓** at the end of each flagged row when it looks correct
5. Click **Check duplicates & write →** to push new transactions to your Google Sheet

## Project Structure

```
budget-app/
├── index.html   # Browser UI — parses CSVs, classifies, renders table, calls Apps Script
└── Code.gs      # Google Apps Script — reads/writes to the Google Sheet
```

## Google Sheet Setup

The spreadsheet must have **12 tabs** named `January` through `December`. Each tab uses this column layout:

| Column | Field              |
|--------|--------------------|
| A      | Date               |
| B      | Business / Merchant |
| C      | Amount             |
| D      | Category           |
| E      | Spending Category  |
| F      | Notes              |

Row 1 is a header row and is skipped on all reads and writes.

## Deploying the Apps Script

1. Open your Google Sheet → **Extensions → Apps Script**
2. Paste the contents of `Code.gs` into the editor
3. Click **Deploy → New deployment** → type: **Web app**
4. Set **Who has access** to `Anyone`
5. Copy the web app URL and paste it into the Setup field in `index.html`

## Categorization Logic

Transactions are classified in priority order:

1. **User-saved rules** — stored in `localStorage`; created by clicking "Save rule" on any row
2. **Built-in rules** — hardcoded merchant-to-category mappings (e.g. `LYFT → Ride Share`, `TRADER JOE → Grocery`)
3. **Sheet history** — the most recent category used for that merchant in the sheet (fetched on CSV load)
4. **Capital One bank category** — maps Capital One's own categories (Dining, Grocery, Health Care, etc.) to the app's categories
5. **Default fallback** — `Shopping / Fun Spending`, flagged for review

Rows flagged for review (unknown merchants, sheet-history suggestions, and non-grocery Amazon purchases) show a **✓** button in the rightmost column. Click it after you adjust categories to confirm the row; editing a confirmed row clears the checkmark until you confirm again. Use **Needs review** in the filter bar to focus on unconfirmed rows.

## Duplicate Detection

Before writing, the app fetches all existing row keys from the sheet. A key is:

```
{date}|{MERCHANT UPPERCASE}|{amount to 2 decimals}
```

Rows that match an existing key are marked as duplicates and skipped on write.

## Row Coloring

- **Amazon rows** (merchant contains `AMAZON`): yellow background `#fff2cc`
- **All other rows**: inherit the color already used for that category in the sheet

## Categories

| Category | Spending Category |
|---|---|
| Housing | Necessary Costs |
| Transportation | Necessary Costs |
| Ride Share | Fun Spending |
| Grocery | Necessary Costs |
| Health | Necessary Costs |
| Workout | Necessary Costs |
| Dining Out | Fun Spending |
| Activity | Fun Spending |
| Gift | Fun Spending |
| Shopping | Fun Spending |
| Travel | Fun Spending |
| Essentials | Necessary Costs |
| Work Lunch | Necessary Costs |
| Work Coffee | Necessary Costs |
| Pre-Tax Health Insurance | Pre-Tax |

Spending categories also include: `From Savings`, `Investments`, `Saving Goals`, `Post Tax Income`, `Bonus!`

## Supported CSV Formats

**Capital One** — standard transaction export with columns: Transaction Date, Description, Category, Debit, Credit. Payment rows are automatically skipped.

**Venmo** — statement CSV downloaded from the Venmo app. Bank transfers and incomplete transactions are skipped. Sent and received payments are both imported.
