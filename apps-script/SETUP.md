# Setup Instructions

One-time setup — takes about 5 minutes.

## Step 1: Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new sheet
2. Name it **Dividend Tracker**
3. Set up the headers in Row 1:

| A | B | C | D | E |
|---|---|---|---|---|
| Ticker | Shares | Price | Cost Basis | Dividends |

4. Add your INDIV positions starting in Row 2 (A–Z order):

| Ticker | Shares | Price | Cost Basis | Dividends |
|--------|--------|-------|------------|-----------|
| BABO | 1000 | 10.59 | 16751.52 | 6148.59 |
| CHPY | 200 | 74.90 | 10954.50 | 3447.03 |
| LFGY | 200 | 24.36 | 8118.00 | 3056.52 |
| NVDY | 1500 | 14.08 | 25828.42 | 10682.72 |
| PLTY | 100 | 34.70 | 7719.10 | 3019.78 |

5. Add IRA positions below (leave a blank row between):

| Ticker | Shares | Price | Cost Basis | Dividends |
|--------|--------|-------|------------|-----------|
| APLY | 450 | 12.53 | 5851.00 | 1670.13 |
| CONY | 55 | 27.19 | 4348.00 | 3032.07 |
| NVDY | 300 | 14.08 | 4753.09 | 2393.31 |

---

## Step 2: Open the Script Editor

1. In your Google Sheet, click **Extensions → Apps Script**
2. Delete the existing empty function
3. Copy the entire contents of `Code.gs` and paste it in
4. Click the 💾 save icon — name the project **Dividend Tracker**

---

## Step 3: Authorize

1. In the Script Editor, select `runWeeklyUpdate` from the function dropdown
2. Click **Run**
3. Click **Review permissions → Advanced → Go to Dividend Tracker (unsafe)**
4. Click **Allow**

---

## Step 4: Create the automatic triggers

1. Still in the Script Editor, select `createTriggers` from the function dropdown
2. Click **Run**
3. That's it — the script will now run automatically every Wednesday and Thursday at 7–8 AM

---

## Step 5: Verify

Check the **Executions** log (left sidebar clock icon) after the next Wednesday/Thursday to confirm it ran successfully. You should also receive an email at nickt8r@gmail.com.

---

## Updating positions

When you buy more shares or open a new position, just update the relevant row in the Google Sheet directly. The script reads from the sheet each time it runs.
