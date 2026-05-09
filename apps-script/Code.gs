// ============================================================
// DIVIDEND TRACKER — Google Apps Script
// Runs automatically every Wednesday and Thursday morning
// Updates dividend totals, recalculates metrics, emails summary
// ============================================================

const CONFIG = {
  SHEET_NAME:    'Dividend Tracker',
  EMAIL:         'nickt8r@gmail.com',
  TICKER_GROUPS: {
    // GlobeNewsWire search terms per group
    GROUP1: ['CHPY', 'LFGY', 'GOOW', 'HOOW', 'PLTW', 'WPAY'],  // Wed ex-div, Thu pay
    GROUP2: ['BABO', 'NVDY', 'PLTY', 'CONY', 'APLY'],            // Thu ex-div, Fri pay
  },
  INDIV_TICKERS: ['BABO','CHPY','LFGY','NVDY','PLTY'],
  IRA_TICKERS:   ['APLY','CONY','NVDY'],
  WATCHLIST:     ['GOOW','HOOW','PLTW','WPAY'],
};

// ── ENTRY POINT — set this as your trigger ──────────────────
function runWeeklyUpdate() {
  Logger.log('=== Dividend Tracker Update Starting ===');
  const ss    = getOrCreateSheet();
  const divs  = fetchLatestDividends();

  if (Object.keys(divs).length === 0) {
    Logger.log('No new dividends found — skipping update');
    return;
  }

  updateDividendTotals(ss, divs);
  updatePrices(ss);
  const summary = buildSummary(ss, divs);
  sendEmail(summary);
  Logger.log('=== Update Complete ===');
}

// ── FETCH DIVIDENDS from GlobeNewsWire ──────────────────────
function fetchLatestDividends() {
  const results = {};
  const today   = new Date();
  const dayOfWeek = today.getDay(); // 0=Sun,1=Mon,2=Tue,3=Wed,4=Thu,5=Fri

  // Only run on Wed (3) or Thu (4)
  if (dayOfWeek !== 3 && dayOfWeek !== 4) {
    Logger.log('Not a Wed/Thu — manual run detected, continuing anyway');
  }

  // Search GlobeNewsWire for latest YieldMax distribution announcements
  const searchUrl = 'https://www.globenewswire.com/RssFeed/company/yieldmax-etfs';
  try {
    const response = UrlFetchApp.fetch(searchUrl, {muteHttpExceptions: true});
    const xml      = response.getContentText();

    // Parse each ticker we care about
    const allTickers = [...CONFIG.INDIV_TICKERS, ...CONFIG.IRA_TICKERS, ...CONFIG.WATCHLIST];
    allTickers.forEach(ticker => {
      const amount = extractDividendFromXml(xml, ticker);
      if (amount > 0) {
        results[ticker] = amount;
        Logger.log(`${ticker}: $${amount}/share`);
      }
    });
  } catch(e) {
    Logger.log('GlobeNewsWire fetch error: ' + e.message);
    // Fallback: try Market Chameleon
    fetchFromMarketChameleon(results);
  }

  return results;
}

function extractDividendFromXml(xml, ticker) {
  // Look for the ticker name and a dollar amount near it
  const pattern = new RegExp(ticker + '[^$]*\\$([0-9]+\\.[0-9]{4})', 'i');
  const match   = xml.match(pattern);
  if (match) return parseFloat(match[1]);

  // Also try "0.XXXX per share" pattern near ticker
  const idx = xml.indexOf(ticker);
  if (idx > -1) {
    const nearby = xml.substring(idx, idx + 500);
    const m2 = nearby.match(/\$?([0-9]+\.[0-9]{3,4})\s*per\s*share/i);
    if (m2) return parseFloat(m2[1]);
  }
  return 0;
}

function fetchFromMarketChameleon(results) {
  const allTickers = [...CONFIG.INDIV_TICKERS, ...CONFIG.IRA_TICKERS];
  allTickers.forEach(ticker => {
    if (results[ticker]) return; // already found
    try {
      const url      = `https://marketchameleon.com/Overview/${ticker}/Dividends/`;
      const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      const html     = response.getContentText();
      // Find most recent dividend amount in the table
      const match    = html.match(/most.recent[^$]*\$([0-9]+\.[0-9]{3,4})/i) ||
                       html.match(/([0-9]+\.[0-9]{4})\s*<\/td>\s*<td[^>]*>\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i);
      if (match) {
        results[ticker] = parseFloat(match[1]);
        Logger.log(`${ticker} (MC fallback): $${results[ticker]}/share`);
      }
    } catch(e) {
      Logger.log(`Market Chameleon error for ${ticker}: ${e.message}`);
    }
  });
}

// ── SHEET MANAGEMENT ────────────────────────────────────────
function getOrCreateSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // Create new spreadsheet in Drive root
    ss = SpreadsheetApp.create('Dividend Tracker — ' + new Date().getFullYear());
    Logger.log('Created new spreadsheet: ' + ss.getUrl());
  }
  return ss;
}

// ── UPDATE DIVIDEND TOTALS ───────────────────────────────────
function updateDividendTotals(ss, divs) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
  const data  = sheet.getDataRange().getValues();

  // Find rows by ticker and update dividends column
  // Assumes columns: A=Ticker, B=Shares, C=Price, D=CostBasis, E=Dividends, ...
  for (let i = 1; i < data.length; i++) {
    const ticker = data[i][0];
    if (divs[ticker] && data[i][1] > 0) {
      const shares     = data[i][1];
      const oldDivs    = data[i][4] || 0;
      const newPayment = divs[ticker] * shares;
      const newTotal   = oldDivs + newPayment;
      sheet.getRange(i + 1, 5).setValue(newTotal);
      Logger.log(`Updated ${ticker}: +$${newPayment.toFixed(2)} → total $${newTotal.toFixed(2)}`);

      // Log to weekly history sheet
      logWeeklyPayment(ss, ticker, divs[ticker], shares, newPayment);
    }
  }
}

function logWeeklyPayment(ss, ticker, perShare, shares, total) {
  let histSheet = ss.getSheetByName('Dividend History');
  if (!histSheet) {
    histSheet = ss.insertSheet('Dividend History');
    histSheet.appendRow(['Date','Ticker','Per Share','Shares','Total Paid']);
  }
  histSheet.appendRow([new Date(), ticker, perShare, shares, total]);
}

// ── UPDATE PRICES ────────────────────────────────────────────
function updatePrices(ss) {
  const sheet  = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
  const data   = sheet.getDataRange().getValues();
  const allTickers = [...CONFIG.INDIV_TICKERS, ...CONFIG.IRA_TICKERS, ...CONFIG.WATCHLIST];

  for (let i = 1; i < data.length; i++) {
    const ticker = data[i][0];
    if (!allTickers.includes(ticker)) continue;
    try {
      // Use Google Finance formula (works natively in Sheets)
      // We set the formula so the cell auto-updates
      sheet.getRange(i + 1, 3).setFormula(`=GOOGLEFINANCE("${ticker}")`);
    } catch(e) {
      Logger.log(`Price update failed for ${ticker}: ${e.message}`);
    }
  }
}

// ── BUILD EMAIL SUMMARY ──────────────────────────────────────
function buildSummary(ss, divs) {
  const today    = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'EEEE, MMMM d yyyy');
  const sheet    = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
  const data     = sheet.getDataRange().getValues();

  // Build position map from sheet
  const positions = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) positions[data[i][0]] = {
      shares:   data[i][1],
      price:    data[i][2],
      cost:     data[i][3],
      divs:     data[i][4],
    };
  }

  // This week's payments
  let weekTotal  = 0;
  let payRows    = '';
  Object.entries(divs).sort((a,b) => a[0].localeCompare(b[0])).forEach(([tk, pps]) => {
    const p = positions[tk];
    if (!p || p.shares === 0) return;
    const paid = pps * p.shares;
    weekTotal += paid;
    payRows += `
      <tr>
        <td style="padding:8px 12px;font-family:monospace;font-weight:bold;color:#fff">${tk}</td>
        <td style="padding:8px 12px;color:#aaa">$${pps.toFixed(4)}/sh × ${p.shares}</td>
        <td style="padding:8px 12px;color:#00e5a0;text-align:right;font-weight:bold">+$${paid.toFixed(2)}</td>
      </tr>`;
  });

  // Portfolio totals
  let indivVal=0, indivCost=0, indivDivs=0, iraVal=0, iraCost=0, iraDivs=0;
  CONFIG.INDIV_TICKERS.forEach(tk => {
    const p = positions[tk]; if(!p) return;
    indivVal  += p.shares * p.price;
    indivCost += p.cost;
    indivDivs += p.divs;
  });
  CONFIG.IRA_TICKERS.forEach(tk => {
    const p = positions[tk]; if(!p) return;
    iraVal  += p.shares * p.price;
    iraCost += p.cost;
    iraDivs += p.divs;
  });

  const indivNet = (indivVal - indivCost) + indivDivs;
  const iraNet   = (iraVal - iraCost) + iraDivs;

  // YTD avg forecasts (from sheet or hardcoded fallback)
  const indivFcst = 471; // updated by script from YTD avg calc
  const iraFcst   = 89;

  return {
    subject: `Dividend Update — ${today} — $${weekTotal.toFixed(2)} received`,
    html: `
<!DOCTYPE html>
<html>
<body style="background:#0a0e0f;color:#cfe4e8;font-family:'Helvetica Neue',Arial,sans-serif;padding:24px;max-width:600px;margin:0 auto">
  <h2 style="font-size:22px;color:#fff;margin-bottom:4px">Dividend Portfolio Update</h2>
  <p style="color:#6e8f96;font-size:13px;margin-top:0">${today}</p>

  <div style="background:#0f1517;border:1px solid #1f2d30;border-radius:10px;padding:16px 20px;margin:16px 0">
    <p style="font-size:11px;color:#6e8f96;text-transform:uppercase;letter-spacing:.1em;margin:0 0 6px">This Week's Payments</p>
    <table width="100%" cellpadding="0" cellspacing="0">
      ${payRows}
      <tr style="border-top:1px solid #253235">
        <td style="padding:10px 12px;color:#fff;font-weight:bold">TOTAL</td>
        <td></td>
        <td style="padding:10px 12px;color:#00e5a0;text-align:right;font-size:18px;font-weight:bold">$${weekTotal.toFixed(2)}</td>
      </tr>
    </table>
  </div>

  <table width="100%" cellpadding="0" cellspacing="8" style="margin:16px 0">
    <tr>
      <td width="48%" style="background:#0f1517;border:1px solid #1f2d30;border-radius:10px;padding:14px 16px;vertical-align:top">
        <p style="font-size:10px;color:#6e8f96;text-transform:uppercase;letter-spacing:.1em;margin:0 0 6px">INDIV Total Return</p>
        <p style="font-size:20px;font-weight:bold;color:${indivNet>=0?'#00e5a0':'#ef5350'};margin:0">$${indivNet>=0?'+':''}${indivNet.toFixed(0)}</p>
        <p style="font-size:11px;color:#3d5a60;margin:4px 0 0">Forecast $${indivFcst}/wk</p>
      </td>
      <td width="4%"></td>
      <td width="48%" style="background:#0f1517;border:1px solid #1f2d30;border-radius:10px;padding:14px 16px;vertical-align:top">
        <p style="font-size:10px;color:#6e8f96;text-transform:uppercase;letter-spacing:.1em;margin:0 0 6px">IRA Total Return</p>
        <p style="font-size:20px;font-weight:bold;color:${iraNet>=0?'#00e5a0':'#ef5350'};margin:0">$${iraNet>=0?'+':''}${iraNet.toFixed(0)}</p>
        <p style="font-size:11px;color:#3d5a60;margin:4px 0 0">Forecast $${iraFcst}/wk</p>
      </td>
    </tr>
  </table>

  <p style="font-size:11px;color:#3d5a60;margin-top:20px;text-align:center">
    Auto-generated by Dividend Tracker · <a href="https://github.com/nickt8r/dividend-tracker" style="color:#29b6f6">View on GitHub</a>
  </p>
</body>
</html>`
  };
}

// ── SEND EMAIL ───────────────────────────────────────────────
function sendEmail(summary) {
  GmailApp.sendEmail(CONFIG.EMAIL, summary.subject, '', {
    htmlBody: summary.html,
    name:     'Dividend Tracker'
  });
  Logger.log('Email sent to ' + CONFIG.EMAIL);
}

// ── TRIGGER SETUP HELPER ─────────────────────────────────────
// Run this ONCE manually to create the Wed + Thu triggers
function createTriggers() {
  // Delete existing triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Wednesday 7-8 AM
  ScriptApp.newTrigger('runWeeklyUpdate')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(7)
    .create();

  // Thursday 7-8 AM
  ScriptApp.newTrigger('runWeeklyUpdate')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(7)
    .create();

  Logger.log('Triggers created: Wednesday + Thursday 7-8 AM');
}
