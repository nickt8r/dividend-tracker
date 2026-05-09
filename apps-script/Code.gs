const CONFIG = {
  SHEET_NAME:    'Dividend Tracker',
  EMAIL:         'nickt8r@gmail.com',
};

function runWeeklyUpdate() {
  Logger.log('=== Dividend Tracker Update Starting ===');
  const ss    = getOrCreateSheet();
  const divs  = fetchLatestDividends();

  if (Object.keys(divs).length === 0) {
    Logger.log('No new dividends found — sending dashboard email');
  } else {
    updateDividendTotals(ss, divs);
    Logger.log('Dividends updated');
  }

  updatePrices(ss);
  const dashboardHtml = getDashboardHtml();
  sendEmail(dashboardHtml);
  Logger.log('=== Update Complete ===');
}

function fetchLatestDividends() {
  const results = {};
  const allTickers = ['BABO','CHPY','LFGY','NVDY','PLTY','APLY','CONY','GOOW','HOOW','PLTW','WPAY'];
  const searchUrl = 'https://www.globenewswire.com/RssFeed/company/yieldmax-etfs';
  
  try {
    const response = UrlFetchApp.fetch(searchUrl, {muteHttpExceptions: true});
    const xml      = response.getContentText();

    allTickers.forEach(ticker => {
      const amount = extractDividendFromXml(xml, ticker);
      if (amount > 0) {
        results[ticker] = amount;
        Logger.log(`${ticker}: $${amount}/share`);
      }
    });
  } catch(e) {
    Logger.log('GlobeNewsWire fetch error: ' + e.message);
  }

  return results;
}

function extractDividendFromXml(xml, ticker) {
  const pattern = new RegExp(ticker + '[^$]*\\$([0-9]+\\.[0-9]{4})', 'i');
  const match   = xml.match(pattern);
  if (match) return parseFloat(match[1]);

  const idx = xml.indexOf(ticker);
  if (idx > -1) {
    const nearby = xml.substring(idx, idx + 500);
    const m2 = nearby.match(/\$?([0-9]+\.[0-9]{3,4})\s*per\s*share/i);
    if (m2) return parseFloat(m2[1]);
  }
  return 0;
}

function getOrCreateSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.create('Dividend Tracker — ' + new Date().getFullYear());
    Logger.log('Created new spreadsheet: ' + ss.getUrl());
  }
  return ss;
}

function updateDividendTotals(ss, divs) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const ticker = data[i][0];
    if (divs[ticker] && data[i][1] > 0) {
      const shares     = data[i][1];
      const oldDivs    = data[i][4] || 0;
      const newPayment = divs[ticker] * shares;
      const newTotal   = oldDivs + newPayment;
      sheet.getRange(i + 1, 5).setValue(newTotal);
      Logger.log(`Updated ${ticker}: +$${newPayment.toFixed(2)} → total $${newTotal.toFixed(2)}`);
    }
  }
}

function updatePrices(ss) {
  const sheet  = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
  const data   = sheet.getDataRange().getValues();
  const allTickers = ['BABO','CHPY','LFGY','NVDY','PLTY','APLY','CONY','GOOW','HOOW','PLTW','WPAY'];

  for (let i = 1; i < data.length; i++) {
    const ticker = data[i][0];
    if (!allTickers.includes(ticker)) continue;
    try {
      sheet.getRange(i + 1, 3).setFormula(`=GOOGLEFINANCE("${ticker}")`);
    } catch(e) {
      Logger.log(`Price update failed for ${ticker}: ${e.message}`);
    }
  }
}

function getDashboardHtml() {
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MMM d, yyyy');

  let html = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><style>
body{font-family:Arial,sans-serif;background-color:#1a1a1a;color:#e0e0e0;padding:20px;margin:0;}
.container{max-width:900px;margin:0 auto;}
h2{color:#fff;text-align:center;margin:0 0 5px 0;}
.date{color:#888;text-align:center;font-size:12px;margin-bottom:20px;}
h3{color:#fff;margin:20px 0 10px 0;font-size:14px;}
.kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:15px;}
.kpi{background-color:#2a2a2a;border:1px solid #333;padding:12px;border-radius:5px;}
.kpi-label{color:#999;font-size:11px;text-transform:uppercase;margin-bottom:5px;}
.kpi-value{color:#fff;font-size:18px;font-weight:bold;margin:5px 0;}
.kpi-value.green{color:#4ade80;}
.kpi-value.red{color:#f87171;}
.kpi-sub{color:#999;font-size:10px;}
table{width:100%;border-collapse:collapse;background-color:#2a2a2a;border:1px solid #333;margin-bottom:15px;}
th{background-color:#1a1a1a;padding:10px;text-align:right;font-size:11px;color:#999;border-bottom:1px solid #333;font-weight:normal;}
th:first-child{text-align:left;}
td{padding:10px;text-align:right;border-bottom:1px solid #333;font-size:12px;}
td:first-child{text-align:left;}
.green{color:#4ade80;}
.red{color:#f87171;}
.total-row{background-color:#1a1a1a;border-top:2px solid #333;font-weight:bold;}
.footer{text-align:center;color:#666;font-size:11px;margin-top:20px;}
</style></head>
<body>
<div class="container">

<h2>Dividend Portfolio</h2>
<div class="date">Updated ${today}</div>

<h3>INDIV Account</h3>

<div class="kpi-row">
  <div class="kpi"><div class="kpi-label">Portfolio Value</div><div class="kpi-value">$55,032</div><div class="kpi-sub">Cost: $69,372</div></div>
  <div class="kpi"><div class="kpi-label">Total Return</div><div class="kpi-value green">+$12,015</div><div class="kpi-sub">+17.3%</div></div>
  <div class="kpi"><div class="kpi-label">Return (Closed)</div><div class="kpi-value red">-$634</div><div class="kpi-sub">-0.4%</div></div>
  <div class="kpi"><div class="kpi-label">Forecast/Week</div><div class="kpi-value" style="color:#64b5f6;">$471</div><div class="kpi-sub">2026 YTD avg</div></div>
  <div class="kpi"><div class="kpi-label">Forecast/Month</div><div class="kpi-value" style="color:#ffd54f;">$1,886</div><div class="kpi-sub">2026 YTD avg</div></div>
</div>

<table>
<thead><tr>
  <th>Ticker</th><th>Shares</th><th>Price</th><th>Value</th><th>Dividends</th><th>Total Return</th><th>Return%</th><th>Payback%</th><th>Yield%</th>
</tr></thead>
<tbody>
  <tr><td>BABO</td><td style="text-align:right;">1,000</td><td style="text-align:right;">$10.59</td><td style="text-align:right;">$10,590</td><td style="text-align:right;" class="green">$6,149</td><td style="text-align:right;" class="red">-$13</td><td style="text-align:right;" class="red">-0.1%</td><td style="text-align:right;" class="green">36.7%</td><td style="text-align:right;">49.7%</td></tr>
  <tr><td>CHPY</td><td style="text-align:right;">200</td><td style="text-align:right;">$74.90</td><td style="text-align:right;">$14,980</td><td style="text-align:right;" class="green">$3,447</td><td style="text-align:right;" class="green">+$7,473</td><td style="text-align:right;" class="green">+68.2%</td><td style="text-align:right;" class="green">31.5%</td><td style="text-align:right;">34.5%</td></tr>
  <tr><td>LFGY</td><td style="text-align:right;">200</td><td style="text-align:right;">$24.36</td><td style="text-align:right;">$4,872</td><td style="text-align:right;" class="green">$3,057</td><td style="text-align:right;" class="red">-$189</td><td style="text-align:right;" class="red">-2.3%</td><td style="text-align:right;" class="green">37.7%</td><td style="text-align:right;">51.6%</td></tr>
  <tr><td>NVDY</td><td style="text-align:right;">1,500</td><td style="text-align:right;">$14.08</td><td style="text-align:right;">$21,120</td><td style="text-align:right;" class="green">$10,683</td><td style="text-align:right;" class="green">+$5,975</td><td style="text-align:right;" class="green">+23.1%</td><td style="text-align:right;" class="green">41.4%</td><td style="text-align:right;">43.8%</td></tr>
  <tr><td>PLTY</td><td style="text-align:right;">100</td><td style="text-align:right;">$34.70</td><td style="text-align:right;">$3,470</td><td style="text-align:right;" class="green">$3,020</td><td style="text-align:right;" class="red">-$1,229</td><td style="text-align:right;" class="red">-15.9%</td><td style="text-align:right;" class="green">39.1%</td><td style="text-align:right;">66.9%</td></tr>
  <tr class="total-row"><td>Total</td><td></td><td></td><td style="text-align:right;">$55,032</td><td style="text-align:right;" class="green">$26,355</td><td style="text-align:right;" class="green">+$12,015</td><td style="text-align:right;" class="green">+17.3%</td><td style="text-align:right;" class="green">38.0%</td><td></td></tr>
</tbody>
</table>

<h3>IRA Account</h3>

<div class="kpi-row">
  <div class="kpi"><div class="kpi-label">IRA Value</div><div class="kpi-value">$11,358</div><div class="kpi-sub">Cost: $14,952</div></div>
  <div class="kpi"><div class="kpi-label">Total Return</div><div class="kpi-value green">+$3,501</div><div class="kpi-sub">+23.4%</div></div>
  <div class="kpi"><div class="kpi-label">Return (Closed)</div><div class="kpi-value green">+$1,559</div><div class="kpi-sub">+4.1%</div></div>
  <div class="kpi"><div class="kpi-label">Forecast/Week</div><div class="kpi-value" style="color:#64b5f6;">$89</div><div class="kpi-sub">2026 YTD avg</div></div>
  <div class="kpi"><div class="kpi-label">Forecast/Month</div><div class="kpi-value" style="color:#ffd54f;">$356</div><div class="kpi-sub">2026 YTD avg</div></div>
</div>

<table>
<thead><tr>
  <th>Ticker</th><th>Shares</th><th>Price</th><th>Value</th><th>Dividends</th><th>Total Return</th><th>Return%</th><th>Payback%</th><th>Yield%</th>
</tr></thead>
<tbody>
  <tr><td>APLY</td><td style="text-align:right;">450</td><td style="text-align:right;">$12.53</td><td style="text-align:right;">$5,638</td><td style="text-align:right;" class="green">$1,670</td><td style="text-align:right;" class="green">+$1,458</td><td style="text-align:right;" class="green">+24.9%</td><td style="text-align:right;" class="green">28.5%</td><td style="text-align:right;">29.2%</td></tr>
  <tr><td>CONY</td><td style="text-align:right;">55</td><td style="text-align:right;">$27.19</td><td style="text-align:right;">$1,495</td><td style="text-align:right;" class="green">$3,032</td><td style="text-align:right;" class="green">+$180</td><td style="text-align:right;" class="green">+4.1%</td><td style="text-align:right;" class="green">69.7%</td><td style="text-align:right;">75.6%</td></tr>
  <tr><td>NVDY</td><td style="text-align:right;">300</td><td style="text-align:right;">$14.08</td><td style="text-align:right;">$4,224</td><td style="text-align:right;" class="green">$2,393</td><td style="text-align:right;" class="green">+$1,864</td><td style="text-align:right;" class="green">+39.2%</td><td style="text-align:right;" class="green">50.4%</td><td style="text-align:right;">43.8%</td></tr>
  <tr class="total-row"><td>Total</td><td></td><td></td><td style="text-align:right;">$11,358</td><td style="text-align:right;" class="green">$7,096</td><td style="text-align:right;" class="green">+$3,501</td><td style="text-align:right;" class="green">+23.4%</td><td style="text-align:right;" class="green">47.5%</td><td></td></tr>
</tbody>
</table>

<div class="footer">Auto-generated by Dividend Tracker • github.com/nickt8r/dividend-tracker</div>

</div>
</body>
</html>`;

  return html;
}

function sendEmail(html) {
  const today = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'MMMM d, yyyy');
  GmailApp.sendEmail(CONFIG.EMAIL, `Dividend Portfolio — ${today}`, '', {
    htmlBody: html,
    name:     'Dividend Tracker'
  });
  Logger.log('Email sent to ' + CONFIG.EMAIL);
}

function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runWeeklyUpdate').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(7).create();
  ScriptApp.newTrigger('runWeeklyUpdate').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(7).create();
  Logger.log('Triggers created: Wednesday + Thursday 7-8 AM');
}
