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
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
:root{--bg:#080c0d;--s1:#0f1517;--s2:#141b1d;--s3:#1a2325;--bdr:#1f2d30;--g:#00e5a0;--b:#29b6f6;--am:#ffb74d;--rd:#ef5350;--tx:#cfe4e8;--tx2:#6e8f96;--tx3:#3d5a60;}
body{background:var(--bg);color:var(--tx);font-family:monospace;font-size:12px;padding:20px;max-width:900px;margin:0 auto;}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px;}
.kpi{background:var(--s1);border:1px solid var(--bdr);border-radius:9px;padding:14px 16px;}
.kl{font-size:9px;color:var(--tx2);text-transform:uppercase;margin-bottom:7px;}
.kv{font-size:20px;font-weight:bold;color:#fff;}
.kv.g{color:var(--g);}.kv.r{color:var(--rd);}.kv.b{color:var(--b);}.kv.am{color:var(--am);}
.ks{font-size:10px;color:var(--tx3);margin-top:4px;}
.card{background:var(--s1);border:1px solid var(--bdr);border-radius:9px;margin-bottom:12px;overflow:hidden;}
.ch{padding:12px 16px;border-bottom:1px solid var(--bdr);font-weight:bold;font-size:12px;}
table{width:100%;border-collapse:collapse;}
th{padding:7px 9px;text-align:right;font-size:9px;color:var(--tx3);border-bottom:1px solid var(--bdr);background:var(--s2);}
th:first-child{text-align:left;}
td{padding:7px 9px;text-align:right;font-size:11px;}
td:first-child{text-align:left;}
.g{color:var(--g);}.r{color:var(--rd);}
.tft td{background:var(--s2);border-top:1px solid var(--bdr2);font-size:10px;color:var(--tx2);}
h3{font-size:16px;color:#fff;margin:20px 0 10px;font-weight:bold;}
</style>
</head>
<body>
<div style="text-align:center;margin-bottom:20px;"><h2 style="margin:0;color:#fff;">Dividend Portfolio</h2><p style="color:var(--tx3);margin:5px 0 0;">Updated ${today}</p></div>
`;

  html += `<h3>INDIV</h3>
<div class="kpis">
  <div class="kpi"><div class="kl">Portfolio Value</div><div class="kv">$55,032</div><div class="ks">Cost $69,372</div></div>
  <div class="kpi"><div class="kl">Total Return</div><div class="kv g">+$12,015</div><div class="ks">+17.3%</div></div>
  <div class="kpi" style="background:#0f2218;"><div class="kl">Total Return — Closed</div><div class="kv r">-$634</div><div class="ks">-0.4%</div></div>
  <div class="kpi"><div class="kl">Forecast / Week</div><div class="kv b">$471</div></div>
  <div class="kpi"><div class="kl">Forecast / Month</div><div class="kv am">$1,886</div></div>
</div>
<div class="card">
  <div class="ch">Positions (5 active)</div>
  <table>
    <thead><tr><th>Ticker</th><th>Shares</th><th>Price</th><th>Val</th><th>P/L</th><th>Divs</th><th>Ret%</th><th>PB%</th><th>Yield%</th></tr></thead>
    <tbody>
      <tr><td>BABO</td><td>1,000</td><td>$10.59</td><td>$10,590</td><td class="r">-$6,162</td><td class="g">$6,149</td><td class="r">-0.1%</td><td class="g">36.7%</td><td>49.7%</td></tr>
      <tr><td>CHPY</td><td>200</td><td>$74.90</td><td>$14,980</td><td class="g">+$4,026</td><td class="g">$3,447</td><td class="g">+68.2%</td><td class="g">31.5%</td><td>34.5%</td></tr>
      <tr><td>LFGY</td><td>200</td><td>$24.36</td><td>$4,872</td><td class="r">-$3,246</td><td class="g">$3,057</td><td class="r">-2.3%</td><td class="g">37.7%</td><td>51.6%</td></tr>
      <tr><td>NVDY</td><td>1,500</td><td>$14.08</td><td>$21,120</td><td class="r">-$4,708</td><td class="g">$10,683</td><td class="g">+23.1%</td><td class="g">41.4%</td><td>43.8%</td></tr>
      <tr><td>PLTY</td><td>100</td><td>$34.70</td><td>$3,470</td><td class="r">-$4,249</td><td class="g">$3,020</td><td class="r">-15.9%</td><td class="g">39.1%</td><td>66.9%</td></tr>
    </tbody>
    <tbody style="border-top:1px solid var(--bdr2);">
      <tr style="background:var(--s2);"><td><strong>Total</strong></td><td colspan="2"></td><td>$55,032</td><td class="r">-$14,340</td><td class="g">$26,355</td><td class="g">+17.3%</td><td class="g">38.0%</td><td></td></tr>
    </tbody>
  </table>
</div>

<h3>IRA</h3>
<div class="kpis">
  <div class="kpi"><div class="kl">IRA Value</div><div class="kv">$11,358</div><div class="ks">Cost $14,952</div></div>
  <div class="kpi"><div class="kl">Total Return</div><div class="kv g">+$3,501</div><div class="ks">+23.4%</div></div>
  <div class="kpi" style="background:#0f2218;"><div class="kl">Total Return — Closed</div><div class="kv g">+$1,559</div><div class="ks">+4.1%</div></div>
  <div class="kpi"><div class="kl">Forecast / Week</div><div class="kv b">$89</div></div>
  <div class="kpi"><div class="kl">Forecast / Month</div><div class="kv am">$356</div></div>
</div>
<div class="card">
  <div class="ch">IRA Positions (3 active)</div>
  <table>
    <thead><tr><th>Ticker</th><th>Shares</th><th>Price</th><th>Val</th><th>P/L</th><th>Divs</th><th>Ret%</th><th>PB%</th><th>Yield%</th></tr></thead>
    <tbody>
      <tr><td>APLY</td><td>450</td><td>$12.53</td><td>$5,638</td><td class="r">-$212</td><td class="g">$1,670</td><td class="g">+24.9%</td><td class="g">28.5%</td><td>29.2%</td></tr>
      <tr><td>CONY</td><td>55</td><td>$27.19</td><td>$1,495</td><td class="r">-$2,853</td><td class="g">$3,032</td><td class="g">+4.1%</td><td class="g">69.7%</td><td>75.6%</td></tr>
      <tr><td>NVDY</td><td>300</td><td>$14.08</td><td>$4,224</td><td class="r">-$529</td><td class="g">$2,393</td><td class="g">+39.2%</td><td class="g">50.4%</td><td>43.8%</td></tr>
    </tbody>
    <tbody style="border-top:1px solid var(--bdr2);">
      <tr style="background:var(--s2);"><td><strong>Total</strong></td><td colspan="2"></td><td>$11,358</td><td class="r">-$3,594</td><td class="g">$7,096</td><td class="g">+23.4%</td><td class="g">47.5%</td><td></td></tr>
    </tbody>
  </table>
</div>

<p style="text-align:center;color:var(--tx3);font-size:10px;margin-top:20px;">Auto-generated by Dividend Tracker</p>
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
