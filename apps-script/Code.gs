const CONFIG = {
  SHEET_NAME: 'Dividend Tracker',
  EMAIL:      'nickt8r@gmail.com',
};

function runWeeklyUpdate() {
  Logger.log('=== Dividend Tracker Update Starting ===');
  const ss   = getOrCreateSheet();
  const divs = fetchLatestDividends();
  if (Object.keys(divs).length === 0) {
    Logger.log('No new dividends — sending current snapshot');
  } else {
    updateDividendTotals(ss, divs);
  }
  updatePrices(ss);
  sendEmail(getDashboardHtml());
  Logger.log('=== Done ===');
}

function fetchLatestDividends() {
  const results = {};
  const tickers = ['BABO','CHPY','LFGY','NVDY','PLTY','APLY','CONY','GOOW','HOOW','PLTW','WPAY'];
  try {
    const xml = UrlFetchApp.fetch('https://www.globenewswire.com/RssFeed/company/yieldmax-etfs',{muteHttpExceptions:true}).getContentText();
    tickers.forEach(tk => {
      const a = extractDiv(xml, tk);
      if (a > 0) { results[tk] = a; Logger.log(`${tk}: $${a}`); }
    });
  } catch(e) { Logger.log('Fetch error: '+e.message); }
  return results;
}

function extractDiv(xml, tk) {
  const m = xml.match(new RegExp(tk+'[^$]*\\$([0-9]+\\.[0-9]{4})','i'));
  if (m) return parseFloat(m[1]);
  const idx = xml.indexOf(tk);
  if (idx>-1) { const m2=(xml.substring(idx,idx+500)).match(/\$?([0-9]+\.[0-9]{3,4})\s*per\s*share/i); if(m2) return parseFloat(m2[1]); }
  return 0;
}

function getOrCreateSheet() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Dividend Tracker '+new Date().getFullYear());
}

function updateDividendTotals(ss, divs) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME)||ss.getActiveSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i=1;i<data.length;i++) {
    const tk=data[i][0];
    if (divs[tk]&&data[i][1]>0) {
      const total=(data[i][4]||0)+divs[tk]*data[i][1];
      sheet.getRange(i+1,5).setValue(total);
      Logger.log(`${tk} → $${total.toFixed(2)}`);
    }
  }
}

function updatePrices(ss) {
  const sheet=ss.getSheetByName(CONFIG.SHEET_NAME)||ss.getActiveSheet();
  const data=sheet.getDataRange().getValues();
  const tks=['BABO','CHPY','LFGY','NVDY','PLTY','APLY','CONY','GOOW','HOOW','PLTW','WPAY'];
  for(let i=1;i<data.length;i++) {
    if(tks.includes(data[i][0])) try{sheet.getRange(i+1,3).setFormula(`=GOOGLEFINANCE("${data[i][0]}")`)}catch(e){}
  }
}

// ─── HELPERS ────────────────────────────────────────────────
function kpi(label, value, sub, color) {
  return `<td width="20%" style="padding:4px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
      <tr><td style="padding:12px;">
        <div style="color:#aaa;font-size:10px;font-family:Arial,sans-serif;text-transform:uppercase;margin-bottom:6px;">${label}</div>
        <div style="color:${color||'#ffffff'};font-size:20px;font-weight:bold;font-family:Arial,sans-serif;margin-bottom:4px;">${value}</div>
        <div style="color:#888;font-size:10px;font-family:Arial,sans-serif;">${sub}</div>
      </td></tr>
    </table>
  </td>`;
}

function posRow(ticker, shares, price, val, pl, divs, ret, pb, exDay, bg) {
  const plColor = pl.startsWith('-') ? '#f87171' : '#4ade80';
  const retColor = ret.startsWith('-') ? '#f87171' : '#4ade80';
  const exBg = exDay==='WED' ? '#1a3a4a' : '#1a4a3a';
  const exColor = exDay==='WED' ? '#64b5f6' : '#4ade80';
  return `<tr style="background-color:${bg};">
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:#fff;border-bottom:1px solid #333;">
      ${ticker} <span style="background-color:${exBg};color:${exColor};font-size:9px;padding:2px 5px;border-radius:3px;">${exDay}</span>
    </td>
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#ddd;text-align:right;border-bottom:1px solid #333;">${shares}</td>
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#ddd;text-align:right;border-bottom:1px solid #333;">${price}</td>
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#ddd;text-align:right;border-bottom:1px solid #333;">${val}</td>
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:${plColor};text-align:right;border-bottom:1px solid #333;">${pl}</td>
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;text-align:right;border-bottom:1px solid #333;">${divs}</td>
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:${retColor};text-align:right;border-bottom:1px solid #333;">${ret}</td>
    <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;text-align:right;border-bottom:1px solid #333;">${pb}</td>
  </tr>`;
}

function tableHeader(cols) {
  return `<tr style="background-color:#1a1a1a;">${cols.map(c=>`<th style="padding:8px 10px;font-family:Arial,sans-serif;font-size:10px;color:#aaa;text-align:${c.left?'left':'right'};border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">${c.label}</th>`).join('')}</tr>`;
}

function pbBar(ticker, pct, recv, cost) {
  const filled = Math.round(pct);
  return `<td style="padding:6px 4px;width:${100/5}%">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
      <tr><td style="padding:10px 12px;">
        <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
          <span style="font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:#fff;">${ticker}</span>
          <span style="font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#4ade80;">${pct}%</span>
        </div>
        <div style="font-family:Arial,sans-serif;font-size:9px;color:#888;margin-bottom:6px;">$${recv.toLocaleString()} of $${cost.toLocaleString()}</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#1a1a1a;border-radius:3px;height:5px;">
          <tr>
            <td width="${filled}%" style="background:linear-gradient(90deg,#64b5f6,#4ade80);border-radius:3px;height:5px;"></td>
            <td width="${100-filled}%" style="height:5px;"></td>
          </tr>
        </table>
      </td></tr>
    </table>
  </td>`;
}

function calRow(ticker, date, pps, amt) {
  return `<tr>
    <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:#fff;border-bottom:1px solid #333;width:60px;">${ticker}</td>
    <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;border-bottom:1px solid #333;">${date} · ${pps}</td>
    <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;text-align:right;border-bottom:1px solid #333;">${amt}</td>
    <td style="padding:8px 12px;border-bottom:1px solid #333;text-align:right;">
      <span style="background-color:#1a3a2a;color:#4ade80;font-size:9px;padding:2px 6px;border-radius:3px;border:1px solid #2a5a3a;">PAID</span>
    </td>
  </tr>`;
}

function barCell(heightPct, label, isSpike) {
  const color = isSpike ? '#ffd54f' : '#4ade80';
  const opacity = isSpike ? '0.9' : '0.65';
  return `<td style="text-align:center;vertical-align:bottom;padding:0 3px;width:12%;">
    <table width="100%" cellpadding="0" cellspacing="0" style="height:100px;vertical-align:bottom;">
      <tr><td style="vertical-align:bottom;height:100px;">
        <div style="background-color:${color};opacity:${opacity};width:100%;height:${heightPct}px;border-radius:2px 2px 0 0;"></div>
      </td></tr>
    </table>
    <div style="font-family:Arial,sans-serif;font-size:9px;color:#888;margin-top:3px;">${label}</div>
  </td>`;
}

function sectionHeader(text) {
  return `<tr><td colspan="2" style="padding:20px 0 10px 0;">
    <div style="font-family:Arial,sans-serif;font-size:16px;font-weight:bold;color:#fff;border-left:3px solid #4ade80;padding-left:10px;">${text}</div>
  </td></tr>`;
}

// ─── MAIN HTML ───────────────────────────────────────────────
function getDashboardHtml() {
  const today = Utilities.formatDate(new Date(),'America/Los_Angeles','MMM d, yyyy');

  const html = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="background-color:#111;margin:0;padding:20px;">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:900px;margin:0 auto;">

  <!-- HEADER -->
  <tr><td colspan="2" style="padding-bottom:20px;text-align:center;">
    <div style="font-family:Arial,sans-serif;font-size:26px;font-weight:bold;color:#fff;">Dividend Portfolio</div>
    <div style="font-family:Arial,sans-serif;font-size:12px;color:#888;margin-top:4px;">Updated ${today}</div>
  </td></tr>

  ${sectionHeader('INDIV')}

  <!-- INDIV KPIs -->
  <tr><td colspan="2" style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        ${kpi('Portfolio Value','$55,032','Cost: $69,372','')}
        ${kpi('Total Return','+$12,015','+17.3%','#4ade80')}
        ${kpi('Closed Positions','-$634','-0.4%','#f87171')}
        ${kpi('Forecast / Week','$471','2026 YTD avg','#64b5f6')}
        ${kpi('Forecast / Month','$1,886','2026 YTD avg','#ffd54f')}
      </tr>
    </table>
  </td></tr>

  <!-- INDIV POSITIONS TABLE -->
  <tr><td colspan="2" style="padding-bottom:6px;">
    <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Positions</div>
  </td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;overflow:hidden;">
      ${tableHeader([{label:'Ticker',left:true},{label:'Shares'},{label:'Price'},{label:'Value'},{label:'P/L'},{label:'Dividends'},{label:'Total Ret%'},{label:'Payback%'}])}
      ${posRow('BABO','1,000','$10.59','$10,590','-$6,162','$6,149','-0.1%','36.7%','THU','#2a2a2a')}
      ${posRow('CHPY','200','$74.90','$14,980','+$4,026','$3,447','+68.2%','31.5%','WED','#252525')}
      ${posRow('LFGY','200','$24.36','$4,872','-$3,246','$3,057','-2.3%','37.7%','WED','#2a2a2a')}
      ${posRow('NVDY','1,500','$14.08','$21,120','-$4,708','$10,683','+23.1%','41.4%','THU','#252525')}
      ${posRow('PLTY','100','$34.70','$3,470','-$4,249','$3,020','-15.9%','39.1%','THU','#2a2a2a')}
      <tr style="background-color:#1a1a1a;">
        <td colspan="3" style="padding:9px 10px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;font-weight:bold;border-top:2px solid #444;">TOTAL</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#fff;font-weight:bold;text-align:right;border-top:2px solid #444;">$55,032</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#f87171;font-weight:bold;text-align:right;border-top:2px solid #444;">-$14,340</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">$26,355</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">+17.3%</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">38.0%</td>
      </tr>
    </table>
  </td></tr>

  <!-- INDIV PAYBACK -->
  <tr><td colspan="2" style="padding-bottom:6px;">
    <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Payback Progress</div>
  </td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        ${pbBar('NVDY',41.4,10683,25828)}
        ${pbBar('PLTY',39.1,3020,7719)}
        ${pbBar('LFGY',37.7,3057,8118)}
        ${pbBar('BABO',36.7,6149,16752)}
        ${pbBar('CHPY',31.5,3447,10955)}
      </tr>
    </table>
  </td></tr>

  <!-- INDIV BOTTOM ROW: chart + calendar -->
  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <!-- WEEKLY DIV RECEIVED -->
        <td width="55%" style="vertical-align:top;padding-right:8px;">
          <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Weekly Div Received</div>
          <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
            <tr><td style="padding:14px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #444;">
                <tr>
                  ${barCell(74,'Mar17',false)}
                  ${barCell(79,'Mar24',false)}
                  ${barCell(74,'Mar31',false)}
                  ${barCell(69,'Apr7',false)}
                  ${barCell(71,'Apr14',false)}
                  ${barCell(74,'Apr21',false)}
                  ${barCell(79,'Apr28',false)}
                  ${barCell(100,'May5',true)}
                </tr>
              </table>
              <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;border-top:1px solid #333;padding-top:10px;">
                <tr>
                  <td style="font-family:Arial,sans-serif;font-size:10px;color:#aaa;">This week<br><span style="font-size:16px;font-weight:bold;color:#4ade80;">$482</span></td>
                  <td style="text-align:right;font-family:Arial,sans-serif;font-size:10px;color:#aaa;">8-week avg<br><span style="font-size:16px;font-weight:bold;color:#fff;">$476</span></td>
                </tr>
              </table>
            </td></tr>
          </table>
        </td>
        <!-- DIV FORECAST -->
        <td width="45%" style="vertical-align:top;padding-left:8px;">
          <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Div Forecast</div>
          <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
            <tr><td style="padding:6px 0;">
              <table width="100%" cellpadding="0" cellspacing="0">
                ${calRow('CHPY','Thu May 7','$0.6024/sh','+$120.48')}
                ${calRow('LFGY','Thu May 7','$0.2513/sh','+$50.26')}
                ${calRow('BABO','Fri May 8','$0.0924/sh','+$92.40')}
                ${calRow('NVDY','Fri May 8','$0.1281/sh','+$192.15')}
                ${calRow('PLTY','Fri May 8','$0.2708/sh','+$27.08')}
              </table>
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#1a1a1a;border-top:1px solid #444;">
                <tr>
                  <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;">Week total</td>
                  <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#4ade80;text-align:right;">$482.37</td>
                </tr>
              </table>
            </td></tr>
          </table>
        </td>
      </tr>
    </table>
  </td></tr>

  ${sectionHeader('IRA')}

  <!-- IRA KPIs -->
  <tr><td colspan="2" style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        ${kpi('IRA Value','$11,358','Cost: $14,952','')}
        ${kpi('Total Return','+$3,501','+23.4%','#4ade80')}
        ${kpi('Closed Positions','+$1,559','+4.1%','#4ade80')}
        ${kpi('Forecast / Week','$89','2026 YTD avg','#64b5f6')}
        ${kpi('Forecast / Month','$356','2026 YTD avg','#ffd54f')}
      </tr>
    </table>
  </td></tr>

  <!-- IRA POSITIONS TABLE -->
  <tr><td colspan="2" style="padding-bottom:6px;">
    <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Positions</div>
  </td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;overflow:hidden;">
      ${tableHeader([{label:'Ticker',left:true},{label:'Shares'},{label:'Price'},{label:'Value'},{label:'P/L'},{label:'Dividends'},{label:'Total Ret%'},{label:'Payback%'}])}
      ${posRow('APLY','450','$12.53','$5,638','-$212','$1,670','+24.9%','28.5%','THU','#2a2a2a')}
      ${posRow('CONY','55','$27.19','$1,495','-$2,853','$3,032','+4.1%','69.7%','THU','#252525')}
      ${posRow('NVDY','300','$14.08','$4,224','-$529','$2,393','+39.2%','50.4%','THU','#2a2a2a')}
      <tr style="background-color:#1a1a1a;">
        <td colspan="3" style="padding:9px 10px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;font-weight:bold;border-top:2px solid #444;">TOTAL</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#fff;font-weight:bold;text-align:right;border-top:2px solid #444;">$11,358</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#f87171;font-weight:bold;text-align:right;border-top:2px solid #444;">-$3,594</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">$7,096</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">+23.4%</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">47.5%</td>
      </tr>
    </table>
  </td></tr>

  <!-- IRA PAYBACK -->
  <tr><td colspan="2" style="padding-bottom:6px;">
    <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Payback Progress</div>
  </td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        ${pbBar('CONY',69.7,3032,4348)}
        ${pbBar('NVDY',50.4,2393,4753)}
        ${pbBar('APLY',28.5,1670,5851)}
        <td width="40%" style="padding:4px;"></td>
      </tr>
    </table>
  </td></tr>

  <!-- IRA BOTTOM ROW: chart + calendar -->
  <tr><td colspan="2" style="padding-bottom:30px;">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td width="55%" style="vertical-align:top;padding-right:8px;">
          <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Weekly Div Received</div>
          <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
            <tr><td style="padding:14px;">
              <table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #444;">
                <tr>
                  ${barCell(77,'Mar17',false)}
                  ${barCell(81,'Mar24',false)}
                  ${barCell(68,'Mar31',false)}
                  ${barCell(69,'Apr7',false)}
                  ${barCell(72,'Apr14',false)}
                  ${barCell(80,'Apr21',false)}
                  ${barCell(90,'Apr28',false)}
                  ${barCell(100,'May5',true)}
                </tr>
              </table>
              <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:10px;border-top:1px solid #333;padding-top:10px;">
                <tr>
                  <td style="font-family:Arial,sans-serif;font-size:10px;color:#aaa;">This week<br><span style="font-size:16px;font-weight:bold;color:#4ade80;">$109</span></td>
                  <td style="text-align:right;font-family:Arial,sans-serif;font-size:10px;color:#aaa;">8-week avg<br><span style="font-size:16px;font-weight:bold;color:#fff;">$97</span></td>
                </tr>
              </table>
            </td></tr>
          </table>
        </td>
        <td width="45%" style="vertical-align:top;padding-left:8px;">
          <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding:0 0 6px 0;">Div Forecast</div>
          <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
            <tr><td style="padding:6px 0;">
              <table width="100%" cellpadding="0" cellspacing="0">
                ${calRow('APLY','Fri May 8','$0.1033/sh','+$46.49')}
                ${calRow('CONY','Fri May 8','$0.4464/sh','+$24.55')}
                ${calRow('NVDY','Fri May 8','$0.1281/sh','+$38.43')}
              </table>
              <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#1a1a1a;border-top:1px solid #444;">
                <tr>
                  <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;">Week total</td>
                  <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#4ade80;text-align:right;">$109.47</td>
                </tr>
              </table>
            </td></tr>
          </table>
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- FOOTER -->
  <tr><td colspan="2" style="text-align:center;padding-top:10px;border-top:1px solid #333;">
    <div style="font-family:Arial,sans-serif;font-size:11px;color:#555;">Auto-generated by Dividend Tracker</div>
  </td></tr>

</table>
</body>
</html>`;

  return html;
}

function sendEmail(html) {
  const today = Utilities.formatDate(new Date(),'America/Los_Angeles','MMMM d, yyyy');
  GmailApp.sendEmail(CONFIG.EMAIL, `Dividend Portfolio — ${today}`, '', {htmlBody:html, name:'Dividend Tracker'});
  Logger.log('Email sent to '+CONFIG.EMAIL);
}

function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runWeeklyUpdate').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(7).create();
  ScriptApp.newTrigger('runWeeklyUpdate').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(7).create();
  Logger.log('Triggers created');
}
