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
  const avgs = getYtdAverages(ss);
  sendEmail(getDashboardHtml(avgs));
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
      logYtdHistory(ss, tk, divs[tk]);
    }
  }
}

function logYtdHistory(ss, ticker, perShare) {
  let hist = ss.getSheetByName('YTD History');
  if (!hist) {
    hist = ss.insertSheet('YTD History');
    hist.appendRow(['Date','Ticker','Per Share']);
    hist.getRange(1,1,1,3).setFontWeight('bold');
  }
  hist.appendRow([new Date(), ticker, perShare]);
  Logger.log(`YTD History: ${ticker} $${perShare}/sh`);
}

function getYtdAverages(ss) {
  // Seed with known 2026 history so averages work from day one
  const seed = {
    BABO: [0.0915,0.1166,0.0864,0.1277,0.1695,0.1190,0.1138,0.0965,0.0984,0.0908,0.0892,0.0928,0.0870,0.0866,0.0913,0.0906,0.0833,0.1012,0.0924],
    CHPY: [0.5040,0.5196,0.5267,0.4826,0.5253,0.5293,0.5259,0.4686,0.4577,0.4384,0.3851,0.4912,0.4089,0.4406,0.5384,0.5008,0.6041,0.6024],
    LFGY: [0.2885,0.2835,0.2822,0.2741,0.2353,0.2170,0.2294,0.2209,0.2356,0.2303,0.2369,0.2253,0.2033,0.2203,0.2356,0.2562,0.2228,0.2513],
    NVDY: [0.1435,0.1054,0.0950,0.0848,0.1076,0.0939,0.1057,0.0944,0.1151,0.1162,0.1197,0.1332,0.1195,0.1148,0.1111,0.1161,0.1401,0.2072,0.1281],
    PLTY: [0.5130,0.4508,0.4130,0.3791,0.3688,0.3591,0.3845,0.3865,0.3933,0.4781,0.7999,0.8018,0.4779,0.4497,0.4548,0.3556,0.3832,0.3607,0.2708],
    CONY: [0.4342,0.4091,0.3965,0.2219,0.3089,0.2838,0.2556,0.2994,0.3177,0.3115,0.5942,0.6138,0.5332,0.3763,0.3767,0.3833,0.4161,0.5307,0.4464],
    APLY: [0.0532,0.0481,0.0415,0.0473,0.0494,0.0584,0.1774,0.0498,0.0911,0.0649,0.0622,0.0471,0.0606,0.0617,0.0613,0.0717,0.0902,0.0958,0.1033],
  };

  // Merge with any new entries from YTD History sheet
  const hist = ss.getSheetByName('YTD History');
  if (hist) {
    const data = hist.getDataRange().getValues();
    const ytdStart = new Date('2026-05-09'); // only count new entries from when sheet was created
    for (let i=1;i<data.length;i++) {
      const date=new Date(data[i][0]), tk=data[i][1], pps=parseFloat(data[i][2]);
      if (!tk||isNaN(pps)||date<ytdStart) continue;
      if (!seed[tk]) seed[tk]=[];
      seed[tk].push(pps);
    }
  }

  // Compute averages
  const avgs = {};
  Object.keys(seed).forEach(tk => {
    const arr = seed[tk];
    avgs[tk] = arr.reduce((s,v)=>s+v,0) / arr.length;
  });
  return avgs;
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
  const filled = Math.round(0);
  return `<td width="20%" style="padding:4px;vertical-align:top;height:1px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;height:100%;">
      <tr><td style="padding:12px;vertical-align:top;">
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
  return `<td style="padding:6px 4px;width:${100/5}%;vertical-align:top;height:1px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;height:100%;">
      <tr><td style="padding:10px 12px;vertical-align:top;">
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
function getDashboardHtml(avgs) {
  const today = Utilities.formatDate(new Date(),'America/Los_Angeles','MMM d, yyyy');

  // Live YTD forecasts from averages
  const shares = {BABO:1000,CHPY:200,LFGY:200,NVDY:1500,PLTY:100,APLY:450,CONY:55,NVDY_IRA:300};
  const a = avgs || {};
  const indivFcstWk  = Math.round((a.BABO||0.1013)*1000 + (a.CHPY||0.4972)*200 + (a.LFGY||0.2416)*200 + (a.NVDY||0.1185)*1500 + (a.PLTY||0.4463)*100);
  const indivFcstMo  = Math.round(indivFcstWk * 4);
  const iraFcstWk    = Math.round((a.APLY||0.0703)*450 + (a.CONY||0.3952)*55 + (a.NVDY||0.1185)*300);
  const iraFcstMo    = Math.round(iraFcstWk * 4);

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
        ${kpi('Forecast / Week','$'+indivFcstWk,'2026 YTD avg','#64b5f6')}
        ${kpi('Forecast / Month','$'+indivFcstMo,'2026 YTD avg','#ffd54f')}
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
        ${kpi('Portfolio Value','$11,358','Cost: $14,952','')}
        ${kpi('Total Return','+$3,501','+23.4%','#4ade80')}
        ${kpi('Closed Positions','+$1,559','+4.1%','#4ade80')}
        ${kpi('Forecast / Week','$'+iraFcstWk,'2026 YTD avg','#64b5f6')}
        ${kpi('Forecast / Month','$'+iraFcstMo,'2026 YTD avg','#ffd54f')}
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
  <tr><td colspan="2" style="text-align:center;padding-top:20px;border-top:1px solid #333;">
    <a href="https://script.google.com/macros/s/AKfycbzHsDpaNec2_fpgthmbHudzFYFnvxw6DvjJzdwMZx8PC6sW-Mk7zwrTgluMnwYX068S9A/exec" style="display:inline-block;background-color:#4ade80;color:#000;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;padding:12px 28px;border-radius:6px;text-decoration:none;">View Full Dashboard →</a>
    <div style="font-family:Arial,sans-serif;font-size:11px;color:#555;margin-top:12px;">Auto-generated by Dividend Tracker</div>
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

// ─── WEB APP — serves interactive dashboard ─────────────────
function doGet() {
  const ss   = getOrCreateSheet();
  const avgs = getYtdAverages(ss);
  const html = HtmlService.createHtmlOutput(getInteractiveDashboard(avgs));
  html.setTitle('Dividend Portfolio');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getInteractiveDashboard(avgs) {
  const a = avgs || {};
  const indivFcstWk = Math.round((a.BABO||0.1013)*1000+(a.CHPY||0.4972)*200+(a.LFGY||0.2416)*200+(a.NVDY||0.1185)*1500+(a.PLTY||0.4463)*100);
  const indivFcstMo = Math.round(indivFcstWk*4);
  const iraFcstWk   = Math.round((a.APLY||0.0703)*450+(a.CONY||0.3952)*55+(a.NVDY||0.1185)*300);
  const iraFcstMo   = Math.round(iraFcstWk*4);
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dividend Portfolio Dashboard</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&family=Syne:wght@400;600;700;800&display=swap');
*{margin:0;padding:0;box-sizing:border-box;}
:root{--bg:#080c0d;--s1:#0f1517;--s2:#141b1d;--s3:#1a2325;--bdr:#1f2d30;--bdr2:#253235;--g:#00e5a0;--b:#29b6f6;--am:#ffb74d;--rd:#ef5350;--pu:#b39ddb;--tx:#cfe4e8;--tx2:#6e8f96;--tx3:#3d5a60;--f:'DM Mono',monospace;--h:'Syne',sans-serif;}
body{background:var(--bg);color:var(--tx);font-family:var(--f);font-size:12px;min-height:100vh;padding:18px 20px 32px;overflow-x:hidden;}
.hdr{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:20px;}
.hdr-t{font-family:var(--h);font-size:28px;font-weight:800;color:#fff;letter-spacing:-.5px;}
.hdr-s{font-size:10px;color:var(--tx2);margin-top:5px;letter-spacing:.12em;text-transform:uppercase;display:flex;align-items:center;gap:6px;}
.dot{width:6px;height:6px;background:var(--g);border-radius:50%;animation:pulse 2s infinite;}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.3;}}
.tabs{display:flex;gap:4px;}
.tab{padding:6px 14px;border-radius:5px;font-family:var(--h);font-size:11px;font-weight:700;letter-spacing:.07em;border:1px solid var(--bdr2);color:var(--tx2);background:var(--s1);cursor:pointer;transition:all .15s;}
.tab.on{background:var(--g);color:#000;border-color:var(--g);}
.ts{font-size:10px;color:var(--tx3);}
.panel{display:none;}.panel.on{display:block;}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px;}
.kpi{background:var(--s1);border:1px solid var(--bdr);border-radius:9px;padding:14px 16px;position:relative;overflow:hidden;}
.kpi::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--g);opacity:.5;}
.kpi.kr::after{background:var(--rd);}.kpi.kb::after{background:var(--b);}.kpi.ka::after{background:var(--am);}.kpi.kp::after{background:var(--pu);}
.kl{font-size:9px;color:var(--tx2);text-transform:uppercase;letter-spacing:.12em;margin-bottom:7px;}
.kv{font-family:var(--h);font-size:20px;font-weight:700;line-height:1;color:#fff;}
.kv.g{color:var(--g);}.kv.r{color:var(--rd);}.kv.b{color:var(--b);}.kv.am{color:var(--am);}
.ks{font-size:10px;color:var(--tx3);margin-top:4px;}.kn{font-size:9px;color:var(--tx3);margin-top:3px;font-style:italic;}
.g2{display:grid;grid-template-columns:1fr 330px;gap:12px;margin-bottom:12px;}
.card{background:var(--s1);border:1px solid var(--bdr);border-radius:9px;overflow:hidden;}
.table-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;width:100%;max-width:calc(100vw - 40px);}
table{border-collapse:collapse;width:750px;}
thead th{padding:7px 9px;text-align:right;font-size:9px;color:var(--tx3);text-transform:uppercase;letter-spacing:.08em;border-bottom:1px solid var(--bdr);background:var(--s2);}
thead th:first-child{text-align:left;}
tbody tr{border-bottom:1px solid #0d1314;transition:background .1s;}
tbody tr:hover{background:#0e1617;}
tbody tr.z{opacity:.35;}
td{padding:7px 9px;text-align:right;font-size:11px;}td:first-child{text-align:left;}
.ch{display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border-bottom:1px solid var(--bdr);}
.ct{font-family:var(--h);font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#fff;}
.bdg{font-size:9px;padding:3px 7px;border-radius:4px;background:var(--s2);color:var(--tx2);border:1px solid var(--bdr2);}
.sort-btns{display:flex;gap:4px;}
.sb{padding:3px 8px;border-radius:4px;font-family:var(--f);font-size:9px;border:1px solid var(--bdr2);color:var(--tx3);background:transparent;cursor:pointer;transition:all .15s;}
.sb:hover{border-color:var(--tx2);color:var(--tx2);}.sb.active{border-color:var(--g);color:var(--g);background:rgba(0,229,160,.08);}
.card{background:var(--s1);border:1px solid var(--bdr);border-radius:9px;overflow:hidden;}
.table-scroll{overflow-x:scroll;-webkit-overflow-scrolling:touch;width:100%;display:block;}
table{border-collapse:collapse;min-width:750px;width:750px;}
thead th{padding:7px 9px;text-align:right;font-size:9px;color:var(--tx3);text-transform:uppercase;letter-spacing:.08em;border-bottom:1px solid var(--bdr);background:var(--s2);}
thead th:first-child{text-align:left;}
tbody tr{border-bottom:1px solid #0d1314;transition:background .1s;}
tbody tr:hover{background:#0e1617;}
tbody tr.z{opacity:.35;}
td{padding:7px 9px;text-align:right;font-size:11px;}td:first-child{text-align:left;}
.tk{font-family:var(--h);font-weight:700;font-size:12px;color:#fff;}
.exd{display:inline-block;font-size:8px;padding:1px 4px;border-radius:3px;margin-left:4px;vertical-align:middle;}
.thu{background:rgba(0,229,160,.1);color:var(--g);border:1px solid rgba(0,229,160,.25);}
.wed{background:rgba(41,182,246,.1);color:var(--b);border:1px solid rgba(41,182,246,.25);}
.mon{background:rgba(255,183,77,.1);color:var(--am);border:1px solid rgba(255,183,77,.25);}
.tue{background:rgba(239,83,80,.1);color:#ff8a80;border:1px solid rgba(239,83,80,.25);}
.g{color:var(--g);}.r{color:var(--rd);}.dm{color:var(--tx3);}
.tft td{padding:7px 9px;border-top:1px solid var(--bdr2);background:var(--s2);font-size:10px;color:var(--tx2);}
.pb-section{border-top:1px solid var(--bdr2);}
.pb-header{display:flex;justify-content:space-between;align-items:center;padding:10px 16px;border-bottom:1px solid var(--bdr2);}
.pb-grid{display:flex;}.pb-col{flex:1;padding:10px 12px;border-right:1px solid var(--bdr2);}.pb-col:last-child{border-right:none;}
.pb-tk{font-family:var(--h);font-size:11px;font-weight:700;color:#fff;}
.pb-nums{font-size:9px;color:var(--tx3);margin:3px 0 5px;display:flex;justify-content:space-between;}
.pb-pct{font-family:var(--h);font-size:13px;font-weight:700;color:var(--g);}
.pb-track{height:4px;background:var(--s3);border-radius:2px;overflow:hidden;margin-top:5px;}
.pb-fill{height:100%;border-radius:2px;background:linear-gradient(90deg,var(--b),var(--g));}
.rp{display:flex;flex-direction:column;gap:12px;}
.chart-outer{padding:14px 16px 12px;}
.chart-inner{display:flex;gap:10px;}
.y-axis{display:flex;flex-direction:column;justify-content:space-between;align-items:flex-end;padding-bottom:18px;flex-shrink:0;}
.y-label{font-size:9px;color:var(--tx3);}
.chart-body{flex:1;display:flex;flex-direction:column;gap:4px;position:relative;}
.bars-area{position:relative;height:120px;display:flex;align-items:flex-end;gap:5px;border-left:1px solid var(--bdr2);border-bottom:1px solid var(--bdr2);}
.bars-area::before{content:'';position:absolute;left:0;right:0;top:0;height:1px;background:var(--bdr2);opacity:.4;}
.bars-area::after{content:'';position:absolute;left:0;right:0;top:50%;height:1px;background:var(--bdr2);opacity:.4;}
.bar-col{flex:1;height:100%;display:flex;align-items:flex-end;}
.bar{width:100%;border-radius:2px 2px 0 0;background:var(--g);opacity:.6;}
.chart-svg{position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;}
.x-labels{display:flex;gap:5px;}.x-lbl{flex:1;text-align:center;font-size:8px;color:var(--tx3);padding-top:3px;}
.chart-foot{display:flex;justify-content:space-between;margin-top:10px;padding-top:10px;border-top:1px solid var(--bdr);}
.cfl{font-size:9px;color:var(--tx3);margin-bottom:3px;text-transform:uppercase;letter-spacing:.08em;}
.cfv{font-family:var(--h);font-size:16px;font-weight:700;}
.trend-note{font-size:9px;color:var(--tx3);margin-top:5px;display:flex;align-items:center;gap:5px;}
.trend-swatch{width:20px;height:2px;background:var(--am);display:inline-block;border-radius:1px;}
.calnote{font-size:9px;color:var(--tx3);padding:7px 16px 0;font-style:italic;}
.cal-body{padding:10px 16px 12px;}
.cr{display:flex;align-items:center;gap:6px;padding:6px 0;border-bottom:1px solid var(--bdr);}.cr:last-child{border:none;}
.ctk{font-family:var(--h);font-weight:700;font-size:12px;color:#fff;width:46px;flex-shrink:0;}
.cd{font-size:10px;color:var(--tx2);flex:1;}.ca{font-size:11px;color:var(--g);}
.cs{font-size:8px;padding:2px 6px;border-radius:3px;margin-left:5px;flex-shrink:0;}
.paid{background:rgba(0,229,160,.1);color:var(--g);border:1px solid rgba(0,229,160,.25);}
.pend{background:rgba(255,183,77,.1);color:var(--am);border:1px solid rgba(255,183,77,.25);}
.caltot{display:flex;justify-content:space-between;padding:8px 16px;background:var(--s2);border-top:1px solid var(--bdr2);font-size:10px;color:var(--tx2);}
.caltot span{color:var(--g);font-family:var(--h);font-weight:700;}
</style>
</head>
<body>
<div class="hdr">
  <div>
    <div class="hdr-t">Dividend Portfolio</div>
    <div class="hdr-s"><span class="dot"></span>Auto-updated · Pay dates: Thu &amp; Fri</div>
  </div>
  <div style="display:flex;align-items:center;gap:8px;">
    <div class="tabs">
      <div class="tab on" onclick="switchTab('indiv',this)">INDIV</div>
      <div class="tab" onclick="switchTab('ira',this)">IRA</div>
    </div>
    <span class="ts">Last: Fri May 8</span>
  </div>
</div>

<div id="panel-indiv" class="panel on">
  <div class="kpis">
    <div class="kpi"><div class="kl">Portfolio Value</div><div class="kv">$55,032</div><div class="ks">Cost basis $69,372</div></div>
    <div class="kpi"><div class="kl">Total Return</div><div class="kv g">+$12,015</div><div class="ks">+17.3%</div></div>
    <div class="kpi kp"><div class="kl">Total Return — Closed</div><div class="kv r">-$634</div><div class="ks">-0.4%</div></div>
    <div class="kpi kb"><div class="kl">Forecast / Week</div><div class="kv b">$${indivFcstWk}</div><div class="ks">2026 YTD avg × shares</div></div>
    <div class="kpi ka"><div class="kl">Forecast / Month</div><div class="kv am">$${indivFcstMo}</div><div class="ks">2026 YTD avg × shares</div></div>
  </div>
  <div class="g2">
    <div class="card">
      <div class="ch"><span class="ct">Positions</span><span class="bdg">5 active · 4 watchlist</span></div>
      <div class="table-scroll">
      <table>
        <thead><tr><th>Ticker</th><th>Shares</th><th>Price</th><th>Curr Val</th><th>P/L</th><th>Dividends</th><th>P/L+Div</th><th>Tot Ret%</th><th>Payback%</th><th>YTD Avg/Wk</th><th>Fcst/Wk</th><th>Div Yield%</th></tr></thead>
        <tbody>
          <tr><td><span class="tk">BABO</span><span class="exd thu">THU</span></td><td>1,000</td><td>$10.59</td><td>$10,590</td><td class="r">-$6,162</td><td class="g">$6,149</td><td class="r">-$13</td><td class="r">-0.1%</td><td class="g">36.7%</td><td>$0.1013</td><td class="g">$101.30</td><td>49.7%</td></tr>
          <tr><td><span class="tk">CHPY</span><span class="exd wed">WED</span></td><td>200</td><td>$74.90</td><td>$14,980</td><td class="g">+$4,026</td><td class="g">$3,447</td><td class="g">+$7,473</td><td class="g">+68.2%</td><td class="g">31.5%</td><td>$0.4972</td><td class="g">$99.44</td><td>34.5%</td></tr>
          <tr><td><span class="tk">LFGY</span><span class="exd wed">WED</span></td><td>200</td><td>$24.36</td><td>$4,872</td><td class="r">-$3,246</td><td class="g">$3,057</td><td class="r">-$189</td><td class="r">-2.3%</td><td class="g">37.7%</td><td>$0.2416</td><td class="g">$48.32</td><td>51.6%</td></tr>
          <tr><td><span class="tk">NVDY</span><span class="exd thu">THU</span></td><td>1,500</td><td>$14.08</td><td>$21,120</td><td class="r">-$4,708</td><td class="g">$10,683</td><td class="g">+$5,975</td><td class="g">+23.1%</td><td class="g">41.4%</td><td>$0.1185</td><td class="g">$177.75</td><td>43.8%</td></tr>
          <tr><td><span class="tk">PLTY</span><span class="exd thu">THU</span></td><td>100</td><td>$34.70</td><td>$3,470</td><td class="r">-$4,249</td><td class="g">$3,020</td><td class="r">-$1,229</td><td class="r">-15.9%</td><td class="g">39.1%</td><td>$0.4463</td><td class="g">$44.63</td><td>66.9%</td></tr>
          <tr class="z"><td><span class="tk">GOOW</span><span class="exd mon">MON</span></td><td class="dm">—</td><td>$82.35</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td>$0.5971</td><td class="dm">—</td><td>37.7%</td></tr>
          <tr class="z"><td><span class="tk">HOOW</span><span class="exd mon">MON</span></td><td class="dm">—</td><td>$23.46</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td>$0.9722</td><td class="dm">—</td><td>215.5%</td></tr>
          <tr class="z"><td><span class="tk">PLTW</span><span class="exd mon">MON</span></td><td class="dm">—</td><td>$22.24</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td>$0.5584</td><td class="dm">—</td><td>130.6%</td></tr>
          <tr class="z"><td><span class="tk">WPAY</span><span class="exd tue">TUE</span></td><td class="dm">—</td><td>$39.31</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td class="dm">—</td><td>$0.4730</td><td class="dm">—</td><td>62.6%</td></tr>
        </tbody>
        <tbody class="tft"><tr><td>Total</td><td colspan="2"></td><td>$55,032</td><td class="r">-$14,340</td><td class="g">$26,355</td><td class="g">+$12,015</td><td class="g">+17.3%</td><td class="g">38.0%</td><td></td><td class="g">$471.44</td><td></td></tr></tbody>
      </table>
      </div>
      <div class="pb-section">
        <div class="pb-header"><span class="ct" style="font-size:11px;">Payback Progress</span>
          <div class="sort-btns">
            <button class="sb" onclick="sortPB('indiv','alpha',this)">A→Z</button>
            <button class="sb active" onclick="sortPB('indiv','pct-desc',this)">% High→Low</button>
            <button class="sb" onclick="sortPB('indiv','pct-asc',this)">% Low→High</button>
          </div>
        </div>
        <div class="pb-grid" id="pb-indiv"></div>
      </div>
    </div>
    <div class="rp">
      <div class="card">
        <div class="ch"><span class="ct">Weekly Div Received</span><span class="bdg">Actual · 8 weeks</span></div>
        <div class="chart-outer">
          <div class="chart-inner">
            <div class="y-axis"><span class="y-label">$614</span><span class="y-label">$460</span><span class="y-label">$307</span><span class="y-label">$153</span><span class="y-label">$0</span></div>
            <div class="chart-body">
              <div class="bars-area">
                <div class="bar-col"><div class="bar" style="height:74.4%"></div></div>
                <div class="bar-col"><div class="bar" style="height:78.5%"></div></div>
                <div class="bar-col"><div class="bar" style="height:73.5%"></div></div>
                <div class="bar-col"><div class="bar" style="height:69.4%"></div></div>
                <div class="bar-col"><div class="bar" style="height:70.9%"></div></div>
                <div class="bar-col"><div class="bar" style="height:74.1%"></div></div>
                <div class="bar-col"><div class="bar" style="height:78.7%"></div></div>
                <div class="bar-col"><div class="bar" style="height:99.8%"></div></div>
                <svg class="chart-svg" id="trend-svg-indiv"></svg>
              </div>
              <div class="x-labels"><div class="x-lbl">Mar17</div><div class="x-lbl">Mar24</div><div class="x-lbl">Mar31</div><div class="x-lbl">Apr7</div><div class="x-lbl">Apr14</div><div class="x-lbl">Apr21</div><div class="x-lbl">Apr28</div><div class="x-lbl">May5</div></div>
            </div>
          </div>
          <div class="trend-note"><span class="trend-swatch"></span> 8-week average</div>
          <div class="chart-foot">
            <div><div class="cfl">This week</div><div class="cfv g">$482.37</div></div>
            <div style="text-align:right"><div class="cfl">8-week avg</div><div class="cfv" style="color:#fff">$475.62</div></div>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="ch"><span class="ct">Div Forecast</span>
          <div class="sort-btns">
            <button class="sb" onclick="sortCal('indiv','alpha',this)">A→Z</button>
            <button class="sb active" onclick="sortCal('indiv','date',this)">By Date</button>
          </div>
        </div>
        <div class="calnote">Dates = day cash deposits to account</div>
        <div class="cal-body" id="cal-indiv"></div>
        <div class="caltot">Week total <span>$482.37</span></div>
      </div>
    </div>
  </div>
</div>

<div id="panel-ira" class="panel">
  <div class="kpis">
    <div class="kpi"><div class="kl">Portfolio Value</div><div class="kv">$11,358</div><div class="ks">Cost basis $14,952</div></div>
    <div class="kpi"><div class="kl">Total Return</div><div class="kv g">+$3,501</div><div class="ks">+23.4%</div></div>
    <div class="kpi kp"><div class="kl">Total Return — Closed</div><div class="kv g">+$1,559</div><div class="ks">+4.1%</div></div>
    <div class="kpi kb"><div class="kl">Forecast / Week</div><div class="kv b">$${iraFcstWk}</div><div class="ks">2026 YTD avg × shares</div></div>
    <div class="kpi ka"><div class="kl">Forecast / Month</div><div class="kv am">$${iraFcstMo}</div><div class="ks">2026 YTD avg × shares</div></div>
  </div>
  <div class="g2">
    <div class="card">
      <div class="ch"><span class="ct">Positions</span><span class="bdg">3 active</span></div>
      <div class="table-scroll">
      <table>
        <thead><tr><th>Ticker</th><th>Shares</th><th>Price</th><th>Curr Val</th><th>P/L</th><th>Dividends</th><th>P/L+Div</th><th>Tot Ret%</th><th>Payback%</th><th>YTD Avg/Wk</th><th>Fcst/Wk</th><th>Div Yield%</th></tr></thead>
        <tbody>
          <tr><td><span class="tk">APLY</span><span class="exd thu">THU</span></td><td>450</td><td>$12.53</td><td>$5,638</td><td class="r">-$212</td><td class="g">$1,670</td><td class="g">+$1,458</td><td class="g">+24.9%</td><td class="g">28.5%</td><td>$0.0703</td><td class="g">$31.64</td><td>29.2%</td></tr>
          <tr><td><span class="tk">CONY</span><span class="exd thu">THU</span></td><td>55</td><td>$27.19</td><td>$1,495</td><td class="r">-$2,853</td><td class="g">$3,032</td><td class="g">+$180</td><td class="g">+4.1%</td><td class="g">69.7%</td><td>$0.3952</td><td class="g">$21.74</td><td>75.6%</td></tr>
          <tr><td><span class="tk">NVDY</span><span class="exd thu">THU</span></td><td>300</td><td>$14.08</td><td>$4,224</td><td class="r">-$529</td><td class="g">$2,393</td><td class="g">+$1,864</td><td class="g">+39.2%</td><td class="g">50.4%</td><td>$0.1185</td><td class="g">$35.55</td><td>43.8%</td></tr>
        </tbody>
        <tbody class="tft"><tr><td>IRA total</td><td colspan="2"></td><td>$11,358</td><td class="r">-$3,594</td><td class="g">$7,096</td><td class="g">+$3,501</td><td class="g">+23.4%</td><td class="g">47.5%</td><td></td><td class="g">$88.92</td><td></td></tr></tbody>
      </table>
      </div>
      <div class="pb-section">
        <div class="pb-header"><span class="ct" style="font-size:11px;">Payback Progress</span>
          <div class="sort-btns">
            <button class="sb active" onclick="sortPB('ira','alpha',this)">A→Z</button>
            <button class="sb" onclick="sortPB('ira','pct-desc',this)">% High→Low</button>
            <button class="sb" onclick="sortPB('ira','pct-asc',this)">% Low→High</button>
          </div>
        </div>
        <div class="pb-grid" id="pb-ira"></div>
      </div>
    </div>
    <div class="rp">
      <div class="card">
        <div class="ch"><span class="ct">Weekly Div Received</span><span class="bdg">Actual · 8 weeks</span></div>
        <div class="chart-outer">
          <div class="chart-inner">
            <div class="y-axis"><span class="y-label">$120</span><span class="y-label">$90</span><span class="y-label">$60</span><span class="y-label">$30</span><span class="y-label">$0</span></div>
            <div class="chart-body">
              <div class="bars-area">
                <div class="bar-col"><div class="bar" style="height:77%"></div></div>
                <div class="bar-col"><div class="bar" style="height:81%"></div></div>
                <div class="bar-col"><div class="bar" style="height:68%"></div></div>
                <div class="bar-col"><div class="bar" style="height:69%"></div></div>
                <div class="bar-col"><div class="bar" style="height:72%"></div></div>
                <div class="bar-col"><div class="bar" style="height:80%"></div></div>
                <div class="bar-col"><div class="bar" style="height:90%"></div></div>
                <div class="bar-col"><div class="bar" style="height:100%"></div></div>
                <svg class="chart-svg" id="trend-svg-ira"></svg>
              </div>
              <div class="x-labels"><div class="x-lbl">Mar17</div><div class="x-lbl">Mar24</div><div class="x-lbl">Mar31</div><div class="x-lbl">Apr7</div><div class="x-lbl">Apr14</div><div class="x-lbl">Apr21</div><div class="x-lbl">Apr28</div><div class="x-lbl">May5</div></div>
            </div>
          </div>
          <div class="trend-note"><span class="trend-swatch"></span> 8-week average</div>
          <div class="chart-foot">
            <div><div class="cfl">This week</div><div class="cfv g">$109.47</div></div>
            <div style="text-align:right"><div class="cfl">8-week avg</div><div class="cfv" style="color:#fff">$97.44</div></div>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="ch"><span class="ct">Div Forecast</span>
          <div class="sort-btns">
            <button class="sb active" onclick="sortCal('ira','alpha',this)">A→Z</button>
            <button class="sb" onclick="sortCal('ira','date',this)">By Date</button>
          </div>
        </div>
        <div class="calnote">Dates = day cash deposits to account</div>
        <div class="cal-body" id="cal-ira"></div>
        <div class="caltot">Week total <span>$109.47</span></div>
      </div>
    </div>
  </div>
</div>

<script>
const pbData={indiv:[{tk:'NVDY',recv:10683,cost:25828,pct:41.4},{tk:'PLTY',recv:3020,cost:7719,pct:39.1},{tk:'LFGY',recv:3057,cost:8118,pct:37.7},{tk:'BABO',recv:6149,cost:16752,pct:36.7},{tk:'CHPY',recv:3447,cost:10955,pct:31.5}],ira:[{tk:'APLY',recv:1670,cost:5851,pct:28.5},{tk:'CONY',recv:3032,cost:4348,pct:69.7},{tk:'NVDY',recv:2393,cost:4753,pct:50.4}]};
// ── DYNAMIC CALENDAR DATA ───────────────────────────────────
function getUpcomingDates() {
  const now = new Date();
  const day = now.getDay(); // 0=Sun,1=Mon,2=Tue,3=Wed,4=Thu,5=Fri,6=Sat
  // Find the upcoming Wednesday (Group1 pay day = Thu)
  const daysToWed = day <= 3 ? 3 - day : 10 - day;
  const daysToThu = day <= 4 ? 4 - day : 11 - day;
  const daysToFri = day <= 5 ? 5 - day : 12 - day;
  const wed = new Date(now); wed.setDate(now.getDate() + daysToWed);
  const thu = new Date(now); thu.setDate(now.getDate() + daysToThu);
  const fri = new Date(now); fri.setDate(now.getDate() + daysToFri);
  const fmt = d => d.toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric'});
  const isPaid = d => d <= now;
  return { wed, thu, fri, fmt, isPaid };
}

const { wed, thu, fri, fmt, isPaid } = getUpcomingDates();

const calData = {
  indiv: [
    {tk:'CHPY', date:fmt(thu), paid:isPaid(thu), pps:'$0.6024/sh', amt:'+$120.48'},
    {tk:'LFGY', date:fmt(thu), paid:isPaid(thu), pps:'$0.2513/sh', amt:'+$50.26'},
    {tk:'BABO', date:fmt(fri), paid:isPaid(fri), pps:'$0.0924/sh', amt:'+$92.40'},
    {tk:'NVDY', date:fmt(fri), paid:isPaid(fri), pps:'$0.1281/sh', amt:'+$192.15'},
    {tk:'PLTY', date:fmt(fri), paid:isPaid(fri), pps:'$0.2708/sh', amt:'+$27.08'},
  ],
  ira: [
    {tk:'APLY', date:fmt(fri), paid:isPaid(fri), pps:'$0.1033/sh', amt:'+$46.49'},
    {tk:'CONY', date:fmt(fri), paid:isPaid(fri), pps:'$0.4464/sh', amt:'+$24.55'},
    {tk:'NVDY', date:fmt(fri), paid:isPaid(fri), pps:'$0.1281/sh', amt:'+$38.43'},
  ]
};
function renderPB(a,items){document.getElementById('pb-'+a).innerHTML=items.map(d=>'<div class="pb-col"><div style="display:flex;justify-content:space-between;align-items:baseline;"><span class="pb-tk">'+d.tk+'</span><span class="pb-pct">'+d.pct+'%</span></div><div class="pb-nums"><span>$'+d.recv.toLocaleString()+'</span><span>of $'+d.cost.toLocaleString()+'</span></div><div class="pb-track"><div class="pb-fill" style="width:'+d.pct+'%"></div></div></div>').join('');}
function sortPB(a,mode,btn){btn.closest('.sort-btns').querySelectorAll('.sb').forEach(b=>b.classList.remove('active'));btn.classList.add('active');const items=[...pbData[a]];if(mode==='pct-desc')items.sort((a,b)=>b.pct-a.pct);else if(mode==='pct-asc')items.sort((a,b)=>a.pct-b.pct);else items.sort((a,b)=>a.tk.localeCompare(b.tk));renderPB(a,items);}
function renderCal(a,items){document.getElementById('cal-'+a).innerHTML=items.map(d=>'<div class="cr"><span class="ctk">'+d.tk+'</span><span class="cd">'+d.date+' · '+d.pps+'</span><span class="ca">'+d.amt+'</span><span class="cs '+(d.paid?'paid':'pend')+'">'+(d.paid?'PAID':'PENDING')+'</span></div>').join('');}
function sortCal(a,mode,btn){btn.closest('.sort-btns').querySelectorAll('.sb').forEach(b=>b.classList.remove('active'));btn.classList.add('active');const items=[...calData[a]];if(mode==='date')items.sort((a,b)=>a.ds-b.ds||a.tk.localeCompare(b.tk));else items.sort((a,b)=>a.tk.localeCompare(b.tk));renderCal(a,items);}
function drawTrend(id,pct){const s=document.getElementById(id);if(!s)return;const y=(1-pct/100)*100;s.innerHTML='<line x1="0" y1="'+y+'%" x2="100%" y2="'+y+'%" stroke="#ffb74d" stroke-width="1.5" stroke-dasharray="4,3" opacity="0.75"/>';}
function switchTab(n,el){document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));document.getElementById('panel-'+n).classList.add('on');el.classList.add('on');}
renderPB('indiv',[...pbData.indiv].sort((a,b)=>b.pct-a.pct));
renderPB('ira',[...pbData.ira].sort((a,b)=>a.tk.localeCompare(b.tk)));
renderCal('indiv',[...calData.indiv].sort((a,b)=>a.ds-b.ds||a.tk.localeCompare(b.tk)));
renderCal('ira',[...calData.ira].sort((a,b)=>a.tk.localeCompare(b.tk)));
window.addEventListener('load',()=>{drawTrend('trend-svg-indiv',75.5);drawTrend('trend-svg-ira',81.2);});
</script>
</body>
</html>`;
}
