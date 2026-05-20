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
  const data = getPortfolioData(ss, avgs);
  sendEmail(getDashboardHtml(data, avgs));
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
  // Wait for GOOGLEFINANCE to resolve
  SpreadsheetApp.flush();
  Utilities.sleep(3000);
}

// ─── READ LIVE PORTFOLIO DATA FROM SHEET ────────────────────
// Sheet columns: A=Ticker, B=Shares, C=Price, D=CostBasis, E=Dividends
// INDIV rows: BABO,CHPY,LFGY,NVDY,PLTY (A-Z)
// IRA rows: APLY,CONY,NVDY (after blank row)
// Static data that never changes:
const STATIC = {
  INDIV: [
    {tk:'BABO', shares:1000, cost:16751.52, exDay:'THU', ytdAvgKey:'BABO'},
    {tk:'CHPY', shares:200,  cost:10954.50, exDay:'WED', ytdAvgKey:'CHPY'},
    {tk:'LFGY', shares:200,  cost:8118.00,  exDay:'WED', ytdAvgKey:'LFGY'},
    {tk:'NVDY', shares:1500, cost:25828.42, exDay:'THU', ytdAvgKey:'NVDY'},
    {tk:'PLTY', shares:100,  cost:7719.10,  exDay:'THU', ytdAvgKey:'PLTY'},
  ],
  WATCHLIST: [
    {tk:'GOOW', exDay:'MON'}, {tk:'HOOW', exDay:'MON'},
    {tk:'PLTW', exDay:'MON'}, {tk:'WPAY', exDay:'TUE'},
  ],
  IRA: [
    {tk:'APLY', shares:450, cost:5851.00,  exDay:'THU', ytdAvgKey:'APLY'},
    {tk:'CONY', shares:55,  cost:4348.00,  exDay:'THU', ytdAvgKey:'CONY'},
    {tk:'NVDY', shares:300, cost:4753.09,  exDay:'THU', ytdAvgKey:'NVDY'},
  ],
  INDIV_CLOSED_NET: -634,
  IRA_CLOSED_NET:   1559,
};

function getPortfolioData(ss, avgs) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
  const rows  = sheet.getDataRange().getValues();

  // Build a map of ticker → {price, dividends} from sheet
  const sheetData = {};
  for (let i = 1; i < rows.length; i++) {
    const tk = rows[i][0];
    if (!tk) continue;
    sheetData[tk] = sheetData[tk] || [];
    sheetData[tk].push({
      shares: rows[i][1],
      price:  parseFloat(rows[i][2]) || 0,
      cost:   parseFloat(rows[i][3]) || 0,
      divs:   parseFloat(rows[i][4]) || 0,
    });
  }

  function calcPos(p, avgs) {
    const price  = sheetData[p.tk] ? sheetData[p.tk].find(r=>r.shares===p.shares)?.price || 0 : 0;
    const divs   = sheetData[p.tk] ? sheetData[p.tk].find(r=>r.shares===p.shares)?.divs  || 0 : 0;
    const val    = p.shares * price;
    const pl     = val - p.cost;
    const net    = pl + divs;
    const retPct = (net / p.cost) * 100;
    const pbPct  = (divs / p.cost) * 100;
    const avg    = avgs[p.ytdAvgKey] || 0;
    const fcstWk = avg * p.shares;
    const yield_ = price > 0 ? (avg * 52 / price) * 100 : 0;
    return { ...p, price, divs, val, pl, net, retPct, pbPct, avg, fcstWk, yield_ };
  }

  const a = avgs || {};
  const indiv = STATIC.INDIV.map(p => calcPos(p, a));
  const ira   = STATIC.IRA.map(p => calcPos(p, a));

  // Watchlist prices
  const watchlist = STATIC.WATCHLIST.map(p => ({
    ...p,
    price: sheetData[p.tk] ? (sheetData[p.tk][0]?.price || 0) : 0,
    avg: a[p.tk] || 0,
  }));

  // Totals
  const sum = arr => arr.reduce((s,p) => ({
    val:  s.val  + p.val,
    cost: s.cost + p.cost,
    divs: s.divs + p.divs,
    net:  s.net  + p.net,
    fcstWk: s.fcstWk + p.fcstWk,
  }), {val:0,cost:0,divs:0,net:0,fcstWk:0});

  const indivTot = sum(indiv);
  const iraTot   = sum(ira);

  return { indiv, ira, watchlist, indivTot, iraTot };
}

// ─── FORMATTING HELPERS ──────────────────────────────────────
function fmt$(n)  { return (n<0?'-$':'$') + Math.abs(Math.round(n)).toLocaleString(); }
function fmtP(n)  { return (n>=0?'+':'')+n.toFixed(1)+'%'; }
function fmtD(n)  { return '$'+n.toFixed(4); }
function fmtWk(n) { return '$'+Math.round(n).toLocaleString(); }

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

// ─── MAIN EMAIL HTML — reads live from sheet ─────────────────
function getDashboardHtml(data, avgs) {
  const today = Utilities.formatDate(new Date(),'America/Los_Angeles','MMM d, yyyy');
  const d = data;
  const it = d.indivTot, rt = d.iraTot;

  // Helper: position rows from live data
  function liveRows(positions) {
    return positions.map((p,i) => posRow(
      p.tk,
      p.shares.toLocaleString(),
      '$'+p.price.toFixed(2),
      fmt$(p.val),
      fmt$(p.pl),
      fmt$(p.divs),
      fmtP(p.retPct),
      p.pbPct.toFixed(1)+'%',
      p.exDay,
      i%2===0?'#2a2a2a':'#252525'
    )).join('');
  }

  // Payback bars sorted by pct desc
  function livePbBars(positions, padTo) {
    const sorted = [...positions].sort((a,b)=>b.pbPct-a.pbPct);
    let rows = sorted.map(p=>pbBar(p.tk, parseFloat(p.pbPct.toFixed(1)), Math.round(p.divs), Math.round(p.cost))).join('');
    if (padTo && sorted.length < padTo) rows += `<td width="${(padTo-sorted.length)*20}%" style="padding:4px;"></td>`;
    return rows;
  }

  const indivFcstWk = Math.round(it.fcstWk);
  const iraFcstWk   = Math.round(rt.fcstWk);

  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="background-color:#111;margin:0;padding:20px;">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:900px;margin:0 auto;">
  <tr><td colspan="2" style="padding-bottom:20px;text-align:center;">
    <div style="font-family:Arial,sans-serif;font-size:26px;font-weight:bold;color:#fff;">Dividend Portfolio</div>
    <div style="font-family:Arial,sans-serif;font-size:12px;color:#888;margin-top:4px;">Updated ${today}</div>
  </td></tr>

  ${sectionHeader('INDIV')}
  <tr><td colspan="2" style="padding-bottom:12px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    ${kpi('Portfolio Value',fmt$(it.val),'Cost: '+fmt$(it.cost),'')}
    ${kpi('Total Return',(it.net>=0?'+':'')+fmt$(it.net),fmtP(it.net/it.cost*100),it.net>=0?'#4ade80':'#f87171')}
    ${kpi('Closed Positions',fmt$(STATIC.INDIV_CLOSED_NET),fmtP(STATIC.INDIV_CLOSED_NET/it.cost*100),STATIC.INDIV_CLOSED_NET>=0?'#4ade80':'#f87171')}
    ${kpi('Forecast / Week',fmtWk(it.fcstWk),'2026 YTD avg','#64b5f6')}
    ${kpi('Forecast / Month',fmtWk(it.fcstWk*4),'2026 YTD avg','#ffd54f')}
  </tr></table></td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Positions</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;overflow:hidden;">
      ${tableHeader([{label:'Ticker',left:true},{label:'Shares'},{label:'Price'},{label:'Value'},{label:'P/L'},{label:'Dividends'},{label:'Total Ret%'},{label:'Payback%'}])}
      ${liveRows(d.indiv)}
      <tr style="background-color:#1a1a1a;">
        <td colspan="3" style="padding:9px 10px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;font-weight:bold;border-top:2px solid #444;">TOTAL</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#fff;font-weight:bold;text-align:right;border-top:2px solid #444;">${fmt$(it.val)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:${it.pl>=0?'#4ade80':'#f87171'};font-weight:bold;text-align:right;border-top:2px solid #444;">${fmt$(it.pl)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">${fmt$(it.divs)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:${it.net>=0?'#4ade80':'#f87171'};font-weight:bold;text-align:right;border-top:2px solid #444;">${fmtP(it.net/it.cost*100)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">${(it.divs/it.cost*100).toFixed(1)}%</td>
      </tr>
    </table>
  </td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Payback Progress</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>${livePbBars(d.indiv,5)}</tr></table></td></tr>

  ${sectionHeader('IRA')}
  <tr><td colspan="2" style="padding-bottom:12px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    ${kpi('Portfolio Value',fmt$(rt.val),'Cost: '+fmt$(rt.cost),'')}
    ${kpi('Total Return',(rt.net>=0?'+':'')+fmt$(rt.net),fmtP(rt.net/rt.cost*100),rt.net>=0?'#4ade80':'#f87171')}
    ${kpi('Closed Positions',fmt$(STATIC.IRA_CLOSED_NET),fmtP(STATIC.IRA_CLOSED_NET/rt.cost*100),STATIC.IRA_CLOSED_NET>=0?'#4ade80':'#f87171')}
    ${kpi('Forecast / Week',fmtWk(rt.fcstWk),'2026 YTD avg','#64b5f6')}
    ${kpi('Forecast / Month',fmtWk(rt.fcstWk*4),'2026 YTD avg','#ffd54f')}
  </tr></table></td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Positions</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;overflow:hidden;">
      ${tableHeader([{label:'Ticker',left:true},{label:'Shares'},{label:'Price'},{label:'Value'},{label:'P/L'},{label:'Dividends'},{label:'Total Ret%'},{label:'Payback%'}])}
      ${liveRows(d.ira)}
      <tr style="background-color:#1a1a1a;">
        <td colspan="3" style="padding:9px 10px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;font-weight:bold;border-top:2px solid #444;">TOTAL</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#fff;font-weight:bold;text-align:right;border-top:2px solid #444;">${fmt$(rt.val)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:${rt.pl>=0?'#4ade80':'#f87171'};font-weight:bold;text-align:right;border-top:2px solid #444;">${fmt$(rt.pl)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">${fmt$(rt.divs)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:${rt.net>=0?'#4ade80':'#f87171'};font-weight:bold;text-align:right;border-top:2px solid #444;">${fmtP(rt.net/rt.cost*100)}</td>
        <td style="padding:9px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">${(rt.divs/rt.cost*100).toFixed(1)}%</td>
      </tr>
    </table>
  </td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Payback Progress</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>${livePbBars(d.ira,5)}</tr></table></td></tr>

  <!-- FOOTER -->
  <tr><td colspan="2" style="text-align:center;padding-top:20px;border-top:1px solid #333;">
    <a href="https://script.google.com/macros/s/AKfycbzHsDpaNec2_fpgthmbHudzFYFnvxw6DvjJzdwMZx8PC6sW-Mk7zwrTgluMnwYX068S9A/exec" style="display:inline-block;background-color:#4ade80;color:#000;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;padding:12px 28px;border-radius:6px;text-decoration:none;">View Full Dashboard →</a>
    <div style="font-family:Arial,sans-serif;font-size:11px;color:#555;margin-top:12px;">Auto-generated by Dividend Tracker</div>
  </td></tr>
</table></body></html>`;
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
  const data = getPortfolioData(ss, avgs);
  const html = HtmlService.createHtmlOutput(getInteractiveDashboard(data, avgs));
  html.setTitle('Dividend Portfolio');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getInteractiveDashboard(data, avgs) {
  const d  = data || {};
  const it = d.indivTot || {val:55032,cost:69372,divs:26355,net:12015,fcstWk:471};
  const rt = d.iraTot   || {val:11358,cost:14952,divs:7096, net:3501, fcstWk:89};

  // Build position data for JS injection
  function posJson(positions) {
    return JSON.stringify((positions||[]).map(p=>({
      tk:p.tk, shares:p.shares, price:p.price, val:p.val, pl:p.pl,
      divs:p.divs, net:p.net, retPct:p.retPct, pbPct:p.pbPct,
      avg:p.avg, fcstWk:p.fcstWk, yield_:p.yield_, exDay:p.exDay,
      cost:p.cost
    })));
  }

  const indivJson    = posJson(d.indiv);
  const iraJson      = posJson(d.ira);
  const watchJson    = JSON.stringify((d.watchlist||[]).map(w=>({tk:w.tk,price:w.price,avg:w.avg,exDay:w.exDay})));
  const indivClosedNet = STATIC.INDIV_CLOSED_NET;
  const iraClosedNet   = STATIC.IRA_CLOSED_NET;

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
.ks{font-size:10px;color:var(--tx3);margin-top:4px;}
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
    <div class="hdr-s"><span class="dot"></span>Live data</div>
  </div>
  <div style="display:flex;align-items:center;gap:8px;">
    <div class="tabs">
      <div class="tab on" onclick="switchTab('indiv',this)">INDIV</div>
      <div class="tab" onclick="switchTab('ira',this)">IRA</div>
    </div>
    <span class="ts" id="lastUpdate"></span>
  </div>
</div>

<div id="panel-indiv" class="panel on"></div>
<div id="panel-ira" class="panel"></div>

<script>
// ── LIVE DATA INJECTED FROM SHEET ───────────────────────────
const LIVE = {
  indiv:         ${indivJson},
  ira:           ${iraJson},
  watchlist:     ${watchJson},
  indivTot:      ${JSON.stringify(it)},
  iraTot:        ${JSON.stringify(rt)},
  indivClosedNet:${indivClosedNet},
  iraClosedNet:  ${iraClosedNet},
};

// ── HELPERS ─────────────────────────────────────────────────
function f$(n)  { return (n<0?'-$':'$')+Math.abs(Math.round(n)).toLocaleString(); }
function fp(n)  { return (n>=0?'+':'')+n.toFixed(1)+'%'; }
function gc(n)  { return n>=0?'var(--g)':'var(--rd)'; }

// ── RENDER POSITIONS TABLE ───────────────────────────────────
function renderPositions(positions, watchlist) {
  const rows = positions.map((p,i) => {
    const exCls = p.exDay==='WED'?'wed':p.exDay==='MON'?'mon':p.exDay==='TUE'?'tue':'thu';
    return '<tr>'
      +'<td><span class="tk">'+p.tk+'</span><span class="exd '+exCls+'">'+p.exDay+'</span></td>'
      +'<td>'+p.shares.toLocaleString()+'</td>'
      +'<td>$'+p.price.toFixed(2)+'</td>'
      +'<td>'+f$(p.val)+'</td>'
      +'<td style="color:'+gc(p.pl)+'">'+f$(p.pl)+'</td>'
      +'<td class="g">'+f$(p.divs)+'</td>'
      +'<td style="color:'+gc(p.net)+'">'+f$(p.net)+'</td>'
      +'<td style="color:'+gc(p.retPct)+'">'+fp(p.retPct)+'</td>'
      +'<td class="g">'+p.pbPct.toFixed(1)+'%</td>'
      +'<td>$'+p.avg.toFixed(4)+'</td>'
      +'<td class="g">$'+Math.round(p.fcstWk)+'</td>'
      +'<td>'+p.yield_.toFixed(1)+'%</td>'
      +'</tr>';
  }).join('');
  const watchRows = (watchlist||[]).map(w => {
    const exCls = w.exDay==='WED'?'wed':w.exDay==='MON'?'mon':w.exDay==='TUE'?'tue':'thu';
    return '<tr class="z">'
      +'<td><span class="tk">'+w.tk+'</span><span class="exd '+exCls+'">'+w.exDay+'</span></td>'
      +'<td class="dm">—</td><td>$'+w.price.toFixed(2)+'</td>'
      +'<td class="dm">—</td><td class="dm">—</td><td class="dm">—</td>'
      +'<td class="dm">—</td><td class="dm">—</td><td class="dm">—</td>'
      +'<td>$'+w.avg.toFixed(4)+'</td><td class="dm">—</td><td></td>'
      +'</tr>';
  }).join('');
  return rows + watchRows;
}

function renderTotRow(tot) {
  return '<tr style="background:var(--s2);border-top:1px solid var(--bdr2);">'
    +'<td colspan="3" style="color:var(--tx2);font-size:9px;text-transform:uppercase;">Total</td>'
    +'<td>'+f$(tot.val)+'</td>'
    +'<td style="color:'+gc(tot.pl||0)+'">'+f$(tot.pl||0)+'</td>'
    +'<td class="g">'+f$(tot.divs)+'</td>'
    +'<td style="color:'+gc(tot.net)+'">'+f$(tot.net)+'</td>'
    +'<td style="color:'+gc(tot.net)+'">'+fp(tot.net/tot.cost*100)+'</td>'
    +'<td class="g">'+(tot.divs/tot.cost*100).toFixed(1)+'%</td>'
    +'<td></td><td class="g">$'+Math.round(tot.fcstWk)+'</td><td></td>'
    +'</tr>';
}

// ── RENDER KPIs ──────────────────────────────────────────────
function kpi(label, value, sub, accent) {
  return '<div class="kpi '+(accent||'')+'">'
    +'<div class="kl">'+label+'</div>'
    +'<div class="kv '+(accent?accent.replace('k',''):'')+'">'+value+'</div>'
    +'<div class="ks">'+sub+'</div>'
    +'</div>';
}

// ── RENDER PAYBACK ───────────────────────────────────────────
function renderPB(acct, items) {
  document.getElementById('pb-'+acct).innerHTML = items.map(p =>
    '<div class="pb-col">'
    +'<div style="display:flex;justify-content:space-between;align-items:baseline;">'
    +'<span class="pb-tk">'+p.tk+'</span><span class="pb-pct">'+p.pbPct.toFixed(1)+'%</span></div>'
    +'<div class="pb-nums"><span>'+f$(p.divs)+'</span><span>of '+f$(p.cost)+'</span></div>'
    +'<div class="pb-track"><div class="pb-fill" style="width:'+p.pbPct.toFixed(1)+'%"></div></div>'
    +'</div>'
  ).join('');
}
function sortPB(acct,mode,btn){
  btn.closest('.sort-btns').querySelectorAll('.sb').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const items=[...LIVE[acct]];
  if(mode==='pct-desc')items.sort((a,b)=>b.pbPct-a.pbPct);
  else if(mode==='pct-asc')items.sort((a,b)=>a.pbPct-b.pbPct);
  else items.sort((a,b)=>a.tk.localeCompare(b.tk));
  renderPB(acct,items);
}

// ── RENDER CALENDAR ──────────────────────────────────────────
function getUpcomingDates() {
  const now=new Date(), day=now.getDay();
  const dToWed=day<=3?3-day:10-day, dToThu=day<=4?4-day:11-day, dToFri=day<=5?5-day:12-day;
  const wed=new Date(now),thu=new Date(now),fri=new Date(now);
  wed.setDate(now.getDate()+dToWed);
  thu.setDate(now.getDate()+dToThu);
  fri.setDate(now.getDate()+dToFri);
  const fmt=d=>d.toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric'});
  const paid=d=>d<=now;
  return {wed,thu,fri,fmt,paid};
}

const calData = (() => {
  const {thu,fri,fmt,paid} = getUpcomingDates();
  const lastAvg = p => LIVE.indiv.find(x=>x.tk===p)||LIVE.ira.find(x=>x.tk===p)||{avg:0,shares:0};
  const row = (tk,d,acct) => {
    const p = (acct==='indiv'?LIVE.indiv:LIVE.ira).find(x=>x.tk===tk)||{avg:0,shares:0};
    return {tk, date:fmt(d), paid:paid(d), pps:'$'+p.avg.toFixed(4)+'/sh', amt:'+$'+(p.avg*p.shares).toFixed(2)};
  };
  return {
    indiv: [row('CHPY',thu,'indiv'),row('LFGY',thu,'indiv'),row('BABO',fri,'indiv'),row('NVDY',fri,'indiv'),row('PLTY',fri,'indiv')],
    ira:   [row('APLY',fri,'ira'),row('CONY',fri,'ira'),row('NVDY',fri,'ira')]
  };
})();

function renderCal(acct,items){
  document.getElementById('cal-'+acct).innerHTML=items.map(d=>
    '<div class="cr"><span class="ctk">'+d.tk+'</span><span class="cd">'+d.date+' · '+d.pps+'</span>'
    +'<span class="ca">'+d.amt+'</span>'
    +'<span class="cs '+(d.paid?'paid':'pend')+'">'+(d.paid?'PAID':'PENDING')+'</span></div>'
  ).join('');
}
function sortCal(acct,mode,btn){
  btn.closest('.sort-btns').querySelectorAll('.sb').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const items=[...calData[acct]];
  if(mode==='date')items.sort((a,b)=>new Date(a.date)-new Date(b.date)||a.tk.localeCompare(b.tk));
  else items.sort((a,b)=>a.tk.localeCompare(b.tk));
  renderCal(acct,items);
}

// ── DRAW TREND ───────────────────────────────────────────────
function drawTrend(id,pct){
  const s=document.getElementById(id);if(!s)return;
  const y=(1-pct/100)*100;
  s.innerHTML='<line x1="0" y1="'+y+'%" x2="100%" y2="'+y+'%" stroke="#ffb74d" stroke-width="1.5" stroke-dasharray="4,3" opacity="0.75"/>';
}

// ── BUILD PANELS ─────────────────────────────────────────────
function buildPanel(acct, positions, watchlist, tot, closedNet) {
  const tRows = renderPositions(positions, watchlist);
  const tTot  = renderTotRow(tot);
  const indivFcstWk = Math.round(tot.fcstWk);

  const panel = document.getElementById('panel-'+acct);
  panel.innerHTML =
    '<div class="kpis">'
    +kpi('Portfolio Value',f$(tot.val),'Cost '+f$(tot.cost),'')
    +kpi('Total Return',(tot.net>=0?'+':'')+f$(tot.net),fp(tot.net/tot.cost*100),'kb')
    +kpi('Total Return — Closed',(closedNet>=0?'+':'')+f$(closedNet),fp(closedNet/tot.cost*100),'kp')
    +kpi('Forecast / Week','$'+indivFcstWk,'2026 YTD avg','kb')
    +kpi('Forecast / Month','$'+Math.round(indivFcstWk*4),'2026 YTD avg','ka')
    +'</div>'
    +'<div class="g2">'
    +  '<div class="card">'
    +    '<div class="ch"><span class="ct">Positions</span><span class="bdg">'+positions.length+' active'+(watchlist?(' · '+watchlist.length+' watchlist'):'')+'</span></div>'
    +    '<div class="table-scroll"><table>'
    +      '<thead><tr><th style="text-align:left">Ticker</th><th>Shares</th><th>Price</th><th>Curr Val</th><th>P/L</th><th>Dividends</th><th>P/L+Div</th><th>Tot Ret%</th><th>Payback%</th><th>YTD Avg/Wk</th><th>Fcst/Wk</th><th>Div Yield%</th></tr></thead>'
    +      '<tbody>'+tRows+'</tbody>'
    +      '<tbody class="tft">'+tTot+'</tbody>'
    +    '</table></div>'
    +    '<div class="pb-section">'
    +      '<div class="pb-header"><span class="ct" style="font-size:11px;">Payback Progress</span>'
    +        '<div class="sort-btns">'
    +          '<button class="sb" onclick="sortPB(\''+acct+'\',\'alpha\',this)">A→Z</button>'
    +          '<button class="sb active" onclick="sortPB(\''+acct+'\',\'pct-desc\',this)">% High→Low</button>'
    +          '<button class="sb" onclick="sortPB(\''+acct+'\',\'pct-asc\',this)">% Low→High</button>'
    +        '</div>'
    +      '</div>'
    +      '<div class="pb-grid" id="pb-'+acct+'"></div>'
    +    '</div>'
    +  '</div>'
    +  '<div class="rp">'
    +    '<div class="card">'
    +      '<div class="ch"><span class="ct">Weekly Div Received</span><span class="bdg">Actual · 8 weeks</span></div>'
    +      '<div class="chart-outer"><div class="chart-inner">'
    +        '<div class="y-axis"><span class="y-label">High</span><span class="y-label"></span><span class="y-label">Avg</span><span class="y-label"></span><span class="y-label">$0</span></div>'
    +        '<div class="chart-body"><div class="bars-area" id="bars-'+acct+'">'
    +          '<div class="bar-col"><div class="bar" style="height:74%"></div></div>'
    +          '<div class="bar-col"><div class="bar" style="height:79%"></div></div>'
    +          '<div class="bar-col"><div class="bar" style="height:74%"></div></div>'
    +          '<div class="bar-col"><div class="bar" style="height:69%"></div></div>'
    +          '<div class="bar-col"><div class="bar" style="height:71%"></div></div>'
    +          '<div class="bar-col"><div class="bar" style="height:74%"></div></div>'
    +          '<div class="bar-col"><div class="bar" style="height:79%"></div></div>'
    +          '<div class="bar-col"><div class="bar" style="height:100%"></div></div>'
    +          '<svg class="chart-svg" id="trend-svg-'+acct+'"></svg>'
    +        '</div>'
    +        '<div class="x-labels"><div class="x-lbl">Mar17</div><div class="x-lbl">Mar24</div><div class="x-lbl">Mar31</div><div class="x-lbl">Apr7</div><div class="x-lbl">Apr14</div><div class="x-lbl">Apr21</div><div class="x-lbl">Apr28</div><div class="x-lbl">May5</div></div>'
    +        '</div></div>'
    +      '<div class="trend-note"><span class="trend-swatch"></span> 8-week average</div>'
    +      '<div class="chart-foot"><div><div class="cfl">8-wk avg</div><div class="cfv g">$'+indivFcstWk+'</div></div></div>'
    +      '</div>'
    +    '</div>'
    +    '<div class="card">'
    +      '<div class="ch"><span class="ct">Div Forecast</span>'
    +        '<div class="sort-btns">'
    +          '<button class="sb active" onclick="sortCal(\''+acct+'\',\'alpha\',this)">A→Z</button>'
    +          '<button class="sb" onclick="sortCal(\''+acct+'\',\'date\',this)">By Date</button>'
    +        '</div>'
    +      '</div>'
    +      '<div class="calnote">Dates = day cash deposits to account</div>'
    +      '<div class="cal-body" id="cal-'+acct+'"></div>'
    +      '<div class="caltot">Week total <span>$'+Math.round(tot.fcstWk)+'</span></div>'
    +    '</div>'
    +  '</div>'
    +'</div>';

  renderPB(acct, [...positions].sort((a,b)=>b.pbPct-a.pbPct));
  renderCal(acct, [...calData[acct]].sort((a,b)=>a.tk.localeCompare(b.tk)));
}

function switchTab(n,el){
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  document.getElementById('panel-'+n).classList.add('on');
  el.classList.add('on');
}

// ── INIT ─────────────────────────────────────────────────────
document.getElementById('lastUpdate').textContent = 'Updated '+new Date().toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
buildPanel('indiv', LIVE.indiv, LIVE.watchlist, LIVE.indivTot, LIVE.indivClosedNet);
buildPanel('ira',   LIVE.ira,   null,           LIVE.iraTot,   LIVE.iraClosedNet);
window.addEventListener('load', () => {
  drawTrend('trend-svg-indiv', 75);
  drawTrend('trend-svg-ira',   75);
});
</script>
</body>
</html>`;
}
