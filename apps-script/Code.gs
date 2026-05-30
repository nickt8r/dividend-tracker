// ============================================================
// DIVIDEND TRACKER — Google Apps Script
// Uses Yahoo Finance API for live prices + dividend history
// Runs automatically every Wednesday and Thursday morning
// ============================================================

const CONFIG = {
  SHEET_NAME: 'Dividend Tracker',
  EMAIL:      'nickt8r@gmail.com',
  DASHBOARD:  'https://script.google.com/macros/s/AKfycbzHsDpaNec2_fpgthmbHudzFYFnvxw6DvjJzdwMZx8PC6sW-Mk7zwrTgluMnwYX068S9A/exec',
  YTD_START:  '2026-01-01',
  SEED_CUTOFF:'2026-05-09', // dividends before this date are in seed data below
};

// ── STATIC POSITION DATA (never changes unless you buy/sell) ─
const POSITIONS = {
  INDIV: [
    {tk:'BABO', shares:1000, cost:16751.52, exDay:'THU'},
    {tk:'CHPY', shares:200,  cost:10954.50, exDay:'WED'},
    {tk:'LFGY', shares:200,  cost:8118.00,  exDay:'WED'},
    {tk:'NVDY', shares:1500, cost:25828.42, exDay:'THU'},
    {tk:'PLTY', shares:100,  cost:7719.10,  exDay:'THU'},
  ],
  WATCHLIST: [
    {tk:'GOOW', exDay:'MON'}, {tk:'HOOW', exDay:'MON'},
    {tk:'PLTW', exDay:'MON'}, {tk:'WPAY', exDay:'TUE'},
  ],
  IRA: [
    {tk:'APLY', shares:450, cost:5851.00,  exDay:'THU'},
    {tk:'CONY', shares:55,  cost:4348.00,  exDay:'THU'},
    {tk:'NVDY', shares:300, cost:4753.09,  exDay:'THU'},
  ],
  INDIV_CLOSED_LOSS: 12648,
  IRA_CLOSED_LOSS:   1943,
};

// ── SEED: known dividends paid per share Jan 1 – May 8 2026 ──
// Source: Digrin / YieldMax press releases
const SEED_PPS = {
  BABO: [0.0915,0.1166,0.0864,0.1277,0.1695,0.1190,0.1138,0.0965,0.0984,0.0908,0.0892,0.0928,0.0870,0.0866,0.0913,0.0906,0.0833,0.1012,0.0924],
  CHPY: [0.5040,0.5196,0.5267,0.4826,0.5253,0.5293,0.5259,0.4686,0.4577,0.4384,0.3851,0.4912,0.4089,0.4406,0.5384,0.5008,0.6041,0.6024],
  LFGY: [0.2885,0.2835,0.2822,0.2741,0.2353,0.2170,0.2294,0.2209,0.2356,0.2303,0.2369,0.2253,0.2033,0.2203,0.2356,0.2562,0.2228,0.2513],
  NVDY: [0.1435,0.1054,0.0950,0.0848,0.1076,0.0939,0.1057,0.0944,0.1151,0.1162,0.1197,0.1332,0.1195,0.1148,0.1111,0.1161,0.1401,0.2072,0.1281],
  PLTY: [0.5130,0.4508,0.4130,0.3791,0.3688,0.3591,0.3845,0.3865,0.3933,0.4781,0.7999,0.8018,0.4779,0.4497,0.4548,0.3556,0.3832,0.3607,0.2708],
  CONY: [0.4342,0.4091,0.3965,0.2219,0.3089,0.2838,0.2556,0.2994,0.3177,0.3115,0.5942,0.6138,0.5332,0.3763,0.3767,0.3833,0.4161,0.5307,0.4464],
  APLY: [0.0532,0.0481,0.0415,0.0473,0.0494,0.0584,0.1774,0.0498,0.0911,0.0649,0.0622,0.0471,0.0606,0.0617,0.0613,0.0717,0.0902,0.0958,0.1033],
};

// ── ENTRY POINT ──────────────────────────────────────────────
function runWeeklyUpdate() {
  Logger.log('=== Dividend Tracker Update Starting ===');
  try {
    const data = buildPortfolioData();
    Logger.log('Portfolio data built. INDIV positions: ' + data.indiv.length);
    Logger.log('INDIV total val: ' + data.indivTot.val);
    Logger.log('BABO price: ' + data.indiv[0].price + ' divs: ' + data.indiv[0].divs);
    const html = getDashboardHtml(data);
    Logger.log('HTML length: ' + html.length);
    if (html.length < 500) {
      Logger.log('ERROR: HTML too short, something went wrong');
      Logger.log('HTML content: ' + html);
      return;
    }
    sendEmail(html);
    Logger.log('=== Done ===');
  } catch(e) {
    Logger.log('FATAL ERROR: ' + e.message + '\n' + e.stack);
    // Send error notification email
    GmailApp.sendEmail(CONFIG.EMAIL, 'Dividend Tracker ERROR', 'Script failed: ' + e.message + '\n\n' + e.stack);
  }
}

// ── DIAGNOSTIC — run this first to check data ────────────────
function testDataFetch() {
  Logger.log('=== Testing Yahoo Finance fetch ===');
  try {
    const d = yahooFetch('BABO');
    const price = d.chart.result[0].meta.regularMarketPrice;
    const divs  = Object.values(d.chart.result[0].events?.dividends || {});
    Logger.log('BABO price: $' + price);
    Logger.log('BABO dividends found: ' + divs.length);
    Logger.log('Most recent: $' + divs[divs.length-1]?.amount + ' on ' + new Date(divs[divs.length-1]?.date*1000).toDateString());
  } catch(e) {
    Logger.log('Yahoo fetch FAILED: ' + e.message);
  }

  Logger.log('=== Testing buildPortfolioData ===');
  try {
    const data = buildPortfolioData();
    Logger.log('indiv positions: ' + data.indiv.length);
    data.indiv.forEach(p => Logger.log(p.tk + ': price=$' + p.price + ' divs=$' + p.divs.toFixed(2) + ' val=$' + p.val.toFixed(2)));
    Logger.log('indivTot val: $' + data.indivTot.val.toFixed(2));
  } catch(e) {
    Logger.log('buildPortfolioData FAILED: ' + e.message + '\n' + e.stack);
  }

  Logger.log('=== Testing getDashboardHtml ===');
  try {
    const data = buildPortfolioData();
    const html = getDashboardHtml(data);
    Logger.log('HTML length: ' + html.length + ' chars');
    Logger.log('First 200 chars: ' + html.substring(0, 200));
  } catch(e) {
    Logger.log('getDashboardHtml FAILED: ' + e.message + '\n' + e.stack);
  }
}

// ── YAHOO FINANCE FETCH ───────────────────────────────────────
function yahooFetch(ticker) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${ticker}?events=dividends&range=1y&interval=1d`;
  const resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
  });
  return JSON.parse(resp.getContentText());
}

// ── BUILD ALL PORTFOLIO DATA ──────────────────────────────────
function buildPortfolioData() {
  const cutoff  = new Date(CONFIG.SEED_CUTOFF);
  const ytdStart= new Date(CONFIG.YTD_START);
  const allTickers = [...new Set([
    ...POSITIONS.INDIV.map(p=>p.tk),
    ...POSITIONS.IRA.map(p=>p.tk),
    ...POSITIONS.WATCHLIST.map(p=>p.tk),
  ])];

  // Fetch Yahoo data for all tickers
  const yahoo = {};
  allTickers.forEach(tk => {
    try {
      const d    = yahooFetch(tk);
      const meta = d.chart.result[0].meta;
      const divs = d.chart.result[0].events?.dividends || {};
      yahoo[tk] = {
        price: parseFloat(meta.regularMarketPrice) || 0,
        divs:  Object.values(divs).map(v => ({date: new Date(v.date*1000), amount: v.amount})),
      };
      Logger.log(`${tk}: $${yahoo[tk].price} | ${yahoo[tk].divs.length} dividends`);
    } catch(e) {
      Logger.log(`Error fetching ${tk}: ${e.message}`);
      yahoo[tk] = {price:0, divs:[]};
    }
  });

  // Compute YTD avg per share (seed + new payments since cutoff)
  function getYtdAvg(tk) {
    const seed    = SEED_PPS[tk] || [];
    const newDivs = (yahoo[tk]?.divs || [])
      .filter(d => d.date >= cutoff)
      .map(d => d.amount);
    const all = [...seed, ...newDivs];
    return all.length ? all.reduce((s,v)=>s+v,0)/all.length : 0;
  }

  // Compute total dividends paid (seed dollars + new payments since cutoff)
  const SEED_DOLLARS = {
    'BABO_1000': 6148.59, 'CHPY_200': 3447.03, 'LFGY_200': 3056.52,
    'NVDY_1500': 10682.72,'PLTY_100': 3019.78,
    'APLY_450':  1670.13, 'CONY_55':  3032.07, 'NVDY_300': 2393.31,
  };

  function getTotalDivs(tk, shares) {
    const seed     = SEED_DOLLARS[`${tk}_${shares}`] || 0;
    const newPaid  = (yahoo[tk]?.divs || [])
      .filter(d => d.date >= cutoff)
      .reduce((s,d) => s + d.amount * shares, 0);
    return seed + newPaid;
  }

  // Calculate position metrics
  function calcPos(p) {
    const price   = yahoo[p.tk]?.price || 0;
    const divs    = getTotalDivs(p.tk, p.shares);
    const ytdAvg  = getYtdAvg(p.tk);
    const val     = p.shares * price;
    const pl      = val - p.cost;
    const net     = pl + divs;
    const retPct  = (net / p.cost) * 100;
    const pbPct   = (divs / p.cost) * 100;
    const fcstWk  = ytdAvg * p.shares;
    const yield_  = price > 0 ? (ytdAvg * 52 / price) * 100 : 0;
    return {...p, price, divs, ytdAvg, val, pl, net, retPct, pbPct, fcstWk, yield_};
  }

  const indiv    = POSITIONS.INDIV.map(calcPos);
  const ira      = POSITIONS.IRA.map(calcPos);
  const watchlist= POSITIONS.WATCHLIST.map(p => ({
    ...p, price: yahoo[p.tk]?.price||0, ytdAvg: getYtdAvg(p.tk),
  }));

  // Totals
  function sum(arr) {
    return arr.reduce((s,p)=>({
      val:s.val+p.val, cost:s.cost+p.cost, divs:s.divs+p.divs,
      pl:s.pl+p.pl, net:s.net+p.net, fcstWk:s.fcstWk+p.fcstWk,
    }),{val:0,cost:0,divs:0,pl:0,net:0,fcstWk:0});
  }

  // Most recent week's payments
  const oneWeekAgo = new Date(); oneWeekAgo.setDate(oneWeekAgo.getDate()-8);
  function recentPayments(positions) {
    const payments = [];
    positions.forEach(p => {
      const recent = (yahoo[p.tk]?.divs||[]).filter(d=>d.date>=oneWeekAgo);
      recent.forEach(d => payments.push({
        tk:p.tk, date:d.date, pps:d.amount, amt:d.amount*p.shares, exDay:p.exDay
      }));
    });
    return payments.sort((a,b)=>a.tk.localeCompare(b.tk));
  }

  // 8-week history — group all payments by week (Mon–Sun buckets)
  function eightWeekHistory(positions) {
    // Build weekly buckets for last 8 weeks
    const weeks = [];
    const now = new Date();
    for (let i = 7; i >= 0; i--) {
      const weekEnd = new Date(now);
      weekEnd.setDate(now.getDate() - (i * 7));
      weekEnd.setHours(23,59,59,999);
      const weekStart = new Date(weekEnd);
      weekStart.setDate(weekEnd.getDate() - 6);
      weekStart.setHours(0,0,0,0);
      weeks.push({start: weekStart, end: weekEnd, total: 0, label: ''});
    }

    // Sum payments into buckets
    positions.forEach(p => {
      (yahoo[p.tk]?.divs || []).forEach(d => {
        weeks.forEach(w => {
          if (d.date >= w.start && d.date <= w.end) {
            w.total += d.amount * p.shares;
          }
        });
      });
    });

    // Label each week by its pay date
    weeks.forEach(w => {
      const d = new Date(w.end);
      w.label = d.toLocaleDateString('en-US',{month:'short',day:'numeric'});
    });

    return weeks;
  }

  const indivTot = sum(indiv);
  const iraTot   = sum(ira);

  // Total Return incl. closed = live open net minus locked-in closed losses
  const indivClosedNet = indivTot.net - POSITIONS.INDIV_CLOSED_LOSS;
  const iraClosedNet   = iraTot.net   - POSITIONS.IRA_CLOSED_LOSS;

  return {
    indiv, ira, watchlist,
    indivTot, iraTot,
    indivRecent:  recentPayments(POSITIONS.INDIV),
    iraRecent:    recentPayments(POSITIONS.IRA),
    indivWeeks:   eightWeekHistory(POSITIONS.INDIV),
    iraWeeks:     eightWeekHistory(POSITIONS.IRA),
    indivClosedNet,
    iraClosedNet,
  };
}

// ── FORMATTING HELPERS ────────────────────────────────────────
function f$(n)  { return (n<0?'-$':'+'+'$')+Math.abs(Math.round(n)).toLocaleString(); }
function fv(n)  { return '$'+Math.abs(Math.round(n)).toLocaleString(); }
function fp(n)  { return (n>=0?'+':'')+n.toFixed(1)+'%'; }
function gc(n)  { return n>=0?'#4ade80':'#f87171'; }
function fd(d)  { return d.toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric'}); }

// ── EMAIL HELPERS ─────────────────────────────────────────────
function kpi(label, value, sub, color) {
  return `<td width="20%" style="padding:4px;vertical-align:top;height:1px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;height:100%;">
      <tr><td style="padding:12px;vertical-align:top;">
        <div style="color:#aaa;font-size:10px;font-family:Arial,sans-serif;text-transform:uppercase;margin-bottom:6px;">${label}</div>
        <div style="color:${color||'#fff'};font-size:20px;font-weight:bold;font-family:Arial,sans-serif;margin-bottom:4px;">${value}</div>
        <div style="color:#888;font-size:10px;font-family:Arial,sans-serif;">${sub}</div>
      </td></tr>
    </table>
  </td>`;
}

function posRow(p, bg) {
  return `<tr style="background-color:${bg};">
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:#fff;border-bottom:1px solid #333;">${p.tk} <span style="font-size:9px;padding:1px 4px;border-radius:3px;background:${p.exDay==='WED'?'#1a3a4a':'#1a4a3a'};color:${p.exDay==='WED'?'#64b5f6':'#4ade80'};">${p.exDay}</span></td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#ddd;text-align:right;border-bottom:1px solid #333;">${p.shares.toLocaleString()}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#ddd;text-align:right;border-bottom:1px solid #333;">$${p.price.toFixed(2)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#ddd;text-align:right;border-bottom:1px solid #333;">${fv(p.val)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:${gc(p.pl)};text-align:right;border-bottom:1px solid #333;">${f$(p.pl)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;text-align:right;border-bottom:1px solid #333;">${fv(p.divs)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:${gc(p.retPct)};text-align:right;border-bottom:1px solid #333;">${fp(p.retPct)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;text-align:right;border-bottom:1px solid #333;">${p.pbPct.toFixed(1)}%</td>
  </tr>`;
}

function totRow(tot) {
  return `<tr style="background-color:#1a1a1a;">
    <td colspan="3" style="padding:8px 10px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;font-weight:bold;border-top:2px solid #444;">TOTAL</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#fff;font-weight:bold;text-align:right;border-top:2px solid #444;">${fv(tot.val)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:${gc(tot.pl)};font-weight:bold;text-align:right;border-top:2px solid #444;">${f$(tot.pl)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">${fv(tot.divs)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:${gc(tot.net)};font-weight:bold;text-align:right;border-top:2px solid #444;">${fp(tot.net/tot.cost*100)}</td>
    <td style="padding:8px 10px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;font-weight:bold;text-align:right;border-top:2px solid #444;">${(tot.divs/tot.cost*100).toFixed(1)}%</td>
  </tr>`;
}

function pbBar(p) {
  const pct = Math.min(Math.round(p.pbPct), 100);
  return `<td style="padding:6px 4px;width:20%;vertical-align:top;height:1px;">
    <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;height:100%;">
      <tr><td style="padding:10px 12px;vertical-align:top;">
        <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
          <span style="font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:#fff;">${p.tk}</span>
          <span style="font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#4ade80;">${p.pbPct.toFixed(1)}%</span>
        </div>
        <div style="font-family:Arial,sans-serif;font-size:9px;color:#888;margin-bottom:6px;">${fv(p.divs)} of ${fv(p.cost)}</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#1a1a1a;border-radius:3px;">
          <tr>
            <td width="${pct}%" style="background:linear-gradient(90deg,#64b5f6,#4ade80);border-radius:3px;height:5px;"></td>
            <td width="${100-pct}%" style="height:5px;"></td>
          </tr>
        </table>
      </td></tr>
    </table>
  </td>`;
}

function calRow(p) {
  return `<tr>
    <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:12px;font-weight:bold;color:#fff;border-bottom:1px solid #333;width:55px;">${p.tk}</td>
    <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;border-bottom:1px solid #333;">${fd(p.date)} · $${p.pps.toFixed(4)}/sh</td>
    <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:12px;color:#4ade80;text-align:right;border-bottom:1px solid #333;">+${fv(p.amt)}</td>
    <td style="padding:8px 12px;border-bottom:1px solid #333;text-align:right;"><span style="background-color:#1a3a2a;color:#4ade80;font-size:9px;padding:2px 6px;border-radius:3px;border:1px solid #2a5a3a;">PAID</span></td>
  </tr>`;
}

function section(title) {
  return `<tr><td colspan="2" style="padding:20px 0 10px 0;">
    <div style="font-family:Arial,sans-serif;font-size:16px;font-weight:bold;color:#fff;border-left:3px solid #4ade80;padding-left:10px;">${title}</div>
  </td></tr>`;
}

function posTable(positions) {
  const rows = positions.map((p,i)=>posRow(p,i%2===0?'#2a2a2a':'#252525')).join('');
  const tot  = positions.reduce((s,p)=>({val:s.val+p.val,cost:s.cost+p.cost,divs:s.divs+p.divs,pl:s.pl+p.pl,net:s.net+p.net}),{val:0,cost:0,divs:0,pl:0,net:0});
  return `<table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;overflow:hidden;">
    <tr style="background-color:#1a1a1a;">
      <th style="padding:8px 10px;text-align:left;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">Ticker</th>
      <th style="padding:8px 10px;text-align:right;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">Shares</th>
      <th style="padding:8px 10px;text-align:right;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">Price</th>
      <th style="padding:8px 10px;text-align:right;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">Value</th>
      <th style="padding:8px 10px;text-align:right;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">P/L</th>
      <th style="padding:8px 10px;text-align:right;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">Dividends</th>
      <th style="padding:8px 10px;text-align:right;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">Total Ret%</th>
      <th style="padding:8px 10px;text-align:right;font-size:10px;color:#aaa;border-bottom:1px solid #444;font-weight:normal;text-transform:uppercase;">Payback%</th>
    </tr>
    ${rows}
    ${totRow(tot)}
  </table>`;
}

// ── MAIN EMAIL HTML ───────────────────────────────────────────
function getDashboardHtml(d) {
  const today = Utilities.formatDate(new Date(),'America/Los_Angeles','MMM d, yyyy');
  const it = d.indivTot, rt = d.iraTot;

  const indivPbSorted = [...d.indiv].sort((a,b)=>b.pbPct-a.pbPct);
  const iraPbSorted   = [...d.ira].sort((a,b)=>b.pbPct-a.pbPct);

  const indivCalRows = d.indivRecent.map(calRow).join('') || '<tr><td colspan="4" style="padding:12px;color:#888;font-family:Arial,sans-serif;font-size:11px;">No payments this week</td></tr>';
  const iraCalRows   = d.iraRecent.map(calRow).join('')   || '<tr><td colspan="4" style="padding:12px;color:#888;font-family:Arial,sans-serif;font-size:11px;">No payments this week</td></tr>';

  const indivWeekTotal = d.indivRecent.reduce((s,p)=>s+p.amt,0);
  const iraWeekTotal   = d.iraRecent.reduce((s,p)=>s+p.amt,0);

  return `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background-color:#111111;">
<div style="background-color:#111111;padding:20px;">
<table width="100%" cellpadding="0" cellspacing="0" style="max-width:900px;margin:0 auto;background-color:#111111;">

  <tr><td colspan="2" style="padding-bottom:20px;text-align:center;">
    <div style="font-family:Arial,sans-serif;font-size:26px;font-weight:bold;color:#fff;">Dividend Portfolio</div>
    <div style="font-family:Arial,sans-serif;font-size:12px;color:#888;margin-top:4px;">Updated ${today}</div>
  </td></tr>

  ${section('INDIV')}

  <tr><td colspan="2" style="padding-bottom:12px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    ${kpi('Portfolio Value',fv(it.val),'Cost: '+fv(it.cost),'')}
    ${kpi('Total Return',f$(it.net),fp(it.net/it.cost*100),gc(it.net))}
    ${kpi('Total Return (All-In)',f$(d.indivClosedNet),fp(d.indivClosedNet/it.cost*100),gc(d.indivClosedNet))}
    ${kpi('Forecast / Week',fv(it.fcstWk),'2026 YTD avg','#64b5f6')}
    ${kpi('Forecast / Month',fv(it.fcstWk*4),'2026 YTD avg','#ffd54f')}
  </tr></table></td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Positions</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">${posTable(d.indiv)}</td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Payback Progress</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    ${indivPbSorted.map(pbBar).join('')}
  </tr></table></td></tr>

  <tr><td colspan="2" style="padding-bottom:15px;">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="55%" style="vertical-align:top;padding-right:8px;">
        <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding-bottom:6px;">Weekly Div Received — 8 Weeks</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
          <tr><td style="padding:14px;">
            ${(()=>{
              const weeks = d.indivWeeks;
              const max = Math.max(...weeks.map(w=>w.total), 1);
              const avg = weeks.reduce((s,w)=>s+w.total,0)/weeks.length;
              const bars = weeks.map(w=>{
                const h = Math.round((w.total/max)*80);
                return `<td style="text-align:center;vertical-align:bottom;padding:0 3px;width:12%;">
                  <div style="background-color:#4ade80;opacity:0.65;width:100%;height:${h}px;border-radius:2px 2px 0 0;min-height:2px;"></div>
                  <div style="font-family:Arial,sans-serif;font-size:8px;color:#888;margin-top:3px;">${w.label}</div>
                </td>`;
              }).join('');
              const avgPct = Math.round((avg/max)*80);
              return `<table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #444;">
                <tr style="height:80px;vertical-align:bottom;">${bars}</tr>
              </table>
              <div style="font-family:Arial,sans-serif;font-size:9px;color:#aaa;margin-top:6px;">8-wk avg: <span style="color:#fff;font-weight:bold;">+${fv(avg)}</span></div>`;
            })()}
          </td></tr>
        </table>
      </td>
      <td width="45%" style="vertical-align:top;padding-left:8px;">
        <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding-bottom:6px;">This Week's Payments</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
          <tr><td style="padding:6px 0;">
            <table width="100%" cellpadding="0" cellspacing="0">${indivCalRows}</table>
            <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#1a1a1a;border-top:1px solid #444;">
              <tr><td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;">Week total</td>
              <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#4ade80;text-align:right;">+${fv(indivWeekTotal)}</td></tr>
            </table>
          </td></tr>
        </table>
      </td>
    </tr></table>
  </td></tr>

  ${section('IRA')}

  <tr><td colspan="2" style="padding-bottom:12px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    ${kpi('Portfolio Value',fv(rt.val),'Cost: '+fv(rt.cost),'')}
    ${kpi('Total Return',f$(rt.net),fp(rt.net/rt.cost*100),gc(rt.net))}
    ${kpi('Total Return (All-In)',f$(d.iraClosedNet),fp(d.iraClosedNet/rt.cost*100),gc(d.iraClosedNet))}
    ${kpi('Forecast / Week',fv(rt.fcstWk),'2026 YTD avg','#64b5f6')}
    ${kpi('Forecast / Month',fv(rt.fcstWk*4),'2026 YTD avg','#ffd54f')}
  </tr></table></td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Positions</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;">${posTable(d.ira)}</td></tr>

  <tr><td colspan="2" style="padding-bottom:6px;"><div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;">Payback Progress</div></td></tr>
  <tr><td colspan="2" style="padding-bottom:15px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    ${iraPbSorted.map(pbBar).join('')}
    <td width="${(5-d.ira.length)*20}%" style="padding:4px;"></td>
  </tr></table></td></tr>

  <tr><td colspan="2" style="padding-bottom:30px;">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="55%" style="vertical-align:top;padding-right:8px;">
        <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding-bottom:6px;">Weekly Div Received — 8 Weeks</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
          <tr><td style="padding:14px;">
            ${(()=>{
              const weeks = d.iraWeeks;
              const max = Math.max(...weeks.map(w=>w.total), 1);
              const avg = weeks.reduce((s,w)=>s+w.total,0)/weeks.length;
              const bars = weeks.map(w=>{
                const h = Math.round((w.total/max)*80);
                return `<td style="text-align:center;vertical-align:bottom;padding:0 3px;width:12%;">
                  <div style="background-color:#4ade80;opacity:0.65;width:100%;height:${h}px;border-radius:2px 2px 0 0;min-height:2px;"></div>
                  <div style="font-family:Arial,sans-serif;font-size:8px;color:#888;margin-top:3px;">${w.label}</div>
                </td>`;
              }).join('');
              return `<table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #444;">
                <tr style="height:80px;vertical-align:bottom;">${bars}</tr>
              </table>
              <div style="font-family:Arial,sans-serif;font-size:9px;color:#aaa;margin-top:6px;">8-wk avg: <span style="color:#fff;font-weight:bold;">+${fv(avg)}</span></div>`;
            })()}
          </td></tr>
        </table>
      </td>
      <td width="45%" style="vertical-align:top;padding-left:8px;">
        <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:bold;color:#aaa;text-transform:uppercase;padding-bottom:6px;">This Week's Payments</div>
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#2a2a2a;border:1px solid #444;border-radius:6px;">
          <tr><td style="padding:6px 0;">
            <table width="100%" cellpadding="0" cellspacing="0">${iraCalRows}</table>
            <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#1a1a1a;border-top:1px solid #444;">
              <tr><td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:11px;color:#aaa;">Week total</td>
              <td style="padding:8px 12px;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;color:#4ade80;text-align:right;">+${fv(iraWeekTotal)}</td></tr>
            </table>
          </td></tr>
        </table>
      </td>
    </tr></table>
  </td></tr>

  <tr><td colspan="2" style="text-align:center;padding-top:20px;border-top:1px solid #333;">
    <a href="${CONFIG.DASHBOARD}" style="display:inline-block;background-color:#4ade80;color:#000;font-family:Arial,sans-serif;font-size:13px;font-weight:bold;padding:12px 28px;border-radius:6px;text-decoration:none;">View Full Dashboard →</a>
    <div style="font-family:Arial,sans-serif;font-size:11px;color:#555;margin-top:12px;">Auto-generated by Dividend Tracker</div>
  </td></tr>

</table></div></body></html>`;
}

// ── SEND EMAIL ────────────────────────────────────────────────
function sendEmail(html) {
  const today = Utilities.formatDate(new Date(),'America/Los_Angeles','MMMM d, yyyy');
  GmailApp.sendEmail(CONFIG.EMAIL, `Dividend Portfolio — ${today}`, '', {
    htmlBody: html, name: 'Dividend Tracker'
  });
  Logger.log('Email sent to '+CONFIG.EMAIL);
}

// ── WEB APP ───────────────────────────────────────────────────
function doGet() {
  const data = buildPortfolioData();
  const html = HtmlService.createHtmlOutput(getInteractiveDashboard(data));
  html.setTitle('Dividend Portfolio');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function getInteractiveDashboard(data) {
  const d  = data;
  const it = d.indivTot, rt = d.iraTot;
  const t  = HtmlService.createTemplateFromFile('Dashboard');
  t.indivJson        = JSON.stringify(d.indiv);
  t.iraJson          = JSON.stringify(d.ira);
  t.watchJson        = JSON.stringify(d.watchlist);
  t.indivTotJson     = JSON.stringify(it);
  t.iraTotJson       = JSON.stringify(rt);
  t.indivClosed      = d.indivClosedNet;
  t.iraClosed        = d.iraClosedNet;
  t.indivWeeksJson   = JSON.stringify(d.indivWeeks);
  t.iraWeeksJson     = JSON.stringify(d.iraWeeks);
  t.indivRecentJson  = JSON.stringify(d.indivRecent);
  t.iraRecentJson    = JSON.stringify(d.iraRecent);
  return t.evaluate().getContent();
}

// ── TRIGGERS ─────────────────────────────────────────────────
function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runWeeklyUpdate').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(7).create();
  ScriptApp.newTrigger('runWeeklyUpdate').timeBased().onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(7).create();
  Logger.log('Triggers created: Wednesday + Thursday 7AM');
}
