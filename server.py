"""
SAP Clearing Dashboard — servidor local Flask
Requiere: pip install flask
"""

import os
import sys
import shutil
import tempfile
import subprocess
from flask import Flask, Response, render_template_string, jsonify
import pandas as pd

app = Flask(__name__)

SCRIPT_DIR = os.path.join(
    os.path.expanduser("~"),
    "OneDrive - SAP SE", "Varios SAP - Daily", "JE-Sofi", "Clearing"
)

SCRIPTS = {
    "us":     "Clearing US Complete v2.py",
    "canada": "Clearing Canada Complete v2.py",
    "concur": "Clearing Concur Complete v2.py",
}

OUTPUT_FILES = {
    "us":     "Clearing US Pivot.xlsx",
    "canada": "Clearing Canada Pivot.xlsx",
    "concur": "Clearing Concur Pivot.xlsx",
}

INPUT_FILES = {
    "us":     "Clearing US Input.xlsx",
    "canada": "Clearing Canada Input.xlsx",
    "concur": "Clearing Concur Input.xlsx",
}

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SAP Clearing</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {
    --shell: #1D2D3E; --blue: #0070F2; --blue-dk: #0057C2;
    --bg: #F5F6F7; --card: #FFFFFF; --text: #1D2D3E; --sub: #556B82;
    --ok: #188918; --err: #BB0000; --warn: #E9730C;
    --border: #E0E0E0; --r: 8px; --sh: 0 1px 4px rgba(0,0,0,.10);
    --font: '72','72full','SAP 72',Arial,sans-serif;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:var(--font);background:var(--bg);min-height:100vh;color:var(--text);}

  .shell{background:var(--shell);height:44px;display:flex;align-items:center;padding:0 16px;gap:14px;}
  .shell-brand{color:#fff;font-size:.7rem;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;padding-right:14px;border-right:1px solid rgba(255,255,255,.2);}
  .shell-title{color:rgba(255,255,255,.8);font-size:.85rem;}

  .page-header{background:var(--card);border-bottom:1px solid var(--border);padding:18px 32px;display:flex;align-items:center;gap:14px;}
  .joule-badge{width:38px;height:38px;background:linear-gradient(135deg,#0070F2,#003CB4);border-radius:10px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:1.1rem;flex-shrink:0;box-shadow:0 2px 8px rgba(0,112,242,.35);}
  .page-header h1{font-size:1.15rem;font-weight:700;}
  .page-header p{font-size:.78rem;color:var(--sub);margin-top:2px;}

  .content{padding:24px 32px;max-width:1200px;margin:0 auto;}
  .section-label{font-size:.7rem;font-weight:700;color:var(--sub);text-transform:uppercase;letter-spacing:.9px;margin-bottom:14px;}

  /* Execution cards */
  .exec-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px;margin-bottom:28px;}
  .card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh);overflow:hidden;}
  .card-hdr{padding:14px 18px;display:flex;align-items:center;gap:12px;border-bottom:1px solid var(--border);}
  .card-icon{width:34px;height:34px;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:1rem;flex-shrink:0;}
  .icon-us{background:#EBF3FF;} .icon-ca{background:#FFF4E5;} .icon-co{background:#EDFBEE;}
  .card-hdr h2{font-size:.9rem;font-weight:700;}
  .card-hdr .sub{font-size:.72rem;color:var(--sub);margin-top:1px;}
  .card-body{padding:14px 18px;}
  .card-actions{display:flex;align-items:center;gap:10px;}
  .btn{display:inline-flex;align-items:center;gap:6px;padding:6px 16px;border:1px solid var(--blue);border-radius:4px;background:var(--blue);color:#fff;font-size:.85rem;font-weight:600;cursor:pointer;transition:background .12s;font-family:var(--font);}
  .btn:not(:disabled):hover{background:var(--blue-dk);border-color:var(--blue-dk);}
  .btn:disabled{opacity:.4;cursor:not-allowed;}
  .status{display:inline-flex;align-items:center;gap:5px;font-size:.73rem;font-weight:600;padding:3px 10px;border-radius:100px;}
  .dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;}
  .status.idle{background:#EEE;color:#757575;} .status.idle .dot{background:#9E9E9E;}
  .status.running{background:#EBF3FF;color:var(--blue);} .status.running .dot{background:var(--blue);animation:blink 1s infinite;}
  .status.ok{background:#EDFBEE;color:var(--ok);} .status.ok .dot{background:var(--ok);}
  .status.error{background:#FFF0F0;color:var(--err);} .status.error .dot{background:var(--err);}
  @keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}
  .log-toggle{margin-top:8px;font-size:.72rem;color:var(--blue);cursor:pointer;background:none;border:none;padding:0;font-family:var(--font);}
  .log-toggle:hover{text-decoration:underline;}
  .btn-ghost{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border:1px solid var(--border);border-radius:4px;background:#fff;color:var(--text);font-size:.82rem;font-weight:500;cursor:pointer;transition:background .12s;font-family:var(--font);}
  .btn-ghost:hover{background:var(--bg);}
  .log-wrap{margin-top:8px;border-radius:6px;overflow:hidden;display:none;}
  .log-wrap.visible{display:block;}
  .log-bar{background:#1A2535;padding:5px 12px;font-size:.66rem;color:#546E7A;font-family:Consolas,monospace;text-transform:uppercase;letter-spacing:.5px;}
  .log-body{background:#0D1520;padding:10px 12px;height:130px;overflow-y:auto;font-family:'Cascadia Code',Consolas,monospace;font-size:.73rem;color:#90A4AE;white-space:pre-wrap;word-break:break-all;}
  .line-ok{color:#66BB6A;} .line-err{color:#EF5350;} .line-inf{color:#42A5F5;}

  /* Results section */
  .results-section{display:none;}
  .results-section.visible{display:block;}
  .result-panel{background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh);margin-bottom:16px;overflow:hidden;display:none;}
  .result-panel.visible{display:block;}
  .panel-hdr{padding:14px 20px;display:flex;align-items:center;gap:12px;border-bottom:2px solid var(--border);background:#FAFBFC;}
  .panel-hdr .flag{font-size:1.2rem;}
  .panel-hdr h3{font-size:.95rem;font-weight:700;flex:1;}
  .panel-badge{background:#FFF0F0;color:var(--err);padding:3px 10px;border-radius:100px;font-size:.72rem;font-weight:700;}
  .panel-badge.all-clear{background:#EDFBEE;color:var(--ok);}

  /* Results table */
  .rtbl{width:100%;border-collapse:collapse;font-size:.82rem;}
  .rtbl th{background:#F5F6F7;color:var(--sub);font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;padding:8px 16px;text-align:left;border-bottom:2px solid var(--border);position:sticky;top:0;z-index:1;}
  .rtbl th.r{text-align:right;}
  .rtbl td{padding:7px 16px;border-bottom:1px solid #F2F2F2;vertical-align:middle;}
  /* Account group header row */
  .acc-hdr td{background:#EBF3FF;font-weight:700;font-size:.84rem;padding:8px 16px;border-bottom:1px solid #C8DEF7;border-top:2px solid #B8D0F0;}
  .acc-hdr .acc-num{color:var(--blue);}
  .acc-hdr .acc-bal{text-align:right;color:var(--err);}
  .acc-hdr .acc-cc{color:var(--text);}
  /* Document rows */
  .doc-row td{background:#fff;font-size:.8rem;}
  .doc-row:hover td{background:#FAFBFF;}
  .doc-row td:first-child{padding-left:32px;color:var(--sub);}
  .doc-row .amt{text-align:right;font-variant-numeric:tabular-nums;}
  .doc-row .amt.neg{color:var(--err);font-weight:600;}
  .doc-row .amt.pos{color:var(--ok);font-weight:600;}
  .doc-row .txt{color:var(--sub);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .rtbl tr:last-child td{border-bottom:none;}
  .tbl-scroll{max-height:500px;overflow-y:auto;}

  /* ── Analytics ─────────────────────────────────── */
  .analytics-section{margin-bottom:28px;}
  .flow-tabs{display:flex;gap:6px;flex-shrink:0;}
  .flow-tab{padding:5px 14px;border:1px solid var(--border);border-radius:100px;background:#fff;font-size:.78rem;font-weight:600;color:var(--sub);cursor:pointer;font-family:var(--font);transition:all .12s;}
  .flow-tab:hover{border-color:var(--blue);color:var(--blue);}
  .flow-tab.active{background:var(--blue);border-color:var(--blue);color:#fff;}
  .kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:14px;}
  @media(max-width:800px){.kpi-row{grid-template-columns:repeat(2,1fr);}}
  .kpi-tile{background:var(--card);border:1px solid var(--border);border-radius:var(--r);padding:14px 18px;box-shadow:var(--sh);}
  .kpi-tile.kpi-green{border-left:3px solid #0ca30c;}
  .kpi-tile.kpi-red{border-left:3px solid #d03b3b;}
  .kpi-tile.kpi-blue{border-left:3px solid var(--blue);}
  .kpi-label{font-size:.66rem;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--sub);margin-bottom:6px;}
  .kpi-value{font-size:1.7rem;font-weight:700;color:var(--text);line-height:1;}
  .kpi-sub{font-size:.71rem;color:var(--sub);margin-top:4px;}
  .charts-row{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;}
  @media(max-width:800px){.charts-row{grid-template-columns:1fr;}}
  .chart-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r);box-shadow:var(--sh);padding:16px 20px;}
  .chart-card-full{margin-bottom:14px;}
  .chart-card-title{font-size:.78rem;font-weight:700;color:var(--text);margin-bottom:12px;display:flex;align-items:center;gap:8px;}
  .chart-card-title .ct-sub{font-weight:400;color:var(--sub);font-size:.72rem;}
  .chart-wrap{position:relative;}
</style>
</head>
<body>

<div class="shell">
  <span class="shell-brand">SAP</span>
  <span class="shell-title">Clearing Automation</span>
</div>

<div class="page-header">
  <div class="joule-badge">✦</div>
  <div>
    <h1>Clearing Dashboard</h1>
    <p>FAGLL03 report automation — ISP &amp; I4P</p>
  </div>
</div>

<div class="content">
  <div class="section-label">Run Report</div>
  <div class="exec-grid">

    <div class="card">
      <div class="card-hdr">
        <div class="card-icon icon-us">🇺🇸</div>
        <div><h2>United States</h2><div class="sub">Variant /GL ACC SA · ISP</div></div>
      </div>
      <div class="card-body">
        <div class="card-actions">
          <button class="btn" onclick="run('us',this)">▶ Run</button>
          <button class="btn-ghost" onclick="loadResults('us')">↻ Load Results</button>
          <span id="status-us" class="status idle"><span class="dot"></span>Idle</span>
        </div>
        <button class="log-toggle" onclick="toggleLog('us',this)" style="display:none" id="logtgl-us">Show output ▾</button>
        <div id="log-us" class="log-wrap"><div class="log-bar">output</div><div id="term-us" class="log-body"></div></div>
      </div>
    </div>

    <div class="card">
      <div class="card-hdr">
        <div class="card-icon icon-ca">🇨🇦</div>
        <div><h2>Canada</h2><div class="sub">Variant /GLCANADA · ISP</div></div>
      </div>
      <div class="card-body">
        <div class="card-actions">
          <button class="btn" onclick="run('canada',this)">▶ Run</button>
          <button class="btn-ghost" onclick="loadResults('canada')">↻ Load Results</button>
          <span id="status-canada" class="status idle"><span class="dot"></span>Idle</span>
        </div>
        <button class="log-toggle" onclick="toggleLog('canada',this)" style="display:none" id="logtgl-canada">Show output ▾</button>
        <div id="log-canada" class="log-wrap"><div class="log-bar">output</div><div id="term-canada" class="log-body"></div></div>
      </div>
    </div>

    <div class="card">
      <div class="card-hdr">
        <div class="card-icon icon-co">📋</div>
        <div><h2>Concur</h2><div class="sub">Variant PAYROLLGLCLEAR · I4P</div></div>
      </div>
      <div class="card-body">
        <div class="card-actions">
          <button class="btn" onclick="run('concur',this)">▶ Run</button>
          <button class="btn-ghost" onclick="loadResults('concur')">↻ Load Results</button>
          <span id="status-concur" class="status idle"><span class="dot"></span>Idle</span>
        </div>
        <button class="log-toggle" onclick="toggleLog('concur',this)" style="display:none" id="logtgl-concur">Show output ▾</button>
        <div id="log-concur" class="log-wrap"><div class="log-bar">output</div><div id="term-concur" class="log-body"></div></div>
      </div>
    </div>

  </div>

  <!-- Analytics section -->
  <div id="analytics-section" class="analytics-section" style="display:none;">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;">
      <div class="section-label" style="margin:0;">Analytics</div>
      <div class="flow-tabs">
        <button class="flow-tab" id="atab-us"     onclick="switchAnalytics('us',this)">🇺🇸 US</button>
        <button class="flow-tab" id="atab-canada" onclick="switchAnalytics('canada',this)">🇨🇦 Canada</button>
        <button class="flow-tab" id="atab-concur" onclick="switchAnalytics('concur',this)">📋 Concur</button>
      </div>
    </div>

    <!-- KPI row -->
    <div class="kpi-row">
      <div class="kpi-tile">
        <div class="kpi-label">Total Combinations</div>
        <div class="kpi-value" id="kpi-total">—</div>
        <div class="kpi-sub">Account × Company Code</div>
      </div>
      <div class="kpi-tile kpi-green">
        <div class="kpi-label">Clearable</div>
        <div class="kpi-value" id="kpi-clear">—</div>
        <div class="kpi-sub" id="kpi-clear-sub">Balance = 0</div>
      </div>
      <div class="kpi-tile kpi-red">
        <div class="kpi-label">Open Items</div>
        <div class="kpi-value" id="kpi-open">—</div>
        <div class="kpi-sub">Balance ≠ 0</div>
      </div>
      <div class="kpi-tile kpi-blue">
        <div class="kpi-label">Docs to Clear</div>
        <div class="kpi-value" id="kpi-docs">—</div>
        <div class="kpi-sub">clearable documents</div>
      </div>
    </div>

    <!-- Charts row 1 -->
    <div class="charts-row">
      <div class="chart-card">
        <div class="chart-card-title">Clear vs Open Items
          <span class="ct-sub">by Account</span>
        </div>
        <div class="chart-wrap" id="wrap-status"><canvas id="chart-status"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-card-title">Clearable Documents
          <span class="ct-sub">by Account</span>
        </div>
        <div class="chart-wrap" id="wrap-docs"><canvas id="chart-docs"></canvas></div>
      </div>
    </div>

    <!-- Charts row 2: period (only when Year/Month available) -->
    <div class="chart-card chart-card-full" id="period-card" style="display:none;">
      <div class="chart-card-title">Clearable Documents
        <span class="ct-sub">by Period</span>
      </div>
      <div style="position:relative;height:180px;"><canvas id="chart-period"></canvas></div>
    </div>
  </div>

  <!-- Results section -->
  <div id="results-section" class="results-section">
    <div class="section-label">Clearable Documents</div>

    <div id="panel-us" class="result-panel">
      <div class="panel-hdr">
        <span class="flag">🇺🇸</span>
        <h3>United States</h3>
        <span id="badge-us" class="panel-badge"></span>
      </div>
      <div class="tbl-scroll">
        <table class="rtbl">
          <thead><tr><th>G/L Account</th><th>Company Code</th><th>Doc Number</th><th>Year/Month</th><th class="r">Amount</th><th>Text</th></tr></thead>
          <tbody id="tbody-us"></tbody>
        </table>
      </div>
    </div>

    <div id="panel-canada" class="result-panel">
      <div class="panel-hdr">
        <span class="flag">🇨🇦</span>
        <h3>Canada</h3>
        <span id="badge-canada" class="panel-badge"></span>
      </div>
      <div class="tbl-scroll">
        <table class="rtbl">
          <thead><tr><th>G/L Account</th><th>Company Code</th><th>Doc Number</th><th>Year/Month</th><th class="r">Amount</th><th>Text</th></tr></thead>
          <tbody id="tbody-canada"></tbody>
        </table>
      </div>
    </div>

    <div id="panel-concur" class="result-panel">
      <div class="panel-hdr">
        <span class="flag">📋</span>
        <h3>Concur</h3>
        <span id="badge-concur" class="panel-badge"></span>
      </div>
      <div class="tbl-scroll">
        <table class="rtbl">
          <thead><tr><th>G/L Account</th><th>Company Code</th><th>Doc Number</th><th>Year/Month</th><th class="r">Amount</th><th>Text</th></tr></thead>
          <tbody id="tbody-concur"></tbody>
        </table>
      </div>
    </div>

  </div>
</div>

<script>
function fmt(n){return Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});}

function toggleLog(key,btn){
  const w=document.getElementById('log-'+key);
  const open=w.classList.toggle('visible');
  btn.textContent=open?'Hide output ▴':'Show output ▾';
}

function run(key,btn){
  const term=document.getElementById('term-'+key);
  const status=document.getElementById('status-'+key);
  const logtgl=document.getElementById('logtgl-'+key);
  btn.disabled=true; term.innerHTML='';
  status.className='status running';
  status.innerHTML='<span class="dot"></span>Running...';
  logtgl.style.display='none';

  const es=new EventSource('/run/'+key);
  es.onmessage=function(e){
    const msg=e.data;
    if(msg.startsWith('__DONE__:')){
      es.close(); btn.disabled=false;
      const ok=msg.includes(':OK');
      status.className='status '+(ok?'ok':'error');
      status.innerHTML='<span class="dot"></span>'+(ok?'Completed':'Error');
      logtgl.style.display='inline';
      if(ok){ loadResults(key); loadAnalytics(key); }
      return;
    }
    const d=document.createElement('div');
    const l=msg.toLowerCase();
    if(l.includes('error')) d.className='line-err';
    else if(l.startsWith('  ')) d.className='line-inf';
    else if(l.includes('downloaded')||l.includes('generated')||l.includes('completed')||l.includes('successfully')) d.className='line-ok';
    d.textContent=msg; term.appendChild(d); term.scrollTop=term.scrollHeight;
  };
  es.onerror=function(){
    es.close(); btn.disabled=false;
    status.className='status error';
    status.innerHTML='<span class="dot"></span>No connection';
  };
}

function loadResults(key){
  fetch('/results/'+key).then(r=>r.json()).then(data=>{
    const panel=document.getElementById('panel-'+key);
    const tbody=document.getElementById('tbody-'+key);
    const badge=document.getElementById('badge-'+key);
    const section=document.getElementById('results-section');

    if(data && data.error){
      badge.textContent='Read error';
      badge.className='panel-badge';
      panel.classList.add('visible');
      section.classList.add('visible');
      return;
    }
    if(!Array.isArray(data)||!data.length){
      badge.textContent='No open items';
      badge.className='panel-badge all-clear';
      panel.classList.add('visible');
      section.classList.add('visible');
      return;
    }
    tbody.innerHTML='';
    data.forEach(grp=>{
      const hdr=document.createElement('tr');
      hdr.className='acc-hdr';
      hdr.innerHTML=
        '<td class="acc-num">'+grp.account+'</td>'+
        '<td class="acc-cc">'+grp.cc+'</td>'+
        '<td colspan="3"></td>'+
        '<td class="acc-bal"><span style="background:#EDFBEE;color:#188918;padding:2px 10px;border-radius:100px;font-size:.75rem;font-weight:700;">✓ CLEARABLE</span></td>';
      tbody.appendChild(hdr);
      grp.documents.forEach(doc=>{
        const tr=document.createElement('tr');
        tr.className='doc-row';
        tr.innerHTML=
          '<td></td>'+
          '<td></td>'+
          '<td>'+doc.doc+'</td>'+
          '<td>'+doc.ym+'</td>'+
          '<td class="amt '+(doc.amount<0?'neg':'pos')+'">'+fmt(doc.amount)+'</td>'+
          '<td class="txt">'+(doc.text||'')+'</td>';
        tbody.appendChild(tr);
      });
    });
    badge.textContent=data.length+(data.length===1?' account':' accounts');
    badge.className='panel-badge';
    panel.classList.add('visible');
    section.classList.add('visible');
  }).catch(()=>{
    const badge=document.getElementById('badge-'+key);
    badge.textContent='File not found';
    badge.className='panel-badge';
    document.getElementById('panel-'+key).classList.add('visible');
    document.getElementById('results-section').classList.add('visible');
  });
}
// ── Analytics ────────────────────────────────────────────────
const analyticsCache={};
let activeKey=null;
let cStatus=null,cDocs=null,cPeriod=null;

// Palette — reference instance (palette.md) + status tokens
const P={
  clear:'#0ca30c', open:'#d03b3b', blue:'#2a78d6',
  grid:'#e1e0d9', muted:'#898781', sub:'#52514e', surface:'#ffffff'
};
const BASE_OPT={
  responsive:true, maintainAspectRatio:false,
  animation:{duration:300},
  plugins:{tooltip:{cornerRadius:4,padding:8,
    callbacks:{label:ctx=>' '+ctx.dataset.label+': '+ctx.formattedValue}}}
};

function fetchAnalytics(key){
  return fetch('/summary/'+key).then(r=>r.json()).then(d=>{
    if(!d||d.error||Array.isArray(d)) return false;
    analyticsCache[key]=d; return true;
  }).catch(()=>false);
}

function showAnalytics(key){
  activeKey=key;
  document.getElementById('analytics-section').style.display='block';
  document.querySelectorAll('.flow-tab').forEach(b=>b.classList.remove('active'));
  const t=document.getElementById('atab-'+key); if(t) t.classList.add('active');
  renderAnalytics(key);
}

function loadAnalytics(key){
  if(analyticsCache[key]){showAnalytics(key);return;}
  fetchAnalytics(key).then(ok=>{if(ok) showAnalytics(key);});
}

function switchAnalytics(key,btn){
  activeKey=key;
  document.querySelectorAll('.flow-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  if(analyticsCache[key]) renderAnalytics(key);
  else fetchAnalytics(key).then(ok=>{if(ok) renderAnalytics(key);});
}

function renderAnalytics(key){
  const d=analyticsCache[key]; if(!d) return;
  // KPIs
  document.getElementById('kpi-total').textContent=d.total_combos;
  document.getElementById('kpi-clear').textContent=d.clear_combos+' ('+d.clear_pct+'%)';
  document.getElementById('kpi-clear-sub').textContent=
    d.open_combos>0?'of '+d.total_combos+' combinations':'All accounts at zero';
  document.getElementById('kpi-open').textContent=d.open_combos;
  document.getElementById('kpi-docs').textContent=d.total_docs_clearable;
  // Charts
  buildStatusChart(d.by_account);
  buildDocsChart(d.docs_by_account);
  if(d.docs_by_period&&d.docs_by_period.length>0){
    document.getElementById('period-card').style.display='block';
    buildPeriodChart(d.docs_by_period);
  } else {
    document.getElementById('period-card').style.display='none';
  }
}

function dynHeight(rows,rowPx,min,max){
  return Math.min(max,Math.max(min,rows*rowPx+60));
}

function hAxis(){ return {
  grid:{color:P.grid,lineWidth:1},border:{display:false},
  ticks:{color:P.muted,font:{size:11},precision:0}
};}
function vAxis(){ return {
  grid:{display:false},border:{display:false},
  ticks:{color:P.sub,font:{size:11}}
};}

function buildStatusChart(byAcc){
  const wrap=document.getElementById('wrap-status');
  wrap.style.height=dynHeight(byAcc.length,28,160,420)+'px';
  if(cStatus){cStatus.destroy();cStatus=null;}
  cStatus=new Chart(document.getElementById('chart-status'),{
    type:'bar',
    data:{
      labels:byAcc.map(d=>d.account),
      datasets:[
        {label:'Clearable',data:byAcc.map(d=>d.clear),
         backgroundColor:P.clear,maxBarThickness:20,
         borderRadius:{topRight:4,bottomRight:4},borderWidth:2,borderColor:P.surface},
        {label:'Open Items',data:byAcc.map(d=>d.open),
         backgroundColor:P.open,maxBarThickness:20,
         borderRadius:{topRight:4,bottomRight:4},borderWidth:2,borderColor:P.surface}
      ]
    },
    options:{...BASE_OPT,
      indexAxis:'y',
      plugins:{...BASE_OPT.plugins,
        legend:{position:'top',align:'end',
          labels:{boxWidth:12,boxHeight:12,font:{size:11},color:P.sub,padding:14}}},
      scales:{x:{...hAxis(),stacked:true},y:{...vAxis(),stacked:true}}
    }
  });
}

function buildDocsChart(docsAcc){
  const sorted=[...docsAcc].sort((a,b)=>a.count-b.count);
  const wrap=document.getElementById('wrap-docs');
  wrap.style.height=dynHeight(sorted.length,28,160,420)+'px';
  if(cDocs){cDocs.destroy();cDocs=null;}
  cDocs=new Chart(document.getElementById('chart-docs'),{
    type:'bar',
    data:{
      labels:sorted.map(d=>d.account),
      datasets:[{label:'Documents',data:sorted.map(d=>d.count),
        backgroundColor:P.blue,maxBarThickness:20,
        borderRadius:{topRight:4,bottomRight:4},borderWidth:0}]
    },
    options:{...BASE_OPT,
      indexAxis:'y',
      plugins:{...BASE_OPT.plugins,legend:{display:false}},
      scales:{x:hAxis(),y:vAxis()}
    }
  });
}

function buildPeriodChart(byPeriod){
  if(cPeriod){cPeriod.destroy();cPeriod=null;}
  cPeriod=new Chart(document.getElementById('chart-period'),{
    type:'bar',
    data:{
      labels:byPeriod.map(d=>d.ym),
      datasets:[{label:'Documents',data:byPeriod.map(d=>d.count),
        backgroundColor:P.blue,maxBarThickness:24,
        borderRadius:{topLeft:4,topRight:4},borderWidth:0}]
    },
    options:{...BASE_OPT,
      plugins:{...BASE_OPT.plugins,legend:{display:false}},
      scales:{
        x:{grid:{display:false},border:{display:false},ticks:{color:P.muted,font:{size:11}}},
        y:{grid:{color:P.grid,lineWidth:1},border:{display:false},
           ticks:{color:P.muted,font:{size:11},precision:0}}
      }
    }
  });
}

// Auto-load on page load if pivot files exist
window.addEventListener('load',function(){
  ['us','canada','concur'].forEach(k=>{
    fetchAnalytics(k).then(ok=>{if(ok&&!activeKey) showAnalytics(k);});
  });
});
</script>
</body>
</html>"""


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/run/<key>")
def run_script(key):
    if key not in SCRIPTS:
        return "Script no encontrado", 404

    script_path = os.path.join(SCRIPT_DIR, SCRIPTS[key])

    def generate():
        proc = subprocess.Popen(
            [sys.executable, script_path],
            cwd=SCRIPT_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        for line in proc.stdout:
            yield f"data: {line.rstrip()}\n\n"
        proc.wait()
        status = "OK" if proc.returncode == 0 else f"ERROR:{proc.returncode}"
        yield f"data: __DONE__:{status}\n\n"

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/results/<key>")
def get_results(key):
    if key not in OUTPUT_FILES:
        return jsonify([]), 404
    pivot_path = os.path.join(SCRIPT_DIR, OUTPUT_FILES[key])
    if not os.path.exists(pivot_path):
        return jsonify([]), 404
    try:
        def ns(v):
            try:
                f = float(v)
                return str(int(f)) if f == int(f) else str(f)
            except:
                return str(v).strip()

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
        shutil.copy2(pivot_path, tmp)
        try:
            docs_df = pd.read_excel(tmp, sheet_name="Docs a Clearear")
        finally:
            os.unlink(tmp)

        # Detectar columna de importe (puede llamarse "Amount in LC" o "Amount in Local Currency")
        amt_col = next(
            (c for c in docs_df.columns if "amount" in c.lower()),
            None
        )
        has_ym = "Year/Month" in docs_df.columns

        result = []
        for (acc, cc), group in docs_df.groupby(["G/L Account", "Company Code"]):
            doc_list = []
            sort_cols = ["Document Number"] if "Document Number" in group.columns else []
            if sort_cols:
                group = group.sort_values(sort_cols)
            for _, row in group.iterrows():
                raw_amt = row.get(amt_col, 0) if amt_col else 0
                try:
                    amt = round(float(raw_amt), 2)
                except (TypeError, ValueError):
                    amt = 0.0
                doc_list.append({
                    "doc":    str(row.get("Document Number", "")),
                    "amount": amt,
                    "text":   str(row.get("Text", "")) if pd.notna(row.get("Text", "")) else "",
                    "ym":     str(row.get("Year/Month", "")) if has_ym and pd.notna(row.get("Year/Month", "")) else "",
                    "assign": str(row.get("Assignment", "")) if pd.notna(row.get("Assignment", "")) else "",
                })
            result.append({
                "account":   ns(acc),
                "cc":        ns(cc),
                "balance":   0.0,
                "documents": doc_list,
            })
        result.sort(key=lambda x: (x["account"], x["cc"]))
        return jsonify(result)
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


@app.route("/summary/<key>")
def get_summary(key):
    if key not in OUTPUT_FILES:
        return jsonify({"error": "not found"}), 404
    pivot_path = os.path.join(SCRIPT_DIR, OUTPUT_FILES[key])
    if not os.path.exists(pivot_path):
        return jsonify({"error": "no file"}), 404
    try:
        def ns(v):
            try:
                f = float(v)
                return str(int(f)) if f == int(f) else str(f)
            except:
                return str(v).strip()

        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False).name
        shutil.copy2(pivot_path, tmp)
        try:
            pivot_df = pd.read_excel(tmp, sheet_name="Pivot Clearing")
            docs_df  = pd.read_excel(tmp, sheet_name="Docs a Clearear")
        finally:
            os.unlink(tmp)

        # KPI counts
        total     = len(pivot_df)
        clear_ct  = int((pivot_df["Balance"] == 0).sum())
        open_ct   = total - clear_ct
        clear_pct = round(clear_ct / total * 100, 1) if total > 0 else 0.0

        # By-account aggregation for status chart
        pivot_df["_acc"] = pivot_df["G/L Account"].apply(ns)
        by_account = []
        for acc, grp in pivot_df.groupby("_acc"):
            by_account.append({
                "account": acc,
                "clear":   int((grp["Balance"] == 0).sum()),
                "open":    int((grp["Balance"] != 0).sum()),
            })
        by_account.sort(key=lambda x: x["account"])

        # Docs by account (clearable)
        amt_col = next((c for c in docs_df.columns if "amount" in c.lower()), None)
        docs_by_acc = []
        if "G/L Account" in docs_df.columns:
            for acc, grp in docs_df.groupby("G/L Account"):
                amt = 0.0
                if amt_col:
                    try: amt = round(float(grp[amt_col].abs().sum()), 2)
                    except: pass
                docs_by_acc.append({"account": ns(acc), "count": len(grp), "amount": amt})
            docs_by_acc.sort(key=lambda x: x["count"], reverse=True)
            docs_by_acc = docs_by_acc[:15]

        # Docs by period
        docs_by_period = []
        if "Year/Month" in docs_df.columns:
            for ym, grp in sorted(
                docs_df.dropna(subset=["Year/Month"]).groupby("Year/Month"),
                key=lambda x: str(x[0])
            ):
                amt = 0.0
                if amt_col:
                    try: amt = round(float(grp[amt_col].abs().sum()), 2)
                    except: pass
                docs_by_period.append({"ym": str(ym), "count": len(grp), "amount": amt})

        return jsonify({
            "total_combos":         total,
            "clear_combos":         clear_ct,
            "open_combos":          open_ct,
            "clear_pct":            clear_pct,
            "total_docs_clearable": len(docs_df),
            "by_account":           by_account,
            "docs_by_account":      docs_by_acc,
            "docs_by_period":       docs_by_period,
        })
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


if __name__ == "__main__":
    import webbrowser
    print("Iniciando SAP Clearing Dashboard en http://localhost:5001")
    webbrowser.open("http://localhost:5001")
    app.run(host="0.0.0.0", port=5001, threaded=True)