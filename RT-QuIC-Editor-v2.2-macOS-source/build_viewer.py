#!/usr/bin/env python3
"""
RT-QuIC Visualization Builder
Reads Excel file, generates a standalone interactive HTML viewer.
Supports BMG Omega/MARS plate exports and grouped mean curves.
"""
import openpyxl, json, os, re, webbrowser, sys

# ============================================================
# HTML TEMPLATE (minimal, clean, battle-tested)
# ============================================================
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<!-- RTQUIC-ENGINE-V16 -->
<head>
<meta charset="UTF-8">
<title>RT-QuIC Visualizer — __TITLE__</title>
<script src="plotly.min.js"></script>
<style>
html,body{overscroll-behavior:none;}body{font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','Helvetica Neue',Arial,sans-serif;margin:0;padding:16px;background:#ffffff;color:#000;}
body.bare{padding:0;margin:0;background:#fff;overflow:hidden;}
body.bare .sticky-top{display:none;}
body.bare .ov-toolbar{display:none;}
body.bare #tab-overlay{height:100vh;display:flex;align-items:center;justify-content:center;}
body.bare #overlay-wrapper{max-width:96vw;max-height:94vh;}
body.bare .chart-box{border:none;border-radius:0;padding:0;margin:0;min-width:0;min-height:0;}
body.bare #overlay-display-legend{display:none;}
h2{margin:0 0 8px 0;}
.tabs{display:flex;gap:4px;margin-bottom:12px;}
.tab-btn{padding:8px 16px;border:1px solid #ddd;border-radius:6px;cursor:pointer;background:#fff;}
.tab-btn.active{background:#007AFF;color:#fff;border-color:#007AFF;}
.tab-content{display:none;}
.tab-content.active{display:block;}
.chart-box{background:#fff;border-radius:8px;border:1px solid #ddd;padding:8px;margin-bottom:12px;}
.controls{padding:8px 12px;background:#fff;border-radius:8px;border:1px solid #ddd;margin-bottom:12px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;font-size:0.85rem;}
select,input{padding:4px 8px;border:1px solid #ddd;border-radius:6px;font-size:0.85rem;}
.plate-grid{display:grid;gap:4px;}
.plate-well{border:2px solid #e5e7eb;border-radius:6px;cursor:pointer;background:#fafafa;aspect-ratio:1.4;overflow:hidden;position:relative;}
.plate-well:hover{border-color:#007AFF;}
.plate-well.selected{border-color:#007AFF;box-shadow:0 0 0 2px rgba(0,122,255,0.25);}
.plate-well .lbl{position:absolute;top:1px;left:4px;font-size:0.6rem;font-weight:700;color:#555;z-index:2;background:rgba(255,255,255,0.8);padding:0 2px;border-radius:3px;}
.plate-well .smp{position:absolute;top:1px;right:4px;font-size:0.55rem;color:#999;z-index:2;max-width:72%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;background:rgba(255,255,255,0.8);padding:0 2px;border-radius:3px;}
.plate-well svg.spark{position:absolute;top:0;left:0;width:100%;height:100%;z-index:1;}
.plate-well.excluded{opacity:0.45;border-style:dashed;}
.plate-hdr{font-weight:600;color:#888;font-size:0.78rem;display:flex;align-items:center;justify-content:center;}
/* Excel-style plate grid (sparkline sheet look) */
.xl-grid{border:1px solid #b7b7b7;background:#fff;}
.xl-hd{background:#f5f5f5;border-right:1px solid #d9d9d9;border-bottom:1px solid #d9d9d9;font-weight:600;color:#555;font-size:0.8rem;display:flex;align-items:center;justify-content:center;min-height:22px;}
.xl-cell{position:relative;height:90px;background:#fff;border-right:1px solid #d9d9d9;border-bottom:1px solid #d9d9d9;cursor:pointer;overflow:hidden;}
.xl-cell:hover{background:#f6fcf6;}
.xl-cell.excluded{background:#fbfbfb;}
.xl-cell.excluded svg.spark{opacity:0.4;}
.xl-cell svg.spark{position:absolute;top:0;left:0;width:100%;height:100%;}
.well-detail{display:none;background:#fff;border-radius:8px;padding:12px;border:1px solid #ddd;margin-bottom:12px;}
.well-detail.active{display:block;}

.ov-toolbar{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap;}
.ov-toolbar button{padding:4px 10px;font-size:12px;border:1px solid #ccc;border-radius:6px;background:#fff;color:#000;cursor:pointer;}
.ov-toolbar button:hover{background:#f0f0f0;}
.ov-toolbar select{padding:4px 8px;font-size:12px;border:1px solid #ccc;border-radius:6px;background:#fff;color:#000;cursor:pointer;}
.se-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.35);z-index:1000;display:none;align-items:center;justify-content:center;}
.se-panel{background:#fff;border-radius:6px;padding:16px 18px;width:400px;max-width:94vw;box-shadow:0 8px 30px rgba(0,0,0,0.3);color:#000;font-size:13px;max-height:90vh;overflow-y:auto;}
.se-panel h3{margin:0 0 12px;font-size:15px;font-weight:bold;}
.se-row{display:flex;align-items:center;gap:8px;margin-bottom:8px;}
.se-row label{width:84px;font-size:12px;flex-shrink:0;color:#333;}
.se-row input[type=color]{width:42px;height:26px;border:none;padding:0;cursor:pointer;background:none;}
.se-row input[type=number]{width:76px;padding:3px 6px;border:1px solid #ccc;border-radius:6px;}
.se-row select{padding:3px 6px;border:1px solid #ccc;border-radius:6px;background:#fff;}
.se-row input[type=text]{flex:1;padding:3px 6px;border:1px solid #ccc;border-radius:6px;}
.se-row input[type=checkbox]{width:16px;height:16px;}
.se-btns{display:flex;gap:8px;justify-content:flex-end;margin-top:14px;}
.se-btns button{padding:5px 12px;border:1px solid #ccc;border-radius:6px;cursor:pointer;background:#f0f0f0;color:#000;}
.se-btns button.primary{background:#007AFF;color:#fff;border-color:#007AFF;}
.resize-handle{position:absolute;z-index:20;}
.resize-b{left:0;right:0;bottom:-4px;height:8px;cursor:ns-resize;}
.resize-r{top:0;bottom:0;right:-4px;width:8px;cursor:ew-resize;}
#toast{position:fixed;bottom:26px;left:50%;transform:translateX(-50%);background:rgba(0,0,0,0.78);color:#fff;padding:9px 18px;border-radius:6px;font-size:13px;z-index:2000;display:none;pointer-events:none;font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','Helvetica Neue',Arial,sans-serif;}
.sticky-top{position:sticky;top:0;background:#ffffff;z-index:100;padding-bottom:6px;margin-bottom:8px;border-bottom:1px solid #e5e7eb;}
</style>
</head>
<body>
<div class="sticky-top">
<h2>__HEADER__<span style="color:#888;font-size:0.8rem;font-weight:normal;margin-left:10px;">__META__</span></h2>

<div class="controls">
  <label>阈值:</label>
  <input type="range" id="th" min="100" max="5000" value="1000" oninput="updateTh(this.value)">
  <input type="number" id="thn" value="1000" min="100" max="50000" onchange="updateTh(this.value)" style="width:70px;">
  <span>RFU</span>
  <span id="th-scale-hint" style="color:#888;font-size:0.78rem;"></span>
  <button onclick="autoThreshold()" title="按基线均值+3×SD 自动推荐阈值" style="padding:3px 8px;font-size:0.8rem;cursor:pointer;background:#fff;border:1px solid #ccc;border-radius:6px;">✨ 自动阈值</button>
  <button onclick="toggleDispOpts()" title="展开/收起显示选项" style="padding:3px 8px;font-size:0.8rem;cursor:pointer;background:#fff;border:1px solid #ccc;border-radius:6px;">⚙ 显示选项</button>
  <span id="disp-opts" style="display:none;">
    <span>|</span>
    <label>归一化:</label>
    <select id="norm" onchange="redraw()">
      <option value="raw" id="norm-raw-opt">原始 (×10³)</option>
      <option value="firstpt">首点归一</option>
      <option value="zscore">Z-score</option>
      <option value="minmax">Min-Max</option>
      <option value="bgsub">减背景</option>
    </select>
    <label>Y轴单位:</label>
    <select id="scale" onchange="setScale()" title="纵坐标显示倍率（所有 ×10^k 显示跟随变化）">
      <option value="0">×10⁰</option>
      <option value="3" selected>×10³</option>
      <option value="6">×10⁶</option>
    </select>
    <label>着色:</label>
    <select id="color" onchange="redraw()">
      <option value="sample">按样本</option>
      <option value="row">按行</option>
    </select>
    <label>±SD:</label>
    <select id="sd" onchange="sdChanged()" title="复孔均值±SD 显示方式">
      <option value="off">关闭</option>
      <option value="shade">同色阴影</option>
      <option value="tlines">T线</option>
    </select>
    <span id="errtype-wrap" style="display:none;">
      <label>误差:</label>
      <select id="errtype" onchange="redraw()" title="误差棒数值类型">
        <option value="sd" selected>SD</option>
        <option value="sem">SEM</option>
        <option value="ci">95%CI</option>
      </select>
    </span>
    <label>复孔:</label>
    <input type="checkbox" id="rawrep" onchange="redraw()" title="显示复孔原始曲线（虚线）">
    <span id="tstep-wrap" style="display:none;">
      <label>T线抽稀:</label>
      <select id="tstep" onchange="redraw()" title="每 N 个时间点画一根误差棒（仅T线模式）">
        <option value="1">每1点</option>
        <option value="2">每2点</option>
        <option value="3">每3点</option>
        <option value="4" selected>每4点</option>
        <option value="6">每6点</option>
        <option value="8">每8点</option>
        <option value="10">每10点</option>
      </select>
    </span>
  </span>
  <span style="margin-left:auto;color:#888;font-size:0.8rem;" id="info"></span>
</div>

<div class="tabs">
  <button class="tab-btn" onclick="showTab('plate')">🧫 Plate View</button>
  <button class="tab-btn active" onclick="showTab('overlay')">📈 Overlay</button>
  <button class="tab-btn" onclick="showTab('metrics')">📊 Metrics</button>
</div>
</div>

<div id="tab-plate" class="tab-content">
  <div id="plate-caption" style="color:#666;font-size:0.8rem;margin:0 2px 6px;"></div>
  <div class="chart-box"><div id="plate-grid" class="plate-grid"></div></div>
  <div class="well-detail" id="detail">
    <h3 id="detail-title"></h3>
    <div id="detail-chart" style="height:350px;"></div>
    <div id="detail-metrics" style="display:flex;gap:16px;flex-wrap:wrap;margin-top:8px;font-size:0.85rem;"></div>
  </div>
</div>

<div id="tab-overlay" class="tab-content active">
  <div class="ov-toolbar">
    <select id="ov-tools" onchange="ovToolAction(this)" title="编辑工具">
      <option value="">工具 ▾</option>
      <option value="text">➕ 添加文字</option>
      <option value="axistitle">✏️ 坐标标题</option>
      <option value="axisrange">📏 坐标范围</option>
    </select>
    <select id="ov-export" onchange="ovExportAction(this)" title="导出图片">
      <option value="">导出 ▾</option>
      <option value="png">📷 PNG</option>
      <option value="svg">🖼 SVG</option>
    </select>
    <button onclick="openInBrowser()">🌐 Open in Browser</button>
    <button onclick="openCombined()">🖼 拼版图</button>
    <button onclick="location.reload()" title="Refresh (reload file from disk)" style="margin-left:auto;padding:4px 10px;font-size:12px;cursor:pointer;background:#007AFF;color:#fff;border:none;border-radius:6px;font-weight:600;">🔄 Refresh</button>
  </div>
  <div class="chart-box" id="overlay-chart" style="min-height:500px;"></div>
</div>

<div id="tab-metrics" class="tab-content">
  <div class="chart-box" id="metrics-table-box" style="padding:12px;margin-bottom:12px;"><h3 style="margin:0 0 8px 0;">📊 Per-Well Metrics</h3><div id="metrics-table" style="overflow-x:auto;"></div></div>
  <div class="chart-box" id="metrics-bar-box"><div id="metrics-bar-chart" style="min-height:400px;"></div></div>
  <div class="chart-box" id="metrics-titer-box" style="display:none;"><div id="metrics-titer-chart" style="min-height:350px;"></div></div>
  <div class="chart-box" id="metrics-survival-box"><div id="metrics-survival-chart" style="min-height:350px;"></div></div>
  <div class="chart-box" id="metrics-heatmap-box"><div id="metrics-heatmap-chart" style="min-height:400px;"></div></div>
</div>


<div id="toast"></div>

<div class="se-overlay" id="style-editor">
  <div class="se-panel">
    <h3>曲线样式 Curve Style — <span id="se-name"></span></h3>
    <div class="se-row"><label>颜色 Color</label><input type="color" id="se-color" value="#000000"></div>
    <div class="se-row"><label>粗细 Width</label><input type="number" id="se-width" min="0.5" max="12" step="0.5" value="1.8"></div>
    <div class="se-row"><label>线型 Line</label><select id="se-dash">
      <option value="solid">实线 Solid</option>
      <option value="dot">圆点线 Dot</option>
      <option value="dash">短虚线 Dash</option>
      <option value="longdash">长虚线 Long dash</option>
      <option value="dashdot">长短交替 Dash-dot</option>
      <option value="longdashdot">长-点交替 Long dash-dot</option>
    </select></div>
    <div class="se-row"><label>标记形状 Symbol</label><select id="se-symbol">
      <option value="circle">圆 Circle</option><option value="square">方 Square</option>
      <option value="diamond">菱形 Diamond</option><option value="triangle-up">上三角 Up</option>
      <option value="triangle-down">倒三角 Down</option><option value="cross">十字 Cross</option>
      <option value="x">叉 X</option><option value="star">星 Star</option>
      <option value="hexagon">六边形 Hexagon</option><option value="pentagon">五边形 Pentagon</option>
    </select></div>
    <div class="se-row"><label>标记大小 Size</label><input type="number" id="se-size" min="1" max="24" step="0.5" value="6"></div>
    <div class="se-row"><label>标记填充 Fill</label><select id="se-fill">
      <option value="filled">实心 Filled</option><option value="hollow">空心 Hollow</option>
    </select></div>
    <div class="se-btns"><button onclick="closeModal('style-editor')">取消 Cancel</button><button class="primary" onclick="applyCurveStyle()">应用 Apply</button></div>
  </div>
</div>

<div class="se-overlay" id="anno-editor">
  <div class="se-panel">
    <h3>文字样式 Text Style</h3>
    <div class="se-row"><label>文字 Text</label><input type="text" id="anno-text"></div>
    <div class="se-row"><label>位置 X</label><input type="number" id="anno-x" min="0" max="1" step="0.01" value="0.5"></div>
    <div class="se-row"><label>位置 Y</label><input type="number" id="anno-y" min="0" max="1" step="0.01" value="0.92"></div>
    <div class="se-row"><label>字体 Font</label><select id="anno-family">
      <option value="Arial">Arial</option><option value="Times New Roman">Times New Roman</option>
      <option value="Courier New">Courier New</option><option value="Helvetica">Helvetica</option>
      <option value="Georgia">Georgia</option>
    </select></div>
    <div class="se-row"><label>字号 Size</label><input type="number" id="anno-size" min="6" max="72" step="1" value="16"></div>
    <div class="se-row"><label>颜色 Color</label><input type="color" id="anno-color" value="#000000"></div>
    <div class="se-row"><label>加粗 Bold</label><input type="checkbox" id="anno-bold"></div>
    <div class="se-btns">
      <button onclick="deleteAnnotation()" style="color:#c00;">删除 Delete</button>
      <button onclick="closeModal('anno-editor')">取消 Cancel</button>
      <button class="primary" onclick="applyAnnoEditor()">应用 Apply</button>
    </div>
  </div>
</div>

<div class="se-overlay" id="axis-editor">
  <div class="se-panel">
    <h3>坐标标题 Axis Titles</h3>
    <div class="se-row"><label>横轴 X</label><input type="text" id="axis-x"></div>
    <div class="se-row"><label>纵轴 Y</label><input type="text" id="axis-y"></div>
    <div class="se-row"><label>字体 Font</label><select id="axis-family">
      <option value="Arial">Arial</option><option value="Times New Roman">Times New Roman</option>
      <option value="Courier New">Courier New</option><option value="Helvetica">Helvetica</option>
      <option value="Georgia">Georgia</option>
    </select></div>
    <div class="se-row"><label>字号 Size</label><input type="number" id="axis-size" min="6" max="72" step="1" value="13"></div>
    <div class="se-row"><label>颜色 Color</label><input type="color" id="axis-color" value="#000000"></div>
    <div class="se-row"><label>加粗 Bold</label><input type="checkbox" id="axis-bold"></div>
    <div class="se-btns"><button onclick="closeModal('axis-editor')">取消 Cancel</button><button class="primary" onclick="applyAxisTitles()">应用 Apply</button></div>
  </div>
</div>

<div class="se-overlay" id="axis-range-editor">
  <div class="se-panel">
    <h3>坐标范围 Axis Ranges</h3>
    <div class="se-row"><label>X 最小 Min</label><input type="number" id="ar-xmin" step="any" placeholder="自动"></div>
    <div class="se-row"><label>X 最大 Max</label><input type="number" id="ar-xmax" step="any" placeholder="自动"></div>
    <div class="se-row"><label>Y 最小 Min</label><input type="number" id="ar-ymin" step="any" placeholder="自动"></div>
    <div class="se-row"><label>Y 最大 Max</label><input type="number" id="ar-ymax" step="any" placeholder="自动"></div>
    <div class="se-btns">
      <button onclick="resetAxisRange()">自动 Auto</button>
      <button onclick="closeModal('axis-range-editor')">取消 Cancel</button>
      <button class="primary" onclick="applyAxisRange()">应用 Apply</button>
    </div>
  </div>
</div>

<div class="se-overlay" id="combined-modal">
  <div class="se-panel" style="width:880px;max-width:96vw;">
    <h3>组合图 Combined</h3>
    <div id="combined-chart" style="height:640px;"></div>
    <div class="se-btns">
      <button onclick="exportCombinedPng()">📷 导出PNG</button>
      <button class="primary" onclick="closeModal('combined-modal')">关闭 Close</button>
    </div>
  </div>
</div>

<script>
var D = __DATA__;
var BARE = (window.location.search||'').indexOf('bare') >= 0;
if(BARE){ document.body.className += ' bare'; }
var threshold = 1000, selectedWell = null, allM = {};

function normVals(w){
  var v=w.values, m=document.getElementById('norm').value;
  if(m==='raw')return v;
  if(m==='firstpt'){var f0=v[0]||1;return v.map(function(x){return x/f0;});}
  if(m==='zscore'){var mn=0,sd=0;v.forEach(function(x){mn+=x;});mn/=v.length||1;v.forEach(function(x){sd+=(x-mn)*(x-mn);});sd=Math.sqrt(sd/(v.length||1));return v.map(function(x){return sd? (x-mn)/sd : 0;});}
  if(m==='minmax'){var mn2=Math.min.apply(null,v),mx=Math.max.apply(null,v),r=mx-mn2||1;return v.map(function(x){return(x-mn2)/r;});}
  var bg=v.slice(0,5).reduce(function(a,b){return a+b;},0)/5;
  return v.map(function(x){return x-bg;});
}
// Display values in ×10^scaleExp units for raw/bgsub; 0-1 for minmax
var scaleExp = 3;
function scaleFactor(){ return Math.pow(10, scaleExp); }
function scaleTag(){ var s='⁰¹²³⁴⁵⁶⁷⁸⁹',e=String(scaleExp),r=''; for(var i=0;i<e.length;i++){r+=s[parseInt(e[i],10)]||e[i];} return '×10'+r; }
function scaleSuffix(){ return scaleExp===6?'M':(scaleExp===3?'K':''); }
function syncScaleLabel(){
  var o=document.getElementById('norm-raw-opt');
  if(o) o.textContent='原始 ('+scaleTag()+')';
}
function setScale(){
  scaleExp=parseInt(document.getElementById('scale').value,10)||3;
  syncScaleLabel();
  updateThHint();
  redraw();
}
function dispVals(w){
  var v=w.values, m=document.getElementById('norm').value;
  if(m==='firstpt'){var f0=v[0]||1;return v.map(function(x){return x/f0;});}
  if(m==='zscore'){var mn=0,sd=0;v.forEach(function(x){mn+=x;});mn/=v.length||1;v.forEach(function(x){sd+=(x-mn)*(x-mn);});sd=Math.sqrt(sd/(v.length||1));return v.map(function(x){return sd? (x-mn)/sd : 0;});}
  if(m==='minmax'){var mn2=Math.min.apply(null,v),mx=Math.max.apply(null,v),r=mx-mn2||1;return v.map(function(x){return(x-mn2)/r;});}
  var bg=v.slice(0,5).reduce(function(a,b){return a+b;},0)/5;
  return v.map(function(x){return m==='bgsub'?(x-bg)/scaleFactor():x/scaleFactor();});
}
function kRFU(x){ return (x/scaleFactor()).toFixed(2); }
function yTitle(){
  var m=document.getElementById('norm').value;
  if(m==='firstpt') return 'Normalized (first point = 1)';
  if(m==='zscore') return 'Z-score';
  if(m==='minmax') return 'Normalized fluorescence';
  return 'THT fluorescence ('+scaleTag()+')';
}

function calcM(w){
  var v=w.values, t=D.time_hours;
  var mx=Math.max.apply(null,v), mxi=v.indexOf(mx);
  var ttt=null;
  for(var i=0;i<v.length;i++){if(v[i]>=threshold){ttt=t[i];break;}}
  // t50: time to reach 50% of max fluorescence (linear interpolation)
  var t50=null, halfMx=mx/2;
  for(var i=0;i<v.length;i++){if(v[i]>=halfMx){if(i>0&&v[i-1]<halfMx){t50=t[i-1]+(t[i]-t[i-1])*(halfMx-v[i-1])/(v[i]-v[i-1]);}else{t50=t[i];}break;}}
  var ms=0,mst=0,msv=0;
  for(var i=0;i<v.length-5;i++){var dy=v[i+5]-v[i],dx=t[i+5]-t[i],s=dy/dx;if(s>ms){ms=s;mst=t[i+2];msv=v[i+2];}}
  var bg=v.slice(0,5).reduce(function(a,b){return a+b;},0)/5;
  var mpr=bg>0?mx/bg:0;
  var pos=ttt!==null&&mx>threshold*1.2;
  var auc=0;
  for(var i=1;i<v.length;i++){var y0=Math.max(0,v[i-1]-bg),y1=Math.max(0,v[i]-bg);auc+=(y0+y1)/2*(t[i]-t[i-1]);}
  var lag=null;
  if(ms>0&&msv>bg){var lg=mst-(msv-bg)/ms;if(lg>=0&&lg<=mst){lag=lg;}}
  var plt=null,pc=0,ps=0;
  for(var i=0;i<v.length;i++){if(v[i]>=0.95*mx){ps+=v[i];pc++;}}
  if(pc>0)plt=ps/pc;
  return {maxRFU:mx,maxTime:t[mxi],ttt:ttt,t50:t50,maxSlope:ms,maxSlopeTime:mst,bg:bg,mpr:mpr,auc:auc,lag:lag,plateau:plt,isPositive:pos};
}
function calcAll(){D.wells.forEach(function(w){allM[w.well]=calcM(w);}); for(var k in plateById){allM[k]=calcM(plateById[k]);}}

// ===== PLATE VIEW (96-well layout, raw per-well curves, unified axes) =====
var plateWells=[], plateById={}, plateByPos={};
var ROWIDX={}; (function(){for(var i=0;i<26;i++)ROWIDX[String.fromCharCode(65+i)]=i;})();
function plateLabel(w){ return (w.row||'?')+(w.col!=null?w.col:''); }
function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function buildPlateWells(){
  plateWells=[]; plateById={}; plateByPos={};
  function add(w){
    if(!w.values||!w.values.length) return;
    if(!allM[w.id]) allM[w.id]=calcM(w);
    plateWells.push(w); plateById[w.id]=w;
    if(w.row&&w.col) plateByPos[w.row+''+w.col]=w;
  }
  if(D.raws&&D.raws.length){
    D.raws.forEach(function(r,i){
      var pw={id:'R'+(i+1), sample:r.group, values:r.values.slice(0,D.time_hours.length), excluded:!!r.excluded, raw:true};
      var g=String(r.group||'').trim();
      var m=g.match(/^([A-Ha-h])(\d{1,2})$/);
      if(m){ pw.row=m[1].toUpperCase(); pw.col=parseInt(m[2],10); }
      else { pw.row=String.fromCharCode(65+Math.floor(i/12)); pw.col=(i%12)+1; }
      add(pw);
    });
  } else {
    D.wells.forEach(function(w0){
      add({id:w0.well, sample:w0.sample, values:w0.values, excluded:false, raw:false, row:w0.row_letter, col:w0.col_num});
    });
  }
}
function plateYRange(){
  var mn=Infinity, mx=-Infinity;
  plateWells.forEach(function(w){
    if(w.excluded) return;
    var v=dispVals(w);
    for(var i=0;i<v.length;i++){ if(v[i]<mn)mn=v[i]; if(v[i]>mx)mx=v[i]; }
  });
  if(!isFinite(mn)){ mn=0; mx=1; }
  if(mx===mn){ mx=mn+1; }
  var pad=(mx-mn)*0.05;
  return [mn-pad, mx+pad];
}
function sparkSVG(w, xr, yr){
  var v=dispVals(w), t=D.time_hours;
  var W=100, H=60, x0=xr[0], x1=xr[1], y0=yr[0], y1=yr[1];
  var dx=(x1-x0)||1, dy=(y1-y0)||1, pts='';
  for(var i=0;i<v.length;i++){
    var px=((t[i]-x0)/dx*W), py=(H-(v[i]-y0)/dy*H);
    pts+=px.toFixed(2)+','+py.toFixed(2)+' ';
  }
  var col=w.excluded?'#b0b0b0':'#00A651';
  var dash=w.excluded?' stroke-dasharray="3,2"':'';
  return '<svg class="spark" viewBox="0 0 '+W+' '+H+'" preserveAspectRatio="none"><polyline points="'+pts+'" fill="none" stroke="'+col+'" stroke-width="1.5" vector-effect="non-scaling-stroke"'+dash+'/></svg>';
}
function drawPlate(){
  var grid=document.getElementById('plate-grid');
  if(!plateWells.length){ grid.innerHTML='<div style="color:#999;padding:20px;">暂无数据</div>'; return; }
  var maxRow=8, maxCol=12, hasExcl=false;
  plateWells.forEach(function(w){
    var ri=ROWIDX[w.row]; if(ri!==undefined&&ri+1>maxRow)maxRow=ri+1;
    if(w.col>maxCol)maxCol=w.col;
    if(w.excluded)hasExcl=true;
  });
  var xr=[0, D.time_hours[D.time_hours.length-1]||66], yr=plateYRange();
  grid.className='xl-grid';
  grid.style.display='grid';
  grid.style.gridTemplateColumns='40px repeat('+maxCol+', minmax(0,1fr))';
  grid.style.gap='0';
  var h='<div class="xl-hd"></div>';
  for(var c=1;c<=maxCol;c++) h+='<div class="xl-hd">'+c+'</div>';
  for(var r=0;r<maxRow;r++){
    var rl=String.fromCharCode(65+r);
    h+='<div class="xl-hd">'+rl+'</div>';
    for(var c=1;c<=maxCol;c++){
      var w=plateByPos[rl+''+c];
      if(!w){ h+='<div class="xl-cell"></div>'; continue; }
      var m=allM[w.id];
      var tip=esc(rl+''+c+' — '+w.sample+(w.excluded?'（已排除）':'')+' | Pk:'+kRFU(m.maxRFU)+scaleSuffix()+' TtT:'+(m.ttt!==null?m.ttt.toFixed(1)+'h':'—'));
      h+='<div class="xl-cell'+(w.excluded?' excluded':'')+'" id="w-'+w.id+'" title="'+tip+'" onclick="selWell(\''+w.id+'\')">'+sparkSVG(w,xr,yr)+'</div>';
    }
  }
  grid.innerHTML=h;
  grid.setAttribute('data-rows', String(maxRow));
  var cap=document.getElementById('plate-caption');
  if(cap) cap.innerHTML='🧫 96孔板 · 每格显示该孔原始曲线（未取平均）· 所有格统一坐标范围'+(hasExcl?' · 灰色虚线=已排除':'');
  fitPlate();
}
function fitPlate(){
  var grid=document.getElementById('plate-grid');
  if(!grid||!plateWells.length) return;
  var rows=parseInt(grid.getAttribute('data-rows'),10)||8;
  var rect=grid.getBoundingClientRect();
  if(rect.width<10) return;
  var avail=window.innerHeight - Math.max(rect.top,0) - 14;
  if(avail<200) avail=200;
  var hd=grid.querySelector('.xl-hd');
  var hdH=hd?hd.getBoundingClientRect().height:22;
  var ch=Math.floor((avail-hdH-rows)/rows);
  if(!isFinite(ch)||ch<36)ch=36;
  if(ch>150)ch=150;
  var cells=grid.querySelectorAll('.xl-cell');
  for(var i=0;i<cells.length;i++)cells[i].style.height=ch+'px';
}
window.addEventListener('resize', fitPlate);

function selWell(wid){
  // Restore previous selection
  if(selectedWell){
    var prev=document.getElementById('w-'+selectedWell);
    if(prev){
      prev.style.outline='';
      prev.style.boxShadow='';
      prev.style.background='';
      prev.style.zIndex='';
    }
  }
  selectedWell = selectedWell===wid ? null : wid;
  var selW=plateById[selectedWell];
  document.getElementById('info').textContent = selectedWell ? ('已选: '+(selW?(plateLabel(selW)+' · '+selW.sample):wid)) : '';

  if(selectedWell){
    var el=document.getElementById('w-'+selectedWell);
    if(el){
      el.style.outline='2px solid #000';
      el.style.outlineOffset='-2px';
      el.style.boxShadow='';
      el.style.background='';
      el.style.zIndex='';
    }
  }

  var dt=document.getElementById('detail');
  if(selectedWell){dt.classList.add('active');drawDetail(selectedWell);}
  else dt.classList.remove('active');
}

function drawDetail(wid){
  var pw=plateById[wid];
  var w=pw||D.wells.filter(function(x){return x.well===wid;})[0], m=allM[wid];
  if(!w||!m) return;
  document.getElementById('detail-title').textContent=(pw&&pw.raw)?(plateLabel(w)+' — '+w.sample+(w.excluded?'（已排除）':'')):(wid+' — '+w.sample);
  var tMax=D.time_hours[D.time_hours.length-1]||66;
  Plotly.newPlot('detail-chart',[{x:D.time_hours,y:dispVals(w),type:'scatter',mode:'lines+markers',line:{width:2.5,color:m.isPositive?'#00A651':'#D00000'},marker:{symbol:'circle',size:6,color:m.isPositive?'#00A651':'#D00000',line:{width:0}},hoverlabel:{namelength:-1}}],{margin:{l:50,r:20,t:10,b:40},font:{family:'Arial', color:'#000'},hoverlabel:{namelength:-1,font:{family:'Arial',size:12}},xaxis:{title:'<b>Time (h)</b>',titlefont:{family:'Arial', size:18, color:'#000'},showgrid:false,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false},yaxis:{title:'<b>'+yTitle()+'</b>',titlefont:{family:'Arial', size:18, color:'#000'},showgrid:false,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false}},{responsive:true,displayModeBar:false});
  document.getElementById('detail-metrics').innerHTML=
    '<div><b>峰值:</b> '+kRFU(m.maxRFU)+' '+scaleTag()+' RFU @ '+m.maxTime.toFixed(1)+'h</div>'+
    '<div><b>Lag Time:</b> '+(m.lag!==null?m.lag.toFixed(1)+' h':'—')+'</div>'+
    '<div><b>TtT:</b> '+(m.ttt!==null?m.ttt.toFixed(1)+' h':'未达')+'</div>'+
    '<div><b>t₅₀:</b> '+(m.t50!==null?m.t50.toFixed(1)+' h':'—')+'</div>'+
    '<div><b>AUC:</b> '+kRFU(m.auc)+' '+scaleTag()+' RFU·h</div>'+
    '<div><b>最大斜率:</b> '+kRFU(m.maxSlope)+' '+scaleTag()+' RFU/h</div>'+
    '<div><b>MPR:</b> '+m.mpr.toFixed(1)+'× | <b>背景:</b> '+kRFU(m.bg)+'</div>'+
    '<div><b>判定:</b> <span style="color:'+(m.isPositive?'#00A651':'#D00000')+';font-weight:600">'+(m.isPositive?'✅ 阳性':'❌ 阴性')+'</span></div>';
  makeResizable('detail-chart');
}

// ===== MEAN±SD (shade band or T-line error bars) =====
function errType(){ var e=document.getElementById('errtype'); return e?e.value:'sd'; }
function errValue(w, j){
  var sd=w.sd[j];
  if(sd==null) return 0;
  var m=document.getElementById('norm').value;
  var n=w.n_replicates||1;
  var e=errType();
  var base = e==='sem' ? sd/Math.sqrt(n) : (e==='ci' ? 1.96*sd/Math.sqrt(n) : sd);
  if(m==='minmax'){var mn=Math.min.apply(null,w.values),mx=Math.max.apply(null,w.values),r=mx-mn||1;return base/r;}
  if(m==='firstpt'){var f0=w.values[0]||1;return base/Math.abs(f0);}
  if(m==='zscore'){var v=w.values,mm=0,ss=0;v.forEach(function(x){mm+=x;});mm/=v.length||1;v.forEach(function(x){ss+=(x-mm)*(x-mm);});ss=Math.sqrt(ss/(v.length||1));return ss? base/ss : 0;}
  return base/scaleFactor();
}
function sdMode(){
  var sdEl=document.getElementById('sd');
  return sdEl ? sdEl.value : 'off';
}
function syncSdControls(){
  var wrap=document.getElementById('tstep-wrap');
  if(wrap) wrap.style.display=(sdMode()==='tlines')?'inline':'none';
  var ew=document.getElementById('errtype-wrap');
  if(ew) ew.style.display=(sdMode()==='off')?'none':'inline';
}
function sdChanged(){
  syncSdControls();
  redraw();
}
function hexToRgba(hex, a){
  if(!hex||hex.length<7) return 'rgba(128,128,128,'+a+')';
  var r=parseInt(hex.slice(1,3),16),g=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16);
  if(isNaN(r)) return 'rgba(128,128,128,'+a+')';
  return 'rgba('+r+','+g+','+b+','+a+')';
}
function buildSdTraces(){
  var mode=sdMode();
  sdTraceIdx = {};
  if(mode==='off') return [];
  var base=overlayTraces.length;
  var out=[];
  overlayTraces.forEach(function(t, i){
    var w=D.wells[t._idx];
    if(!w || !w.sd) return;
    var parentVis = (t.visible!==false && t.visible!=='legendonly') ? true : 'legendonly';
    var idxs=[];
    if(mode==='shade'){
      var up=t.x.map(function(x,j){return (t.y[j]||0)+errValue(w,j);});
      var lo=t.x.map(function(x,j){return (t.y[j]||0)-errValue(w,j);});
      idxs.push(base+out.length);
      out.push({x:t.x, y:up, type:'scatter', mode:'lines', line:{width:0}, visible:parentVis, hoverinfo:'none', showlegend:false, _band:true});
      idxs.push(base+out.length);
      out.push({x:t.x, y:lo, type:'scatter', mode:'lines', line:{width:0}, fill:'tonexty', fillcolor:hexToRgba(t._color,0.25), visible:parentVis, hoverinfo:'none', showlegend:false, _band:true});
    } else if(mode==='tlines'){
      // Reduce cross-curve overlap: subsample + horizontal dodge + transparency.
      // - error bars only at ~1/step of the time points
      // - each curve's bars shifted slightly in x (within one time step)
      // - bars drawn semi-transparent so any residual overlap stays readable
      var step=parseInt((document.getElementById('tstep')||{}).value,10)||4;
      step=Math.max(1,step);
      var nC=overlayTraces.length;
      var dt=(t.x.length>1)?(t.x[t.x.length-1]-t.x[0])/(t.x.length-1):1;
      var dodge=(nC>1)?(i-(nC-1)/2)*(0.8*dt)/Math.max(1,nC-1):0;
      var xs=[], ys=[], ers=[];
      for(var j=0;j<t.x.length;j+=step){
        xs.push(t.x[j]+dodge); ys.push(t.y[j]); ers.push(errValue(w,j));
      }
      idxs.push(base+out.length);
      out.push({
        x:xs, y:ys, type:'scatter', mode:'lines',
        line:{width:0, color:t._color},
        error_y:{type:'data', array:ers, visible:true, symmetric:true, thickness:1, width:4, color:hexToRgba(t._color,0.7)},
        visible:parentVis, hoverinfo:'none', showlegend:false, _band:true
      });
    }
    sdTraceIdx[i]=idxs;
  });
  return out;
}
function restyleSdVisibility(){
  var idxs=[], vis=[];
  overlayTraces.forEach(function(t, i){
    var v = (t.visible!==false && t.visible!=='legendonly') ? true : 'legendonly';
    var sd = sdTraceIdx[i];
    if(sd && sd.length){ sd.forEach(function(pi){ idxs.push(pi); vis.push(v); }); }
    var rw = rawTraceIdx[i];
    if(rw && rw.length){ rw.forEach(function(pi){ idxs.push(pi); vis.push(v); }); }
  });
  if(idxs.length) Plotly.restyle('overlay-chart', {visible: vis}, idxs);
}

// ===== RAW REPLICATE CURVES =====
function buildRawTraces(sdCount){
  rawTraceIdx = {};
  if(!D.raws || !D.raws.length) return [];
  var cb=document.getElementById('rawrep');
  if(!cb || !cb.checked) return [];
  var base=overlayTraces.length + (sdCount||0);
  var out=[];
  D.raws.forEach(function(r){
    var t=null, ti=-1;
    for(var k=0;k<overlayTraces.length;k++){
      var w0=D.wells[overlayTraces[k]._idx];
      if(w0 && w0.sample===r.group){ t=overlayTraces[k]; ti=k; break; }
    }
    if(!t) return;
    var parentVis = (t.visible!==false && t.visible!=='legendonly') ? true : 'legendonly';
    out.push({
      x:D.time_hours, y:dispVals({values:r.values}), type:'scatter', mode:'lines',
      line:{width:1, color:t._color, dash:r.excluded?'dot':'dash'}, opacity:0.45,
      visible:parentVis,
      hoverinfo:'none', showlegend:false, _raw:true
    });
    (rawTraceIdx[ti] = rawTraceIdx[ti] || []).push(base + out.length - 1);
  });
  return out;
}

// ===== OVERLAY =====
var overlayTraces = [], overlayVis = [], overlayVisState = null, sdTraceIdx = {}, rawTraceIdx = {};
function visKey(){ try { return 'rtquic_vis_' + (window.location.pathname.split('/').pop() || 'view'); } catch(e){ return 'rtquic_vis'; } }
function saveVisState(){
  if(!overlayTraces.length) return;
  var st = {};
  overlayTraces.forEach(function(t){ st[t.name] = t.visible; });
  try { localStorage.setItem(visKey(), JSON.stringify(st)); } catch(e){}
}
function loadVisState(){
  try {
    var s = localStorage.getItem(visKey());
    if(s){ var o = JSON.parse(s); return (o && typeof o === 'object') ? o : null; }
  } catch(e){}
  return null;
}
function drawOverlay(){
  // 1. Setup flex layout BEFORE Plotly renders
  var chartEl = document.getElementById('overlay-chart');
  var wrapper = document.getElementById('overlay-wrapper');
  if (!wrapper) {
    var parent = chartEl.parentElement;
    wrapper = document.createElement('div');
    wrapper.id = 'overlay-wrapper';
    wrapper.style.cssText = 'display:flex;gap:8px;align-items:stretch;' + (BARE ? 'flex-direction:column;' : '');
    parent.insertBefore(wrapper, chartEl);
    wrapper.appendChild(chartEl);
    // Display legend INSIDE the chart box (overlaid on the right, within the frame)
    if(!BARE){
      var dispEl = document.createElement('div');
      dispEl.id = 'overlay-display-legend';
      dispEl.style.cssText = 'position:absolute;top:50%;right:10px;transform:translateY(-50%);width:150px;max-height:80%;font-size:11px;overflow-y:auto;z-index:15;';
      chartEl.appendChild(dispEl);
    }
    // Legend: right column (normal) or horizontal bar ABOVE the chart (bare mode)
    var legendEl = document.createElement('div');
    legendEl.id = 'overlay-legend';
    if(BARE){
      legendEl.style.cssText = 'display:flex;flex-direction:row;flex-wrap:wrap;gap:4px 12px;padding:6px 10px;background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;font-size:11px;width:100%;box-sizing:border-box;';
      wrapper.insertBefore(legendEl, chartEl);
    } else {
      legendEl.style.cssText = 'display:flex;flex-direction:column;gap:6px;padding:8px 10px;background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;font-size:10px;max-height:520px;overflow-y:auto;width:230px;flex-shrink:0;';
      wrapper.appendChild(legendEl);
    }
  }
  // Chart fills remaining space; the display legend is inside this box
  chartEl.style.position = 'relative';
  chartEl.style.flex = '1';
  chartEl.style.minWidth = '0';

  // Save current visibility so toggles persist across redraws and view switches
  if (overlayTraces.length > 0) {
    overlayVisState = {};
    overlayTraces.forEach(function(t){ overlayVisState[t.name] = t.visible; });
    saveVisState();
  } else if (!overlayVisState) {
    var saved = loadVisState();
    if (saved) overlayVisState = saved;
  }

  // 2. Build traces
  var cm=document.getElementById('color').value;
  var PRISM_COLORS=['#000000','#0050B0','#D00000','#00A651','#FF8000','#7030A0','#008080','#8B5A2B','#C000C0','#708090','#006B8F','#B8860B','#4B0082','#2E8B57','#8B0000','#C71585'];
  var rows={A:PRISM_COLORS[0],B:PRISM_COLORS[1],C:PRISM_COLORS[2],D:PRISM_COLORS[3],E:PRISM_COLORS[4],F:PRISM_COLORS[5],G:PRISM_COLORS[6],H:PRISM_COLORS[7],I:PRISM_COLORS[8],J:PRISM_COLORS[9],K:PRISM_COLORS[10],L:PRISM_COLORS[11],M:PRISM_COLORS[12],N:PRISM_COLORS[13],O:PRISM_COLORS[14],P:PRISM_COLORS[15]};
  var sc={},sci=0;
  function gc(s){if(!sc[s]){sc[s]=PRISM_COLORS[sci%PRISM_COLORS.length];sci++;}return sc[s];}
  var markers=['circle','square','diamond','cross','triangle-up','triangle-down','star','x','triangle-left','triangle-right','hexagon','pentagon','bowtie','hourglass'];
  var nm=document.getElementById('norm').value;
  var yFmt = (nm==='minmax'||nm==='firstpt'||nm==='zscore') ? '%{y:.2f}' : '%{y:.1f}'+scaleSuffix();
  var mi=0;
  overlayTraces = D.wells.map(function(w, idx){
    var cl=cm==='row'?(rows[w.row_letter]||'#888'):gc(w.sample);
    var mk=markers[mi % markers.length]; mi++;
    var n=D.time_hours.length;
    var step=Math.max(1,Math.floor(n/40));
    var nv=dispVals(w);
    return {
      x:D.time_hours, y:nv, type:'scatter', mode:'lines+markers',
      line:{width:1.8, color:cl},
      marker:{symbol:mk, size:6, color:cl, line:{width:0}},
      name:(D.isGrouped ? w.sample : w.well+' ('+w.sample.replace('Sample ','')+')'),
      showlegend:false, hovertemplate:'<b>%{data.name}</b><br>Time: %{x:.1f} h<br>THT fluorescence: '+yFmt+'<extra></extra>',
      _color:cl, _marker:mk, _idx:idx, _hollow:false
    };
  });
  overlayVis = overlayTraces.map(function(t){ return true; });
  // Restore previous visibility state, or apply default (first 8 visible)
  if (overlayVisState) {
    overlayTraces.forEach(function(t) {
      if (overlayVisState[t.name] !== undefined) t.visible = overlayVisState[t.name];
    });
  } else {
    overlayTraces.forEach(function(t, i) { if (i >= 8) t.visible = 'legendonly'; });
  }

  // 3. Render Plotly into the already-flexed chart div
  var tMax=D.time_hours[D.time_hours.length-1]||67;
  var sdTraces = buildSdTraces();
  var plotData = overlayTraces.concat(sdTraces).concat(buildRawTraces(sdTraces.length));
  Plotly.newPlot('overlay-chart', plotData, {
    title:'Mean Fluorescence Kinetics (markers = groups)',
    xaxis:{title:'<b>Time (h)</b>', titlefont:{family:'Arial', size:18, color:'#000'}, showgrid:false, showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false, range:[0,tMax+1]},
    yaxis:{title:'<b>'+yTitle()+'</b>', titlefont:{family:'Arial', size:18, color:'#000'}, showgrid:false, showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false},
    hovermode:'closest',
    hoverlabel:{namelength:-1, font:{family:'Arial', size:12, color:'#333'}, bordercolor:'#ccc', bgcolor:'#fff'},
    showlegend:false,
    font:{family:'Arial', color:'#000'},
    margin:{l:40, r:BARE?8:175, t:20, b:40}
  }, {responsive:true, displaylogo:false, editable:true, edits:{annotationPosition:false, annotationTail:false, annotationText:false, axisTitleText:true, colorbarPosition:false, colorbarTitleText:false, legendPosition:false, legendText:false, shapePosition:false, titleText:true}});
  applyUserRanges();

  // 4. Populate legend
  buildCustomLegend();
  applyAnnotations();
  attachOverlayEvents();
  makeResizable('overlay-chart');
}


// Inline SVG marker paths (filled shapes use fill; line-based shapes use stroke)
var markerSVGs = {
  'circle':       '<circle cx="6" cy="6" r="4" fill="C"/>',
  'square':       '<rect x="2" y="2" width="8" height="8" fill="C"/>',
  'diamond':      '<polygon points="6,1 11,6 6,11 1,6" fill="C"/>',
  'triangle-up':  '<polygon points="6,1.5 11,10.5 1,10.5" fill="C"/>',
  'triangle-down':'<polygon points="1,1.5 11,1.5 6,10.5" fill="C"/>',
  'triangle-left':'<polygon points="1.5,6 10.5,1 10.5,11" fill="C"/>',
  'triangle-right':'<polygon points="10.5,6 1.5,1 1.5,11" fill="C"/>',
  'star':         '<polygon points="6,1 7.6,4.5 11.4,4.9 8.7,7.6 9.3,11.4 6,9.5 2.7,11.4 3.3,7.6 0.6,4.9 4.4,4.5" fill="C"/>',
  'hexagon':      '<polygon points="6,1 10.5,3.5 10.5,8.5 6,11 1.5,8.5 1.5,3.5" fill="C"/>',
  'pentagon':     '<polygon points="6,1 11,4.5 9.5,10 2.5,10 1,4.5" fill="C"/>',
  'cross':        '<path d="M3,3 L9,9 M9,3 L3,9" fill="none" stroke="C" stroke-width="1.8"/>',
  'x':            '<path d="M3,3 L9,9 M9,3 L3,9" fill="none" stroke="C" stroke-width="1.8"/>',
  'bowtie':       '<path d="M1,1 L6,6 L11,1 L11,11 L6,6 L1,11 Z" fill="C" fill-opacity="0.35" stroke="C" stroke-width="1.2"/>',
  'hourglass':    '<path d="M2,1 L10,1 L6,6 L10,11 L2,11 L6,6 Z" fill="C" fill-opacity="0.35" stroke="C" stroke-width="1.2"/>'
};
function markerSVG(marker, color, hollow){
  var raw = markerSVGs[marker] || markerSVGs['circle'];
  if (hollow) { raw = raw.replace(/fill="C"/g, 'fill="none" stroke="C" stroke-width="1.3"'); }
  return raw.replace(/C/g, color);
}

function buildSimpleLegend(){
  var container = document.getElementById('overlay-legend');
  if(!container) return;
  var h='<span style="font-weight:700;font-size:11px;margin-right:8px;">图例 Legend</span>';
  overlayTraces.forEach(function(t){
    var vis = (t.visible!==false && t.visible!=='legendonly');
    if(!vis) return;
    var svg = (markerSVGs[t._marker] || markerSVGs['circle']).replace(/C/g, t._color);
    h += '<span style="display:inline-flex;align-items:center;gap:4px;margin-right:10px;font-size:11px;">'
       + '<svg width="12" height="12" viewBox="0 0 12 12">'+svg+'</svg>'
       + '<span style="width:14px;height:2px;background:'+t._color+';display:inline-block;"></span>'
       + '<span>'+t.name+'</span></span>';
  });
  container.innerHTML=h;
}
function buildCustomLegend() {
  if(BARE){ buildSimpleLegend(); return; }
  var wrapper = document.getElementById('overlay-wrapper');
  if (!wrapper) {
    // Wrap chart + legend in a flex row
    var chartEl = document.getElementById('overlay-chart');
    var parent = chartEl.parentElement;
    wrapper = document.createElement('div');
    wrapper.id = 'overlay-wrapper';
    wrapper.style.cssText = 'display:flex;gap:8px;align-items:stretch;';
    parent.insertBefore(wrapper, chartEl);
    wrapper.appendChild(chartEl);
    chartEl.style.flex = '1';
    chartEl.style.minWidth = '0';
  }
  var container = document.getElementById('overlay-legend');
  if (!container) {
    container = document.createElement('div');
    container.id = 'overlay-legend';
    container.style.cssText = 'display:flex;flex-direction:column;gap:6px;padding:8px 10px;background:#fafafa;border:1px solid #e5e7eb;border-radius:8px;font-size:10px;max-height:500px;overflow-y:auto;min-width:200px;max-width:320px;flex-shrink:0;';
    wrapper.appendChild(container);
  }
  // Natural sort: split names into text/number segments, compare numerically
  function naturalParts(name) {
    var parts = [], re = /(\d+(?:\.\d+)?)/g, li = 0, m;
    while ((m = re.exec(name)) !== null) {
      if (m.index > li) parts.push(name.slice(li, m.index));
      parts.push(parseFloat(m[1]));
      li = re.lastIndex;
    }
    if (li < name.length) parts.push(name.slice(li));
    return parts;
  }
  function naturalCompare(a, b) {
    var pa = naturalParts(a), pb = naturalParts(b);
    for (var i = 0; i < Math.min(pa.length, pb.length); i++) {
      var na = typeof pa[i] === 'number', nb = typeof pb[i] === 'number';
      if (na && nb && pa[i] !== pb[i]) return pa[i] - pb[i];
      if (na !== nb) return na ? -1 : 1;
      if (String(pa[i]) !== String(pb[i])) return String(pa[i]) < String(pb[i]) ? -1 : 1;
    }
    return pa.length - pb.length;
  }
  var sortedIndices = overlayTraces.map(function(t, i) { return i; });
  sortedIndices.sort(function(a, b) {
    return naturalCompare(overlayTraces[a].name, overlayTraces[b].name);
  });
  var allShapes = ['circle','square','diamond','triangle-up','triangle-down','cross','x','star','triangle-left','triangle-right','hexagon','pentagon','bowtie','hourglass'];
  // Mini SVG icons for shape picker (16x16 viewBox for better proportions)
  var shapeIcons = {
    'circle':       '<circle cx="8" cy="8" r="5.5" fill="C"/>',
    'square':       '<rect x="2.5" y="2.5" width="11" height="11" fill="C"/>',
    'diamond':      '<polygon points="8,1.5 14.5,8 8,14.5 1.5,8" fill="C"/>',
    'triangle-up':  '<polygon points="8,2 15,14 1,14" fill="C"/>',
    'triangle-down':'<polygon points="1,2 15,2 8,14" fill="C"/>',
    'triangle-left':'<polygon points="2,8 14,1 14,15" fill="C"/>',
    'triangle-right':'<polygon points="14,8 2,1 2,15" fill="C"/>',
    'star':         '<polygon points="8,1.5 10,6 15,6.5 11.5,10 12.5,15 8,12.5 3.5,15 4.5,10 1,6.5 6,6" fill="C"/>',
    'hexagon':      '<polygon points="8,1.5 14,4.8 14,11.2 8,14.5 2,11.2 2,4.8" fill="C"/>',
    'pentagon':     '<polygon points="8,1.5 15,6 13,14 3,14 1,6" fill="C"/>',
    'cross':        '<path d="M4,4 L12,12 M12,4 L4,12" fill="none" stroke="C" stroke-width="2.2"/>',
    'x':            '<path d="M4,4 L12,12 M12,4 L4,12" fill="none" stroke="C" stroke-width="2.2"/>',
    'bowtie':       '<path d="M1.5,1.5 L8,8 L14.5,1.5 L14.5,14.5 L8,8 L1.5,14.5 Z" fill="C" fill-opacity="0.35" stroke="C" stroke-width="1.2"/>',
    'hourglass':    '<path d="M3,1.5 L13,1.5 L8,8 L13,14.5 L3,14.5 L8,8 Z" fill="C" fill-opacity="0.35" stroke="C" stroke-width="1.2"/>'
  };

  // Legend header with Select All / Deselect All
  var h = '<div style="font-weight:bold;font-size:12px;margin-bottom:4px;">显示开关 Show/Hide</div>';
  h += '<div style="display:flex;gap:4px;margin-bottom:6px;padding-bottom:6px;border-bottom:1px solid #e5e7eb;">';
  h += '<button onclick="showAllTraces()" style="flex:1;padding:2px 6px;font-size:10px;font-weight:600;cursor:pointer;background:#007AFF;color:#fff;border:none;border-radius:3px;" onmouseover="this.style.background=\'#0071e3\'" onmouseout="this.style.background=\'#007AFF\'">Select All</button>';
  h += '<button onclick="deselectAllTraces()" style="flex:1;padding:2px 6px;font-size:10px;font-weight:600;cursor:pointer;background:#f0f0f0;color:#666;border:1px solid #ddd;border-radius:3px;" onmouseover="this.style.background=\'#e0e0e0\'" onmouseout="this.style.background=\'#f0f0f0\'">Deselect</button>';
  h += '</div>';

  sortedIndices.forEach(function(i) {
    var t = overlayTraces[i];
    var vis = t.visible !== false && t.visible !== 'legendonly';
    var fw = vis ? '900' : '400';
    var op = vis ? '1' : '0.4';
    var sz = vis ? 14 : 11;
    var svg = markerSVG(t._marker, t._color, t._hollow);

    // Legend item + inline edit (wrapped for absolute editor positioning)
    h += '<div style="position:relative;">';
    h += '<div style="display:flex;align-items:center;gap:2px;">';

    // Main clickable legend item
    h += '<span onclick="handleLegendClick(event,'+i+')" style="flex:1;cursor:pointer;opacity:'+op+';font-weight:'+fw+';font-size:12px;white-space:nowrap;user-select:none;display:inline-flex;align-items:center;gap:4px;padding:3px 5px;border-radius:3px;line-height:1.4;min-width:0;" onmouseover="this.style.background=\'#e8e8e8\'" onmouseout="this.style.background=\'transparent\'">';
    h += '<svg width="'+sz+'" height="'+sz+'" viewBox="0 0 12 12" style="flex-shrink:0;">'+svg+'</svg>';
    h += '<span style="overflow:hidden;text-overflow:ellipsis;">'+t.name+'</span>';
    h += '</span>';

    // Edit button for this item
    h += '<span onclick="event.stopPropagation();openStyleEditor('+i+')" style="cursor:pointer;font-size:11px;padding:2px 3px;border-radius:3px;flex-shrink:0;color:#999;" title="Edit style" onmouseover="this.style.background=\'#e0e0e0\';this.style.color=\'#333\'" onmouseout="this.style.background=\'transparent\';this.style.color=\'#999\'">✎</span>';

    h += '</div>';  // close item row

    // Inline editor popup (absolute, hidden by default)
    h += '<div id="editor-'+i+'" style="display:none;position:absolute;left:0;top:100%;z-index:100;align-items:center;gap:4px;padding:4px 6px;background:#fff;border:1px solid #d5d8de;border-radius:6px;font-size:10px;box-shadow:0 4px 12px rgba(0,0,0,0.15);white-space:nowrap;">';
    h += '<input type="color" value="'+t._color+'" onchange="updateTraceStyle('+i+',this.value,null)" style="width:22px;height:20px;border:none;cursor:pointer;padding:0;border-radius:3px;">';
    h += '<div style="display:flex;gap:2px;flex-wrap:wrap;max-width:196px;">';
    allShapes.forEach(function(s) {
      var icon = (shapeIcons[s] || shapeIcons['circle']).replace(/C/g, t._color);
      var sel = t._marker === s ? '2px solid #007AFF' : '1px solid #ccc';
      h += '<span onclick="updateTraceStyle('+i+',null,\''+s+'\')" style="cursor:pointer;border:'+sel+';border-radius:2px;padding:1px;background:#fff;" title="'+s+'">';
      h += '<svg width="16" height="16" viewBox="0 0 16 16">'+icon+'</svg>';
      h += '</span>';
    });
    h += '</div>';  // close shape grid
    h += '</div>';  // close editor popup
    h += '</div>';  // close position:relative wrapper
  });
  container.innerHTML = h;
  buildDisplayLegend();
}



function buildDisplayLegend(){
  var container = document.getElementById('overlay-display-legend');
  if(!container) return;
  var vis = overlayTraces.filter(function(t){ return t.visible !== false && t.visible !== 'legendonly'; });
  var h = '';
  if(vis.length === 0){
    h += '<div style="color:#999;font-size:11px;">（当前无显示曲线）</div>';
  } else {
    vis.forEach(function(t){
      var svg = markerSVG(t._marker, t._color, t._hollow);
      h += '<div style="display:flex;align-items:center;gap:6px;padding:2px 0;font-size:11px;color:#000;">';
      h += '<svg width="14" height="14" viewBox="0 0 12 12" style="flex-shrink:0;">'+svg+'</svg>';
      h += '<span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'+t.name+'</span>';
      h += '</div>';
    });
  }
  container.innerHTML = h;
}
var legendClickTimer = null, legendLastIdx = -1;
function handleLegendClick(ev, i) {
  if (legendClickTimer && legendLastIdx === i) {
    // Double-click: isolate this trace
    clearTimeout(legendClickTimer);
    legendClickTimer = null;
    legendLastIdx = -1;
    isolateTrace(i);
    return;
  }
  legendLastIdx = i;
  legendClickTimer = setTimeout(function() {
    legendClickTimer = null;
    legendLastIdx = -1;
    toggleOverlayTrace(i);
  }, 300);
}

function isolateTrace(i) {
  overlayTraces.forEach(function(t, j) {
    t.visible = (j === i) ? true : 'legendonly';
  });
  Plotly.restyle('overlay-chart', {
    visible: overlayTraces.map(function(t) { return t.visible; })
  });
  restyleSdVisibility();
  buildCustomLegend();
  saveVisState();
}

function showAllTraces() {
  overlayTraces.forEach(function(t) {
    t.visible = true;
  });
  Plotly.restyle('overlay-chart', {
    visible: overlayTraces.map(function(t) { return true; })
  });
  restyleSdVisibility();
  buildCustomLegend();
  saveVisState();
}

function deselectAllTraces() {
  overlayTraces.forEach(function(t) {
    t.visible = 'legendonly';
  });
  Plotly.restyle('overlay-chart', {
    visible: overlayTraces.map(function(t) { return 'legendonly'; })
  });
  restyleSdVisibility();
  buildCustomLegend();
  saveVisState();
}

function toggleItemEditor(ev, idx) {
  var panel = document.getElementById('editor-'+idx);
  if (!panel) return;
  var isOpening = panel.style.display === 'none' || panel.style.display === '';
  // Close all other open editors
  overlayTraces.forEach(function(_, j) {
    var ed = document.getElementById('editor-'+j);
    if (ed && j !== idx) ed.style.display = 'none';
  });
  if (isOpening) {
    // Smart position: above if near bottom, below otherwise
    var itemRow = panel.previousElementSibling;
    var legendEl = document.getElementById('overlay-legend');
    if (itemRow && legendEl) {
      var ir = itemRow.getBoundingClientRect();
      var lr = legendEl.getBoundingClientRect();
      if (ir.bottom + 140 > lr.bottom) {
        panel.style.top = 'auto'; panel.style.bottom = '100%';
      } else {
        panel.style.top = '100%'; panel.style.bottom = 'auto';
      }
    }
    panel.style.display = 'flex';
    // Click outside to close
    setTimeout(function() {
      document.addEventListener('click', function closeEditor(e) {
        if (!panel.contains(e.target) && e.target !== ev.target) {
          panel.style.display = 'none';
          document.removeEventListener('click', closeEditor);
        }
      });
    }, 50);
  } else {
    panel.style.display = 'none';
  }
}

function updateTraceStyle(idx, color, shape) {
  var gd = document.getElementById('overlay-chart');
  var t = overlayTraces[idx];
  t.line = t.line || {};
  t.marker = t.marker || {};
  if (color) {
    t._color = color;
    t.line.color = color;
    t.marker.color = color;
  }
  if (shape) {
    t._marker = shape;
    t.marker.symbol = shape;
  }
  var update = {};
  if (color) { update['line.color'] = color; update['marker.color'] = color; }
  if (shape) { update['marker.symbol'] = shape; }
  Plotly.restyle(gd, update, [idx]);
  buildCustomLegend();
}

function toggleOverlayTrace(i) {
  var currentlyVis = overlayTraces[i].visible !== false && overlayTraces[i].visible !== 'legendonly';
  overlayTraces[i].visible = currentlyVis ? 'legendonly' : true;
  Plotly.restyle('overlay-chart', {visible: overlayTraces[i].visible}, [i]);
  restyleSdVisibility();
  buildCustomLegend();
  saveVisState();
}


// ===== MODAL HELPERS =====
function showModal(id){ var m=document.getElementById(id); if(m) m.style.display='flex'; }
function closeModal(id){ var m=document.getElementById(id); if(m) m.style.display='none'; }
function showToast(msg){
  var el=document.getElementById('toast');
  if(!el) return;
  el.textContent=msg; el.style.display='block';
  clearTimeout(el._t);
  el._t=setTimeout(function(){ el.style.display='none'; }, 2400);
}
function openInBrowser(){ window.open(window.location.href, '_blank'); }
function openCombined(){
  var tr=[];
  D.wells.forEach(function(w){
    tr.push({x:D.time_hours, y:dispVals(w), type:'scatter', mode:'lines+markers',
      name:w.sample, line:{width:1.5}, marker:{size:4}, showlegend:false, hoverinfo:'skip'});
  });
  var names=D.wells.map(function(w){return w.sample;});
  var ttt=D.wells.map(function(w){var m=calcM(w); return m.ttt;});
  var auc=D.wells.map(function(w){var m=calcM(w); return m.auc/scaleFactor();});
  tr.push({x:names, y:ttt, type:'bar', name:'TtT', xaxis:'x2', yaxis:'y2', marker:{color:'#007AFF'}});
  tr.push({x:names, y:auc, type:'bar', name:'AUC', xaxis:'x3', yaxis:'y3', marker:{color:'#D00000'}});
  var dls={};
  D.wells.forEach(function(w){var m=/^(\d+(?:\.\d+)?)\s*%/.exec(w.sample); if(m){var p=parseFloat(m[1]); if(!dls[p])dls[p]={t:0,pos:0}; dls[p].t++; if(allM[w.well].isPositive)dls[p].pos++;}});
  var pks=Object.keys(dls).map(Number).sort(function(a,b){return b-a;});
  if(pks.length>=2){
    tr.push({x:pks.map(function(p){return p+'%';}), y:pks.map(function(p){return dls[p].t?dls[p].pos/dls[p].t*100:0;}), type:'bar', name:'%Positive', xaxis:'x4', yaxis:'y4', marker:{color:'#00A651'}});
  } else {
    var pt=[];
    D.wells.forEach(function(w){var m=allM[w.well]; if(m.isPositive&&m.ttt!=null) pt.push(m.ttt);});
    pt.sort(function(a,b){return a-b;});
    tr.push({x:pt, y:pt.map(function(_,i){return (i+1)/Math.max(1,pt.length)*100;}), type:'scatter', mode:'lines', name:'Cumulative %', xaxis:'x4', yaxis:'y4', line:{color:'#00A651', shape:'hv'}});
  }
  Plotly.newPlot('combined-chart', tr, {
    title:'Combined View',
    font:{family:'Arial', size:11},
    // 四宫格用显式 domain 定位（替代 grid 自动布局）：左右间隔 35%，上下间隔 22%，
    // 上排图的 X 轴标题落在行间空隙内，不会与下排曲线重叠。
    xaxis:{title:'Time (h)', domain:[0, 0.325]}, yaxis:{title:yTitle(), domain:[0.60, 1]},
    xaxis2:{title:'Sample', domain:[0.675, 1]}, yaxis2:{title:'TtT (h)', domain:[0.60, 1]},
    xaxis3:{title:'Sample', domain:[0, 0.325]}, yaxis3:{title:'AUC ('+scaleTag()+'·h)', side:'right', domain:[0, 0.38]},
    xaxis4:{title:'Dilution', domain:[0.675, 1]}, yaxis4:{title:'% Positive', range:[0,105], domain:[0, 0.38]},
    margin:{l:50,r:70,t:40,b:50}
  }, {responsive:true, displaylogo:false});
  showModal('combined-modal');
}
function exportCombinedPng(){
  var gd=document.getElementById('combined-chart');
  if(!gd || !gd._fullLayout) return;
  Plotly.toImage(gd, {format:'png', width:1500, height:1100, scale:2}).then(function(url){
    var a=document.createElement('a'); a.href=url; a.download='combined.png';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  });
}
function exportChart(fmt){
  var gd=document.getElementById('overlay-chart');
  if(!gd || !gd._fullLayout) return;
  var f = (fmt==='svg') ? 'svg' : 'png';
  var scale = (fmt==='svg') ? 1 : 3;
  var idx = overlayTraces.map(function(_,i){ return i; });
  // Export file name matches the view file name
  var fname;
  try {
    var p = window.location.pathname.split('/').pop() || '';
    p = p.replace(/\.html?$/i, '');
    fname = (p || 'rtquic_chart') + '.' + f;
  } catch(e){ fname = 'rtquic_chart.' + f; }
  // Legend in the exported image: only the VISIBLE (selected) curves
  var showMap = overlayTraces.map(function(t){ return (t.visible!==false && t.visible!=='legendonly'); });
  showToast('正在导出 '+(f==='svg'?'SVG':'PNG')+' ...');
  Plotly.relayout(gd, {showlegend:true, legend:{x:1.02, y:0.5, xanchor:'left', yanchor:'middle', font:{family:'Arial', size:12}}})
    .then(function(){ return Plotly.restyle(gd, {showlegend:showMap}, idx); })
    .then(function(){ return Plotly.toImage(gd, {format:f, width:gd._fullLayout.width, height:gd._fullLayout.height, scale:scale}); })
    .then(function(url){
      var a=document.createElement('a');
      a.href=url;
      a.download=fname;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      showToast('✅ 导出完成: '+fname);
      // restore original legend-hidden state
      return Plotly.relayout(gd, {showlegend:false}).then(function(){ return Plotly.restyle(gd, {showlegend:false}, idx); });
    })
    .catch(function(e){ showToast('导出失败: '+(e && e.message ? e.message : e)); });
}

// ===== RESIZE HANDLES: drag to change axis length (task 4) =====
function makeResizable(boxId){
  var box=document.getElementById(boxId);
  if(!box || box.querySelector('.resize-handle')) return;
  box.style.position='relative';
  var bh=document.createElement('div'); bh.className='resize-handle resize-b'; box.appendChild(bh);
  var rh=document.createElement('div'); rh.className='resize-handle resize-r'; box.appendChild(rh);
  function bind(handle, axis){
    handle.addEventListener('mousedown', function(e){
      e.preventDefault(); e.stopPropagation();
      var sx=e.clientX, sy=e.clientY;
      var w0=box.getBoundingClientRect().width, h0=box.getBoundingClientRect().height;
      var resizeRAF=null;
      function mv(ev){
        if(axis==='x'){
          var nw=Math.max(240, w0+(ev.clientX-sx));
          box.style.width=nw+'px'; box.style.flex='0 0 '+nw+'px';
        } else {
          var nh=Math.max(180, h0+(ev.clientY-sy));
          box.style.height=nh+'px';
        }
        if(!resizeRAF){
          resizeRAF=requestAnimationFrame(function(){ resizeRAF=null; Plotly.Plots.resize(box); });
        }
      }
      function up(){ document.removeEventListener('mousemove',mv); document.removeEventListener('mouseup',up); }
      document.addEventListener('mousemove',mv); document.addEventListener('mouseup',up);
    });
  }
  bind(bh,'y'); bind(rh,'x');
}

// ===== DOUBLE-CLICK + CURVE STYLE EDITOR (task 3) =====
var styleEditIdx=-1;
var plotClickTimer=null, plotClickTrace=-1;
function onPlotClick(ev){
  if(!ev || !ev.points || !ev.points.length) return;
  var ti=ev.points[0].curveNumber;
  if(ti==null || ti>=overlayTraces.length) return;
  if(plotClickTimer && plotClickTrace===ti){
    clearTimeout(plotClickTimer); plotClickTimer=null; plotClickTrace=-1;
    openStyleEditor(ti);
  } else {
    plotClickTrace=ti;
    plotClickTimer=setTimeout(function(){ plotClickTimer=null; plotClickTrace=-1; }, 280);
  }
}
function onAnnoClick(ev){
  if(!ev) return;
  var idx = ev.index;
  if(idx!=null && idx>=0 && idx<textAnnotations.length){
    editAnnoIndex=idx;
    openAnnoEditor(idx);
  }
}
function onPlotContextMenu(ev){
  var t = ev.target;
  while(t && t !== document){
    if(t.getAttribute){
      var cls = t.getAttribute('class') || '';
      var di = t.getAttribute('data-index');
      if(di != null && cls.indexOf('annotation') !== -1){
        var idx = parseInt(di, 10);
        if(idx >= 0 && idx < textAnnotations.length){
          ev.preventDefault();
          editAnnoIndex = idx;
          openAnnoEditor(idx);
        }
        return;
      }
      if(cls.indexOf('xtitle') !== -1 || cls.indexOf('ytitle') !== -1){
        ev.preventDefault();
        openAxisTitleEditor();
        return;
      }
    }
    t = t.parentNode;
  }
}
document.getElementById('overlay-chart').addEventListener('contextmenu', onPlotContextMenu);
function attachOverlayEvents(){
  var gd=document.getElementById('overlay-chart');
  if(!gd || !gd.on) return;
  gd.on('plotly_click', onPlotClick);
  gd.on('plotly_clickannotation', onAnnoClick);
}
function openStyleEditor(i){
  if(i<0 || i>=overlayTraces.length) return;
  styleEditIdx=i;
  var t=overlayTraces[i];
  document.getElementById('se-name').textContent=t.name;
  document.getElementById('se-color').value=t._color||'#000000';
  document.getElementById('se-width').value=(t.line&&t.line.width)||1.8;
  document.getElementById('se-dash').value=(t.line&&t.line.dash)||'solid';
  document.getElementById('se-symbol').value=t._marker||'circle';
  document.getElementById('se-size').value=(t.marker&&t.marker.size)||6;
  document.getElementById('se-fill').value=t._hollow?'hollow':'filled';
  showModal('style-editor');
}
function applyCurveStyle(){
  if(styleEditIdx<0) return;
  var t=overlayTraces[styleEditIdx];
  var color=document.getElementById('se-color').value;
  var width=parseFloat(document.getElementById('se-width').value)||1.8;
  var dash=document.getElementById('se-dash').value;
  var symbol=document.getElementById('se-symbol').value;
  var size=parseFloat(document.getElementById('se-size').value)||6;
  var hollow=document.getElementById('se-fill').value==='hollow';
  t._color=color; t._marker=symbol; t._hollow=hollow;
  t.line=t.line||{}; t.marker=t.marker||{};
  t.line.color=color; t.line.width=width; t.line.dash=dash;
  t.marker.symbol=symbol; t.marker.size=size;
  var upd={'line.color':color,'line.width':width,'line.dash':dash,'marker.symbol':symbol,'marker.size':size};
  if(hollow){
    upd['marker.color']='rgba(0,0,0,0)';
    upd['marker.line.color']=color;
    upd['marker.line.width']=1.5;
  } else {
    upd['marker.color']=color;
    upd['marker.line.width']=0;
  }
  Plotly.restyle('overlay-chart', upd, [styleEditIdx]);
  buildCustomLegend();
  closeModal('style-editor');
}

// ===== TEXT ANNOTATIONS (task 1) =====
var textAnnotations=[], editAnnoIndex=-1;
function applyAnnotations(){
  var anns=textAnnotations.map(function(a){
    var txt=a.text;
    if(a.font && a.font.bold) txt='<b>'+txt+'</b>';
    return {text:txt, x:a.x, y:a.y, xref:'paper', yref:'paper', showarrow:a.showarrow||false,
            captureevents:true,
            font:{family:a.font.family, size:a.font.size, color:a.font.color}};
  });
  Plotly.relayout('overlay-chart', {annotations: anns});
}
function addTextAnnotation(){
  var a={text:'Text '+(textAnnotations.length+1), x:0.5, y:0.92, xref:'paper', yref:'paper', showarrow:false,
         font:{family:'Arial', size:16, color:'#000000', bold:false}};
  textAnnotations.push(a);
  applyAnnotations();
  editAnnoIndex=textAnnotations.length-1;
  openAnnoEditor(editAnnoIndex);
}
function openAnnoEditor(i){
  if(i<0||i>=textAnnotations.length) return;
  editAnnoIndex=i;
  var a=textAnnotations[i];
  document.getElementById('anno-text').value=a.text;
  document.getElementById('anno-x').value=a.x;
  document.getElementById('anno-y').value=a.y;
  document.getElementById('anno-family').value=a.font.family||'Arial';
  document.getElementById('anno-size').value=a.font.size||16;
  document.getElementById('anno-color').value=a.font.color||'#000000';
  document.getElementById('anno-bold').checked=!!a.font.bold;
  showModal('anno-editor');
}
function applyAnnoEditor(){
  if(editAnnoIndex<0) return;
  var a=textAnnotations[editAnnoIndex];
  a.text=document.getElementById('anno-text').value||'Text';
  a.x=parseFloat(document.getElementById('anno-x').value);
  a.y=parseFloat(document.getElementById('anno-y').value);
  if(isNaN(a.x)) a.x=0.5;
  if(isNaN(a.y)) a.y=0.5;
  a.font={family:document.getElementById('anno-family').value, size:parseFloat(document.getElementById('anno-size').value)||16,
          color:document.getElementById('anno-color').value, bold:document.getElementById('anno-bold').checked};
  applyAnnotations();
  closeModal('anno-editor');
}
function deleteAnnotation(){
  if(editAnnoIndex<0) return;
  textAnnotations.splice(editAnnoIndex,1);
  editAnnoIndex=-1;
  applyAnnotations();
  closeModal('anno-editor');
}

// ===== AXIS TITLE EDITOR (task 1) =====
function openAxisTitleEditor(){
  var gd=document.getElementById('overlay-chart');
  var la=gd._fullLayout||{};
  var xt=(la.xaxis&&la.xaxis.title&&la.xaxis.title.text)||'Time (h)';
  var yt=(la.yaxis&&la.yaxis.title&&la.yaxis.title.text)||yTitle();
  var xb=xt.indexOf('<b>')===0 && xt.slice(-4)==='</b>';
  var yb=yt.indexOf('<b>')===0 && yt.slice(-4)==='</b>';
  if(xb) xt=xt.slice(3,-4);
  if(yb) yt=yt.slice(3,-4);
  document.getElementById('axis-x').value=xt;
  document.getElementById('axis-y').value=yt;
  document.getElementById('axis-family').value='Arial';
  document.getElementById('axis-size').value=18;
  document.getElementById('axis-color').value='#000000';
  document.getElementById('axis-bold').checked=xb||yb;
  showModal('axis-editor');
}
function applyAxisTitles(){
  var fam=document.getElementById('axis-family').value;
  var size=parseFloat(document.getElementById('axis-size').value)||13;
  var color=document.getElementById('axis-color').value;
  var bold=document.getElementById('axis-bold').checked;
  var xt=document.getElementById('axis-x').value;
  var yt=document.getElementById('axis-y').value;
  Plotly.relayout('overlay-chart', {
    'xaxis.title.text': bold?'<b>'+xt+'</b>':xt,
    'yaxis.title.text': bold?'<b>'+yt+'</b>':yt,
    'xaxis.title.font': {family:fam, size:size, color:color},
    'yaxis.title.font': {family:fam, size:size, color:color}
  });
  closeModal('axis-editor');
}

// ===== AXIS RANGE EDITOR (user-set x/y ranges) =====
var axisRange = {xmin:null, xmax:null, ymin:null, ymax:null};
function applyUserRanges(){
  var gd=document.getElementById('overlay-chart');
  if(!gd) return;
  var u={};
  if(axisRange.xmin!==null) u['xaxis.range[0]']=axisRange.xmin;
  if(axisRange.xmax!==null) u['xaxis.range[1]']=axisRange.xmax;
  if(axisRange.ymin!==null) u['yaxis.range[0]']=axisRange.ymin;
  if(axisRange.ymax!==null) u['yaxis.range[1]']=axisRange.ymax;
  if(Object.keys(u).length) Plotly.relayout(gd, u);
}
function openAxisRangeEditor(){
  document.getElementById('ar-xmin').value = axisRange.xmin!==null ? axisRange.xmin : '';
  document.getElementById('ar-xmax').value = axisRange.xmax!==null ? axisRange.xmax : '';
  document.getElementById('ar-ymin').value = axisRange.ymin!==null ? axisRange.ymin : '';
  document.getElementById('ar-ymax').value = axisRange.ymax!==null ? axisRange.ymax : '';
  showModal('axis-range-editor');
}
function applyAxisRange(){
  function num(v){ var f=parseFloat(v); return isNaN(f)?null:f; }
  axisRange.xmin = num(document.getElementById('ar-xmin').value);
  axisRange.xmax = num(document.getElementById('ar-xmax').value);
  axisRange.ymin = num(document.getElementById('ar-ymin').value);
  axisRange.ymax = num(document.getElementById('ar-ymax').value);
  closeModal('axis-range-editor');
  redraw();
}
function resetAxisRange(){
  axisRange = {xmin:null, xmax:null, ymin:null, ymax:null};
  closeModal('axis-range-editor');
  redraw();
}

// ===== METRICS (comprehensive multi-panel) =====
var metricsBarDrawn=false;
function drawMetrics(){
  var ws=D.wells;
  // ---- 1. Per-Well Metrics Table ----
  var th='<table style="width:100%;border-collapse:collapse;font-size:0.82rem;font-family:-apple-system,BlinkMacSystemFont,Helvetica Neue,Arial,sans-serif;">';
  th+='<tr style="background:#f0f1f5;font-weight:900;"><th style="padding:6px 10px;text-align:left;">Sample</th><th style="padding:6px 10px;cursor:help;" title="成核滞后期：最大斜率点切线外推到基线的时间">Lag Time (h)</th><th style="padding:6px 10px;cursor:help;" title="达阈时间：荧光首次超过阈值的时间">TtT (h)</th><th style="padding:6px 10px;cursor:help;" title="达到最大荧光一半所需时间">t₅₀ (h)</th><th style="padding:6px 10px;cursor:help;" title="曲线下面积：背景扣除后信号的积分，综合反映总扩增活性">AUC ('+scaleTag()+'·h)</th><th style="padding:6px 10px;cursor:help;" title="最大荧光值">Max ('+scaleTag()+')</th><th style="padding:6px 10px;cursor:help;" title="最大斜率：荧光上升最快时的速率">Slope ('+scaleTag()+'/h)</th><th style="padding:6px 10px;">Result</th></tr>';
  ws.forEach(function(w){
    var m=allM[w.well];
    th+='<tr style="border-bottom:1px solid #e5e7eb;">';
    th+='<td style="padding:6px 10px;font-weight:600;">'+w.sample+'</td>';
    th+='<td style="padding:6px 10px;text-align:center;">'+(m.lag!==null?m.lag.toFixed(1):'—')+'</td>';
    th+='<td style="padding:6px 10px;text-align:center;">'+(m.ttt!==null?m.ttt.toFixed(1):'—')+'</td>';
    th+='<td style="padding:6px 10px;text-align:center;">'+(m.t50!==null?m.t50.toFixed(1):'—')+'</td>';
    th+='<td style="padding:6px 10px;text-align:center;">'+kRFU(m.auc)+'</td>';
    th+='<td style="padding:6px 10px;text-align:center;">'+kRFU(m.maxRFU)+'</td>';
    th+='<td style="padding:6px 10px;text-align:center;">'+kRFU(m.maxSlope)+'</td>';
    th+='<td style="padding:6px 10px;text-align:center;"><span style="color:'+(m.isPositive?'#00A651':'#D00000')+';">'+(m.isPositive?'✅ Positive':'❌ Negative')+'</span></td>';
    th+='</tr>';
  });
  th+='</table>';
  document.getElementById('metrics-table').innerHTML=th;

  // ---- 2. Group Comparison Bar Chart ----
  if(!metricsBarDrawn){
    var selHtml='<label style="font-weight:600;margin-right:8px;">Metric:</label><select id="metric-select" onchange="drawMetricsBar()" style="padding:4px 8px;border:1px solid #ddd;border-radius:6px;font-size:0.85rem;">';
    selHtml+='<option value="t50">t₅₀ (h)</option>';
    selHtml+='<option value="ttt">TtT (h)</option>';
    selHtml+='<option value="lag">Lag Time (h)</option>';
    selHtml+='<option value="auc">AUC ('+scaleTag()+'·h)</option>';
    selHtml+='<option value="maxRFU">Max ('+scaleTag()+')</option>';
    selHtml+='<option value="maxSlope">Slope ('+scaleTag()+'/h)</option>';
    selHtml+='</select>';
    var barBox=document.getElementById('metrics-bar-box');
    var selDiv=document.createElement('div');
    selDiv.style.cssText='padding:8px 12px;';
    selDiv.innerHTML=selHtml;
    barBox.insertBefore(selDiv, barBox.firstChild);
    metricsBarDrawn=true;
  }
  drawMetricsBar();
  drawTitration();
  drawSurvival();
  drawHeatmap();
}

function drawMetricsBar(){
  var metric=(document.getElementById('metric-select')||{}).value||'t50';
  var ws=D.wells;
  var names=ws.map(function(w){return D.isGrouped?w.sample:w.well+' ('+w.sample+')';});
  var vals=ws.map(function(w){var m=allM[w.well];if(metric==='t50')return m.t50;if(metric==='ttt')return m.ttt;if(metric==='lag')return m.lag;if(metric==='auc')return m.auc/scaleFactor();if(metric==='maxRFU')return m.maxRFU/scaleFactor();if(metric==='maxSlope')return m.maxSlope/scaleFactor();});
  var colors=ws.map(function(w){return allM[w.well].isPositive?'#00A651':'#D00000';});
  var labels={t50:'t₅₀ (h)',ttt:'TtT (h)',lag:'Lag Time (h)',auc:'AUC ('+scaleTag()+'·h)',maxRFU:'Max ('+scaleTag()+')',maxSlope:'Slope ('+scaleTag()+'/h)'};
  Plotly.newPlot('metrics-bar-chart',[{x:names,y:vals,type:'bar',marker:{color:colors},text:vals.map(function(v){return v!==null?v.toFixed(1):'N/A';}),textposition:'outside'}],{
    title:labels[metric]+' by Sample',font:{family:'Arial', color:'#000'},xaxis:{title:'Sample',tickangle:-45,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false},yaxis:{title:labels[metric],showgrid:false,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false},margin:{l:60,r:20,t:40,b:100}
  },{responsive:true,displaylogo:false});
}

function drawTitration(){
  var ws=D.wells;
  // Detect titration: sample names containing percentage patterns
  var pctRe=/(\d+(?:\.\d+)?)\s*%/;
  var dilutions={}, isTitration=false;
  ws.forEach(function(w){
    var m=pctRe.exec(w.sample);
    if(m){isTitration=true;var pct=parseFloat(m[1]);if(!dilutions[pct])dilutions[pct]={total:0,positive:0,values:[]};dilutions[pct].total++;if(allM[w.well].isPositive)dilutions[pct].positive++;}
  });
  var box=document.getElementById('metrics-titer-box');
  if(!isTitration||Object.keys(dilutions).length<2){box.style.display='none';return;}
  box.style.display='block';
  var pcts=Object.keys(dilutions).map(Number).sort(function(a,b){return b-a;});
  var posRates=[],posCounts=[];
  pcts.forEach(function(p){var d=dilutions[p];posRates.push(d.total>0?d.positive/d.total*100:0);posCounts.push(d.positive);});
  // SD50 estimation (Reed-Muench style)
  var sd50=null;
  for(var i=0;i<pcts.length;i++){
    if(posRates[i]<50&&i>0){
      var above=pcts[i-1],below=pcts[i],ra=posRates[i-1],rb=posRates[i];
      if(ra>rb){sd50=above+(below-above)*(50-ra)/(rb-ra);}else{sd50=(above+below)/2;}
      break;
    }
  }
  if(sd50===null&&posRates.length>0&&posRates[posRates.length-1]>=50)sd50=pcts[pcts.length-1];
  // Show sd50 info
  var ei=document.getElementById('titer-info');
  if(ei)ei.remove();
  if(sd50!==null){
    var idv=document.createElement('div');
    idv.id='titer-info';
    idv.style.cssText='padding:8px 12px;font-weight:600;font-size:0.9rem;';
    idv.innerHTML='Estimated SD₅₀: <span style="color:#007AFF;">'+sd50.toFixed(1)+'%</span> &nbsp;|&nbsp; Seed titer ≈ <span style="color:#007AFF;">'+(100/sd50).toFixed(0)+'</span> SD₅₀ units';
    box.insertBefore(idv,box.firstChild);
  }
  Plotly.newPlot('metrics-titer-chart',[{x:pcts.map(String).map(function(p){return p+'%';}),y:posRates,type:'bar',marker:{color:posRates.map(function(r){return r>=50?'#00A651':'#D00000';})},text:posRates.map(function(r,i){return posCounts[i]+'/'+dilutions[pcts[i]].total+' ('+r.toFixed(0)+'%)';}),textposition:'outside'}],{
    title:'Titration: % Positive Wells by Seed Dilution',font:{family:'Arial', color:'#000'},xaxis:{title:'Seed Dilution',tickangle:-45,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false},yaxis:{title:'Positive Wells (%)',showgrid:false,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false,range:[0,105]},margin:{l:60,r:20,t:40,b:80}
  },{responsive:true,displaylogo:false});
}

function drawSurvival(){
  var ws=D.wells;
  var posTtts=[];
  ws.forEach(function(w){var m=allM[w.well];if(m.isPositive&&m.ttt!==null)posTtts.push(m.ttt);});
  posTtts.sort(function(a,b){return a-b;});
  if(posTtts.length===0){
    Plotly.newPlot('metrics-survival-chart',[{x:[0],y:[0],type:'scatter',mode:'text',text:['No positive wells detected'],textfont:{size:14,color:'#999'}}],{title:'Cumulative Positive Distribution',font:{family:'Arial', color:'#000'},xaxis:{visible:false},yaxis:{visible:false},margin:{l:40,r:20,t:40,b:40}},{responsive:true,displaylogo:false});
    return;
  }
  var n=posTtts.length,total=ws.length;
  var x=[0],y=[0];
  for(var i=0;i<n;i++){x.push(posTtts[i]);y.push(i/n);x.push(posTtts[i]);y.push((i+1)/n);}
  var tMax=D.time_hours[D.time_hours.length-1];
  x.push(tMax);y.push(1);
  // Add scatter of individual TtT values
  var t2={x:posTtts,y:posTtts.map(function(_,i){return(i+0.5)/n*100;}),type:'scatter',mode:'markers',marker:{symbol:'x',size:8,color:'#D00000',line:{width:1.5}},name:'Individual TtT',showlegend:false};
  var t1={x:x,y:y.map(function(v){return v*100;}),type:'scatter',mode:'lines',line:{width:2.5,color:'#000',shape:'hv'},fill:'tozeroy',fillcolor:'rgba(0,0,0,0.05)',name:'Cumulative %',showlegend:false};
  Plotly.newPlot('metrics-survival-chart',[t1,t2],{
    title:'Cumulative Positive Distribution ('+n+'/'+total+' wells, '+ (n/total*100).toFixed(0)+'%)',font:{family:'Arial', color:'#000'},xaxis:{title:'Time (h)',showgrid:false,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false},yaxis:{title:'Positive Wells (%)',showgrid:false,showline:true,linecolor:'#000',linewidth:1,ticks:'outside',tickcolor:'#000',zeroline:false,range:[0,105]},margin:{l:60,r:20,t:40,b:50}
  },{responsive:true,displaylogo:false});
}

// ===== HEATMAP (renders into metrics tab) =====
function drawHeatmap(){
  if(D.isGrouped){document.getElementById('metrics-heatmap-box').style.display='none';return;}
  document.getElementById('metrics-heatmap-box').style.display='block';
  var rows=[],cols=[],wm={};
  D.wells.forEach(function(w){
    if(rows.indexOf(w.row_letter)<0)rows.push(w.row_letter);
    if(cols.indexOf(w.col_num)<0)cols.push(w.col_num);
    wm[w.well]=w;wm[w.row_letter+w.col_num]=w;
  });
  rows.sort();cols.sort(function(a,b){return a-b;});
  function fw(r,c){return wm[r+String(c).padStart(2,'0')]||wm[r+c];}
  var z=rows.map(function(r){return cols.map(function(c){var w=fw(r,c);return w?allM[w.well].maxRFU/scaleFactor():null;});});
  var zt=rows.map(function(r){return cols.map(function(c){var w=fw(r,c);return w?allM[w.well].ttt:null;});});
  var z50=rows.map(function(r){return cols.map(function(c){var w=fw(r,c);return w?allM[w.well].t50:null;});});
  var at=rows.map(function(r){return cols.map(function(c){var w=fw(r,c);return w?w.well+'<br>'+w.sample+'<br>Max:'+kRFU(allM[w.well].maxRFU)+scaleSuffix()+' t₅₀:'+(allM[w.well].t50!==null?allM[w.well].t50.toFixed(1):'—'):'';});});
  var t1={z:z,type:'heatmap',x:cols.map(String),y:rows,colorscale:'Viridis',text:at,texttemplate:'%{text}',textfont:{size:8},hovertemplate:'%{text}<extra></extra>',colorbar:{title:'Max ('+scaleTag()+')',len:0.3,y:0.82}};
  var t2={z:z50,type:'heatmap',x:cols.map(String),y:rows,colorscale:[[0,'#00A651'],[0.5,'#facc15'],[1,'#D00000']],reversescale:true,hovertemplate:'%{text}<extra></extra>',colorbar:{title:'t₅₀ (h)',len:0.3,y:0.5},xaxis:'x2',yaxis:'y2',text:rows.map(function(r,ri){return cols.map(function(c,ci){var w=fw(r,c);if(!w)return'';var t50=allM[w.well].t50;return w.well+'<br>'+w.sample+'<br>t₅₀:'+(t50!==null?t50.toFixed(1)+'h':'N/A');});}),texttemplate:'%{text}',textfont:{size:8}};
  var t3={z:zt,type:'heatmap',x:cols.map(String),y:rows,colorscale:[[0,'#00A651'],[0.5,'#facc15'],[1,'#D00000']],reversescale:true,hovertemplate:'%{text}<extra></extra>',colorbar:{title:'TtT (h)',len:0.3,y:0.18},xaxis:'x3',yaxis:'y3',text:rows.map(function(r,ri){return cols.map(function(c,ci){var w=fw(r,c);if(!w)return'';var t=allM[w.well].ttt;return w.well+'<br>'+w.sample+'<br>TtT:'+(t!==null?t.toFixed(1)+'h':'N/A');});}),texttemplate:'%{text}',textfont:{size:8}};
  Plotly.newPlot('metrics-heatmap-chart',[t1,t2,t3],{title:'Plate Heatmap: Max ('+scaleTag()+') / t₅₀ / TtT',font:{family:'Arial', color:'#000'},grid:{rows:1,columns:3,pattern:'independent'},xaxis:{title:'Column'},yaxis:{title:'Row'},xaxis2:{title:'Column'},yaxis2:{title:'Row'},xaxis3:{title:'Column'},yaxis3:{title:'Row',side:'right'},margin:{l:50,r:120,t:40,b:50}},{responsive:true,displaylogo:false});
}

// ===== TABS =====
function showTab(t){
  // Save overlay visibility state before switching away
  if (overlayTraces.length > 0) {
    overlayVisState = {};
    overlayTraces.forEach(function(tr) { overlayVisState[tr.name] = tr.visible; });
  }
  document.querySelectorAll('.tab-btn').forEach(function(b){b.classList.remove('active');});
  document.querySelectorAll('.tab-content').forEach(function(c){c.classList.remove('active');});
  document.querySelector('.tab-btn[onclick*="'+t+'"]').classList.add('active');
  document.getElementById('tab-'+t).classList.add('active');
  if(t==='overlay')drawOverlay();
  if(t==='metrics')drawMetrics();
  if(t==='plate')drawPlate();
  setTimeout(function(){Plotly.Plots.resize(document.querySelectorAll('.js-plotly-plot'));},100);
}

// ===== CONTROLS =====
function updateTh(v){
  threshold=parseInt(v);
  document.getElementById('th').value=threshold;
  document.getElementById('thn').value=threshold;
  calcAll();
  updateThHint();
  redraw();
}
function updateThHint(){
  var el=document.getElementById('th-scale-hint');
  if(!el) return;
  el.textContent='(= '+(threshold/scaleFactor()).toFixed(3)+' '+scaleTag()+')';
}
function toggleDispOpts(){
  var el=document.getElementById('disp-opts');
  if(!el) return;
  el.style.display = el.style.display==='none' ? 'inline' : 'none';
  setTimeout(function(){ try{ Plotly.Plots.resize(document.querySelectorAll('.js-plotly-plot')); }catch(e){} }, 60);
}
function ovToolAction(sel){
  var v=sel.value; sel.value='';
  if(v==='text')addTextAnnotation();
  else if(v==='axistitle')openAxisTitleEditor();
  else if(v==='axisrange')openAxisRangeEditor();
}
function ovExportAction(sel){
  var v=sel.value; sel.value='';
  if(v==='png')exportChart('png');
  else if(v==='svg')exportChart('svg');
}
function autoThreshold(){
  var bgs=[];
  D.wells.forEach(function(w){
    var v=w.values;
    if(v.length>=5){var b=v.slice(0,5).reduce(function(a,b){return a+b;},0)/5;bgs.push(b);}
  });
  if(!bgs.length){return;}
  var mean=bgs.reduce(function(a,b){return a+b;},0)/bgs.length;
  var sd=Math.sqrt(bgs.reduce(function(a,b){return a+(b-mean)*(b-mean);},0)/bgs.length);
  var th=Math.max(100, Math.round(mean+3*sd));
  var sl=document.getElementById('th');
  if(parseInt(sl.max)<th){sl.max=th;document.getElementById('thn').max=th;}
  updateTh(th);
  var info=document.getElementById('info');
  if(info)info.textContent='自动阈值: '+th+' RFU (= '+(th/scaleFactor()).toFixed(3)+' '+scaleTag()+', 基线均值+3×SD)';
}
function redraw(){
  var tab=document.querySelector('.tab-content.active').id.replace('tab-','');
  showTab(tab);
}

// ===== INIT =====
calcAll();
buildPlateWells();
syncScaleLabel();
updateThHint();
drawOverlay();
syncSdControls();
</script>
</body>
</html>"""

# ============================================================
# EXCEL PARSER
# ============================================================
def parse_excel(filepath):
    """Parse BMG plate export or grouped format Excel."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = None

    for sn in wb.sheetnames:
        s = wb[sn]
        for row in range(1, min(s.max_row + 1, 30)):
            a = str(s.cell(row=row, column=1).value or '').strip()
            if a in ('Well', 'Time'):
                ws = s
                break
        if ws:
            break

    if not ws:
        raise ValueError("未找到 'Well' 或 'Time' 表头")

    a1 = str(ws.cell(row=1, column=1).value or '').strip()

    if a1 == 'Time':
        data = _parse_grouped(ws)
    else:
        data = _parse_bmg(ws)
    _read_sd(wb, data['wells'])
    _read_raw(wb, data)
    _read_notes(wb, data)
    return data


def _read_sd(wb, wells):
    """Attach per-group SD to wells from the 'SD' sheet if present (for ±SD bands)."""
    try:
        if 'SD' not in wb.sheetnames:
            return
        s = wb['SD']
        sd_map = {}
        for row in range(2, s.max_row + 1):
            name = str(s.cell(row=row, column=1).value or '').strip()
            if not name:
                continue
            vals = []
            for col in range(2, s.max_column + 1):
                v = s.cell(row=row, column=col).value
                if v is None:
                    break
                vals.append(float(v))
            if vals:
                sd_map[name] = vals
        for w in wells:
            nm = w.get('sample') or w.get('well')
            if nm in sd_map:
                w['sd'] = sd_map[nm]
    except Exception:
        pass


def _read_raw(wb, data):
    """Attach per-replicate raw curves from the 'RAW' sheet (group, excluded, values)."""
    try:
        if 'RAW' not in wb.sheetnames:
            return
        s = wb['RAW']
        raws = []
        for row in range(2, s.max_row + 1):
            g = str(s.cell(row=row, column=1).value or '').strip()
            if not g:
                continue
            ex = int(s.cell(row=row, column=2).value or 0) or 0
            vals = []
            for col in range(3, s.max_column + 1):
                v = s.cell(row=row, column=col).value
                if v is None:
                    break
                try:
                    vals.append(float(v))
                except (TypeError, ValueError):
                    pass
            if vals:
                raws.append({'group': g, 'excluded': ex, 'values': vals})
        if raws:
            data['raws'] = raws
    except Exception:
        pass


def _read_notes(wb, data):
    """Attach structured experiment notes from the 'Notes' sheet (key, value)."""
    try:
        if 'Notes' not in wb.sheetnames:
            return
        s = wb['Notes']
        items = []
        for row in range(1, s.max_row + 1):
            k = str(s.cell(row=row, column=1).value or '').strip()
            v = s.cell(row=row, column=2).value
            if k and v is not None:
                items.append(f'{k}: {v}')
        if items:
            data['notes_str'] = ' | '.join(items)
    except Exception:
        pass


def _parse_time(ws, time_row, start_col):
    labels, hours = [], []
    for col in range(start_col, ws.max_column + 1):
        v = ws.cell(row=time_row, column=col).value
        if v is None: break
        t = str(v).strip()
        if not t: break
        labels.append(t)
        parts = t.split()
        try:
            h = float(parts[0])
        except ValueError:
            h = 0
        if len(parts) > 2:
            try:
                h += float(parts[2]) / 60
            except ValueError:
                pass
        hours.append(round(h, 4))
    return labels, hours


def _parse_grouped(ws):
    labels, hours = _parse_time(ws, 1, 2)
    wells = []
    for row in range(2, ws.max_row + 1):
        name = str(ws.cell(row=row, column=1).value or '').strip()
        if not name: break
        values = [float(ws.cell(row=row, column=c).value or 0)
                  for c in range(2, 2 + len(hours))]
        wells.append({
            'well': f'G{row - 1:02d}',
            'sample': name,
            'row_letter': 'G',
            'col_num': row - 1,
            'values': values
        })
    return {'time_labels': labels, 'time_hours': hours, 'wells': wells, 'isGrouped': True}


def _parse_bmg(ws):
    time_row = None
    for row in range(1, 20):
        b = str(ws.cell(row=row, column=2).value or '').strip()
        if b == 'Time':
            time_row = row
            break
    if not time_row:
        raise ValueError("未找到 Time 行")

    labels, hours = _parse_time(ws, time_row, 3)
    wells = []
    is_grouped = False
    fallback_idx = 0
    for row in range(time_row + 1, ws.max_row + 1):
        wid = str(ws.cell(row=row, column=1).value or '').strip()
        sample = str(ws.cell(row=row, column=2).value or '').strip()

        # Check if it's a traditional well ID (e.g., A01, B02)
        if wid and re.match(r'^[A-Z]\d{1,2}$', wid):
            # Standard well ID format
            m = re.match(r'^([A-Z]+)(\d+)$', wid)
            values = [float(ws.cell(row=row, column=c).value or 0)
                      for c in range(3, 3 + len(hours))]
            wells.append({
                'well': wid,
                'sample': sample or f'Sample {wid}',
                'row_letter': m.group(1) if m else wid[0],
                'col_num': int(m.group(2)) if m else 0,
                'values': values,
                'group': wid
            })
        elif wid or sample:
            # Group-name format: Col A = group name, Col B = sample ID
            # Use sample ID as well identifier
            is_grouped = True
            if not sample:
                sample = wid
            well_id = sample if sample else f'W{fallback_idx:03d}'
            values = [float(ws.cell(row=row, column=c).value or 0)
                      for c in range(3, 3 + len(hours))]
            wells.append({
                'well': well_id,
                'sample': wid if wid else sample,
                'row_letter': 'G',
                'col_num': fallback_idx,
                'values': values,
                'group': wid
            })
            fallback_idx += 1
        else:
            if not wells:
                continue
            has_data = False
            for c in range(3, min(3 + len(hours), 10)):
                if ws.cell(row=row, column=c).value is not None:
                    has_data = True
                    break
            if not has_data:
                break

    # Auto-mean: if group-name format with uniform replicate count > 1, compute means
    if is_grouped:
        from collections import defaultdict
        groups_dict = defaultdict(list)
        for w in wells:
            groups_dict[w['group']].append(w)
        counts = [len(v) for v in groups_dict.values()]
        if len(set(counts)) == 1 and counts[0] > 1:
            import numpy as np
            n_rep = counts[0]
            mean_wells = []
            for idx, (gname, gwells) in enumerate(groups_dict.items()):
                arr = np.array([w['values'] for w in gwells])
                mean_vals = np.nanmean(arr, axis=0).tolist()
                sd_vals = np.nanstd(arr, axis=0).tolist()
                mean_wells.append({
                    'well': gname,
                    'sample': gname,
                    'row_letter': 'G',
                    'col_num': idx,
                    'values': mean_vals,
                    'sd': sd_vals,
                    'group': gname,
                    'n_replicates': n_rep
                })
            wells = mean_wells
            groups = list(dict.fromkeys(w['group'] for w in wells))
            is_mean = True
        else:
            groups = list(dict.fromkeys(w['group'] for w in wells))
            is_mean = False
    else:
        groups = []
        is_mean = False

    info = {'time_labels': labels, 'time_hours': hours, 'wells': wells,
            'isGrouped': is_grouped, 'groups': groups, 'isMean': is_mean}
    if is_mean:
        info['n_replicates'] = counts[0]
    return info


# ============================================================
# MAIN
# ============================================================
def build(excel_path, output_dir=None, open_browser=True):
    """Main entry point: parse Excel, generate HTML, open browser."""
    data = parse_excel(excel_path)

    basename = os.path.splitext(os.path.basename(excel_path))[0]
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(excel_path))

    # Build metadata strings
    n_wells = len(data['wells'])
    n_tp = len(data['time_hours'])
    t_range = f"{data['time_hours'][0]}–{data['time_hours'][-1]} h" if data['time_hours'] else '—'

    if data['isGrouped']:
        n_rep = data.get('n_replicates', 1)
        if data.get('isMean'):
            meta = f"📊 {n_wells} 组 (均值±SD, n={n_rep}) | ⏱ {n_tp} 时间点 ({t_range})"
            header = f"🔬 {basename} — 分组均值曲线"
        else:
            meta = f"📊 {n_wells} 组 | ⏱ {n_tp} 时间点 ({t_range}) | 分组均值"
            header = f"🔬 {basename} — 分组均值曲线"
    else:
        rows_set = sorted(set(w['row_letter'] for w in data['wells']))
        cols_set = sorted(set(w['col_num'] for w in data['wells']))
        meta = f"🔬 {n_wells} 孔 ({len(rows_set)}×{len(cols_set)}) | ⏱ {n_tp} 时间点 ({t_range})"
        header = f"🔬 {basename} — RT-QuIC 荧光动力学"
    if data.get('notes_str'):
        meta = meta + ' | ' + data['notes_str']

    # Generate HTML
    html = HTML_TEMPLATE
    html = html.replace('__TITLE__', basename)
    html = html.replace('__HEADER__', header)
    html = html.replace('__META__', meta)
    html = html.replace('__DATA__', json.dumps(data, ensure_ascii=False))

    # Save
    out_path = os.path.join(output_dir, f'{basename}_viewer.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"✅ 已生成: {out_path}")
    print(f"   {meta}")

    # Also ensure plotly.min.js is in output dir
    plotly_src = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'static', 'plotly.min.js')
    plotly_dst = os.path.join(output_dir, 'plotly.min.js')
    if os.path.exists(plotly_src) and not os.path.exists(plotly_dst):
        import shutil
        shutil.copy2(plotly_src, plotly_dst)
        print(f"   📦 已复制 plotly.min.js")

    if open_browser:
        webbrowser.open('file://' + out_path)

    return out_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 build_viewer.py <excel_file.xlsx>")
        sys.exit(1)
    build(sys.argv[1])
