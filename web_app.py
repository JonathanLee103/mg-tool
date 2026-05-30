"""MG 订单表 & 库存表 生成工具 — Web 版"""

import os
import sys
import json
import queue
import threading
import webbrowser
from datetime import datetime
from flask import Flask, request, jsonify, Response
from generate_tables import generate, detect_files, FILES_EXPECTED

app = Flask(__name__)

FILE_DISPLAY = {
    "系统导出_采购订单数据.xlsx": "采购订单数据",
    "系统导出_订单汇总.xls": "订单汇总",
    "系统导出_业务库存数据.xlsx": "业务库存数据",
    "系统导出_已发货库存明细.xlsx": "已发货库存明细",
}

HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MG 订单表 & 库存表 生成工具</title>
<style>
:root {
  --bg: #f5f6fa;
  --card: #ffffff;
  --primary: #2d6cdf;
  --primary-hover: #1a56c4;
  --success: #10b981;
  --danger: #ef4444;
  --text: #1e293b;
  --text-secondary: #64748b;
  --border: #e2e8f0;
  --shadow: 0 4px 24px rgba(0,0,0,.06);
  --radius: 12px;
  --font: "Segoe UI","Microsoft YaHei","PingFang SC",sans-serif;
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:var(--font);background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column}
.header{background:linear-gradient(135deg,#1e3a5f 0%,#2d6cdf 100%);color:#fff;padding:20px 24px;text-align:center}
.header h1{font-size:22px;font-weight:700;letter-spacing:.5px}
.header .sub{font-size:13px;opacity:.75;margin-top:4px;font-weight:400}
.container{max-width:760px;margin:0 auto;padding:20px 16px 80px;flex:1}
.card{background:var(--card);border-radius:var(--radius);padding:20px 24px;margin-bottom:14px;box-shadow:var(--shadow);border:1px solid var(--border)}
.card-title{font-size:15px;font-weight:700;color:var(--text);margin-bottom:12px;display:flex;align-items:center;gap:8px}
.card-title .icon{font-size:18px}

/* Directory input */
.dir-input-wrap{display:flex;gap:8px}
.dir-input-wrap input{flex:1;padding:10px 14px;border:2px solid var(--border);border-radius:8px;font-size:14px;font-family:var(--font);outline:none;transition:border-color .2s;background:#f8fafc}
.dir-input-wrap input:focus{border-color:var(--primary);background:#fff}

/* Buttons */
.btn{display:inline-flex;align-items:center;gap:6px;padding:10px 18px;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;transition:all .2s;font-family:var(--font);white-space:nowrap}
.btn-primary{background:var(--primary);color:#fff}
.btn-primary:hover{background:var(--primary-hover);transform:translateY(-1px);box-shadow:0 4px 12px rgba(45,108,223,.3)}
.btn-primary:disabled{background:#93b8e8;cursor:not-allowed;transform:none;box-shadow:none}
.btn-outline{background:#fff;color:var(--primary);border:2px solid var(--primary)}
.btn-outline:hover{background:#eef4fd}
.btn-generate{background:var(--success);color:#fff;font-size:16px;padding:12px 48px;border-radius:10px}
.btn-generate:hover{background:#059669;transform:translateY(-1px);box-shadow:0 4px 16px rgba(16,185,129,.3)}
.btn-generate:disabled{background:#a7f3d0;cursor:not-allowed;transform:none;box-shadow:none}

/* File status */
.file-row{display:flex;align-items:center;padding:6px 0;font-size:14px;gap:10px}
.file-row .fname{width:140px;color:var(--text-secondary);text-align:right}
.file-row .fstatus{font-weight:600;display:flex;align-items:center;gap:4px}
.fok{color:var(--success)} .fmiss{color:var(--danger)}
.fdot{width:8px;height:8px;border-radius:50%;display:inline-block}
.fdot-ok{background:var(--success)} .fdot-miss{background:var(--danger)}

/* Progress */
.progress-row{display:flex;align-items:center;gap:14px}
.progress-bar{flex:1;height:10px;background:#e2e8f0;border-radius:5px;overflow:hidden}
.progress-fill{height:100%;background:linear-gradient(90deg,var(--primary),#60a5fa);border-radius:5px;transition:width .4s ease;width:0%}
.progress-pct{font-size:13px;font-weight:700;color:var(--primary);min-width:40px;text-align:right}
.status-msg{text-align:center;font-size:13px;color:var(--text-secondary);margin:6px 0;font-weight:500}

/* Log */
.log-box{background:#1a1d23;color:#c9d1d9;font-family:"Cascadia Code","Fira Code",Consolas,Menlo,monospace;font-size:12.5px;height:250px;overflow-y:auto;padding:14px 16px;border-radius:8px;line-height:1.7;border:1px solid #2a2d33}
.log-box .ts{color:#58a6ff}
.log-box .err{color:#f85149;font-weight:600}
.log-box .warn{color:#d2991d}

/* Directory browser */
.breadcrumb{font-size:13px;color:var(--text-secondary);margin-bottom:8px;word-break:break-all;display:flex;align-items:center;gap:4px;flex-wrap:wrap}
.breadcrumb span{cursor:pointer;color:var(--primary);font-weight:500}
.breadcrumb span:hover{text-decoration:underline}
.breadcrumb .sep{color:#cbd5e1;cursor:default}
.dir-grid{display:flex;flex-wrap:wrap;gap:6px;max-height:130px;overflow-y:auto;padding:2px 0}
.dir-chip{display:inline-flex;align-items:center;gap:4px;font-size:13px;padding:7px 12px;background:#f1f5f9;border:1px solid var(--border);border-radius:20px;cursor:pointer;transition:all .15s;user-select:none;font-family:var(--font)}
.dir-chip:hover{background:var(--primary);color:#fff;border-color:var(--primary)}
.dir-chip .chip-icon{font-size:16px}

/* Developer credit */
.dev-credit{text-align:center;padding:12px;font-size:13px;color:var(--text-secondary)}
.dev-credit strong{color:var(--primary);font-weight:700}
</style>
</head>
<body>

<div class="header">
  <h1>MG 订单表 &amp; 库存表 生成工具</h1>
  <div class="sub">本地 Web 服务 · 数据不上传 · 无需联网</div>
</div>

<div class="container">

  <!-- 工作目录 -->
  <div class="card">
    <div class="card-title"><span class="icon">📂</span> 选择工作目录</div>
    <div class="dir-input-wrap">
      <input type="text" id="dirPath" placeholder="输入或粘贴目录路径，如 C:\Users\Jonathan\数据" autocomplete="off" spellcheck="false">
      <button class="btn btn-primary" onclick="scanDir()">🔍 扫描目录</button>
    </div>
    <div class="breadcrumb" id="breadcrumb"></div>
    <div class="dir-grid" id="dirGrid"></div>
  </div>

  <!-- 文件状态 -->
  <div class="card" id="fileCard">
    <div class="card-title"><span class="icon">📋</span> 源文件检测</div>
    <div id="fileStatus">
      <div class="file-row"><span class="fname">等待选择目录</span><span class="fstatus" style="color:#94a3b8">— 请先选择包含源文件的工作目录</span></div>
    </div>
  </div>

  <!-- 生成按钮 -->
  <div style="text-align:center;padding:6px 0">
    <button class="btn btn-generate" id="btnGenerate" disabled onclick="startGenerate()">⚡ 开始生成</button>
  </div>

  <!-- 进度 -->
  <div class="card" id="progressCard" style="display:none">
    <div class="progress-row">
      <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
      <span class="progress-pct" id="progressPct">0%</span>
    </div>
    <div class="status-msg" id="statusMsg">正在初始化...</div>
  </div>

  <!-- 日志 -->
  <div class="card">
    <div class="card-title"><span class="icon">📜</span> 运行日志</div>
    <div class="log-box" id="logBox"></div>
  </div>

  <!-- 结果 -->
  <div class="card" id="resultCard" style="display:none;text-align:center">
    <div style="font-size:28px;margin-bottom:8px">✅</div>
    <div style="font-size:16px;font-weight:700;color:var(--success);margin-bottom:4px" id="resultText"></div>
    <div style="font-size:13px;color:var(--text-secondary);margin-bottom:12px">文件已保存至工作目录</div>
    <button class="btn btn-outline" onclick="openFolder()">📂 打开输出文件夹</button>
  </div>

</div>

<div class="dev-credit">Developed by <strong>Jonathan Lee</strong></div>

<script>
var _allReady = false, _currentPath = '';

window.onload = function(){ browseDir(''); };

function browseDir(path){
  _currentPath = path || '';
  fetch('/api/browse?path=' + encodeURIComponent(path||''))
    .then(r => r.json())
    .then(data => {
      var grid = document.getElementById('dirGrid');
      grid.innerHTML = '';
      if(path && data.parent !== undefined && data.parent !== null){
        grid.innerHTML += '<div class="dir-chip" onclick="browseDir(\''+data.parent.replace(/'/g,"\\'")+'\')"><span class="chip-icon">📁</span> ..</div>';
      }
      data.dirs.forEach(function(d){
        grid.innerHTML += '<div class="dir-chip" onclick="browseDir(\''+d.path.replace(/'/g,"\\'")+'\')"><span class="chip-icon">📁</span> '+d.name+'</div>';
      });
      document.getElementById('dirPath').value = path || '';
      updateBreadcrumb(path);
      if(path) scanDir();
    });
}

function updateBreadcrumb(path){
  if(!path){ document.getElementById('breadcrumb').innerHTML='<span>此电脑</span>'; return; }
  var parts = path.replace(/\\/g,'/').split('/').filter(Boolean);
  var html = '<span onclick="browseDir(\'\')">此电脑</span>';
  var cur = '';
  if(sys.platform==='win32' && parts.length>=1){
    cur = parts[0];
    html += ' <span class="sep">/</span> <span onclick="browseDir(\''+cur+'\')">'+cur+'</span>';
    for(var i=1;i<parts.length;i++){
      cur += '/'+parts[i];
      html += ' <span class="sep">/</span> <span onclick="browseDir(\''+cur+'\')">'+parts[i]+'</span>';
    }
  } else {
    for(var i=0;i<parts.length;i++){
      cur += '/'+parts[i];
      html += ' <span class="sep">/</span> <span onclick="browseDir(\''+cur+'\')">'+parts[i]+'</span>';
    }
  }
  document.getElementById('breadcrumb').innerHTML = html;
}

function scanDir(){
  var p = document.getElementById('dirPath').value.trim();
  if(!p){ return; }
  _currentPath = p;
  fetch('/api/scan',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path:p})})
    .then(r => r.json()).then(updateFileStatus);
}

function updateFileStatus(files){
  var html=''; _allReady=true;
  for(var key in files){
    var ok=files[key]; _allReady=_allReady&&ok;
    html+='<div class="file-row"><span class="fname">'+(FILE_NAMES[key]||key)+'</span>'+
      '<span class="fstatus '+ (ok?'fok':'fmiss') +'"><span class="fdot '+(ok?'fdot-ok':'fdot-miss')+'"></span>'+
      (ok?'已找到':'未找到')+'</span></div>';
  }
  document.getElementById('fileStatus').innerHTML=html;
  document.getElementById('btnGenerate').disabled=!_allReady;
}

function startGenerate(){
  if(!_allReady) return;
  var btn=document.getElementById('btnGenerate');
  btn.disabled=true; btn.textContent='⏳ 生成中...';
  document.getElementById('progressCard').style.display='block';
  document.getElementById('progressFill').style.width='0%';
  document.getElementById('progressPct').textContent='0%';
  document.getElementById('statusMsg').textContent='正在初始化...';
  document.getElementById('logBox').innerHTML='';
  document.getElementById('resultCard').style.display='none';

  var es=new EventSource('/api/generate?path='+encodeURIComponent(_currentPath));
  es.onmessage=function(e){
    var d=JSON.parse(e.data);
    var box=document.getElementById('logBox');
    var ts=new Date().toLocaleTimeString('zh-CN',{hour12:false});
    if(d.type==='log'){
      box.innerHTML+='<span class="ts">['+ts+']</span> '+d.msg.replace(/ /g,'&nbsp;')+'\n';
      box.scrollTop=box.scrollHeight;
    }else if(d.type==='progress'){
      document.getElementById('progressFill').style.width=d.pct+'%';
      document.getElementById('progressPct').textContent=d.pct+'%';
      document.getElementById('statusMsg').textContent=d.status;
    }else if(d.type==='done'){
      es.close();
      btn.disabled=false; btn.textContent='⚡ 开始生成';
      document.getElementById('statusMsg').textContent='✅ 生成完成';
      document.getElementById('progressPct').textContent='100%';
      document.getElementById('progressFill').style.width='100%';
      document.getElementById('resultCard').style.display='block';
      document.getElementById('resultText').innerHTML=
        '订单表 <b>'+d.result.order_rows+'</b> 行 &nbsp;|&nbsp; 库存表 <b>'+d.result.inventory_rows+'</b> 行';
    }else if(d.type==='error'){
      es.close();
      btn.disabled=false; btn.textContent='⚡ 开始生成';
      box.innerHTML+='<span class="err">[错误] '+d.msg+'</span>\n';
      box.scrollTop=box.scrollHeight;
      document.getElementById('statusMsg').textContent='❌ 生成失败';
    }
  };
  es.onerror=function(){es.close();btn.disabled=false;btn.textContent='⚡ 开始生成';};
}

function openFolder(){
  fetch('/api/open?path='+encodeURIComponent(_currentPath));
}

var FILE_NAMES = FILE_NAMES_JSON;
var sys = { platform: 'SYS_PLATFORM' };
</script>
</body>
</html>"""

HTML = HTML.replace("FILE_NAMES_JSON", json.dumps(FILE_DISPLAY, ensure_ascii=False))
HTML = HTML.replace("SYS_PLATFORM", sys.platform)


@app.route("/")
def index():
    return HTML


@app.route("/api/scan", methods=["POST"])
def api_scan():
    data = request.get_json()
    path = data.get("path", "")
    if not path or not os.path.isdir(path):
        return jsonify({f: False for f in FILES_EXPECTED})
    return jsonify(detect_files(path))


@app.route("/api/browse")
def api_browse():
    path = request.args.get("path", "")
    dirs = []
    parent = ""

    if not path:
        if sys.platform == "win32":
            import string
            for letter in string.ascii_uppercase:
                p = letter + ":\\"
                if os.path.exists(p):
                    dirs.append({"name": p, "path": p})
        else:
            dirs.append({"name": "/", "path": "/"})
            home = os.path.expanduser("~")
            for label, sub in [("桌面", "Desktop"), ("下载", "Downloads"), ("文档", "Documents")]:
                p = os.path.join(home, sub)
                if os.path.isdir(p):
                    dirs.append({"name": f"🏠 {label}", "path": p})
    else:
        parent_dir = os.path.dirname(path)
        if parent_dir and parent_dir != path:
            parent = parent_dir
        try:
            entries = sorted(os.listdir(path))
            for name in entries:
                full = os.path.join(path, name)
                if os.path.isdir(full) and not name.startswith("."):
                    dirs.append({"name": name, "path": full})
        except PermissionError:
            pass

    return jsonify({"dirs": dirs, "parent": parent})


@app.route("/api/generate")
def api_generate():
    work_dir = request.args.get("path", "")
    if not work_dir or not os.path.isdir(work_dir):
        def _err():
            yield f"data: {json.dumps({'type': 'error', 'msg': '无效的工作目录'})}\n\n"
        return Response(_err(), mimetype="text/event-stream")

    q = queue.Queue()

    def _log(msg):
        q.put(json.dumps({"type": "log", "msg": msg}))

    def _progress(pct, status):
        q.put(json.dumps({"type": "progress", "pct": pct, "status": status}))

    def _run():
        try:
            result = generate(work_dir=work_dir, log_cb=_log, progress_cb=_progress)
            q.put(json.dumps({"type": "done", "result": {
                "order_rows": result["order_rows"],
                "inventory_rows": result["inventory_rows"],
            }}))
        except Exception as e:
            q.put(json.dumps({"type": "error", "msg": str(e)}))

    threading.Thread(target=_run, daemon=True).start()

    def _stream():
        while True:
            msg = q.get()
            yield f"data: {msg}\n\n"
            if json.loads(msg)["type"] in ("done", "error"):
                break

    return Response(_stream(), mimetype="text/event-stream")


@app.route("/api/open")
def api_open():
    path = request.args.get("path", "")
    if path and os.path.isdir(path):
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    return jsonify({"ok": True})


if __name__ == "__main__":
    port = 5800
    url = f"http://127.0.0.1:{port}"
    print(f"\n{'='*50}")
    print(f"  MG 订单表 & 库存表 生成工具")
    print(f"  Developed by Jonathan Lee")
    print(f"  服务已启动: {url}")
    print(f"{'='*50}\n")
    webbrowser.open(url)
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
