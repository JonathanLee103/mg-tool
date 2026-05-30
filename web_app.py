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

# ── HTML 模板（内联，便于 PyInstaller 打包） ──────────────────

HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MG 订单表 & 库存表 生成工具</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"Microsoft YaHei","PingFang SC",sans-serif;background:#f0f2f5;color:#333;min-height:100vh}
.header{background:#1a5fb4;color:#fff;padding:16px 24px;text-align:center;font-size:20px;font-weight:bold;letter-spacing:1px}
.container{max-width:720px;margin:0 auto;padding:16px}

/* 面板 */
.panel{background:#fff;border-radius:8px;padding:16px 20px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,.08)}

/* 目录选择 */
.dir-row{display:flex;gap:8px;align-items:center}
.dir-row input{flex:1;padding:8px 12px;border:1px solid #d0d5dd;border-radius:6px;font-size:14px;outline:none}
.dir-row input:focus{border-color:#1a5fb4}
.btn{padding:8px 16px;border:none;border-radius:6px;font-size:14px;cursor:pointer;white-space:nowrap}
.btn-primary{background:#1a5fb4;color:#fff}
.btn-primary:hover{background:#1650a0}
.btn-primary:disabled{background:#93b8e0;cursor:not-allowed}
.btn-outline{background:#fff;color:#1a5fb4;border:1px solid #1a5fb4}
.btn-outline:hover{background:#e8f0fa}
.btn-success{background:#26a269;color:#fff;font-size:16px;padding:10px 40px}
.btn-success:hover{background:#1e8a56}
.btn-success:disabled{background:#a0d4b8;cursor:not-allowed}

/* 文件状态 */
.file-item{display:flex;align-items:center;padding:4px 0;font-size:14px;gap:6px}
.file-item .name{width:130px;color:#666}
.file-item .status{font-weight:500}
.status-ok{color:#26a269}
.status-missing{color:#c01c28}

/* 进度 */
.progress-wrap{display:flex;align-items:center;gap:12px}
.progress-bar{flex:1;height:8px;background:#e0e0e0;border-radius:4px;overflow:hidden}
.progress-fill{height:100%;background:#1a5fb4;border-radius:4px;transition:width .3s;width:0%}
.progress-pct{font-size:13px;color:#666;min-width:36px}

/* 日志 */
.log-area{background:#1e1e1e;color:#d4d4d4;font-family:Consolas,Menlo,monospace;font-size:12px;height:240px;overflow-y:auto;padding:12px;border-radius:6px;line-height:1.6}
.log-area .ts{color:#6a9955}
.log-area .err{color:#f44747}

/* 浏览 */
.browse-path{font-size:12px;color:#888;margin-bottom:4px;word-break:break-all}
.browse-list{display:flex;flex-wrap:wrap;gap:4px;max-height:120px;overflow-y:auto}
.browse-item{font-size:13px;padding:3px 8px;background:#e8f0fa;border-radius:4px;cursor:pointer;user-select:none}
.browse-item:hover{background:#c8ddf5}
.hint{font-size:12px;color:#999;margin-top:4px}

/* 底部 */
.footer{text-align:center;padding:8px;font-size:12px;color:#999}
</style>
</head>
<body>

<div class="header">MG 订单表 & 库存表 生成工具</div>

<div class="container">

  <!-- 目录选择 -->
  <div class="panel">
    <div style="font-weight:bold;margin-bottom:8px">工作目录</div>
    <div class="dir-row">
      <input type="text" id="dirPath" placeholder="输入工作目录路径，例如 C:\Users\张三\数据" autocomplete="off">
      <button class="btn btn-primary" onclick="scanDir()">扫描</button>
    </div>
    <div class="browse-list" id="browseList" style="margin-top:8px"></div>
    <div class="browse-path" id="browsePath"></div>
    <div class="hint">点击上方目录快速导航，或直接输入路径后点「扫描」</div>
  </div>

  <!-- 文件检测 -->
  <div class="panel" id="filePanel">
    <div style="font-weight:bold;margin-bottom:8px">源文件检测</div>
    <div id="fileStatus"></div>
  </div>

  <!-- 生成按钮 -->
  <div style="text-align:center;padding:8px 0">
    <button class="btn btn-success" id="btnGenerate" disabled onclick="startGenerate()">开始生成</button>
  </div>

  <!-- 进度 -->
  <div class="panel progress-wrap" id="progressWrap" style="display:none">
    <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
    <span class="progress-pct" id="progressPct">0%</span>
  </div>
  <div id="statusText" style="text-align:center;font-size:13px;color:#666;margin-bottom:8px"></div>

  <!-- 日志 -->
  <div class="panel">
    <div style="font-weight:bold;margin-bottom:8px">日志</div>
    <div class="log-area" id="logArea"></div>
  </div>

  <!-- 结果 -->
  <div class="panel" id="resultPanel" style="display:none;text-align:center">
    <div style="font-size:16px;font-weight:bold;color:#26a269;margin-bottom:8px" id="resultText"></div>
    <button class="btn btn-outline" onclick="openFolder()">打开输出文件夹</button>
  </div>

</div>

<div class="footer">本地服务 · 无需联网 · 数据不上传</div>

<script>
var _allReady = false;
var _currentPath = '';

// ── 初始化：加载根目录 ──
window.onload = function(){ browseDir(''); };

// ── 目录浏览（后端 /api/browse） ──
function browseDir(path){
  _currentPath = path || '';
  fetch('/api/browse?path=' + encodeURIComponent(path || ''))
    .then(r => r.json())
    .then(data => {
      var el = document.getElementById('browseList');
      el.innerHTML = '';
      // 上级目录
      if (path) el.innerHTML += '<span class="browse-item" onclick="browseDir(\'' + data.parent + '\')">📁 ..</span>';
      data.dirs.forEach(function(d){
        el.innerHTML += '<span class="browse-item" onclick="browseDir(\'' + d.path + '\')">📁 ' + d.name + '</span>';
      });
      document.getElementById('browsePath').textContent = path || '此电脑';
      // 填入路径并自动扫描
      document.getElementById('dirPath').value = path || '';
      if(path) scanDir();
    });
}

// ── 扫描目录 ──
function scanDir(){
  var p = document.getElementById('dirPath').value.trim();
  if(!p){alert('请输入工作目录路径');return;}
  _currentPath = p;
  fetch('/api/scan', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({path:p})})
    .then(r => r.json())
    .then(updateFileStatus);
}

function updateFileStatus(files){
  var html = '';
  _allReady = true;
  for(var key in files){
    var ok = files[key];
    _allReady = _allReady && ok;
    html += '<div class="file-item">' +
      '<span class="name">' + (FILE_NAMES[key]||key) + '</span>' +
      '<span class="status ' + (ok?'status-ok':'status-missing') + '">' +
      (ok?'✓ 已找到':'✗ 未找到') + '</span></div>';
  }
  document.getElementById('fileStatus').innerHTML = html;
  document.getElementById('btnGenerate').disabled = !_allReady;
}

// ── 开始生成（SSE） ──
function startGenerate(){
  if(!_allReady) return;
  var btn = document.getElementById('btnGenerate');
  btn.disabled = true; btn.textContent = '生成中...';

  document.getElementById('progressWrap').style.display = 'flex';
  document.getElementById('progressFill').style.width = '0%';
  document.getElementById('progressPct').textContent = '0%';
  document.getElementById('statusText').textContent = '正在初始化...';
  document.getElementById('logArea').innerHTML = '';
  document.getElementById('resultPanel').style.display = 'none';

  var es = new EventSource('/api/generate?path=' + encodeURIComponent(_currentPath));
  es.onmessage = function(e){
    var d = JSON.parse(e.data);
    var logArea = document.getElementById('logArea');
    var now = new Date().toLocaleTimeString('zh-CN',{hour12:false});

    if(d.type === 'log'){
      logArea.innerHTML += '<span class="ts">[' + now + ']</span> ' + d.msg + '\n';
      logArea.scrollTop = logArea.scrollHeight;
    } else if(d.type === 'progress'){
      document.getElementById('progressFill').style.width = d.pct + '%';
      document.getElementById('progressPct').textContent = d.pct + '%';
      document.getElementById('statusText').textContent = d.status;
    } else if(d.type === 'done'){
      es.close();
      btn.disabled = false; btn.textContent = '开始生成';
      document.getElementById('statusText').textContent = '完成!';
      document.getElementById('progressPct').textContent = '100%';
      document.getElementById('progressFill').style.width = '100%';
      document.getElementById('resultPanel').style.display = 'block';
      document.getElementById('resultText').innerHTML =
        '订单表 ' + d.result.order_rows + ' 行 &nbsp;|&nbsp; 库存表 ' + d.result.inventory_rows + ' 行';
    } else if(d.type === 'error'){
      es.close();
      btn.disabled = false; btn.textContent = '开始生成';
      logArea.innerHTML += '<span class="err">[错误] ' + d.msg + '</span>\n';
      logArea.scrollTop = logArea.scrollHeight;
      document.getElementById('statusText').textContent = '生成失败';
    }
  };
  es.onerror = function(){
    es.close();
    btn.disabled = false; btn.textContent = '开始生成';
  };
}

// ── 打开文件夹 ──
function openFolder(){
  fetch('/api/open?path=' + encodeURIComponent(_currentPath));
}

// ── 文件名显示 ──
var FILE_NAMES = FILE_NAMES_JSON;
</script>
</body>
</html>"""

# 注入 FILE_DISPLAY 到 HTML
HTML = HTML.replace(
    "FILE_NAMES_JSON",
    json.dumps(FILE_DISPLAY, ensure_ascii=False),
)

# ── 路由 ─────────────────────────────────────────────────────


@app.route("/")
def index():
    return HTML


@app.route("/api/scan", methods=["POST"])
def api_scan():
    data = request.get_json()
    path = data.get("path", "")
    if not path or not os.path.isdir(path):
        return jsonify(
            {f: False for f in FILES_EXPECTED}
        )
    status = detect_files(path)
    return jsonify(status)


@app.route("/api/browse")
def api_browse():
    path = request.args.get("path", "")
    dirs = []
    parent = ""

    if not path:
        # Windows: 列出驱动器盘符
        if sys.platform == "win32":
            import string
            for letter in string.ascii_uppercase:
                p = letter + ":\\"
                if os.path.exists(p):
                    dirs.append({"name": p, "path": p})
        else:
            dirs.append({"name": "/", "path": "/"})
            home = os.path.expanduser("~")
            dirs.append({"name": "🏠 桌面", "path": os.path.join(home, "Desktop")})
            dirs.append({"name": "🏠 下载", "path": os.path.join(home, "Downloads")})
            dirs.append({"name": "🏠 文档", "path": os.path.join(home, "Documents")})
    else:
        # 上级目录
        parent_dir = os.path.dirname(path)
        if parent_dir and parent_dir != path:
            # 对于盘符根目录 (C:\)，parent 会是 C:\ 本身
            parent = parent_dir if parent_dir != path else ""
        else:
            parent = ""

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
            result = generate(
                work_dir=work_dir,
                log_cb=_log,
                progress_cb=_progress,
            )
            q.put(
                json.dumps(
                    {
                        "type": "done",
                        "result": {
                            "order_rows": result["order_rows"],
                            "inventory_rows": result["inventory_rows"],
                        },
                    }
                )
            )
        except Exception as e:
            q.put(
                json.dumps(
                    {"type": "error", "msg": str(e)}
                )
            )

    threading.Thread(target=_run, daemon=True).start()

    def _stream():
        while True:
            msg = q.get()
            yield f"data: {msg}\n\n"
            parsed = json.loads(msg)
            if parsed["type"] in ("done", "error"):
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


# ── 启动 ────────────────────────────────────────────────────

if __name__ == "__main__":
    port = 5800
    url = f"http://127.0.0.1:{port}"
    print(f"\n{'='*50}")
    print(f"  MG 订单表 & 库存表 生成工具")
    print(f"  服务已启动: {url}")
    print(f"  浏览器将自动打开，如未打开请手动访问上述地址")
    print(f"  关闭此窗口即可停止服务")
    print(f"{'='*50}\n")
    webbrowser.open(url)
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
