#!/usr/bin/env python3
"""
Pikmin Patrol - RPi Zero 2W Flask Web UI
手機連 PikminAP 後開 http://10.42.0.1:5000 控制
"""
from flask import Flask, render_template_string, jsonify, request
from flask_socketio import SocketIO, emit
import subprocess, threading, time, os, json

app = Flask(__name__)
app.config['SECRET_KEY'] = 'pikmin'
socketio = SocketIO(app, cors_allowed_origins="*")

patrol_thread = None
patrol_running = False
patrol_log = []

HTML = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pikmin 巡邏控制</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #1a2e1a; color: #90ee90; font-family: monospace; padding: 16px; }
h1 { color: #7fff00; text-align: center; margin-bottom: 20px; font-size: 1.4em; }
.status-box { background: #0d1f0d; border: 2px solid #2d5a2d; border-radius: 8px; padding: 12px; margin-bottom: 16px; }
.status-label { color: #aaa; font-size: 0.85em; }
.status-val { font-size: 1.3em; font-weight: bold; }
.status-val.running { color: #7fff00; }
.status-val.stopped { color: #ff6b6b; }
.btn { display: block; width: 100%; padding: 16px; margin: 8px 0; border: none; border-radius: 8px; font-size: 1.1em; font-weight: bold; cursor: pointer; }
.btn-start { background: #2d8a2d; color: #fff; }
.btn-stop { background: #8a2d2d; color: #fff; }
.btn-start:disabled, .btn-stop:disabled { opacity: 0.4; }
.log-box { background: #0d1f0d; border: 1px solid #2d5a2d; border-radius: 8px; padding: 10px; height: 200px; overflow-y: auto; font-size: 0.8em; margin-top: 16px; }
.log-line { padding: 2px 0; border-bottom: 1px solid #1a3a1a; }
</style>
</head>
<body>
<h1>🌿 Pikmin 巡邏控制台</h1>
<div class="status-box">
  <div class="status-label">狀態</div>
  <div class="status-val stopped" id="status">停止中</div>
</div>
<button class="btn btn-start" id="btnStart" onclick="startPatrol()">▶ 開始巡邏</button>
<button class="btn btn-stop" id="btnStop" onclick="stopPatrol()" disabled>⏹ 停止巡邏</button>
<div class="log-box" id="log"><div class="log-line">等待指令...</div></div>

<script src="https://cdn.socket.io/4.7.2/socket.io.min.js"></script>
<script>
const socket = io();
socket.on('log', function(data) {
  const log = document.getElementById('log');
  const line = document.createElement('div');
  line.className = 'log-line';
  line.textContent = data.msg;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
});
socket.on('status', function(data) {
  const s = document.getElementById('status');
  s.textContent = data.running ? '巡邏中 🟢' : '停止中 🔴';
  s.className = 'status-val ' + (data.running ? 'running' : 'stopped');
  document.getElementById('btnStart').disabled = data.running;
  document.getElementById('btnStop').disabled = !data.running;
});
function startPatrol() {
  fetch('/start', {method:'POST'}).then(r=>r.json()).then(d=>addLog(d.msg));
}
function stopPatrol() {
  fetch('/stop', {method:'POST'}).then(r=>r.json()).then(d=>addLog(d.msg));
}
function addLog(msg) {
  const log = document.getElementById('log');
  const line = document.createElement('div');
  line.className = 'log-line';
  line.textContent = msg;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}
</script>
</body>
</html>
"""

def log(msg):
    ts = time.strftime('%H:%M:%S')
    line = f"[{ts}] {msg}"
    patrol_log.append(line)
    socketio.emit('log', {'msg': line})
    print(line)

def patrol_worker():
    global patrol_running
    log("巡邏開始")
    socketio.emit('status', {'running': True})
    try:
        while patrol_running:
            # 這裡之後接 pymobiledevice3 假定位
            log("巡邏中... (待接假定位)")
            time.sleep(10)
    except Exception as e:
        log(f"錯誤: {e}")
    log("巡邏停止")
    socketio.emit('status', {'running': False})

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/start', methods=['POST'])
def start():
    global patrol_thread, patrol_running
    if patrol_running:
        return jsonify({'msg': '已經在巡邏中'})
    patrol_running = True
    patrol_thread = threading.Thread(target=patrol_worker, daemon=True)
    patrol_thread.start()
    return jsonify({'msg': '開始巡邏'})

@app.route('/stop', methods=['POST'])
def stop():
    global patrol_running
    patrol_running = False
    return jsonify({'msg': '停止巡邏'})

@app.route('/status')
def status():
    return jsonify({'running': patrol_running, 'log': patrol_log[-20:]})

if __name__ == '__main__':
    print("Pikmin Patrol 啟動 http://10.42.0.1:5000")
    socketio.run(app, host='0.0.0.0', port=5000, debug=False)
