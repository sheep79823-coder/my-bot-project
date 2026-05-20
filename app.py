# --- 引用所有必要的函式庫 ---
import os
import re
import json
import datetime
import gc
from datetime import date, timedelta
import gspread
import pandas as pd
from flask import Flask, request, abort, g
from google.oauth2.service_account import Credentials
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import threading
import time
import hashlib
from apscheduler.schedulers.background import BackgroundScheduler

# --- 初始設定 ---
app = Flask(__name__)

YOUR_CHANNEL_ACCESS_TOKEN = os.environ.get('YOUR_CHANNEL_ACCESS_TOKEN')
YOUR_CHANNEL_SECRET = os.environ.get('YOUR_CHANNEL_SECRET')
GOOGLE_SHEETS_CREDENTIALS_JSON = os.environ.get('GOOGLE_SHEETS_CREDENTIALS_JSON')

# HR 系統同步設定
HR_API_URL   = os.environ.get('HR_API_URL', 'https://hr.yitai-c.uk')
HR_API_TOKEN = os.environ.get('HR_API_TOKEN', 'hr-internal-token-2024')

# 權限管理
ADMIN_USER_IDS = ["U724ac19c55418145a5af5aa1af558cbb"]
MANAGER_USER_IDS = [
    "Uc6aab7ac59f36d31c963c8357c0e19da", 
    "Uac143535b8d18cbf93a6fc5f83054e5f", 
    "Uaa8464a6b973709e941e2c6a3fd51441"
]

LIFF_ID = "2010148908-W1ntyx8i"

GOOGLE_SHEET_NAME = "我的工務助理資料庫"
WORKSHEET_NAME = "出勤時數計算"
ATTENDANCE_SHEET_NAME = "出勤時數計算"
DAILY_SUMMARY_SHEET = "每日統整"

# [優化] Session 管理設定
MAX_SESSIONS = 100  # 最多保留 100 個 Session
SESSION_EXPIRE_DAYS = 7  # Session 保留 7 天
CLEANUP_INTERVAL_HOURS = 6  # 每 6 小時清理一次

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)

# [優化] 使用字典而非全局變量
processed_messages = {}
DUPLICATE_CHECK_WINDOW = 300
session_states = {}
session_lock = threading.Lock()  # [新增] 線程安全鎖

# Google Sheets 連線
try:
    creds_json = json.loads(GOOGLE_SHEETS_CREDENTIALS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(creds_json, scopes=scope)
    gsheet_client = gspread.authorize(creds)
    worksheet = gsheet_client.open(GOOGLE_SHEET_NAME).worksheet(WORKSHEET_NAME)
    
    try:
        attendance_sheet = gsheet_client.open(GOOGLE_SHEET_NAME).worksheet(ATTENDANCE_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        workbook = gsheet_client.open(GOOGLE_SHEET_NAME)
        attendance_sheet = workbook.add_worksheet(title=ATTENDANCE_SHEET_NAME, rows=1000, cols=10)
        headers = ["日期", "姓名", "簽到時間", "離場時間", "出勤時數", "備註", "更新時間"]
        attendance_sheet.append_row(headers)
        print("✅ 已建立出勤時數計算表")
    
    try:
        summary_sheet = gsheet_client.open(GOOGLE_SHEET_NAME).worksheet(DAILY_SUMMARY_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        workbook = gsheet_client.open(GOOGLE_SHEET_NAME)
        summary_sheet = workbook.add_worksheet(title=DAILY_SUMMARY_SHEET, rows=1000, cols=10)
        headers = ["統計日期", "姓名", "總出勤天數", "統計時間"]
        summary_sheet.append_row(headers)
        print("✅ 已建立每日統整表")
    
    print("✅ Google Sheets 連線成功！")
except Exception as e:
    print(f"❌ Google Sheets 連線失敗: {e}")
    worksheet = None
    attendance_sheet = None
    summary_sheet = None

# [優化] 主動垃圾回收
@app.after_request
def after_request(response):
    """每次請求後強制垃圾回收"""
    gc.collect()
    return response

# [優化] 清理過期資源
def cleanup_old_sessions():
    """清理過期的 Session 和 processed_messages"""
    with session_lock:
        try:
            current_time = time.time()
            cutoff_date = date.today() - timedelta(days=SESSION_EXPIRE_DAYS)
            
            # 清理過期 Session
            sessions_to_remove = []
            for session_key, session in session_states.items():
                try:
                    session_date = datetime.datetime.strptime(session.work_date, '%Y/%m/%d').date()
                    if session_date < cutoff_date:
                        sessions_to_remove.append(session_key)
                except:
                    # 民國年格式轉換
                    try:
                        parts = session.work_date.split('/')
                        minguo_year, month, day = [int(p) for p in parts]
                        session_date = date(minguo_year + 1911, month, day)
                        if session_date < cutoff_date:
                            sessions_to_remove.append(session_key)
                    except:
                        pass
            
            for key in sessions_to_remove:
                del session_states[key]
            
            # 限制 Session 數量
            if len(session_states) > MAX_SESSIONS:
                sorted_sessions = sorted(
                    session_states.items(),
                    key=lambda x: x[1].created_time
                )
                excess_count = len(session_states) - MAX_SESSIONS
                for i in range(excess_count):
                    del session_states[sorted_sessions[i][0]]
            
            # 清理過期的 processed_messages
            messages_to_remove = [
                k for k, v in processed_messages.items() 
                if current_time - v > DUPLICATE_CHECK_WINDOW
            ]
            for key in messages_to_remove:
                del processed_messages[key]
            
            print(f"🧹 清理完成: 移除 {len(sessions_to_remove)} 個過期 Session")
            print(f"📊 當前 Session 數: {len(session_states)}")
            
            # 強制垃圾回收
            gc.collect()
            
        except Exception as e:
            print(f"❌ 清理失敗: {e}")

def keep_alive():
    """防止服務休眠"""
    while True:
        try:
            time.sleep(840)
            import urllib.request
            render_url = os.environ.get('RENDER_URL', 'https://my-bot-project-1.onrender.com')
            try:
                urllib.request.urlopen(f"{render_url}/health", timeout=5)
                print("[KEEPALIVE] ✅ 防止休眠")
            except:
                print("[KEEPALIVE] ⚠️ Ping 失敗")
        except Exception as e:
            print(f"[KEEPALIVE] ❌ {e}")

keep_alive_thread = threading.Thread(target=keep_alive, daemon=True)
keep_alive_thread.start()

@app.route("/liff", methods=['GET'])
def liff_page():
    """LIFF 日報輸入表單"""
    today_tw = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
    minguo_year = today_tw.year - 1911
    today_str = f"{minguo_year}/{today_tw.month:02d}/{today_tw.day:02d}"
    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>日報輸入</title>
<script charset="utf-8" src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, sans-serif; background: #f0f4f8; padding: 16px; }}
  h2 {{ text-align: center; color: #2d3748; margin-bottom: 16px; font-size: 18px; }}
  .card {{ background: white; border-radius: 12px; padding: 16px; margin-bottom: 12px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }}
  label {{ display: block; font-size: 13px; color: #718096; margin-bottom: 4px; font-weight: 600; }}
  input {{ width: 100%; padding: 10px 12px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 15px; }}
  input:focus {{ outline: none; border-color: #06c755; }}
  .staff-item {{ display: flex; gap: 8px; margin-bottom: 8px; }}
  .staff-item input {{ flex: 1; }}
  .btn-remove {{ background: #fed7d7; color: #c53030; border: none; border-radius: 8px; padding: 0 12px; cursor: pointer; font-size: 16px; }}
  .btn-add {{ width: 100%; padding: 10px; background: #ebf8f0; color: #276749; border: 1px dashed #9ae6b4; border-radius: 8px; font-size: 14px; cursor: pointer; margin-top: 4px; }}
  .btn-submit {{ width: 100%; padding: 14px; background: #06c755; color: white; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 8px; }}
  .btn-submit:active {{ background: #05a847; }}
  #preview {{ background: #f7fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; font-size: 13px; white-space: pre-wrap; color: #2d3748; min-height: 60px; }}
</style>
</head>
<body>
<h2>📋 出工日報</h2>
<div class="card">
  <label>日期（民國年）</label>
  <input type="text" id="date" value="{today_str}" placeholder="115/05/20">
</div>
<div class="card">
  <label>工程名稱</label>
  <input type="text" id="project" placeholder="請輸入工程名稱">
</div>
<div class="card">
  <label>出工人員</label>
  <div id="staff-list">
    <div class="staff-item">
      <input type="text" placeholder="姓名（備註）">
      <button class="btn-remove" onclick="removeStaff(this)">✕</button>
    </div>
  </div>
  <button class="btn-add" onclick="addStaff()">＋ 新增人員</button>
</div>
<div class="card">
  <label>預覽訊息</label>
  <div id="preview">填寫後自動預覽...</div>
</div>
<button class="btn-submit" onclick="sendReport()">✅ 送出日報</button>

<script>
liff.init({{ liffId: "{LIFF_ID}" }});

function addStaff() {{
  const list = document.getElementById('staff-list');
  const div = document.createElement('div');
  div.className = 'staff-item';
  div.innerHTML = '<input type="text" placeholder="姓名（備註）" oninput="updatePreview()"><button class="btn-remove" onclick="removeStaff(this)">✕</button>';
  list.appendChild(div);
}}

function removeStaff(btn) {{
  const items = document.querySelectorAll('.staff-item');
  if (items.length > 1) {{ btn.parentElement.remove(); updatePreview(); }}
}}

function getStaffList() {{
  return Array.from(document.querySelectorAll('.staff-item input'))
    .map(i => i.value.trim()).filter(v => v);
}}

function updatePreview() {{
  const date = document.getElementById('date').value.trim();
  const project = document.getElementById('project').value.trim();
  const staff = getStaffList();
  if (!date || !project) {{ document.getElementById('preview').textContent = '填寫後自動預覽...'; return; }}
  let msg = date + '\\n' + project + '\\n出工人員：\\n';
  staff.forEach((s, i) => msg += (i+1) + '. ' + s + '\\n');
  msg += '共計 ' + staff.length + ' 人';
  document.getElementById('preview').textContent = msg;
}}

document.querySelectorAll('input').forEach(i => i.addEventListener('input', updatePreview));

async function sendReport() {{
  const date = document.getElementById('date').value.trim();
  const project = document.getElementById('project').value.trim();
  const staff = getStaffList();
  if (!date || !project || staff.length === 0) {{
    alert('請填寫日期、工程名稱，並至少新增一位人員');
    return;
  }}
  let msg = date + '\\n' + project + '\\n出工人員：\\n';
  staff.forEach((s, i) => msg += (i+1) + '. ' + s + '\\n');
  msg += '共計 ' + staff.length + ' 人';
  try {{
    await liff.sendMessages([{{ type: 'text', text: msg }}]);
    liff.closeWindow();
  }} catch(e) {{
    alert('送出失敗: ' + e.message);
  }}
}}
</script>
</body>
</html>"""
    return html


def send_line_push(user_id, text):
    """推播訊息給指定用戶"""
    try:
        line_bot_api.push_message(user_id, TextSendMessage(text=text))
        print(f"✅ 推播成功: {user_id[-8:]}")
    except Exception as e:
        print(f"❌ 推播失敗: {e}")


def calculate_period_summary():
    """計算當前結算期間的出勤統計，回傳文字"""
    if not attendance_sheet:
        return None
    try:
        today = date.today()
        if today.day <= 5:
            start_date = (today.replace(day=1) - timedelta(days=1)).replace(day=21)
            end_date = today.replace(day=5)
            period_label = f"{start_date.strftime('%m/%d')} ~ {end_date.strftime('%m/%d')}"
        elif today.day <= 20:
            start_date = today.replace(day=6)
            end_date = today.replace(day=20)
            period_label = f"{start_date.strftime('%m/%d')} ~ {end_date.strftime('%m/%d')}"
        else:
            start_date = today.replace(day=21)
            next_month = (today.replace(day=1) + timedelta(days=32)).replace(day=1)
            end_date = next_month.replace(day=5)
            period_label = f"{start_date.strftime('%m/%d')} ~ {end_date.strftime('%m/%d')}"

        records = attendance_sheet.get_all_records()
        if not records:
            return None
        df = pd.DataFrame(records)
        df['日期_dt'] = pd.to_datetime(df['日期'].apply(minguo_to_gregorian), errors='coerce')
        period_df = df.dropna(subset=['日期_dt'])
        period_df = period_df[
            (period_df['日期_dt'] >= pd.to_datetime(start_date)) &
            (period_df['日期_dt'] <= pd.to_datetime(end_date))
        ]
        if period_df.empty:
            return f"📅 結算期間 {period_label}\n本期無出勤記錄"

        period_df['出勤時數'] = pd.to_numeric(period_df['出勤時數'], errors='coerce')
        summary = period_df.groupby('姓名')['出勤時數'].sum().reset_index()
        summary = summary.sort_values('出勤時數', ascending=False)

        text = f"📊 出勤結算報告\n"
        text += f"📅 期間：{period_label}\n"
        text += f"{'─'*20}\n"
        for _, row in summary.iterrows():
            text += f"• {row['姓名']}: {row['出勤時數']} 天\n"
        text += f"{'─'*20}\n"
        text += f"合計 {len(summary)} 人"
        return text
    except Exception as e:
        print(f"❌ 結算計算失敗: {e}")
        return None


def auto_settlement():
    """每月 5 號、20 號 17:00 自動結算並推播給 ADMIN"""
    print("📊 自動結算開始")
    msg = calculate_period_summary()
    if msg:
        for admin_id in ADMIN_USER_IDS:
            send_line_push(admin_id, msg)
        print(f"✅ 結算報告已推播給 {len(ADMIN_USER_IDS)} 位 ADMIN")
    else:
        print("⚠️ 結算無資料，不推播")


@app.route("/health", methods=['GET'])
def health_check():
    """健康檢查端點"""
    return json.dumps({
        'status': 'ok',
        'sessions': len(session_states),
        'memory_info': f'{gc.get_count()}'
    }), 200, {'Content-Type': 'application/json'}

def is_duplicate_message(user_id, message_text, timestamp):
    """檢查重複訊息"""
    msg_hash = hashlib.md5(f"{user_id}{message_text}{timestamp}".encode()).hexdigest()
    current_time = time.time()
    
    # 清理過期訊息記錄
    to_delete = [k for k, v in processed_messages.items() if current_time - v > DUPLICATE_CHECK_WINDOW]
    for k in to_delete:
        del processed_messages[k]
    
    if msg_hash in processed_messages:
        return True
    
    processed_messages[msg_hash] = current_time
    return False

# [優化] 權限檢查
def get_user_role(user_id):
    """取得用戶權限等級"""
    if user_id in ADMIN_USER_IDS:
        return "ADMIN"
    elif user_id in MANAGER_USER_IDS:
        return "MANAGER"
    return None

def can_access_session(user_id, session):
    """檢查用戶是否有權限存取 Session"""
    role = get_user_role(user_id)
    if role == "ADMIN":
        return True
    elif role == "MANAGER":
        return session.is_authorized(user_id)
    return False

# Google Sheets 操作函式
def sync_to_hr(action, name, work_date_minguo, rec_time, project=""):
    """
    同步簽到/離場到 HR 系統
    work_date_minguo: 民國年字串 e.g. '113/04/24'
    rec_time: datetime.datetime 物件
    """
    try:
        greg_date = minguo_to_gregorian(work_date_minguo)
        if not greg_date:
            print(f"[HR同步] 日期轉換失敗: {work_date_minguo}")
            return False
        payload = {
            "action":  action,
            "name":    name,
            "date":    greg_date.strftime('%Y-%m-%d'),
            "time":    rec_time.strftime('%H:%M'),
            "project": project,
        }
        resp = requests.post(
            f"{HR_API_URL}/api/attendance/sync",
            json=payload,
            headers={"X-API-Token": HR_API_TOKEN},
            timeout=5,
        )
        if resp.status_code == 200:
            print(f"[HR同步] ✅ {action} {name} → {resp.json()}")
            return True
        elif resp.status_code == 404:
            print(f"[HR同步] ⚠️ 員工未在 HR 系統：{name}")
            return False
        else:
            print(f"[HR同步] ❌ HTTP {resp.status_code}: {resp.text[:100]}")
            return False
    except Exception as e:
        print(f"[HR同步] ❌ 例外：{e}")
        return False


def write_person_to_sheet(work_date, project_name, person_name, sign_in_time, note=""):
    """立即寫入簽到記錄"""
    if not attendance_sheet:
        print(f"❌ write_person_to_sheet 失敗: attendance_sheet 未初始化（Sheets連線失敗）")
        return False
    
    try:
        update_time = datetime.datetime.now(
            datetime.timezone(datetime.timedelta(hours=8))
        ).strftime('%Y-%m-%d %H:%M:%S')
        
        new_row = [
            work_date,
            person_name,
            sign_in_time.strftime('%H:%M') if sign_in_time else "",
            "",
            "",
            note if note else f"項目: {project_name}",
            update_time
        ]
        attendance_sheet.append_row(new_row)
        print(f"✅ 已即時寫入 {person_name} 的簽到記錄")
        # 同步到 HR 系統
        sync_to_hr('checkin', person_name, work_date, sign_in_time, project_name)
        return True
    except Exception as e:
        print(f"❌ 寫入失敗: {e}")
        return False

def update_person_checkout(work_date, person_name, checkout_time, sign_in_time):
    """更新離場時間和出勤天數"""
    if not attendance_sheet:
        return False
    
    try:
        records = attendance_sheet.get_all_records()
        target_row = None
        
        for i, record in enumerate(records, start=2):
            if record['日期'] == work_date and record['姓名'] == person_name and not record['離場時間']:
                target_row = i
        
        if target_row:
            checkout_hour = checkout_time.hour
            sign_in_hour = sign_in_time.hour
            
            # 計算出勤天數
            if sign_in_hour < 10:
                days = 1.0
                remark = ""
            elif sign_in_hour < 13:
                days = 1.0
                remark = ""
            else:
                days = 0.5
                remark = "下午簽到"
            
            # 16:00 前早退
            if checkout_hour < 16:
                days = 0.5
                remark = f"早退({checkout_time.strftime('%H:%M')})"
            # 17:00 後加班
            elif checkout_hour >= 17:
                remark = (remark + " " if remark else "") + "加班"
            
            attendance_sheet.update_cell(target_row, 4, checkout_time.strftime('%H:%M'))
            attendance_sheet.update_cell(target_row, 5, days)
            attendance_sheet.update_cell(target_row, 6, remark.strip())
            print(f"✅ 已更新 {person_name} 的離場記錄: {days} 天")
            # 同步到 HR 系統（需要 work_date，從 attendance_sheet 的記錄取）
            record_date = records[target_row - 2].get('日期', '')
            sync_to_hr('checkout', person_name, record_date, checkout_time)
            return True
        
        return False
    except Exception as e:
        print(f"❌ 更新失敗: {e}")
        return False

# 每日統整
def daily_summary():
    """每天 22:00 台灣時間執行統整"""
    print("\n" + "="*50)
    print("🕙 22:00 每日統整開始")
    print("="*50)
    
    if not attendance_sheet or not summary_sheet:
        print("❌ 工作表連線失敗")
        return
    
    try:
        today = date.today()
        minguo_year = today.year - 1911
        today_str = f"{minguo_year:03d}/{today.month:02d}/{today.day:02d}"
        
        records = attendance_sheet.get_all_records()
        df = pd.DataFrame(records)
        
        today_df = df[df['日期'] == today_str]
        
        if today_df.empty:
            print(f"ℹ️ {today_str} 沒有出勤記錄")
            return
        
        # 同一天同一人只計最高時數
        summary_list = []
        for person_name in today_df['姓名'].unique():
            person_records = today_df[today_df['姓名'] == person_name]
            days_list = pd.to_numeric(person_records['出勤時數'], errors='coerce').dropna().tolist()
            
            if days_list:
                max_days = max(days_list)
                summary_list.append({'姓名': person_name, '總出勤天數': max_days})
        
        if not summary_list:
            print(f"ℹ️ {today_str} 沒有有效的出勤時數")
            return
        
        summary_df = pd.DataFrame(summary_list)
        update_time = datetime.datetime.now(
            datetime.timezone(datetime.timedelta(hours=8))
        ).strftime('%Y-%m-%d %H:%M:%S')
        
        for _, row in summary_df.iterrows():
            summary_row = [today_str, row['姓名'], row['總出勤天數'], update_time]
            summary_sheet.append_row(summary_row)
        
        print(f"✅ 已統整 {len(summary_df)} 人的 {today_str} 出勤資料")
        
        # 統整後清理垃圾
        gc.collect()
        
    except Exception as e:
        print(f"❌ 統整失敗: {e}")

# 排程設定
scheduler = BackgroundScheduler(timezone='Asia/Taipei')

def start_scheduler():
    """啟動排程器"""
    # 每日統整
    scheduler.add_job(daily_summary, 'cron', hour=22, minute=0, timezone='Asia/Taipei')
    # 定期清理
    scheduler.add_job(cleanup_old_sessions, 'interval', hours=CLEANUP_INTERVAL_HOURS)
    # 雙週期結算：每月 5 號、20 號 17:00
    scheduler.add_job(auto_settlement, 'cron', day='5,20', hour=17, minute=0, timezone='Asia/Taipei')
    scheduler.start()
    print("✅ 已啟動排程器")
    print(f"   - 每日 22:00 (台灣時間) 統整出勤")
    print(f"   - 每月 5/20 號 17:00 自動結算推播")
    print(f"   - 每 {CLEANUP_INTERVAL_HOURS} 小時清理過期 Session")

start_scheduler()

# Session 管理類別
class DailySession:
    __slots__ = ['work_date', 'project_name', 'staff', 'created_time', 'authorized_users']
    
    def __init__(self, work_date, project_name=""):
        self.work_date = work_date
        self.project_name = project_name
        self.staff = []
        self.created_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        self.authorized_users = set()
    
    def add_authorized_user(self, user_id):
        self.authorized_users.add(user_id)
    
    def is_authorized(self, user_id):
        return user_id in self.authorized_users
    
    def add_staff_and_write(self, name, note=None, add_time=None):
        if add_time is None:
            add_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        
        if name not in [s['name'] for s in self.staff]:
            if write_person_to_sheet(self.work_date, self.project_name, name, add_time, note or ""):
                self.staff.append({"name": name, "add_time": add_time, "note": note})
                return True
        return False
    
    def get_summary(self):
        summary = f"📋 {self.work_date}\n"
        summary += f"👥 目前人數: {len(self.staff)} 人\n"
        summary += "人員:\n"
        for i, person in enumerate(self.staff, 1):
            summary += f"  {i}. {person['name']}\n"
        return summary

def get_or_create_session(work_date, project_name, user_id):
    """取得或建立 Session - 線程安全"""
    with session_lock:
        if project_name is None:
            project_name = ""
        
        session_key = f"{work_date}_{project_name}"
        if session_key not in session_states:
            session_states[session_key] = DailySession(work_date, project_name)
        
        session_states[session_key].add_authorized_user(user_id)
        return session_states[session_key]

def find_session_for_user(user_id, project_name=None, work_date=None):
    """智能找到用戶要操作的 Session"""
    today = date.today()
    minguo_year = today.year - 1911
    today_str = f"{minguo_year:03d}/{today.month:02d}/{today.day:02d}"
    
    if work_date is None:
        work_date = today_str
    
    role = get_user_role(user_id)
    accessible_sessions = []
    
    print(f"[查找] 用戶角色: {role}, 目標日期: {work_date}, 指定專案: {project_name}")
    
    for session_key, session in session_states.items():
        # 確保日期匹配
        if session.work_date == work_date:
            if role == "ADMIN":
                accessible_sessions.append(session)
                print(f"  [管理員] 可存取: {session.project_name}")
            elif role == "MANAGER" and session.is_authorized(user_id):
                accessible_sessions.append(session)
                print(f"  [經理] 可存取: {session.project_name}")
    
    print(f"[查找] 共找到 {len(accessible_sessions)} 個可存取的 Session")
    
    # 情況1: 指定了專案名稱 - 精確匹配
    if project_name:
        for session in accessible_sessions:
            if session.project_name == project_name:
                print(f"[匹配] 精確匹配成功: {project_name}")
                return session
        # 如果精確匹配失敗，嘗試部分匹配
        for session in accessible_sessions:
            if project_name in session.project_name or session.project_name in project_name:
                print(f"[匹配] 部分匹配成功: {session.project_name}")
                return session
        print(f"[匹配] 找不到專案: {project_name}")
        return None
    
    # 情況2: 沒指定專案名稱
    if len(accessible_sessions) == 0:
        print(f"[查找] 沒有可用的 Session")
        return None
    elif len(accessible_sessions) == 1:
        print(f"[查找] 唯一 Session: {accessible_sessions[0].project_name}")
        return accessible_sessions[0]
    else:
        # 多個專案，返回最近的
        latest = max(accessible_sessions, key=lambda s: s.created_time)
        print(f"[查找] 返回最新的 Session: {latest.project_name}")
        return latest

# 解析函式
def parse_full_attendance_report(text):
    """解析完整日報"""
    try:
        lines = text.strip().split('\n')
        if len(lines) < 2:
            return None
        
        date_match = re.match(r"^(0?\d{3}/\d{2}/\d{2})", lines[0])
        if not date_match:
            return None
        # 去掉民國年前導零 (0115 -> 115)
        raw_date = date_match.group(1)
        parts = raw_date.split('/')
        work_date = f"{int(parts[0])}/{parts[1]}/{parts[2]}"
        
        project_name = lines[1].strip()
        if not project_name:
            return None
        
        staff_start_idx = None
        for i, line in enumerate(lines):
            if "人員" in line or "出工" in line:
                staff_start_idx = i + 1
                break
        
        if staff_start_idx is None:
            staff_start_idx = 2
        
        staff_list = []
        for i in range(staff_start_idx, len(lines)):
            line = lines[i].strip()
            if not line or "共計" in line or "便當" in line:
                continue
            
            clean_line = re.sub(r"^\d+[\.\、]", "", line).strip()
            note_match = re.search(r"\((.+)\)", clean_line)
            
            if note_match:
                note = note_match.group(1)
                name = clean_line[:note_match.start()].strip()
                staff_list.append({"name": name, "note": note})
            else:
                if clean_line:
                    staff_list.append({"name": clean_line, "note": None})
        
        if not staff_list:
            return None
        
        return {"date": work_date, "project_name": project_name, "staff": staff_list}
    except Exception as e:
        print(f"❌ 解析日報錯誤: {e}")
        return None

def parse_add_staff(text):
    """解析新增人員指令"""
    match = re.search(r"新增[:：]\s*(.+?)@(.+?)(?:\s*\((.+)\))?$", text.strip())
    if match:
        return {"name": match.group(1).strip(), "project": match.group(2).strip(), 
                "note": match.group(3).strip() if match.group(3) else None}
    
    match = re.search(r"新增[:：]\s*(.+?)(?:\s*\((.+)\))?$", text.strip())
    if match:
        return {"name": match.group(1).strip(), "project": None,
                "note": match.group(2).strip() if match.group(2) else None}
    return None

def parse_checkout_staff(text):
    """解析離場指令"""
    match = re.search(r"(?:離場|下班)[:：]\s*(.+?)@(.+?)$", text.strip())
    if match:
        return {"name": match.group(1).strip(), "project": match.group(2).strip()}
    
    match = re.search(r"(?:離場|下班)[:：]\s*(.+?)$", text.strip())
    if match:
        return {"name": match.group(1).strip(), "project": None}
    return None

def minguo_to_gregorian(minguo_str):
    """民國年轉西元年"""
    try:
        parts = minguo_str.split('/')
        minguo_year, month, day = [int(p) for p in parts]
        return date(minguo_year + 1911, month, day)
    except:
        return None

# Webhook 處理
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    
    try:
        handler.handle(body, signature)
        return 'OK', 200
    except InvalidSignatureError:
        return 'Invalid signature', 403
    except Exception as e:
        print(f"❌ Callback 錯誤: {e}")
        return 'Internal Server Error', 500

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    try:
        user_id = event.source.user_id
        message_text = event.message.text.strip()
        timestamp = event.timestamp / 1000
        message_time = datetime.datetime.fromtimestamp(
            timestamp, tz=datetime.timezone(datetime.timedelta(hours=8))
        )
        
        print(f"\n[訊息] User: {user_id[-8:]}, Text: {message_text[:30]}, Time: {message_time.strftime('%H:%M')}")
        
        # 權限檢查
        user_role = get_user_role(user_id)
        if not user_role:
            print(f"[拒絕] 無權限用戶")
            return

        # 重複檢查
        if is_duplicate_message(user_id, message_text, timestamp):
            print(f"[重複] 已處理過")
            return
        
        reply_text = None
        
        # === 完整日報 ===
        if re.search(r"0?\d{3}/\d{2}/\d{2}", message_text) and any(char in message_text for char in ["人員", "出工"]):
            print("📝 處理日報")
            report_data = parse_full_attendance_report(message_text)
            if report_data:
                print(f"[解析] 日期: {report_data['date']}, 專案: {report_data['project_name']}, 人數: {len(report_data['staff'])}")
                session = get_or_create_session(report_data['date'], report_data['project_name'], user_id)
                session.project_name = report_data['project_name']
                
                success_count = 0
                for staff in report_data['staff']:
                    if session.add_staff_and_write(staff['name'], staff['note'], message_time):
                        success_count += 1
                
                print(f"[Session] 已建立 Key: {report_data['date']}_{report_data['project_name']}")
                print(f"[寫入] 成功: {success_count}/{len(report_data['staff'])}")
                
                reply_text = f"✅ 已記錄 {success_count} 人\n專案: {report_data['project_name'][:20]}...\n日期: {report_data['date']}"
            else:
                print("[解析失敗] 無法解析日報")
                reply_text = "❌ 日報格式錯誤"
        
        # === 新增人員 ===
        elif "新增" in message_text:
            print("➕ 新增人員")
            staff_info = parse_add_staff(message_text)
            if staff_info:
                valid_session = find_session_for_user(user_id, staff_info.get('project'))
                if valid_session:
                    if valid_session.add_staff_and_write(staff_info['name'], staff_info['note'], message_time):
                        reply_text = f"✅ 已新增 {staff_info['name']} ({message_time.strftime('%H:%M')})"
                    else:
                        reply_text = f"⚠️ {staff_info['name']} 已在清單中"
                elif staff_info.get('project'):
                    reply_text = f"❌ 找不到專案「{staff_info['project']}」"
                else:
                    # 多專案情況
                    active_projects = [s.project_name for s in session_states.values() 
                                     if can_access_session(user_id, s) and s.project_name]
                    if len(set(active_projects)) > 1:
                        reply_text = f"⚠️ 有多個專案，請用: 新增：名字@專案名稱\n可用: {', '.join(set(active_projects)[:3])}"
                    else:
                        reply_text = "❌ 請先提交完整日報"
        
        # === 單筆離場 ===
        elif ("離場:" in message_text or "離場：" in message_text or 
              "下班:" in message_text or "下班：" in message_text):
            print("🚶 單筆離場")
            checkout_info = parse_checkout_staff(message_text)
            if checkout_info:
                valid_session = find_session_for_user(user_id, checkout_info.get('project'))
                if valid_session:
                    person_data = next((p for p in valid_session.staff if p['name'] == checkout_info['name']), None)
                    if person_data:
                        if update_person_checkout(valid_session.work_date, checkout_info['name'], 
                                                 message_time, person_data['add_time']):
                            reply_text = f"✅ {checkout_info['name']} 已離場 ({message_time.strftime('%H:%M')})"
                        else:
                            reply_text = f"⚠️ 更新失敗，可能已記錄過"
                    else:
                        reply_text = f"❌ 找不到 {checkout_info['name']} 的簽到記錄"
                elif checkout_info.get('project'):
                    reply_text = f"❌ 找不到專案「{checkout_info['project']}」"
                else:
                    reply_text = "❌ 請指定專案名稱"
        
        # === 通用離場 ===
        elif "人員離場" in message_text or "人員下班" in message_text:
            print("⬜ 全員離場")
            if user_role not in ("ADMIN", "MANAGER"):
                print("[拒絕] 非管理者觸發全員離場")
                return
            project_match = re.search(r"@(.+?)$", message_text)
            project_name = project_match.group(1).strip() if project_match else None
            valid_session = find_session_for_user(user_id, project_name)
            
            if valid_session and valid_session.staff:
                default_checkout_time = message_time.replace(hour=16, minute=50, second=0, microsecond=0)
                count = 0
                for person in valid_session.staff:
                    if update_person_checkout(valid_session.work_date, person['name'], 
                                            default_checkout_time, person['add_time']):
                        count += 1
                reply_text = f"✅ 已記錄 {count} 人離場 (預設 16:50)\n專案: {valid_session.project_name}"
            elif project_name:
                reply_text = f"❌ 找不到專案「{project_name}」"
            else:
                active_projects = [s.project_name for s in session_states.values() 
                                 if can_access_session(user_id, s) and s.project_name]
                if len(set(active_projects)) > 1:
                    reply_text = f"⚠️ 有多個專案，請用: 人員離場@專案名稱"
                else:
                    reply_text = "❌ 找不到有效的日報記錄"
        
        # === 查詢出勤 ===
        elif message_text == "查詢本期出勤":
            print("📊 查詢出勤")
            if attendance_sheet:
                try:
                    today = date.today()
                    if today.day <= 5:
                        start_date = (today.replace(day=1) - timedelta(days=1)).replace(day=21)
                        end_date = today.replace(day=5)
                    elif today.day <= 20:
                        start_date = today.replace(day=6)
                        end_date = today.replace(day=20)
                    else:
                        start_date = today.replace(day=21)
                        next_month = (today.replace(day=1) + timedelta(days=32)).replace(day=1)
                        end_date = next_month.replace(day=5)
                    
                    records = attendance_sheet.get_all_records()
                    if records:
                        df = pd.DataFrame(records)
                        df['日期'] = pd.to_datetime(df['日期'].apply(minguo_to_gregorian), errors='coerce')
                        
                        period_df = df.dropna(subset=['日期'])
                        period_df = period_df[
                            (period_df['日期'] >= pd.to_datetime(start_date)) &
                            (period_df['日期'] <= pd.to_datetime(end_date))
                        ]
                        
                        if not period_df.empty:
                            period_df['出勤時數'] = pd.to_numeric(period_df['出勤時數'], errors='coerce')
                            summary = period_df.groupby('姓名')['出勤時數'].sum().reset_index()
                            reply_text = f"📅 本期 ({start_date.strftime('%m/%d')}-{end_date.strftime('%m/%d')}) 統計：\n"
                            for _, row in summary.iterrows():
                                reply_text += f"• {row['姓名']}: {row['出勤時數']} 天\n"
                        else:
                            reply_text = "本期無出勤記錄"
                    else:
                        reply_text = "試算表無資料"
                except Exception as e:
                    reply_text = f"❌ 查詢失敗: {str(e)[:50]}"
                    print(f"查詢錯誤: {e}")
            else:
                reply_text = "❌ Google Sheets 未連線"
        
        # === 系統狀態查詢 ===
        elif message_text == "系統狀態" and user_role == "ADMIN":
            reply_text = f"📊 系統狀態\n"
            reply_text += f"Session 數: {len(session_states)}\n"
            reply_text += f"今日專案: {len([s for s in session_states.values() if s.work_date == date.today().strftime('%Y/%m/%d')])}"
        
        # === ADMIN 手動結算 ===
        elif message_text == "立即結算" and user_role == "ADMIN":
            print("💰 手動結算")
            msg = calculate_period_summary()
            if msg:
                reply_text = msg
            else:
                reply_text = "❌ 無法取得結算資料"
        
        # === 發送回覆 ===
        if reply_text:
            try:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))
                print(f"✅ 已回覆: {reply_text[:30]}")
            except Exception as e:
                print(f"❌ 回覆失敗: {e}")
        else:
            # 即使沒有處理，也不回覆（避免 reply token 錯誤）
            print(f"⚠️ 未識別的指令，不回覆")
        
    except Exception as e:
        print(f"❌ 處理錯誤: {e}")
        # 發生錯誤時不要嘗試回覆，避免 Invalid reply token

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 啟動伺服器 port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)