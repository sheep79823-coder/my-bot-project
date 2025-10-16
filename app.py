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
GOOGLE_SHEETS_CREDENTIALS_JSON = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')

# 權限管理
ADMIN_USER_IDS = ["U724ac19c55418145a5af5aa1af558cbb"]
MANAGER_USER_IDS = [
    "Uc6aab7ac59f36d31c963c8357c0e19da", 
    "Uac143535b8d18cbf93a6fc5f83054e5f", 
    "Uaa8464a6b973709e941e2c6a3fd51441"
]

GOOGLE_SHEET_NAME = "我的工務助理資料庫"
WORKSHEET_NAME = "出勤總表"
ATTENDANCE_SHEET_NAME = "出勤時數計算"
DAILY_SUMMARY_SHEET = "每日統整"

# [優化] Session 管理設定
MAX_SESSIONS = 100  # 最多保留 100 個 Session
SESSION_EXPIRE_DAYS = 7  # Session 保留 7 天
CLEANUP_INTERVAL_HOURS = 10  # 每 10 小時清理一次

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
def write_person_to_sheet(work_date, project_name, person_name, sign_in_time, note=""):
    """立即寫入簽到記錄"""
    if not attendance_sheet:
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
    scheduler.start()
    print("✅ 已啟動排程器")
    print(f"   - 每日 22:00 (台灣時間) 統整出勤")
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
        summary = f"📋 {self.work_date} - {self.project_name}\n"
        summary += f"👥 目前人數: {len(self.staff)} 人"
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
    
    for session_key, session in session_states.items():
        if session.work_date == work_date:
            if role == "ADMIN":
                accessible_sessions.append(session)
            elif role == "MANAGER" and session.is_authorized(user_id):
                accessible_sessions.append(session)
    
    if project_name:
        for session in accessible_sessions:
            if session.project_name == project_name:
                return session
        return None
    
    if len(accessible_sessions) == 0:
        return None
    elif len(accessible_sessions) == 1:
        return accessible_sessions[0]
    else:
        return max(accessible_sessions, key=lambda s: s.created_time)

# 解析函式
def parse_full_attendance_report(text):
    """解析完整日報"""
    try:
        lines = text.strip().split('\n')
        if len(lines) < 2:
            return None
        
        date_match = re.match(r"^(\d{3}/\d{2}/\d{2})", lines[0])
        if not date_match:
            return None
        work_date = date_match.group(1)
        
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
    match = re.search(r"(?：離場|下班)[:：]\s*(.+?)@(.+?)$", text.strip())
    if match:
        return {"name": match.group(1).strip(), "project": match.group(2).strip()}
    
    match = re.search(r"(?：離場|下班)[:：]\s*(.+?)$", text.strip())
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
        
        # 權限檢查
        user_role = get_user_role(user_id)
        if not user_role:
            return
                
        # 完整日報
        if re.search(r"\d{3}/\d{2}/\d{2}", message_text) and any(char in message_text for char in ["人員", "出工"]):
            report_data = parse_full_attendance_report(message_text)
            if report_data:
                session = get_or_create_session(report_data['date'], report_data['project_name'], user_id)
                session.project_name = report_data['project_name']
                
                for staff in report_data['staff']:
                    session.add_staff_and_write(staff['name'], staff['note'], message_time)
                
                reply_text = session.get_summary()
                reply_text += "\n✅ 已寫入 Google Sheets"
        
        # 新增人員
        elif "新增" in message_text:
            staff_info = parse_add_staff(message_text)
            if staff_info:
                valid_session = find_session_for_user(user_id, staff_info.get('project'))
                if valid_session:
                    if valid_session.add_staff_and_write(staff_info['name'], staff_info['note'], message_time):
                        reply_text = f"✅ 已新增 {staff_info['name']}"
        
        # 單筆離場
        elif ("離場:" in message_text or "離場：" in message_text or 
              "下班:" in message_text or "下班：" in message_text):
            checkout_info = parse_checkout_staff(message_text)
            if checkout_info:
                valid_session = find_session_for_user(user_id, checkout_info.get('project'))
                if valid_session:
                    person_data = next((p for p in valid_session.staff if p['name'] == checkout_info['name']), None)
                    if person_data:
                        if update_person_checkout(valid_session.work_date, checkout_info['name'], 
                                                 message_time, person_data['add_time']):
                            reply_text = f"✅ 已記錄 {checkout_info['name']} 離場"
        
        # 通用離場
        elif "人員離場" in message_text or "人員下班" in message_text:
            project_match = re.search(r"@(.+?)$", message_text)
            project_name = project_match.group(1).strip() if project_match else None
            valid_session = find_session_for_user(user_id, project_name)
            
            if valid_session:
                default_checkout_time = message_time.replace(hour=17, minute=30)
                count = 0
                for person in valid_session.staff:
                    if update_person_checkout(valid_session.work_date, person['name'], 
                                            default_checkout_time, person['add_time']):
                        count += 1
                reply_text = f"✅ 已記錄 {count} 人離場 (17:30)"
        
        # 查詢出勤
        elif message_text == "查詢本期出勤":
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
                        reply_text = f"📅 本期統計：\n"
                        for _, row in summary.iterrows():
                            reply_text += f"• {row['姓名']}: {row['出勤時數']} 天\n"
                    else:
                        reply_text = "查詢範圍內無出勤記錄"
                except Exception as e:
                    reply_text = f"❌ 查詢失敗: {str(e)}"
        
        if reply_text:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))
        
    except Exception as e:
        print(f"❌ 處理錯誤: {e}")

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 啟動伺服器 port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)