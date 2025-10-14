# --- 引用所有必要的函式庫 ---
import os
import re
import json
import datetime
from datetime import date, timedelta
import gspread
import pandas as pd
from flask import Flask, request, abort
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

ALLOWED_USER_IDS = ["U724ac19c55418145a5af5aa1af558cbb",
    "Uc6aab7ac59f36d31c963c8357c0e19da", 
    "Uac143535b8d18cbf93a6fc5f83054e5f", 
    "Uaa8464a6b973709e941e2c6a3fd51441"]
GOOGLE_SHEET_NAME = "我的工務助理資料庫"
WORKSHEET_NAME = "出勤總表"
ATTENDANCE_SHEET_NAME = "出勤時數計算"
DAILY_SUMMARY_SHEET = "每日統整"

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)

processed_messages = {}
DUPLICATE_CHECK_WINDOW = 300
session_states = {}

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

def keep_alive():
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
    return 'OK', 200

def is_duplicate_message(user_id, message_text, timestamp):
    msg_hash = hashlib.md5(f"{user_id}{message_text}{timestamp}".encode()).hexdigest()
    current_time = time.time()
    to_delete = [k for k, v in processed_messages.items() if current_time - v > DUPLICATE_CHECK_WINDOW]
    for k in to_delete:
        del processed_messages[k]
    if msg_hash in processed_messages:
        return True
    processed_messages[msg_hash] = current_time
    return False

# --- 即時寫入 Google Sheets 的函式 ---
def write_person_to_sheet(work_date, project_name, person_name, sign_in_time, note=""):
    """立即將一個人的簽到資料寫入 Google Sheets"""
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
            "",  # 離場時間（先空著）
            "",  # 出勤時數（先空著）
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
    """更新一個人的離場時間和出勤天數"""
    if not attendance_sheet:
        return False
    
    try:
        records = attendance_sheet.get_all_records()
        
        # 找到同一天該人員最後一條未完成的記錄（支持同天多專案）
        target_row = None
        for i, record in enumerate(records, start=2):
            if record['日期'] == work_date and record['姓名'] == person_name and not record['離場時間']:
                target_row = i  # 只更新最後一條未完成的記錄
        
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
            
            # 加班判定（17:00 後離場）
            if checkout_hour >= 17:
                remark = (remark + " " if remark else "") + "加班"
            
            # 更新這一行
            attendance_sheet.update_cell(target_row, 4, checkout_time.strftime('%H:%M'))  # D列 離場時間
            attendance_sheet.update_cell(target_row, 5, days)  # E列 出勤時數
            attendance_sheet.update_cell(target_row, 6, remark.strip())  # F列 備註
            print(f"✅ 已更新 {person_name} 的離場時間和出勤天數")
            return True
        
        print(f"⚠️ 找不到 {person_name} 的簽到記錄")
        return False
    except Exception as e:
        print(f"❌ 更新失敗: {e}")
        return False

# --- 每日統整函式 ---
def daily_summary():
    """每天 22:00 執行統整 - 同一天同一人只計一次"""
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
        
        # 篩選今天的記錄
        today_df = df[df['日期'] == today_str]
        
        if today_df.empty:
            print(f"ℹ️ {today_str} 沒有出勤記錄")
            return
        
        # 關鍵改進：同一天同一人只計算最高的出勤時數
        # 例如在兩個專案都簽到，選擇較高的時數
        summary_list = []
        for person_name in today_df['姓名'].unique():
            person_records = today_df[today_df['姓名'] == person_name]
            
            # 取得該人員該天所有的出勤時數
            days_list = pd.to_numeric(person_records['出勤時數'], errors='coerce').dropna().tolist()
            
            if days_list:
                # 取最高的出勤時數（如果多次簽到，取較多的）
                max_days = max(days_list)
                summary_list.append({'姓名': person_name, '總出勤天數': max_days})
        
        if not summary_list:
            print(f"ℹ️ {today_str} 沒有有效的出勤時數")
            return
        
        summary_df = pd.DataFrame(summary_list)
        
        # 寫入統整表
        update_time = datetime.datetime.now(
            datetime.timezone(datetime.timedelta(hours=8))
        ).strftime('%Y-%m-%d %H:%M:%S')
        
        for _, row in summary_df.iterrows():
            summary_row = [today_str, row['姓名'], row['總出勤天數'], update_time]
            summary_sheet.append_row(summary_row)
        
        print(f"✅ 已統整 {len(summary_df)} 人的 {today_str} 出勤資料")
        print(f"統整內容: {summary_df.to_string()}")
        
    except Exception as e:
        print(f"❌ 統整失敗: {e}")

# --- 排程設定 ---
scheduler = BackgroundScheduler(timezone='Asia/Taipei')

def start_scheduler():
    """啟動排程器 - 每天 22:00 台灣時間執行"""
    scheduler.add_job(daily_summary, 'cron', hour=22, minute=0, timezone='Asia/Taipei')
    scheduler.start()
    print("✅ 已啟動每日 22:00 (台灣時間) 統整排程")

start_scheduler()

# --- 對話狀態管理 ---
class DailySession:
    def __init__(self, user_id, work_date, project_name=""):
        self.user_id = user_id
        self.work_date = work_date
        self.project_name = project_name
        self.staff = []
        self.created_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
    
    def add_staff_and_write(self, name, note=None, add_time=None):
        """新增人員並立即寫入 Google Sheets"""
        if add_time is None:
            add_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        
        if name not in [s['name'] for s in self.staff]:
            # 立即寫入 Google Sheets
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

def get_or_create_session(user_id, work_date, project_name=None):
    """取得或建立該日期和專案的對話狀態"""
    if project_name is None:
        project_name = ""
    # 改用 (user_id, work_date, project_name) 作為 key，支持同一天多專案
    session_key = f"{user_id}_{work_date}_{project_name}"
    if session_key not in session_states:
        session_states[session_key] = DailySession(user_id, work_date, project_name)
    return session_states[session_key]

def parse_full_attendance_report(text):
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
            
            if not line:
                continue
            
            if "共計" in line or "便當" in line or "總計" in line:
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
        
        return {
            "date": work_date,
            "project_name": project_name,
            "staff": staff_list
        }
    except Exception as e:
        print(f"❌ 解析日報錯誤: {e}")
        return None

def parse_add_staff(text):
    match = re.search(r"新增[:：]\s*(.+?)(?:\s*\((.+)\))?$", text.strip())
    if match:
        name = match.group(1).strip()
        note = match.group(2).strip() if match.group(2) else None
        return {"name": name, "note": note}
    return None

def calculate_attendance_days(add_hour):
    """根據簽到時間計算出勤天數"""
    if add_hour < 10:
        return 1.0, ""
    elif add_hour < 13:
        return 1.0, ""
    else:
        return 0.5, "下午半天"

def minguo_to_gregorian(minguo_str):
    try:
        parts = minguo_str.split('/')
        minguo_year, month, day = [int(p) for p in parts]
        gregorian_year = minguo_year + 1911
        return date(gregorian_year, month, day)
    except (ValueError, TypeError):
        return None

# --- [新增] 單獨離場人員的解析函式 ---
def parse_checkout_staff(text):
    """解析 '離場:姓名' 或 '下班:姓名' 的指令"""
    match = re.search(r"(?:離場|下班)[:：]\s*(.+?)(?:\s*\((.+)\))?$"", text.strip())
    if match:
        name = match.group(1).strip()
        return {"name": name}
    return None

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
        message_time = datetime.datetime.fromtimestamp(timestamp, tz=datetime.timezone(datetime.timedelta(hours=8)))
        
        print(f"\n[新訊息] User: {user_id}, Text: {message_text}, Time: {message_time.strftime('%H:%M')}")
        


       # --- 完整日報提交 ---
        if re.search(r"\d{3}/\d{2}/\d{2}", message_text) and any(char in message_text for char in ["人員", "出工"]):
            print("📝 檢測到完整日報")
            report_data = parse_full_attendance_report(message_text)
            
            if report_data:
                session = get_or_create_session(user_id, report_data['date'])
                session.project_name = report_data['project_name']
                
                for staff in report_data['staff']:
                    session.add_staff_and_write(staff['name'], staff['note'], message_time)
                
                reply_text = session.get_summary()
                reply_text += "\n✅ 已記錄初始日報並寫入 Google Sheets\n"
                reply_text += "💡 新增人員或發送 '人員離場' 來更新\n"
                reply_text += "📊 每天 22:00 自動統整出勤時數"
            else:
                reply_text = "❌ 日報格式錯誤"

        # --- 新增人員 ---
        elif "新增" in message_text:
            print("➕ 檢測到新增人員")
            
            valid_session = None
            latest_time = None
            for session_key, session in session_states.items():
                if session.user_id == user_id and session.project_name:
                    if latest_time is None or session.created_time > latest_time:
                        valid_session = session
                        latest_time = session.created_time
            
            if valid_session:
                staff_info = parse_add_staff(message_text)
                if staff_info:
                    if valid_session.add_staff_and_write(staff_info['name'], staff_info['note'], message_time):
                        reply_text = f"✅ 已新增 {staff_info['name']} (時間: {message_time.strftime('%H:%M')})\n"
                        reply_text += f"已立即寫入 Google Sheets\n\n" + valid_session.get_summary()
                    else:
                        reply_text = f"⚠️ {staff_info['name']} 已在清單中"
                else:
                    reply_text = "❌ 新增格式錯誤，請用 '新增：名字'"
            else:
                reply_text = "❌ 請先提交完整日報"

        # --- 單獨人員離場 ---
        elif "離場:" in message_text or "下班:" in message_text:
            print("🚶 檢測到單獨人員離場")
            
            valid_session = None
            latest_time = None
            for session_key, session in session_states.items():
                if session.user_id == user_id and session.project_name:
                    if latest_time is None or session.created_time > latest_time:
                        valid_session = session
                        latest_time = session.created_time
            
            if valid_session:
                checkout_info = parse_checkout_staff(message_text)
                if checkout_info:
                    person_name = checkout_info['name']
                    
                    person_data = next((p for p in valid_session.staff if p['name'] == person_name), None)
                    
                    if person_data:
                        sign_in_time = person_data['add_time']
                        if update_person_checkout(valid_session.work_date, person_name, message_time, sign_in_time):
                            reply_text = f"✅ 已記錄 {person_name} 的離場時間 ({message_time.strftime('%H:%M')})\n"
                            reply_text += "📊 出勤時數已更新至 Google Sheets"
                        else:
                            reply_text = f"⚠️ 更新 {person_name} 的離場記錄失敗，可能已記錄過或找不到簽到資料。"
                    else:
                        reply_text = f"❌ 找不到 {person_name} 的簽到記錄，請確認姓名是否正確。"
                else:
                    reply_text = "❌ 離場格式錯誤，請用 '離場：姓名'"
            else:
                reply_text = "❌ 請先提交完整日報"

        # --- 通用人員離場 ---
        elif "人員離場" in message_text or "人員下班" in message_text:
            print("⬜ 檢測到記錄結束")
            
            valid_session = None
            latest_time = None
            for session_key, session in session_states.items():
                if session.user_id == user_id and session.project_name:
                    if latest_time is None or session.created_time > latest_time:
                        valid_session = session
                        latest_time = session.created_time
            
            if valid_session:
                updated_count = 0
                for person in valid_session.staff:
                    if update_person_checkout(valid_session.work_date, person['name'], message_time, person['add_time']):
                        updated_count += 1
                
                reply_text = f"✅ 已記錄 {updated_count} 人的離場時間\n"
                reply_text += f"專案: {valid_session.project_name}\n"
                reply_text += "📊 出勤時數已寫入 Google Sheets\n"
                reply_text += "🕙 每天 22:00 (台灣時間) 將自動統整每日出勤報告"
            else:
                reply_text = "❌ 找不到有效的日報記錄"

        # --- 查詢本期出勤 ---
        elif message_text == "查詢本期出勤":
            print("📊 查詢本期出勤")
            if attendance_sheet:
                try:
                    today = date.today()
                    if today.day <= 5:
                        start_date = (today.replace(day=1) - timedelta(days=1)).replace(day=21)
                        end_date = today.replace(day=5)
                    elif today.day <= 20:
                        end_date = today.replace(day=20)
                        start_date = today.replace(day=6)
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
                        attendance_summary = period_df.groupby('姓名')['出勤時數'].sum().reset_index()
                        reply_text = f"📅 本期 ({start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}) 出勤時數統計：\n"
                        for _, row in attendance_summary.iterrows():
                            reply_text += f"• {row['姓名']}: {row['出勤時數']} 天\n"
                    else:
                        reply_text = "查詢範圍內無出勤紀錄"
                except Exception as e:
                    reply_text = f"❌ 查詢失敗: {str(e)}"
            else:
                reply_text = "❌ 出勤表連線失敗"
        
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))
        
    except Exception as e:
        print(f"❌ 處理錯誤: {e}")

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 啟動伺服器 port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)