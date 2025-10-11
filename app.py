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

# --- 初始設定 ---
app = Flask(__name__)

YOUR_CHANNEL_ACCESS_TOKEN = os.environ.get('YOUR_CHANNEL_ACCESS_TOKEN')
YOUR_CHANNEL_SECRET = os.environ.get('YOUR_CHANNEL_SECRET')
GOOGLE_SHEETS_CREDENTIALS_JSON = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')

ALLOWED_USER_IDS = ["U724ac19c55418145a5af5aa1af558cbb"]
GOOGLE_SHEET_NAME = "我的工務助理資料庫"
WORKSHEET_NAME = "出勤總表"
ATTENDANCE_SHEET_NAME = "出勤時數計算"  # [新增] 新的工作表用於儲存時數

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
    
    # [新增] 嘗試取得出勤時數表，如果沒有則建立
    try:
        attendance_sheet = gsheet_client.open(GOOGLE_SHEET_NAME).worksheet(ATTENDANCE_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        workbook = gsheet_client.open(GOOGLE_SHEET_NAME)
        attendance_sheet = workbook.add_worksheet(title=ATTENDANCE_SHEET_NAME, rows=1000, cols=10)
        # 設定標題列
        headers = ["日期", "姓名", "簽到時間", "離場時間", "出勤時數", "備註", "更新時間"]
        attendance_sheet.append_row(headers)
        print("✅ 已建立新的出勤時數計算表")
    
    print("✅ Google Sheets 連線成功！")
except Exception as e:
    print(f"❌ Google Sheets 連線失敗: {e}")
    worksheet = None
    attendance_sheet = None

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

# --- [改進] 對話狀態管理 - 加入時間追蹤 ---
class DailySession:
    def __init__(self, user_id, work_date):
        self.user_id = user_id
        self.work_date = work_date
        self.project_name = None
        self.staff = []  # [{"name": "", "add_time": datetime, "remove_time": None, "note": ""}]
        self.head_count = 0
        self.lunch_count = 0
        self.submitted = False
        self.created_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        self.end_time = None
    
    def add_staff(self, name, note=None, add_time=None):
        if add_time is None:
            add_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        
        if name not in [s['name'] for s in self.staff]:
            self.staff.append({
                "name": name,
                "add_time": add_time,
                "remove_time": None,
                "note": note
            })
            self.head_count += 1
            self.lunch_count += 1
            return True
        return False
    
    def remove_staff(self, name, remove_time=None):
        if remove_time is None:
            remove_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        
        for person in self.staff:
            if person['name'] == name:
                person['remove_time'] = remove_time
                return True
        return False
    
    def set_end_time(self, end_time=None):
        if end_time is None:
            end_time = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
        self.end_time = end_time
    
    def calculate_attendance_days(self, person):
        """根據時間計算出勤天數"""
        add_time = person['add_time']
        remove_time = person['remove_time'] if person['remove_time'] else self.end_time
        
        if not add_time or not remove_time:
            return 0, "時間不完整"
        
        add_hour = add_time.hour
        remove_hour = remove_time.hour
        
        # 邏輯：10:00後新增算1天，13:00後新增算0.5天
        #      12:00前離場算0.5天，其他情況特別備註
        
        if add_hour < 10:
            if remove_hour >= 12:
                return 1.0, ""  # 整天
            elif remove_hour >= 10:
                return 0.5, "上班但早退"
            else:
                return 0.5, "上午半天"
        elif add_hour < 13:
            if remove_hour >= 13:
                return 1.0, ""
            else:
                return 0.5, "上午半天"
        else:  # 13:00後新增
            if remove_hour >= 13:
                return 0.5, "下午半天"
            else:
                return 0.5, f"({add_time.strftime('%H:%M')}-{remove_time.strftime('%H:%M')})"
    
    def get_summary(self):
        summary = f"📋 {self.work_date}\n"
        summary += f"👥 目前人數: {self.head_count} 人\n"
        summary += "人員:\n"
        for i, person in enumerate(self.staff, 1):
            remove_status = " ✓已離場" if person['remove_time'] else ""
            summary += f"  {i}. {person['name']}{remove_status}\n"
        return summary

def get_or_create_session(user_id, work_date):
    session_key = f"{user_id}_{work_date}"
    if session_key not in session_states:
        session_states[session_key] = DailySession(user_id, work_date)
    return session_states[session_key]

def parse_full_attendance_report(text):
    try:
        pattern = re.compile(
            r"^(?P<date>\d{3}/\d{2}/\d{2})\n"
            r"(?P<project_name>.+)\n\n"
            r"出工人員：\n"
            r"(?P<names_block>(?:.*\n)+?)\n"
            r"共計：(?P<head_count>\d+)人\n\n"
            r"便當：(?P<lunch_box_count>\d+)個$",
            re.MULTILINE
        )
        match = pattern.search(text.strip())
        if not match:
            return None
            
        data = match.groupdict()
        staff_list = []
        
        for line in data['names_block'].strip().splitlines():
            if not line.strip():
                continue
            clean_line = re.sub(r"^\d+\.", "", line).strip()
            note_match = re.search(r"\((.+)\)", clean_line)
            
            if note_match:
                note = note_match.group(1)
                name = clean_line[:note_match.start()].strip()
                staff_list.append({"name": name, "note": note})
            else:
                staff_list.append({"name": clean_line, "note": None})
        
        return {
            "date": data["date"],
            "project_name": data["project_name"].strip(),
            "head_count": int(data["head_count"]),
            "lunch_box_count": int(data["lunch_box_count"]),
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

def submit_session_to_attendance_sheet(session):
    """將對話狀態中的資料提交到出勤時數計算表"""
    if not attendance_sheet or not session.staff:
        return False
    
    try:
        update_time = datetime.datetime.now(
            datetime.timezone(datetime.timedelta(hours=8))
        ).strftime('%Y-%m-%d %H:%M:%S')
        
        rows_added = 0
        for person in session.staff:
            days, remark = session.calculate_attendance_days(person)
            
            new_row = [
                session.work_date,
                person['name'],
                person['add_time'].strftime('%H:%M') if person['add_time'] else "",
                person['remove_time'].strftime('%H:%M') if person['remove_time'] else "",
                days,
                remark,
                update_time
            ]
            attendance_sheet.append_row(new_row)
            rows_added += 1
        
        session.submitted = True
        print(f"✅ 成功提交 {rows_added} 人的出勤時數")
        return True
    except Exception as e:
        print(f"❌ 提交錯誤: {e}")
        return False

def get_current_period_dates():
    today = date.today()
    if 6 <= today.day <= 20:
        start_date = today.replace(day=6)
        end_date = today.replace(day=20)
    elif today.day >= 21:
        start_date = today.replace(day=21)
        next_month_year = today.year if today.month < 12 else today.year + 1
        next_month = today.month + 1 if today.month < 12 else 1
        end_date = date(next_month_year, next_month, 5)
    else:
        end_date = today.replace(day=5)
        last_month_year = today.year if today.month > 1 else today.year - 1
        last_month = today.month - 1 if today.month > 1 else 12
        start_date = date(last_month_year, last_month, 21)
    return start_date, end_date

def minguo_to_gregorian(minguo_str):
    try:
        parts = minguo_str.split('/')
        minguo_year, month, day = [int(p) for p in parts]
        gregorian_year = minguo_year + 1911
        return date(gregorian_year, month, day)
    except (ValueError, TypeError):
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
        
        if is_duplicate_message(user_id, message_text, timestamp):
            return
        
        if user_id not in ALLOWED_USER_IDS:
            return

        reply_text = "無法識別的指令或格式錯誤。"

        # --- 完整日報提交 ---
        if "出工人員：" in message_text:
            print("📝 檢測到完整日報")
            report_data = parse_full_attendance_report(message_text)
            
            if report_data:
                session = get_or_create_session(user_id, report_data['date'])
                session.project_name = report_data['project_name']
                session.head_count = report_data['head_count']
                session.lunch_count = report_data['lunch_box_count']
                
                for staff in report_data['staff']:
                    session.add_staff(staff['name'], staff['note'], add_time=message_time)
                
                reply_text = session.get_summary()
                reply_text += "\n✅ 已記錄初始日報\n提示: 發送 '新增：名字' 或 '人員離場' 來更新"
            else:
                reply_text = "❌ 日報格式錯誤"

        # --- 新增人員 ---
        elif "新增" in message_text:
            print("➕ 檢測到新增人員")
            date_match = re.search(r"(\d{3}/\d{2}/\d{2})", message_text)
            work_date = date_match.group(1) if date_match else None
            
            if not work_date:
                today = date.today()
                minguo_year = today.year - 1911
                work_date = f"{minguo_year:03d}/{today.month:02d}/{today.day:02d}"
            
            session = get_or_create_session(user_id, work_date)
            staff_info = parse_add_staff(message_text)
            
            if staff_info and session.project_name:
                if session.add_staff(staff_info['name'], staff_info['note'], add_time=message_time):
                    reply_text = f"✅ 已新增 {staff_info['name']} (時間: {message_time.strftime('%H:%M')})\n\n" + session.get_summary()
                else:
                    reply_text = f"⚠️ {staff_info['name']} 已在清單中"
            else:
                reply_text = "❌ 請先提交完整日報或格式錯誤"

        # --- 人員離場/記錄結束 ---
        elif "人員離場" in message_text or "人員下班" in message_text:
            print("⬜ 檢測到記錄結束")
            date_match = re.search(r"(\d{3}/\d{2}/\d{2})", message_text)
            work_date = date_match.group(1) if date_match else None
            
            if not work_date:
                today = date.today()
                minguo_year = today.year - 1911
                work_date = f"{minguo_year:03d}/{today.month:02d}/{today.day:02d}"
            
            session = get_or_create_session(user_id, work_date)
            session.set_end_time(message_time)
            
            if session.staff and not session.submitted:
                if submit_session_to_attendance_sheet(session):
                    summary = "✅ 已提交出勤紀錄\n\n"
                    for person in session.staff:
                        days, remark = session.calculate_attendance_days(person)
                        summary += f"{person['name']}: {days} 天"
                        if remark:
                            summary += f" ({remark})"
                        summary += "\n"
                    reply_text = summary
                else:
                    reply_text = "❌ 提交失敗"
            else:
                reply_text = "❌ 無資料可提交或已提交過"

        # --- 查詢本期出勤 ---
        elif message_text == "查詢本期出勤":
            print("📊 查詢本期出勤")
            if attendance_sheet:
                try:
                    start_date, end_date = get_current_period_dates()
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
                            attendance_summary = period_df.groupby('姓名')['出勤時數'].sum().reset_index()
                            reply_text = f"📅 本期 ({start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}) 出勤時數統計：\n"
                            for _, row in attendance_summary.iterrows():
                                reply_text += f"• {row['姓名']}: {row['出勤時數']} 天\n"
                        else:
                            reply_text = "查詢範圍內無出勤紀錄"
                    else:
                        reply_text = "試算表中沒有任何資料"
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