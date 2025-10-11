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

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)

# [新增] 防重複 + 對話狀態管理
processed_messages = {}
DUPLICATE_CHECK_WINDOW = 300
session_states = {}  # 儲存每個用戶的對話狀態

try:
    creds_json = json.loads(GOOGLE_SHEETS_CREDENTIALS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(creds_json, scopes=scope)
    gsheet_client = gspread.authorize(creds)
    worksheet = gsheet_client.open(GOOGLE_SHEET_NAME).worksheet(WORKSHEET_NAME)
    print("✅ Google Sheets 連線成功！")
except Exception as e:
    print(f"❌ Google Sheets 連線失敗: {e}")
    worksheet = None

# --- Keep-Alive 機制 ---
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

# --- 防重複檢查 ---
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

# --- [新增] 對話狀態管理 ---
class DailySession:
    """管理一天的出勤對話狀態"""
    def __init__(self, user_id, work_date):
        self.user_id = user_id
        self.work_date = work_date
        self.project_name = None
        self.staff = []  # 當前人員清單
        self.head_count = 0
        self.lunch_count = 0
        self.submitted = False
        self.created_time = datetime.datetime.now()
    
    def add_staff(self, name, note=None):
        """增加人員"""
        if name not in [s['name'] for s in self.staff]:
            self.staff.append({"name": name, "note": note})
            self.head_count += 1
            return True
        return False
    
    def remove_staff(self, name):
        """移除人員"""
        self.staff = [s for s in self.staff if s['name'] != name]
        self.head_count = len(self.staff)
        return True
    
    def get_summary(self):
        """取得當前摘要"""
        summary = f"📋 {self.work_date} - {self.project_name}\n"
        summary += f"👥 目前人數: {self.head_count} 人\n"
        summary += f"🍱 便當: {self.lunch_count} 個\n"
        summary += "人員:\n"
        for i, person in enumerate(self.staff, 1):
            note_str = f" ({person['note']})" if person['note'] else ""
            summary += f"  {i}. {person['name']}{note_str}\n"
        return summary

def get_or_create_session(user_id, work_date):
    """取得或建立該日期的對話狀態"""
    session_key = f"{user_id}_{work_date}"
    if session_key not in session_states:
        session_states[session_key] = DailySession(user_id, work_date)
    return session_states[session_key]

# --- 核心解析函式 ---
def parse_full_attendance_report(text):
    """解析完整的出勤日報"""
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
    """解析新增人員訊息 '新增：名字' 或 '新增: 名字 (備註)'"""
    match = re.search(r"新增[:：]\s*(.+?)(?:\s*\((.+)\))?$", text.strip())
    if match:
        name = match.group(1).strip()
        note = match.group(2).strip() if match.group(2) else None
        return {"name": name, "note": note}
    return None

def parse_remove_staff(text):
    """解析移除人員訊息 '人員離場' 或 '移除：名字'"""
    if "人員離場" in text or "人員下班" in text:
        return "all"  # 表示記錄結束
    match = re.search(r"移除[:：]\s*(.+?)$", text.strip())
    if match:
        return match.group(1).strip()
    return None

def submit_session_to_sheet(session):
    """將對話狀態中的資料提交到 Google Sheets"""
    if not worksheet or not session.staff:
        return False
    
    try:
        now_str = datetime.datetime.now(
            datetime.timezone(datetime.timedelta(hours=8))
        ).strftime('%Y-%m-%d %H:%M:%S')
        
        rows_added = 0
        for person in session.staff:
            new_row = [
                now_str,
                session.work_date,
                session.project_name,
                person['name'],
                person['note'] or "",
                session.head_count,
                session.lunch_count
            ]
            worksheet.append_row(new_row)
            rows_added += 1
        
        session.submitted = True
        print(f"✅ 成功提交 {rows_added} 筆資料")
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

# --- LINE Webhook ---
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

# --- 訊息處理 ---
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    try:
        user_id = event.source.user_id
        message_text = event.message.text.strip()
        timestamp = event.timestamp / 1000
        
        print(f"\n[新訊息] User: {user_id}, Text: {message_text}")
        
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
                session.staff = report_data['staff']
                
                reply_text = session.get_summary()
                reply_text += "\n✅ 已記錄初始日報\n提示: 發送 '新增：名字' 或 '人員離場' 來更新"
            else:
                reply_text = "❌ 日報格式錯誤"

        # --- 新增人員 ---
        elif "新增" in message_text:
            print("➕ 檢測到新增人員")
            # 嘗試從訊息中提取日期（如果有）
            date_match = re.search(r"(\d{3}/\d{2}/\d{2})", message_text)
            work_date = date_match.group(1) if date_match else None
            
            if not work_date:
                today = date.today()
                minguo_year = today.year - 1911
                work_date = f"{minguo_year:03d}/{today.month:02d}/{today.day:02d}"
            
            session = get_or_create_session(user_id, work_date)
            staff_info = parse_add_staff(message_text)
            
            if staff_info and session.project_name:
                if session.add_staff(staff_info['name'], staff_info['note']):
                    session.lunch_count += 1
                    reply_text = f"✅ 已新增 {staff_info['name']}\n\n" + session.get_summary()
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
            
            if session.staff and not session.submitted:
                if submit_session_to_sheet(session):
                    reply_text = f"✅ 已提交 {len(session.staff)} 人的出勤紀錄至 Google Sheets"
                else:
                    reply_text = "❌ 提交失敗"
            else:
                reply_text = "❌ 無資料可提交或已提交過"

        # --- 查詢本期出勤 ---
        elif message_text == "查詢本期出勤":
            print("📊 查詢本期出勤")
            if worksheet:
                try:
                    start_date, end_date = get_current_period_dates()
                    records = worksheet.get_all_records()
                    
                    if records:
                        df = pd.DataFrame(records)
                        df['gregorian_date'] = pd.to_datetime(
                            df['日期'].apply(minguo_to_gregorian),
                            errors='coerce'
                        )
                        
                        period_df = df.dropna(subset=['gregorian_date'])
                        period_df = period_df[
                            (period_df['gregorian_date'] >= pd.to_datetime(start_date)) &
                            (period_df['gregorian_date'] <= pd.to_datetime(end_date))
                        ]

                        if not period_df.empty:
                            attendance_count = period_df.groupby('姓名').size().reset_index(name='天數')
                            reply_text = f"📅 本期 ({start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}) 出勤統計：\n"
                            for _, row in attendance_count.iterrows():
                                reply_text += f"• {row['姓名']}: {row['天數']} 天\n"
                        else:
                            reply_text = f"查詢範圍內無出勤紀錄"
                    else:
                        reply_text = "試算表中沒有任何資料"
                except Exception as e:
                    reply_text = f"❌ 查詢失敗: {str(e)}"
            else:
                reply_text = "❌ Google Sheets 連線失敗"
        
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))
        
    except Exception as e:
        print(f"❌ 處理錯誤: {e}")

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 啟動伺服器 port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)