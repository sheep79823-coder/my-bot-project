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

ALLOWED_USER_IDS = ["U724ac19c55418145a5af5aa1af558cbb"]  # ⚠️ 改成你的真實 ID
GOOGLE_SHEET_NAME = "我的工務助理資料庫"
WORKSHEET_NAME = "出勤總表"

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)

# [新增] 防重複機制：記錄已處理過的訊息
processed_messages = {}
DUPLICATE_CHECK_WINDOW = 300  # 5 分鐘內檢查重複

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

# --- [新增] Keep-Alive 機制 ---
def keep_alive():
    """定期 ping 自己以防止服務休眠"""
    while True:
        try:
            time.sleep(840)  # 每 14 分鐘 ping 一次
            import urllib.request
            render_url = os.environ.get('RENDER_URL', 'https://my-bot-project-1.onrender.com')
            try:
                urllib.request.urlopen(f"{render_url}/health", timeout=5)
                print("[KEEPALIVE] ✅ 防止休眠")
            except:
                print("[KEEPALIVE] ⚠️ Ping 失敗，但繼續運行")
        except Exception as e:
            print(f"[KEEPALIVE] ❌ {e}")

# 啟動 Keep-Alive 執行緒
keep_alive_thread = threading.Thread(target=keep_alive, daemon=True)
keep_alive_thread.start()

# --- [新增] 健康檢查端點 ---
@app.route("/health", methods=['GET'])
def health_check():
    return 'OK', 200

# --- [新增] 防重複檢查函式 ---
def is_duplicate_message(user_id, message_text, timestamp):
    """檢查是否為重複訊息"""
    # 建立訊息的唯一識別碼
    msg_hash = hashlib.md5(f"{user_id}{message_text}{timestamp}".encode()).hexdigest()
    
    # 清理過期的記錄
    current_time = time.time()
    to_delete = [k for k, v in processed_messages.items() if current_time - v > DUPLICATE_CHECK_WINDOW]
    for k in to_delete:
        del processed_messages[k]
    
    # 檢查是否已處理過
    if msg_hash in processed_messages:
        print(f"[重複偵測] ⚠️ 檢測到重複訊息: {msg_hash}")
        return True
    
    # 記錄此訊息
    processed_messages[msg_hash] = current_time
    return False

# --- 核心解析函式 ---
def parse_attendance_report(text):
    """解析出勤日報"""
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
        
        result = {
            "date": data["date"],
            "project_name": data["project_name"].strip(),
            "head_count": int(data["head_count"]),
            "lunch_box_count": int(data["lunch_box_count"]),
            "staff": staff_list
        }
        return result
    except Exception as e:
        print(f"❌ 解析日報錯誤: {e}")
        return None

def get_current_period_dates():
    """計算當前統計期間"""
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
    """民國年轉西元年"""
    try:
        parts = minguo_str.split('/')
        minguo_year, month, day = [int(p) for p in parts]
        gregorian_year = minguo_year + 1911
        return date(gregorian_year, month, day)
    except (ValueError, TypeError):
        return None

# --- LINE Webhook 主要進入點 ---
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    
    try:
        handler.handle(body, signature)
        return 'OK', 200  # ✅ 立即回覆 200，告訴 LINE 已收到
    except InvalidSignatureError:
        print("❌ 簽名驗證失敗")
        return 'Invalid signature', 403
    except Exception as e:
        print(f"❌ Callback 處理錯誤: {e}")
        return 'Internal Server Error', 500

# --- 訊息處理總管 ---
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    try:
        user_id = event.source.user_id
        message_text = event.message.text.strip()
        timestamp = event.timestamp / 1000  # LINE 的時間戳是毫秒
        
        print(f"\n[新訊息] User: {user_id}, Text: {message_text}")
        
        # [防重複檢查]
        if is_duplicate_message(user_id, message_text, timestamp):
            print("⚠️ 已跳過重複訊息")
            return
        
        # [白名單檢查]
        if user_id not in ALLOWED_USER_IDS:
            print(f"❌ User {user_id} 未在白名單")
            return

        reply_text = "無法識別的指令或格式錯誤。"

        # --- 日報提交 ---
        if "出工人員：" in message_text:
            print("📝 檢測到日報格式")
            report_data = parse_attendance_report(message_text)
            
            if report_data and worksheet:
                try:
                    now_str = datetime.datetime.now(
                        datetime.timezone(datetime.timedelta(hours=8))
                    ).strftime('%Y-%m-%d %H:%M:%S')
                    
                    rows_added = 0
                    for person in report_data['staff']:
                        new_row = [
                            now_str,
                            report_data['date'],
                            report_data['project_name'],
                            person['name'],
                            person['note'] or "",
                            report_data['head_count'],
                            report_data['lunch_box_count']
                        ]
                        worksheet.append_row(new_row)
                        rows_added += 1
                    
                    reply_text = f"✅ 已成功新增 {rows_added} 筆資料\n日期: {report_data['date']}\n"
                    print(f"✅ 成功新增 {rows_added} 筆日報資料")
                    
                except Exception as e:
                    reply_text = f"❌ 寫入 Google Sheets 失敗: {str(e)}"
                    print(f"❌ 寫入錯誤: {e}")
            else:
                reply_text = "❌ 日報格式錯誤或 Google Sheets 連線失敗"

        # --- 查詢本期出勤 ---
        elif message_text == "查詢本期出勤":
            print("📊 查詢本期出勤")
            if worksheet:
                try:
                    start_date, end_date = get_current_period_dates()
                    records = worksheet.get_all_records()
                    
                    if not records:
                        reply_text = "試算表中沒有任何資料可供統計。"
                    else:
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
                            reply_text = f"查詢範圍 {start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')} 內無出勤紀錄"
                except Exception as e:
                    reply_text = f"❌ 查詢失敗: {str(e)}"
                    print(f"❌ 查詢錯誤: {e}")
            else:
                reply_text = "❌ Google Sheets 連線失敗"
        
        # 發送回覆
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))
        print(f"✅ 已回覆: {reply_text[:50]}...")
        
    except Exception as e:
        print(f"❌ 處理訊息時發生未預期的錯誤: {e}")

# --- 啟動伺服器 ---
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    print(f"🚀 啟動 Flask 伺服器，監聽 port {port}")
    app.run(host='0.0.0.0', port=port, debug=False)