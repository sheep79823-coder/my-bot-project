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

try:
    creds_json = json.loads(GOOGLE_SHEETS_CREDENTIALS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(creds_json, scopes=scope)
    gsheet_client = gspread.authorize(creds)
    worksheet = gsheet_client.open(GOOGLE_SHEET_NAME).worksheet(WORKSHEET_NAME)
    print("Google Sheets 連線成功！")
except Exception as e:
    print(f"Google Sheets 連線失敗: {e}")
    worksheet = None

# --- [新增] Keep-Alive 機制 ---
def keep_alive():
    """定期 ping 自己以防止服務休眠"""
    while True:
        try:
            time.sleep(840)  # 每 14 分鐘 ping 一次（Render 15分鐘休眠前）
            import urllib.request
            render_url = os.environ.get('RENDER_URL', 'https://my-bot-project-1.onrender.com')
            try:
                urllib.request.urlopen(f"{render_url}/health", timeout=5)
                print("[KEEPALIVE] 成功 ping 服務，防止休眠")
            except:
                print("[KEEPALIVE] Ping 失敗，但繼續運行")
        except Exception as e:
            print(f"[KEEPALIVE] 錯誤: {e}")

# 在背景啟動 Keep-Alive 執行緒
keep_alive_thread = threading.Thread(target=keep_alive, daemon=True)
keep_alive_thread.start()

# --- [新增] 健康檢查端點 ---
@app.route("/health", methods=['GET'])
def health_check():
    """健康檢查端點，用於 Keep-Alive"""
    return 'OK', 200

# --- 核心解析函式 ---
def parse_attendance_report(text):
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
        if not match: return None
        data = match.groupdict()
        staff_list = []
        for line in data['names_block'].strip().splitlines():
            if not line.strip(): continue
            clean_line = re.sub(r"^\d+\.", "", line).strip()
            note_match = re.search(r"\((.+)\)", clean_line)
            if note_match:
                note = note_match.group(1)
                name = clean_line[:note_match.start()].strip()
                staff_list.append({"name": name, "note": note})
            else:
                staff_list.append({"name": clean_line, "note": None})
        result = {
            "date": data["date"], "project_name": data["project_name"].strip(),
            "head_count": int(data["head_count"]), "lunch_box_count": int(data["lunch_box_count"]),
            "staff": staff_list
        }
        return result
    except Exception as e:
        print(f"解析完整日報時發生錯誤: {e}")
        return None

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

# --- LINE Webhook 主要進入點 ---
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

# --- 訊息處理總管 ---
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    message_text = event.message.text.strip()
    
    print(f"[DEBUG] 收到訊息 - User ID: {user_id}, Message: {message_text}")
    
    if user_id not in ALLOWED_USER_IDS:
        print(f"[DEBUG] User {user_id} 未在白名單中")
        return

    reply_text = "無法識別的指令或格式錯誤。"

    if "出工人員：" in message_text:
        report_data = parse_attendance_report(message_text)
        if report_data and worksheet:
            now_str = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime('%Y-%m-%d %H:%M:%S')
            for person in report_data['staff']:
                new_row = [
                    now_str, report_data['date'], report_data['project_name'],
                    person['name'], person['note'] or "",
                    report_data['head_count'], report_data['lunch_box_count']
                ]
                worksheet.append_row(new_row)
            reply_text = f"✅ 已成功紀錄 {report_data['date']} 的完整日報至 Google Sheets。"
        else:
            reply_text = "❌ 日報格式錯誤或 Google Sheets 連線失敗。"

    elif message_text == "查詢本期出勤":
        if worksheet:
            start_date, end_date = get_current_period_dates()
            records = worksheet.get_all_records()
            if not records:
                reply_text = "試算表中沒有任何資料可供統計。"
            else:
                df = pd.DataFrame(records)
                df['gregorian_date'] = pd.to_datetime(df['日期'].apply(minguo_to_gregorian), errors='coerce')
                period_df = df.dropna(subset=['gregorian_date'])
                period_df = period_df[
                    (period_df['gregorian_date'] >= pd.to_datetime(start_date)) &
                    (period_df['gregorian_date'] <= pd.to_datetime(end_date))
                ]

                if not period_df.empty:
                    attendance_count = period_df.groupby('姓名').size().reset_index(name='天數')
                    reply_text = f"本期 ({start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}) 出勤統計：\n"
                    for index, row in attendance_count.iterrows():
                        reply_text += f"- {row['姓名']}: {row['天數']} 天\n"
                else:
                    reply_text = f"在 {start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')} 區間內找不到任何出勤紀錄。"
        else:
            reply_text = "Google Sheets 連線失敗，無法查詢。"
    
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

# --- 啟動伺服器 ---
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)