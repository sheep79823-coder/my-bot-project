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

# --- 初始設定 (與之前相同) ---
app = Flask(__name__)

YOUR_CHANNEL_ACCESS_TOKEN = os.environ.get('YOUR_CHANNEL_ACCESS_TOKEN')
YOUR_CHANNEL_SECRET = os.environ.get('YOUR_CHANNEL_SECRET')
GOOGLE_SHEETS_CREDENTIALS_JSON = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')

ALLOWED_USER_IDS = ["請在這裡貼上您自己的LINE_USER_ID"] 
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

# --- 核心解析函式 (與之前相同) ---
def parse_attendance_report(text):
    # ... (此處省略，與之前版本相同) ...
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

# --- [新增] 日期處理輔助函式 ---

def get_current_period_dates():
    """根據今天的日期，計算出當前統計期間的開始與結束日期"""
    today = date.today()
    
    # 情況一：本月 6號 ~ 20號，統計本月 6號到 20號
    if 6 <= today.day <= 20:
        start_date = today.replace(day=6)
        end_date = today.replace(day=20)
    # 情況二：本月 21號 ~ 月底，統計本月 21號到下個月 5號
    elif today.day >= 21:
        start_date = today.replace(day=21)
        # 計算下個月的年份和月份
        next_month_year = today.year if today.month < 12 else today.year + 1
        next_month = today.month + 1 if today.month < 12 else 1
        end_date = date(next_month_year, next_month, 5)
    # 情況三：本月 1號 ~ 5號，統計上個月 21號到這個月 5號
    else: # today.day <= 5
        end_date = today.replace(day=5)
        # 計算上個月的年份和月份
        last_month_year = today.year if today.month > 1 else today.year - 1
        last_month = today.month - 1 if today.month > 1 else 12
        start_date = date(last_month_year, last_month, 21)
        
    return start_date, end_date

def minguo_to_gregorian(minguo_str):
    """將民國年字串 '114/10/10' 轉換為西元年 date 物件"""
    try:
        parts = minguo_str.split('/')
        minguo_year, month, day = [int(p) for p in parts]
        gregorian_year = minguo_year + 1911
        return date(gregorian_year, month, day)
    except (ValueError, TypeError):
        return None


# --- LINE Webhook 主要進入點 (與之前相同) ---
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

# --- [修改] 訊息處理總管 ---
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    message_text = event.message.text.strip()
    
    if user_id not in ALLOWED_USER_IDS:
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
            reply_text = f"已成功紀錄 {report_data['date']} 的完整日報至 Google Sheets。"
        else:
            reply_text = "日報格式錯誤或 Google Sheets 連線失敗。"

    # --- [新增] 新的統計指令 ---
    elif message_text == "查詢本期出勤":
        if worksheet:
            # 1. 取得當期統計區間
            start_date, end_date = get_current_period_dates()
            
            # 2. 從 Google Sheets 讀取所有資料
            records = worksheet.get_all_records()
            if not records:
                reply_text = "試算表中沒有任何資料可供統計。"
            else:
                df = pd.DataFrame(records)
                
                # 3. 將民國日期轉換為可比較的西元日期
                # pd.to_datetime 搭配 errors='coerce' 會在轉換失敗時填入 NaT (Not a Time)
                df['gregorian_date'] = pd.to_datetime(df['日期'].apply(minguo_to_gregorian), errors='coerce')
                
                # 4. 篩選出在統計區間內的紀錄
                period_df = df.dropna(subset=['gregorian_date']) # 移除日期格式錯誤的行
                period_df = period_df[
                    (period_df['gregorian_date'] >= pd.to_datetime(start_date)) &
                    (period_df['gregorian_date'] <= pd.to_datetime(end_date))
                ]

                if not period_df.empty:
                    # 5. 按姓名分組計數
                    attendance_count = period_df.groupby('姓名').size().reset_index(name='天數')
                    
                    # 6. 格式化回覆文字
                    reply_text = f"本期 ({start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')}) 出勤統計：\n"
                    for index, row in attendance_count.iterrows():
                        reply_text += f"- {row['姓名']}: {row['天數']} 天\n"
                else:
                    reply_text = f"在 {start_date.strftime('%Y/%m/%d')} ~ {end_date.strftime('%Y/%m/%d')} 區間內找不到任何出勤紀錄。"
        else:
            reply_text = "Google Sheets 連線失敗，無法查詢。"
    
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

# --- 啟動伺服器 (與之前相同) ---
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)