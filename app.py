import os
import re
import json
from datetime import datetime, timedelta
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import requests
import io
import openpyxl

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
LINE_CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
GOOGLE_DRIVE_API_KEY = os.environ.get("GOOGLE_DRIVE_API_KEY", GEMINI_API_KEY)
SCHEDULE_FOLDER_ID = os.environ.get("SCHEDULE_FOLDER_ID", "1xoPDyT53_5OmrwiNnBmWcZ6t0JfgOCHj")
FALLBACK_SCHEDULE_FILE_ID = os.environ.get("SCHEDULE_FILE_ID", "1JuUl9oPPCbNA-KiGlWAqchatd06Ml4Zk")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)


def call_gemini(prompt):
    """直接用 HTTP 呼叫 Gemini API，不需要 SDK"""
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}]
    }
    # 先試 query param（標準 API key 方式）
    resp = requests.post(
        url,
        params={"key": GEMINI_API_KEY},
        json=payload,
        timeout=30
    )
    if resp.status_code == 200:
        data = resp.json()
        return data["candidates"][0]["content"]["parts"][0]["text"]
    # 如果 query param 失敗，試 Bearer token（AQ.xxx 可能是 OAuth token）
    if resp.status_code in (401, 403):
        resp2 = requests.post(
            url,
            headers={"Authorization": f"Bearer {GEMINI_API_KEY}"},
            json=payload,
            timeout=30
        )
        if resp2.status_code == 200:
            data = resp2.json()
            return data["candidates"][0]["content"]["parts"][0]["text"]
    return None


def get_current_week_pattern():
    today = datetime.now()
    weekday = today.weekday()  # 0=週一, 6=週日
    if weekday == 6:
        monday = today + timedelta(days=1)
    else:
        monday = today - timedelta(days=weekday)
    saturday = monday + timedelta(days=5)
    return f"{monday.strftime('%m%d')}-{saturday.strftime('%m%d')}"


def get_current_week_file_id():
    week_pattern = get_current_week_pattern()
    url = "https://www.googleapis.com/drive/v3/files"
    params = {
        "q": f"'{SCHEDULE_FOLDER_ID}' in parents and trashed=false",
        "key": GOOGLE_DRIVE_API_KEY,
        "orderBy": "createdTime desc",
        "fields": "files(id,name)",
        "pageSize": 20,
    }
    try:
        resp = requests.get(url, params=params, timeout=10)
        if resp.status_code == 200:
            files = resp.json().get("files", [])
            for f in files:
                if week_pattern in f["name"]:
                    return f["id"]
            for f in files:
                if "課程表" in f["name"] or f["name"].lower().endswith(".xlsx"):
                    return f["id"]
    except Exception:
        pass
    return FALLBACK_SCHEDULE_FILE_ID


def get_latest_schedule_text():
    file_id = get_current_week_file_id()
    if not file_id:
        return None
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    resp = requests.get(url, timeout=30)
    if resp.status_code != 200:
        return None

    buf = io.BytesIO(resp.content)
    wb = openpyxl.load_workbook(buf, data_only=True)

    text_parts = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        text_parts.append(f"\n=== {sheet_name} ===")
        for row in ws.iter_rows(values_only=True):
            row_text = "\t".join(str(c) if c is not None else "" for c in row)
            if row_text.strip():
                text_parts.append(row_text)

    return "\n".join(text_parts)


def query_schedule(name, schedule_text):
    prompt = f"""以下是補習班課表資料：

{schedule_text}

請找出「{name}」的相關資訊：
1. 如果「{name}」是學生，列出他/她本週所有上課時間和科目
2. 如果「{name}」是老師，列出這週要教的所有學生、科目和上課時間
3. 如果兩者都有，分開列出
4. 如果完全找不到，回覆「找不到「{name}」的課表資料，請確認姓名是否正確」

請用清楚的格式列出，只回覆課表資訊，不要加其他說明。"""
    return call_gemini(prompt)


def handle_message(text):
    text = text.strip()

    match_schedule = re.match(r"本週上課時間[，,]\s*(.+)", text)
    if match_schedule:
        name = match_schedule.group(1).strip()
        schedule_text = get_latest_schedule_text()
        if not schedule_text:
            return "目前無法讀取課表，請稍後再試。"
        result = query_schedule(name, schedule_text)
        if not result:
            return "Gemini API 暫時無法使用，請稍後再試。"
        return result

    match_material = re.match(r"找講義[，,]\s*(.+)", text)
    if match_material:
        keywords_str = match_material.group(1).strip()
        query = "+".join(keywords_str.split() if " " in keywords_str else [keywords_str])
        search_url = f"https://drive.google.com/drive/search?q={query}"
        return f"📚 請點以下連結搜尋講義：\n{search_url}"

    if re.search(r"(幫助|help|指令|怎麼用)", text, re.IGNORECASE):
        return (
            "📖 使用方式：\n\n"
            "查上課時間：\n「本週上課時間，姓名」\n"
            "例：本週上課時間，王小明\n\n"
            "找講義：\n「找講義，版本書名」\n"
            "例：找講義，翰林國一數學"
        )

    return None


@app.route("/webhook", methods=["POST"])
def webhook():
    signature = request.headers["X-Line-Signature"]
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return "OK"


@handler.add(MessageEvent, message=TextMessage)
def on_message(event):
    reply = handle_message(event.message.text)
    if reply:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=reply)
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
