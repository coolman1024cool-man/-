import os.path
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# 我已經從您的 Google 帳號中抓取出對應的日曆專屬 ID，建立成這個分類對照表
CALENDARS = {
    "私人": "coolman1024cool@gmail.com", # 對應「邱正彥」日曆
    "USR": "2ea36c7b744d4115471af0903de00a63f8190b7e9ac7fc900d02e96074d10b70@group.calendar.google.com", # 對應「USR工作行程」日曆
    "家庭": "bbd711188c250419e86d13b1eca81958536dc5d65f0d5d0725acac7d32b9efbe@group.calendar.google.com"  # 對應「家庭活動行程」日曆
}

SCOPES = ["https://www.googleapis.com/auth/calendar"]

def get_calendar_service():
    """取得授權過的 Google Calendar Service 物件"""
    if not os.path.exists("token.json"):
        raise FileNotFoundError("找不到 token.json，請確認是否已經完成授權。")
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    return build("calendar", "v3", credentials=creds)

def add_smart_event(summary, description, start_time, end_time, category):
    """
    依照您的定義，將行程新增到指定的分類日曆。
    參數 category 必須是 "私人", "USR", 或是 "家庭"。
    """
    service = get_calendar_service()
    
    # 根據分類尋找對應的 ID。如果分類給錯了，預設放進「私人」日曆。
    calendar_id = CALENDARS.get(category, CALENDARS["私人"])
    
    event_body = {
        'summary': summary,
        'description': description,
        'start': {
            'dateTime': start_time,
            'timeZone': 'Asia/Taipei',
        },
        'end': {
            'dateTime': end_time,
            'timeZone': 'Asia/Taipei',
        },
    }
    
    try:
        print(f"正在將行程新增至【{category}】日曆...")
        event = service.events().insert(calendarId=calendar_id, body=event_body).execute()
        print(f"新增成功！連結：{event.get('htmlLink')}")
        return event
    except HttpError as error:
        print(f"發生錯誤：{error}")
        return None

if __name__ == "__main__":
    # 這是一個測試範例：將一筆新的 USR 工作寫入「USR工作行程」日曆
    add_smart_event(
        summary="測試：USR 計畫討論會議",
        description="這是一筆用來測試「分類寫入」功能的行程。",
        start_time="2026-09-02T10:00:00+08:00",
        end_time="2026-09-02T11:00:00+08:00",
        category="USR"
    )
