from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import customtkinter as ctk
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv
import os
from datetime import datetime, timedelta
import pytz
# Webスクレイピング処理を行う関数


load_dotenv()

email = os.getenv('EMAIL')

# サービスアカウントキーファイルのパスを指定する
SERVICE_ACCOUNT_FILE = '/tonton_compaire_app/creditials.json'

SCOPES = ["https://www.googleapis.com/auth/calendar"]

# ファイルパスを引数として渡す
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Google Calendar API を使うための準備
service = build('calendar', 'v3', credentials=credentials)

def insert_event_to_calendar(service, date, start_time, end_time, summary="スケジュール",color_id="6"):
    """Googleカレンダーにイベントを挿入します。"""
    event = {
        'summary': summary,
        'start': {
            'dateTime': start_time.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
        'end': {
            'dateTime': end_time.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
        'colorId': color_id,  # イベントの色を設定
    }
    event_result = service.events().insert(calendarId=email, body=event).execute()
    print(f"イベントを追加しました: {event_result.get('htmlLink')}")



def scrape_data(url):
            # ChromeDriver のパスを指定（ダウンロードしたchromedriverのパスに置き換えてください）
    driver_path = 'C:\chromedriver-win64\chromedriver.exe'  # 例: C:\\path\\to\\chromedriver.exe
    service = Service(driver_path)

    # ブラウザオプションの設定
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # ヘッドレスモードで実行（ブラウザを表示しない）

    # Chrome ブラウザを起動
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    try:
        driver = webdriver.Chrome()  # または使用するブラウザに応じて適切なWebDriverを指定してください
        driver.get(url)

        bodybox_element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'bodybox'))
        )

        inner_html_bodybox = bodybox_element.get_attribute('innerHTML')
        soup_bodybox = BeautifulSoup(inner_html_bodybox, 'html.parser')
        table_elements = soup_bodybox.find_all('table', class_='tablestyle-01')

        labels = []
        for table in table_elements:
            labels.extend(table.find_all('label'))

        label_texts = [label.get_text(strip=True) for label in labels]

        date_time_dict = {}

        for i, label_text in enumerate(label_texts):
            div_id = f'myTimelineDispDiv_{i}'
            div_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, div_id))
            )

            inner_html_div = div_element.get_attribute('innerHTML')
            soup_div = BeautifulSoup(inner_html_div, 'html.parser')
            span_elements = soup_div.find_all('span', class_=['timesel_enabled timesel_00', 'timesel_enabled timesel_30'])

            time_list = []
            if span_elements:
                for span in span_elements:
                    span_id = span.get('id')
                    if span_id:
                        time_part = span_id.split('_')[-1]
                        time_list.append(time_part)
            
            date_time_dict[label_text] = time_list

        result = []
        for date, times in date_time_dict.items():
            formatted_times = []
            for time in times:
                start_hour = int(time[:2])
                start_minute = int(time[2:])

                end_hour = start_hour
                end_minute = start_minute + 30
                if end_minute >= 60:
                    end_hour += 1
                    end_minute -= 60

                start_time = f"{start_hour:02d}{start_minute:02d}"
                end_time = f"{end_hour:02d}{end_minute:02d}"
                formatted_times.append(f"{start_time}-{end_time}")
            
            result.append(f"{date}: {', '.join(formatted_times)}")

        return date_time_dict
    
    except Exception as e:
        return f"エラーが発生しました: {e}"

    finally:
        driver.quit()



def create_gui():
    app = ctk.CTk()
    app.title("Web Scraper GUI")
    app.geometry("600x500")

    # URL入力用のエントリー
    url_label = ctk.CTkLabel(app, text="URLを入力してください:")
    url_label.pack(pady=5)
    url_entry = ctk.CTkEntry(app, width=300)
    url_entry.pack(pady=5)

    # タイトル入力用のエントリー
    title_label = ctk.CTkLabel(app, text="イベントのタイトルを入力してください:")
    title_label.pack(pady=5)
    title_entry = ctk.CTkEntry(app, width=300)
    title_entry.pack(pady=5)

    # 出力エリア
    output_text = ctk.CTkTextbox(app, width=300, height=150)
    output_text.pack(pady=10)

    # 決定ボタンのコールバック関数
    def on_submit():
        url = url_entry.get()
        title = title_entry.get()
        schedule_data = scrape_data(url)
       

        output_text.delete("1.0", ctk.END)
        for date, times in schedule_data.items():
            date_obj = datetime.strptime(date.split('(')[0], "%Y/%m/%d")
            for time in times:
                start_hour = int(time[:2])
                start_minute = int(time[2:])
                start_time = datetime(date_obj.year, date_obj.month, date_obj.day, start_hour, start_minute)
                end_time = start_time + timedelta(minutes=30)
                
                # Googleカレンダーにイベントを挿入
                insert_event_to_calendar(service, date, start_time, end_time, title)

                output_text.insert(ctk.END, f"{date} {start_time.strftime('%H:%M')}-{end_time.strftime('%H:%M')}にイベントを挿入しました。\n")

    # 決定ボタン
    submit_button = ctk.CTkButton(app, text="Google Calendarに追加", command=on_submit)
    submit_button.pack(pady=10)

    app.mainloop()

# GUIアプリケーションを起動
create_gui()