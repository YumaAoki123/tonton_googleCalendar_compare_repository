from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os.path
import pickle
import customtkinter as ctk
import tkinter as tk
import threading
from datetime import datetime
import sqlite3
import uuid


SCOPES = ["https://www.googleapis.com/auth/calendar"]
# トークンを保存するファイルパス
TOKEN_PICKLE = 'token.pickle'

def get_credentials():
    creds = None
    # 既にトークンが存在する場合、それを読み込む
    if os.path.exists(TOKEN_PICKLE):
        with open(TOKEN_PICKLE, 'rb') as token:
            creds = pickle.load(token)
    
    # トークンがないか、無効または期限切れの場合、新しく認証を行う
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('/tonton_googleCalendar_compare_app/creditials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # トークンを保存
        with open(TOKEN_PICKLE, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

# Google Calendar API を使うための準備
creds = get_credentials()
service = build('calendar', 'v3', credentials=creds)

# SQLiteデータベースに接続（ファイルが存在しない場合は作成されます）
conn = sqlite3.connect('tonton_calendar_compare.db')
cursor = conn.cursor()

# テーブルを作成します（存在しない場合のみ）
cursor.execute('''
CREATE TABLE IF NOT EXISTS event_mappings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_uuid TEXT NOT NULL,
    event_id TEXT NOT NULL
)
''')

# テーブルを作成します（存在しない場合のみ）
cursor.execute('''
CREATE TABLE IF NOT EXISTS task_info (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_uuid TEXT NOT NULL,
    task_name TEXT NOT NULL
)
''')

# 接続を閉じます
conn.commit()
conn.close()

def load_tasks():
    conn = sqlite3.connect('tonton_calendar_compare.db')
    cursor = conn.cursor()
    
    # タスク情報をデータベースから読み込む
    cursor.execute('SELECT task_uuid, task_name FROM task_info')
    rows = cursor.fetchall()
    
    global tasks
    tasks = []
    
    for row in rows:
        task = {
            "task_uuid": row[0],
            "task_name": row[1],
        }
        tasks.append(task)
        print(f"Loaded task: {task}")  # デバッグ用の出力
       
    conn.close()

def insert_event_to_calendar(service, date, start_time, end_time, task_uuid, title, conn, color_id="6"):

    """Googleカレンダーにイベントを挿入します。"""
    event = {
        'summary': title,
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
    try:
        # カレンダーIDはユーザーのプライマリカレンダーを指定
        event_result = service.events().insert(calendarId="primary", body=event).execute()
        print(f"イベントを追加しました: {event_result.get('htmlLink')}")
        
        event_id = event_result.get('id')
        save_uuid_event_id_mapping(task_uuid, event_id, conn)
       
    except HttpError as error:
                print(f'An error occurred: {error}')
                raise

def save_uuid_event_id_mapping(uuid, event_id, conn):
    try:
        cursor = conn.cursor()
        cursor.execute('INSERT INTO event_mappings (task_uuid, event_id) VALUES (?, ?)', (uuid, event_id))
    except sqlite3.Error as e:
        print(f"Error saving event ID mapping: {e}")
        raise

def save_uuid_task_name(uuid, task_name, conn):
    try:
        cursor = conn.cursor()
        cursor.execute('INSERT INTO task_info (task_uuid, task_name) VALUES (?, ?)', (uuid, task_name))
    except sqlite3.Error as e:
        print(f"Error saving task: {e}")
        raise

def update_task_listbox():
    task_listbox.delete(0, ctk.END)
    for task in tasks:
        print(f"Current task: {task}")  # デバッグ用の出力
        task_listbox.insert(ctk.END, f"{task['task_name']}")


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
         # 単一のタイトルを取得する
        title_element = soup_bodybox.find('div', class_='titletext')
        if title_element:
            title_text = title_element.get_text(strip=True)
        else:
            title_text = 'No Title Found'  # タイトルが見つからない場合のデフォルト値

        print("タイトル一覧:", title_text)
        print("データ型:", type(title_text))
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
            print(f"timelist for {time_list}")

        result = []
        for date, times in date_time_dict.items():
            if not times:
                continue
            
            # 時間リストをソート
            times.sort()
            print(f"Sorted times for {date}: {times}")
            
            formatted_times = []
            
            # 各時間に対して終了時刻を30分追加する
            time_blocks = []
            for time in times:
                start_hour = int(time[:2])
                start_minute = int(time[2:])
                end_minute = start_minute + 30
                end_hour = start_hour
                if end_minute >= 60:
                    end_hour += 1
                    end_minute -= 60
                end_time = f"{end_hour:02d}{end_minute:02d}"
                time_blocks.append((f"{start_hour:02d}{start_minute:02d}", end_time))
            
            print(f"Time blocks with end times: {time_blocks}")
            
            # 時間ブロックを30分ごとの塊にまとめる
            start_time, end_time = time_blocks[0]
            for current_start_time, current_end_time in time_blocks[1:]:
                end_hour = int(end_time[:2])
                end_minute = int(end_time[2:])
                next_start_hour = int(current_start_time[:2])
                next_start_minute = int(current_start_time[2:])
                
                if end_hour == next_start_hour and end_minute == next_start_minute:
                    # 現在の時間ブロックを延長
                    end_time = current_end_time
                else:
                    # 現在の時間ブロックを追加
                    formatted_times.append(f"{start_time}-{end_time}")
                    print(f"Added time block: {start_time}-{end_time}")
                    
                    # 新しい時間ブロックの開始
                    start_time = current_start_time
                    end_time = current_end_time
            
            # 最後の時間ブロックの追加
            formatted_times.append(f"{start_time}-{end_time}")
            print(f"Added final time block: {start_time}-{end_time}")
            
            result.append(f"{date}: {', '.join(formatted_times)}")
            print(f"Result for {date}: {', '.join(formatted_times)}")

        return result, title_text

    except Exception as e:
        return f"エラーが発生しました: {e}"

    finally:
        driver.quit()

def get_event_ids_by_uuid(uuid):
    """指定されたUUIDに関連するすべてのイベントIDを取得します。"""
    event_ids = []
    try:
        # SQLiteデータベースに接続
        conn = sqlite3.connect('tonton_calendar_compare.db')
        cursor = conn.cursor()
        
        # UUIDに基づくイベントIDの取得
        cursor.execute("SELECT event_id FROM event_mappings WHERE task_uuid = ?", (uuid,))
        event_ids = [row[0] for row in cursor.fetchall()]
        
        conn.close()
    except sqlite3.Error as e:
        print(f"データベースエラー: {e}")
    
    return event_ids

def on_delete():
    progress_bar.set(0)  # プログレスバーをリセット
    # サブミット時の処理を別スレッドで実行する
    threading.Thread(target=delete_selected_task, daemon=True).start()

def delete_selected_task():
    
    selected_task_index = task_listbox.curselection()
    
    if selected_task_index:
        index = selected_task_index[0]  # 選択されたタスクのインデックス
        task_uuid = tasks[index]['task_uuid']  # UUIDを取得
        
        # UUIDに基づいて関連するイベントIDを取得
        event_ids = get_event_ids_by_uuid(task_uuid)
        
        if not event_ids:
            print("UUIDに関連するイベントIDが見つかりませんでした。")
            return
        
        total_events = len(event_ids)  # 総イベント数
        completed_events = 0  # 完了したイベント数

        # Googleカレンダーのイベントを削除
        delete_successful = True
        for event_id in event_ids:
            if delete_google_calendar_event(service, event_id):
                completed_events += 1
                progress = completed_events / total_events
                progress_bar.set(progress)  # プログレスバーを更新
                app.update_idletasks()  # GUIの更新を確実に反映する
            else:
                delete_successful = False
        # 削除に成功したか確認し、データベースからタスク情報を削除
        conn = sqlite3.connect('tonton_calendar_compare.db')
        try:
            cursor = conn.cursor()
            if delete_successful:
                # Googleカレンダーでイベント削除成功した場合
                delete_event_ids_by_uuid(cursor, task_uuid)
                delete_task_info_by_uuid(cursor, task_uuid)
                conn.commit()  # コミット
                print("タスクと関連イベントが削除されました。")
            else:
                # Googleカレンダーでイベント削除失敗した場合
                # イベントが削除された場合もあるため、タスク情報を削除するか確認する
                delete_event_ids_by_uuid(cursor, task_uuid)
                delete_task_info_by_uuid(cursor, task_uuid)
                conn.commit()  # コミット
                print("イベントの削除が一部失敗しましたが、データベースの整合性を保つためにタスク情報も削除しました。")
                
            # タスクをリストから削除
            del tasks[index]

            # リストボックスを更新
            update_task_delete_listbox()

        except Exception as e:
            conn.rollback()  # エラーが発生した場合はロールバック
            print(f"削除中にエラーが発生しました: {e}")
        finally:
            conn.close()
            progress_bar.set(1)  # プログレスバーを完了に設定
    else:
        print("削除するタスクを選択してください。")

#一連のdelete関連のトランザクションを後で実装する
def delete_event_ids_by_uuid(cursor, uuid):
    try:
        cursor.execute('DELETE FROM event_mappings WHERE task_uuid = ?', (uuid,))
    except Exception as e:
        raise e

def delete_task_info_by_uuid(cursor, task_uuid):
    try:
        cursor.execute('DELETE FROM task_info WHERE task_uuid = ?', (task_uuid,))
    except Exception as e:
        raise e

def delete_google_calendar_event(service, event_id):
    try:
        service.events().delete(calendarId="primary", eventId=event_id).execute()
        return True
    except Exception as e:
        print(f"イベント削除中にエラーが発生しました: {e}")
        return False


# タスクリストを更新する関数
def update_task_delete_listbox():
    task_listbox.delete(0, ctk.END)
    for task in tasks:
        task_listbox.insert(ctk.END, f"{task['task_name']}")
def on_submit():
    progress_bar.set(0)  # プログレスバーをリセット
    # サブミット時の処理を別スレッドで実行する
    threading.Thread(target=submit_task, daemon=True).start()

def submit_task():
    url_valid = validate_entry(url_entry)
    
    
    # 両方が有効な場合のみ処理を続行
    if url_valid:
        url = url_entry.get()
        
        schedule_data, title = scrape_data(url)  # タイトルも取得
        # タスクに固有のIDを生成
        task_uuid = str(uuid.uuid4())
        
        # SQLiteデータベース接続
        conn = sqlite3.connect('tonton_calendar_compare.db')
        conn.isolation_level = None  # 自動コミットを無効にする
        cursor = conn.cursor()
        
        try:
            cursor.execute('BEGIN TRANSACTION')  # トランザクション開始
            save_uuid_task_name(task_uuid, title, conn)

            total_events = len(schedule_data)
            completed_events = 0
            
            for entry in schedule_data:
                date, times_str = entry.split(': ')
                date_obj = datetime.strptime(date.split('(')[0], "%Y/%m/%d")
                time_blocks = times_str.split(', ')

                for time_block in time_blocks:
                    start_time_str, end_time_str = time_block.split('-')
                    start_hour = int(start_time_str[:2])
                    start_minute = int(start_time_str[2:])
                    end_hour = int(end_time_str[:2])
                    end_minute = int(end_time_str[2:])
                    
                    start_time = datetime(date_obj.year, date_obj.month, date_obj.day, start_hour, start_minute)
                    end_time = datetime(date_obj.year, date_obj.month, date_obj.day, end_hour, end_minute)

                    
                    # Googleカレンダーにイベントを挿入
                    insert_event_to_calendar(service, date, start_time, end_time, task_uuid, title, conn)
                    completed_events += 1
                    progress = completed_events / total_events
                    progress_bar.set(progress)  # プログレスバーを更新
                    app.update_idletasks()  # GUIの更新を確実に反映する
            cursor.execute('COMMIT')  # コミット
            print("すべてのデータが成功裏に保存されました")
            
        except Exception as e:
            cursor.execute('ROLLBACK')  # ロールバック
            print(f"エラーが発生したため、変更を元に戻しました: {e}")
        
        finally:
            conn.close()
            load_tasks()
            update_task_listbox()
            progress_bar.set(1)  # プログレスバーを完了に設定

def validate_entry(entry):
    """エントリーが空の場合に枠を赤くする関数"""
    if not entry.get().strip():  # 入力が空かどうかをチェック
        entry.configure(border_color="red")
        return False
    else:
        entry.configure(border_color="black")  # 入力があればデフォルトの色に戻す
        return True        

#GUIアプリの設定
app = ctk.CTk()
app.title("Support Schedule Adjustment")
app.geometry("400x450")

# グリッドの列幅を均等に設定
app.grid_columnconfigure(0, weight=1)
app.grid_columnconfigure(1, weight=1)

# URL入力用のエントリー
url_label = ctk.CTkLabel(app, text="URLを入力してください:")
url_label.grid(row=0, column=0, columnspan=2, pady=5)
url_entry = ctk.CTkEntry(app, width=300)
url_entry.grid(row=1, column=0, columnspan=2, pady=5)

# # タイトル入力用のエントリー
# title_label = ctk.CTkLabel(app, text="イベントのタイトルを入力してください:")
# title_label.grid(row=2, column=0, columnspan=2, pady=5)
# title_entry = ctk.CTkEntry(app, width=300)
# title_entry.grid(row=3, column=0, columnspan=2, pady=5)

# タスクリストボックス
task_listbox_label = ctk.CTkLabel(app, text="調整中予定一覧:")
task_listbox_label.grid(row=2, column=0, columnspan=2, pady=5)
task_listbox = tk.Listbox(app, selectmode=tk.MULTIPLE, width=80, height=10)  
task_listbox.grid(row=3, column=0, columnspan=2, pady=5)

# プログレスバー
progress_bar = ctk.CTkProgressBar(app)
progress_bar.grid(row=5, column=0, columnspan=2, pady=20)
progress_bar.set(0)

# 決定ボタン
submit_button = ctk.CTkButton(app, text="カレンダーに追加", command=on_submit)
submit_button.grid(row=6, column=0, padx=20, pady=10)

# 削除ボタン
delete_button = ctk.CTkButton(app, text="日程削除", command=on_delete)
delete_button.grid(row=6, column=1, padx=20, pady=10)

# タスクデータをロードして表示
load_tasks()
update_task_listbox()

app.mainloop()

