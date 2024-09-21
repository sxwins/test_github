from ultralytics import YOLO
import cv2
import math 
import time
import numpy as np
import mss
import sys
import json
import datetime

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.auth.exceptions import GoogleAuthError

############################  Settings  ###########################

camera_idx  = 1             # index for current camera

VIDEO_SOURCE = 2            # 1 for "Camera", 2 for "ScreenShot"

# スクリーンキャプチャする領域を定義（x, y, 幅, 高さ）
# Change settings here to fit your own screen
SCREEN_SHOT_AREA = {"top": 350, "left": 100, "width": 1700, "height": 1150}

SAMPLE_INTERVAL = 10        # time interval for counting people. unit: [s]

spreadsheet_id = "1nbYBbUSOHxSqe0z8aRwh3IUontt37BBU2o8eAGoon8Y"

# Scopes of Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

############################  Initialize  ########################



############################  Functions  #########################

def authenticate_service_account():
    try:
        # Googleサービスアカウントファイルで認証する
        creds = Credentials.from_service_account_file('service_account.json', scopes=SCOPES)
        print("認証に成功しました。")
        return creds

    except FileNotFoundError:
        print("サービスアカウントファイルが見つかりません。ファイル名とパスを確認してください。")
        return None

    except GoogleAuthError as e:
        # Google認証に関するエラーをキャッチする
        print(f"認証エラーが発生しました: {str(e)}")
        return None

    except Exception as e:
        # 他の予期しないエラーをキャッチする
        print(f"予期しないエラーが発生しました: {str(e)}")
        return None

def append_data_to_sheet(sheet, spreadsheet_id, data):

    # # 現在の日付を取得し、'YYYY-MM-DD'形式にフォーマット
    # current_date = datetime.datetime.now().strftime('%Y-%m-%d')

    date_time_str = data["time"]
    current_date, current_time = date_time_str.split(maxsplit=1)

    # 現在のスプレッドシートに存在する全てのシート名を取得
    existing_sheets = sheet.get(spreadsheetId=spreadsheet_id).execute().get('sheets', [])
    sheet_names = [s['properties']['title'] for s in existing_sheets]
    
    # 現在の日付をシート名とするシートが存在するか確認
    if current_date not in sheet_names:
        # 存在しない場合、新しいシートを作成
        requests = [{
            'addSheet': {
                'properties': {
                    'title': current_date,  # シート名を現在の日付に設定
                    'index': 0              # シートを一番前に挿入
                }
            }
        }]
        body = {
            'requests': requests
        }
        # バッチリクエストでシートを作成
        response = sheet.batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        print(f"シート '{current_date}' を作成しました。")

    # name_of_first_sheet = sheet.get(spreadsheetId=spreadsheet_id).execute().get('sheets')[0].get('properties').get('title')

    data_append = [data["index"], current_time, data["density"]]
    insert_range = f"{current_date}!A:C"
    request = sheet.values().append(spreadsheetId=spreadsheet_id, range=insert_range, 
                                    valueInputOption='RAW', 
                                    body={'values': [data_append]})
    request.execute()
    print(f"append_data_to_sheet: {data_append} を追加しました。")
 
##############################  Main  ###########################

if VIDEO_SOURCE == 1:
    # start local camera
    cap = cv2.VideoCapture(0)
    cap.set(3, 640)
    cap.set(4, 480)
    window_name = "Camera"

elif VIDEO_SOURCE == 2:

    # mssの初期化
    sct = mss.mss()

    # ウィンドウ名
    window_name = "ScreenCapture"
else:
    print("Error: unknown video source. quit")
    sys.exit(0)

# Google Sheetsサービスに認証し、サービスを取得    
spreadsheet_creds = authenticate_service_account()
if spreadsheet_creds == None:
    print("Error: fail to authenticate google sheet. quit")
    sys.exit(0)

service = build('sheets', 'v4', credentials=spreadsheet_creds)
sheet = service.spreadsheets()

# OpenCVウィンドウを作成
cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)

# ウィンドウを最前面に設定
cv2.setWindowProperty(window_name, cv2.WND_PROP_TOPMOST, 1)

# Download/load model
model = YOLO("yolo-Weights/yolov8n.pt")

# object classes
classNames = ["person", "bicycle", "car", "motorbike", "aeroplane", "bus", "train", "truck", "boat",
              "traffic light", "fire hydrant", "stop sign", "parking meter", "bench", "bird", "cat",
              "dog", "horse", "sheep", "cow", "elephant", "bear", "zebra", "giraffe", "backpack", "umbrella",
              "handbag", "tie", "suitcase", "frisbee", "skis", "snowboard", "sports ball", "kite", "baseball bat",
              "baseball glove", "skateboard", "surfboard", "tennis racket", "bottle", "wine glass", "cup",
              "fork", "knife", "spoon", "bowl", "banana", "apple", "sandwich", "orange", "broccoli",
              "carrot", "hot dog", "pizza", "donut", "cake", "chair", "sofa", "pottedplant", "bed",
              "diningtable", "toilet", "tvmonitor", "laptop", "mouse", "remote", "keyboard", "cell phone",
              "microwave", "oven", "toaster", "sink", "refrigerator", "book", "clock", "vase", "scissors",
              "teddy bear", "hair drier", "toothbrush"
              ]

# Initialize timer and person count
start_time = time.time()
people_count = 0
frame_count  = 0

str_count_info = ""

while True:

    if VIDEO_SOURCE == 1:
        success, img = cap.read()
    elif VIDEO_SOURCE == 2:        
        screenshot = sct.grab(SCREEN_SHOT_AREA)      # mssを使ってスクリーン領域をキャプチャ
        img = np.array(screenshot)          # 画像をnumpy配列に変換
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)   # BGRフォーマットに変換（mssがキャプチャする画像はBGRAフォーマットなので）

    results = model(img, stream=True)
    frame_count += 1

    # coordinates
    for r in results:
        boxes = r.boxes

        for box in boxes:
            # bounding box
            x1, y1, x2, y2 = box.xyxy[0]
            x1, y1, x2, y2 = int(x1), int(y1), int(x2), int(y2) # convert to int values

            # put box in cam
            cv2.rectangle(img, (x1, y1), (x2, y2), (255, 0, 255), 3)

            # confidence
            confidence = math.ceil((box.conf[0]*100))/100
            print("Confidence --->",confidence)

            # class name
            cls = int(box.cls[0])
            print("Class name -->", classNames[cls])

            # object details
            org = [x1, y1]
            font = cv2.FONT_HERSHEY_SIMPLEX
            fontScale = 1
            color = (255, 0, 0)
            thickness = 2

            cv2.putText(img, classNames[cls], org, font, fontScale, color, thickness)

            # If the class is 'person', increase the person count
            if classNames[cls] == "person":
                people_count += 1

     # Calculate elapsed time
    current_time = time.time()            
    elapsed_time = current_time - start_time

    # Check if 60 seconds have passed
    if elapsed_time >= 10:
        # Get current time
        time_str = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(current_time))
        if frame_count > 0:
            density = people_count / frame_count
        else:
            density = 0.0 
        density = round(density, 2)               
        str_count_info = f"[{time_str}] {density} people per frame"    
        print(str_count_info)
        
        data = {"index": camera_idx, 
                "time": time_str,
                "density": density}
        append_data_to_sheet(sheet, spreadsheet_id, data)

        # Reset the timer and people count
        start_time = current_time
        people_count = 0  
        frame_count  = 0  

    cv2.putText(img, str_count_info, [50, 50], cv2.FONT_HERSHEY_SIMPLEX, 2, (255, 0, 255), 5)
    cv2.imshow(window_name, img)
    if cv2.waitKey(1) == ord('q'):
        break

cv2.destroyAllWindows()   
if VIDEO_SOURCE == 1:
    cap.release()
