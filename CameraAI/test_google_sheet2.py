import datetime

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Google Sheets API的范围（Scopes）
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def authenticate_service_account():
    # Use service account file to authenticate
    creds = Credentials.from_service_account_file('service_account.json', scopes=SCOPES)
    return creds

def append_data_to_sheet(spreadsheet_id):

    # 現在の日付を取得し、'YYYY-MM-DD'形式にフォーマット
    current_date = datetime.datetime.now().strftime('%Y-%m-%d')

    # Google Sheetsサービスに認証し、サービスを取得    
    creds = authenticate_service_account()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

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

    while True:
        input_value = input("追加する値を入力してください (終了するには'quit'を入力): ")
        if input_value.lower() == 'quit':
            break

        insert_range = f"{current_date}!A:B"
        request = sheet.values().append(spreadsheetId=spreadsheet_id, range=insert_range, 
                                        valueInputOption='RAW', 
                                        body={'values': [[input_value, input_value]]})
        request.execute()

        print(f"{input_value} を追加しました。")

if __name__ == '__main__':
    # スプレッドシートのIDを入力
    #spreadsheet_id = input("スプレッドシートのIDを入力してください: ")

    spreadsheet_id = "1nbYBbUSOHxSqe0z8aRwh3IUontt37BBU2o8eAGoon8Y"
    append_data_to_sheet(spreadsheet_id)