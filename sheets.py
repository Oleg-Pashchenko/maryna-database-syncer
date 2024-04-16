import os

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import dotenv

dotenv.load_dotenv()


def write_to_db(data):
    print('Записываю в таблицу')
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)
    sheet = client.open_by_url(os.getenv('SHEET_URL'))
    worksheet = sheet.get_worksheet(0)

    worksheet.update('A1', data)
