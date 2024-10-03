import os.path, pandas

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SAMPLE_SPREADSHEET_ID = '1xPBQvWvHcqJSTwbCksbU23mc_EmydjYvqzizxEygf7k'
SAMPLE_RANGE_NAME = 'sheet'


def connection():
  creds = None
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("sheets", "v4", credentials = creds)
    return service
  except HttpError as err:
    print(err)
    
def read(service, range):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range = range).execute()
    return result['values']

def write(service, value, range):
    sheet = service.spreadsheets()
    value = {'values': value}
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range, valueInputOption='RAW', body=value).execute()

def download(value):
    data = pandas.DataFrame(value[1:], columns = value[0])
    path = os.path.abspath(__file__ + '/../files')
    data.to_csv(path + '/data.csv', index = False)
    data.to_excel(path + '/data.xlsx', index = False)
    data.to_html(path + '/data.html', index = False)
    print(data)
    # data.to_json(path + '/data.json', index = False)

if __name__ == "__main__":
  service = connection()
  value = read(service, 'sheet')
  download(value)
  # result = 0
  # for i in value:
  #     for num in i:
  #       result += int(num)
  #     write(service, [[result]], f'sheet!D{value.index(i) + 1}')
  #     result = 0