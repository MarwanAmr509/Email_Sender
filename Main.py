import pandas as pd
from Service import Create_Service
import re
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

CLIENT_SECRET_FILE = "secret.json"
GMAIL_API_NAME = 'gmail'
GMAIL_API_VERSION = "v1"
GMAIL_SCOPES = ['https://mail.google.com/']

SHEETS_API_NAME= "sheets"
SHEETS_API_VERSION = "v4"
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

gmail_service = Create_Service(CLIENT_SECRET_FILE, GMAIL_API_NAME, GMAIL_API_VERSION, GMAIL_SCOPES)


def read_sheets(google_sheet_id):

    sheets_service = Create_Service(CLIENT_SECRET_FILE,SHEETS_API_NAME,SHEETS_API_VERSION,SHEETS_SCOPES)
    mySpreadSheets = sheets_service.spreadsheets().get(spreadsheetId=google_sheet_id).execute()

    sheets = mySpreadSheets["sheets"]

    for sheet in sheets:
        index = sheet['properties']['index']

        if sheet['properties']['index'] == 0:
            dataset = sheets_service.spreadsheets().values().get(spreadsheetId=google_sheet_id, range = sheet['properties']['title'], majorDimension = "ROWS").execute()
            # print(dataset['values'])
            df = pd.DataFrame(dataset['values'])
            df.columns = df.iloc[0]
            df = df.iloc[1:]




    return df, sheets_service


url = input("Enter the URL of the sheet:")

GOOGLE_SHEET_ID = re.findall('/d/(.*?)/', url)[0]

df, sheets_service = read_sheets(GOOGLE_SHEET_ID)

counter = 2
for index, row in df.iterrows():
    name = row['Name']
    evaluation = row['Feedback']
    evaluation_link = row["Evaluation Link"]
    email = row['Email']

    # arabic
    # name = row['الاسم ثلاثيا باللغة الانجليزية']
    # evaluation = row['المادة العلمية']
    # performance = row["آداء المعلم"]
    # english = row["اللغة الإنجليزية "]
    # email = row['EMAIL']

    # professional teacher
    # rate = row['Rate/10']


    # Arabic Message
 

    mimeMessgae = MIMEMultipart()
    mimeMessgae['to'] = email

    # mimeMessgae['to'] = "marwaamr509@gmail.com"

    # mimeMessgae['cc'] = 'haidy.hassan@ibnsina-academy.com'

    mimeMessgae['subject'] = "تقييم التاسك العاشر مسار القرائية"

    with open('message.txt', 'r',encoding='utf-8') as file:
        message = file.read()

    # arabic
    # message = message.replace("{name}", name)  # Replace "{name}" with the actual name value
    # message = message.replace("{eval}", evaluation) 
    # message = message.replace("{performance}", performance)  # Replace "{email}" with the actual email value
    # message = message.replace("{english}", english)  # Replace "{email}" with the actual email value

    # English
    message = message.replace("{name}", name)  # Replace "{name}" with the actual name value
    message = message.replace("{eval}", evaluation)  # Replace "{email}" with the actual email value
    message = message.replace("{eval_link}", evaluation_link)  # Replace "{email}" with the actual email value

    # message = message.replace("{rate}", rate)


    # message = message.replace("{email}", email)  # Replace "{email}" with the actual email value


    mimeMessgae.attach(MIMEText(message, 'plain','utf-8'))
    raw_string = base64.urlsafe_b64encode(mimeMessgae.as_bytes()).decode()

    message = gmail_service.users().messages().send(userId = 'me', body = {'raw':raw_string}).execute()

    print(counter, message)

    counter +=1



