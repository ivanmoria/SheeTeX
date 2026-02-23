import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", scope)
client = gspread.authorize(creds)

sheet = client.open_by_key("1Kf5_DYjTErSNpVjPoKG0HIyKfAhBfuVM_Qt4a0EOMYQ").sheet1

data = sheet.get_all_records()
df = pd.DataFrame(data)