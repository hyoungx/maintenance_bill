# ==== version history ====
#
# version 1.1
#   - token 만료시 갱신하도록 수정
#   - 저장된 서명을 불러오는 대신 gmail 에서 불러오도록 수정
# version 1.0
#   - first working version. chatGPT 생성.

import os
import shutil
import datetime
import json
from pathlib import Path
from email.message import EmailMessage
import base64

from dateutil.relativedelta import relativedelta

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request


# ===== 설정 불러오기 =====
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

original_filenames = config["original_filenames"]
source_folder = Path(config["source_folder"])
destination_folder = Path(config["destination_folder"])
receiver_email = config["receiver_email"]
new_filename_templates = config["new_filenames"]
email_subject_template = config["email_subject"]
email_text_body_template = config["email_text_body"]
html_body_template = config["html_body_template"]

# ===== 날짜 계산 =====
today = datetime.date.today()
last_month = today - relativedelta(months=1)
month_str_filename = last_month.strftime("%Y년%m월")
month_str_email = last_month.strftime("%Y년 %m월")

# ===== 파일 용량 기준 정렬 =====
file_paths = [source_folder / name for name in original_filenames]
file_sizes = []
for path in file_paths:
    if path.exists():
        file_sizes.append((path.stat().st_size, path.name))
    else:
        file_sizes.append((0, path.name))

sorted_files = sorted(zip(file_sizes, original_filenames), reverse=True)

# ===== 새 파일 이름 구성 =====
new_filenames = [template.replace("{month_str_filename}", month_str_filename) for template in new_filename_templates]

# ===== 파일 이동 및 이름 변경 =====
moved_files = []
for (size_info, original), new_name in zip(sorted_files, new_filenames):
    src = source_folder / original
    dst = destination_folder / new_name

    if src.exists():
        shutil.move(str(src), str(dst))
        moved_files.append(dst)
        print(f"파일 이동 완료: {src.name} -> {dst.name}")
    elif dst.exists():
        moved_files.append(dst)
        print(f"이미 이동된 파일 확인됨: {dst.name}")
    else:
        print(f"[오류] 파일을 찾을 수 없습니다: {src.name} 또는 {dst.name}")

# ===== Gmail 인증 =====
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.settings.basic'
]
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
else:
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('gmail', 'v1', credentials=creds)

# ===== Gmail 서명 가져오기 (sendAsEmail이 chlim@imeco.co.kr인 서명 우선 사용) =====
signature_html = ""
try:
    send_as_list = service.users().settings().sendAs().list(userId="me").execute().get("sendAs", [])

    # 1. chlim@imeco.co.kr 에 해당하는 서명 찾기
    target_entry = next((entry for entry in send_as_list if entry.get("sendAsEmail") == "chlim@imeco.co.kr"), None)

    if target_entry:
        signature_html = target_entry.get("signature", "")
        print("📌 chlim@imeco.co.kr 계정의 서명을 로드했습니다.")
    else:
        print("⚠️ chlim@imeco.co.kr 서명을 찾지 못했습니다. 기본 계정 서명을 사용합니다.")
        # 2. 기본 계정 서명 fallback
        default_entry = next((entry for entry in send_as_list if entry.get("isDefault")), None)
        if default_entry:
            signature_html = default_entry.get("signature", "")
        else:
            print("⚠️ 기본 서명도 존재하지 않습니다.")

except Exception as e:
    print(f"[에러] Gmail 서명 불러오기 실패: {e}")


# ===== 이메일 작성 =====
message = EmailMessage()

email_subject = email_subject_template.replace("{month_str_email}", month_str_email)
email_text_body = email_text_body_template.replace("{month_str_email}", month_str_email)
html_body = html_body_template.replace("{month_str_email}", month_str_email).replace("{signature}", signature_html)

message.set_content(email_text_body)
message.add_alternative(html_body, subtype='html')

message['To'] = receiver_email
message['From'] = "me"
message['Subject'] = email_subject

for path in moved_files:
    with open(path, 'rb') as f:
        file_data = f.read()
        message.add_attachment(file_data, maintype='application', subtype='pdf', filename=path.name)

# ===== 이메일 전송 =====
encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
send_message = {'raw': encoded_message}
service.users().messages().send(userId="me", body=send_message).execute()

print("메일 전송 완료 및 파일 이동 완료")
