import os
import shutil
import datetime
import json
from dateutil.relativedelta import relativedelta
from pathlib import Path
from email.message import EmailMessage
import base64

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ===== 설정 파일 읽어오기 =====
config_file_path = "config.json"
with open(config_file_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

# 설정 값 불러오기
original_filenames = config['original_filenames']
source_folder = config['source_folder']
destination_folder = config['destination_folder']
receiver_email = config['receiver_email']
signature_file_path = config['signature_file_path']
new_filenames_template = config['new_filenames']
email_subject_template = config['email_subject']
email_body_template = config['email_body']
html_body_template = config['html_body']

# ===== 날짜 기반 이름 생성 =====
today = datetime.date.today()
last_month = today - relativedelta(months=1)

# 파일명용 (붙여 쓰기)
month_str_filename = last_month.strftime("%Y년%m월")

# 이메일용 (띄어쓰기 포함)
month_str_email = last_month.strftime("%Y년 %m월")

# ===== 파일 용량 기준으로 이름 매핑 =====
file_paths = [Path(source_folder) / name for name in original_filenames]

# 용량 측정
file_sizes = []
for path in file_paths:
    if path.exists():
        file_sizes.append((path.stat().st_size, path.name))
    else:
        file_sizes.append((0, path.name))  # 없는 파일은 용량 0으로 처리

# 용량 내림차순 정렬 (큰 파일이 먼저)
sorted_files = sorted(zip(file_sizes, original_filenames), reverse=True)

# 정렬된 순서에 따라 이름 지정 (템플릿을 사용하여 이름 생성)
new_filenames = [
    new_filenames_template[0].format(month_str_filename=month_str_filename),
    new_filenames_template[1].format(month_str_filename=month_str_filename)
]

# 파일 이름 변경 및 이동
moved_files = []
for (size_info, original), new in zip(sorted_files, new_filenames):
    src = Path(source_folder) / original
    dst = Path(destination_folder) / new

    if src.exists():
        shutil.move(str(src), str(dst))
        moved_files.append(dst)
        print(f"파일 이동 완료: {src.name} -> {dst.name}")
    elif dst.exists():
        moved_files.append(dst)
        print(f"이미 이동된 파일 확인됨: {dst.name}")
    else:
        print(f"[오류] 파일을 찾을 수 없습니다: {src.name} 또는 {dst.name}")

# ===== Gmail API 인증 =====
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.settings.basic'  # 서명 조회용 추가
]
creds = None

if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
else:
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

service = build('gmail', 'v1', credentials=creds)

# ===== 서명 HTML 파일 읽어오기 =====
signature = ""
if os.path.exists(signature_file_path):
    with open(signature_file_path, 'r', encoding='utf-8') as f:
        signature = f.read()
else:
    print(f"[오류] 서명 HTML 파일을 찾을 수 없습니다: {signature_file_path}")

# ===== 이메일 작성 =====
message = EmailMessage()

# 이메일 본문 내용 작성
email_body = email_body_template.format(month_str_email=month_str_email)
message.set_content(email_body)

# HTML 본문 (서명 포함)
html_body = html_body_template.format(month_str_email=month_str_email)
html_body_with_signature = f"{html_body}{signature}"

message.add_alternative(html_body_with_signature, subtype='html')

# 이메일 필수 헤더 설정
email_subject = email_subject_template.format(month_str_email=month_str_email)
message['To'] = receiver_email
message['From'] = "me"
message['Subject'] = email_subject

# 첨부 파일 추가
for path in moved_files:
    with open(path, 'rb') as f:
        file_data = f.read()
        file_name = path.name
    message.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

# 이메일 전송
encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
send_message = {'raw': encoded_message}
service.users().messages().send(userId="me", body=send_message).execute()

print("메일 전송 완료 및 파일 이동 완료")
