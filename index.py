import gspread
import os
from datetime import datetime
from collections import defaultdict
import re

dir_path = os.path.dirname(__file__)
gc = gspread.oauth(
    credentials_filename=os.path.join(dir_path, "client_secret.json"),
    authorized_user_filename=os.path.join(dir_path, "authorized_user.json"),
)

current_time = datetime.now().strftime("%y%m%d%H%M")
wb = gc.create(f"LINEトーク分析_{current_time}")
print(wb.id)
wb = gc.open_by_key(wb.id)
ws = wb.get_worksheet(0)
monthly_speaker_counts = defaultdict(lambda: defaultdict(int))

with open('line_chat.txt', 'r', encoding='utf-8') as file:
    for line in file:
        date_match = re.match(r'\d{4}/\d{2}/\d{2}', line)
        if date_match:
            current_date = date_match.group().replace('/', '.')  # YYYY.MM形式
        message_match = re.match(r'(\d{2}:\d{2})\s(.+?)\s(.+)', line)
        if message_match:
            time, speaker, message = message_match.groups()
            # スピーカー名を正規化
            normalized_speaker = speaker.split()[0]  
            month = current_date[:7]  # 'YYYY.MM' 形式の月
            monthly_speaker_counts[month][normalized_speaker] += 1

# トーク履歴に登場した全ての発言者を取得
all_speakers = set()
for month_counts in monthly_speaker_counts.values():
    all_speakers.update(month_counts.keys())

all_data = []
for month in sorted(monthly_speaker_counts.keys()):
    row = [month]
    for speaker in sorted(all_speakers):  # すべての発言者に対して
        count = monthly_speaker_counts[month].get(speaker, 0)  # トーク数がない場合は0を返す
        row.append(count)
    all_data.append(row)

headers = ['年月'] + sorted(all_speakers)
ws.append_row(headers)
ws.append_rows(all_data)
