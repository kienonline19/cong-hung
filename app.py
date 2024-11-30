import os
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import datetime
import re
import html
import unicodedata
from bs4 import BeautifulSoup


class GmailSummarizer:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
        self.service = self.gmail_connect()

    def gmail_connect(self):
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file(
                'token.json', self.SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return build('gmail', 'v1', credentials=creds)

    def clean_text(self, text):
        """Làm sạch và format văn bản"""
        # Giải mã HTML entities
        text = html.unescape(text)

        # Loại bỏ ký tự đặc biệt
        text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]', '', text)

        # Loại bỏ ký tự điều khiển
        text = ''.join(char for char in text if not unicodedata.category(
            char).startswith('C'))

        # Xử lý URLs
        urls = re.findall(r'https?://\S+', text)
        for url in urls:
            text = text.replace(url, f'\nURL: {url}\n')

        # Tách câu thành các đoạn
        text = re.sub(r'([.!?]+)(?=\s|\Z)', r'\1\n', text)

        # Loại bỏ khoảng trắng thừa và dòng trống
        lines = [line.strip() for line in text.split('\n')]
        text = '\n'.join(line for line in lines if line)

        # Loại bỏ các dòng trùng lặp liên tiếp
        lines = text.split('\n')
        unique_lines = []
        prev_line = None
        for line in lines:
            if line != prev_line:
                unique_lines.append(line)
                prev_line = line

        return '\n'.join(unique_lines)

    def decode_email_part(self, part):
        try:
            if part.get('body') and part['body'].get('data'):
                content = base64.urlsafe_b64decode(
                    part['body']['data']).decode('utf-8', errors='ignore')

                if 'text/html' in part.get('mimeType', ''):
                    soup = BeautifulSoup(content, 'html.parser')
                    # Loại bỏ các thẻ script và style
                    for script in soup(["script", "style"]):
                        script.decompose()
                    content = soup.get_text(separator='\n', strip=True)

                return self.clean_text(content)
            return ""
        except Exception as e:
            print(f"Lỗi khi giải mã email: {str(e)}")
            return ""

    def format_email_content(self, email_data, index):
        """Format nội dung email với cấu trúc rõ ràng"""
        return f"""
{'-'*80}
Email #{index}
{'-'*80}

Từ: {email_data['from']}
Tiêu đề: {email_data['subject']}
Thời gian: {email_data['date']}

PHÂN TÍCH NỘI DUNG:
{'-'*40}
{self.analyze_content(email_data['body'])}

NỘI DUNG GỐC:
{'-'*40}
{email_data['body']}

{'='*80}
"""

    def analyze_content(self, content):
        """Phân tích nội dung email"""
        analysis = []

        # Đếm số từ
        words = len(content.split())
        analysis.append(f"Số từ: {words}")

        # Đếm số câu
        sentences = len(re.findall(r'[.!?]+', content))
        analysis.append(f"Số câu: {sentences}")

        # Phát hiện URLs
        urls = re.findall(r'https?://\S+', content)
        if urls:
            analysis.append(f"Số liên kết: {len(urls)}")

        # Phát hiện địa chỉ email
        emails = re.findall(r'[\w\.-]+@[\w\.-]+', content)
        if emails:
            analysis.append(f"Số địa chỉ email: {len(emails)}")

        # Phân tích chủ đề chính (dựa trên 50 từ đầu tiên)
        first_50_words = ' '.join(content.split()[:50])
        analysis.append(f"\nĐoạn mở đầu:\n{first_50_words}...")

        return '\n'.join(analysis)

    def get_email_content(self, msg_id):
        try:
            message = self.service.users().messages().get(
                userId='me',
                id=msg_id,
                format='full'
            ).execute()

            headers = message['payload']['headers']
            subject = next((h['value'] for h in headers if h['name'].lower(
            ) == 'subject'), 'Không có tiêu đề')
            sender = next((h['value'] for h in headers if h['name'].lower(
            ) == 'from'), 'Không rõ người gửi')
            date = next((h['value']
                        for h in headers if h['name'].lower() == 'date'), '')

            content = []
            if 'parts' in message['payload']:
                for part in message['payload']['parts']:
                    if part['mimeType'] in ['text/plain', 'text/html']:
                        part_content = self.decode_email_part(part)
                        if part_content:
                            content.append(part_content)
            else:
                part_content = self.decode_email_part(message['payload'])
                if part_content:
                    content.append(part_content)

            return {
                'subject': self.clean_text(subject),
                'from': self.clean_text(sender),
                'date': date,
                'body': '\n'.join(content)
            }
        except Exception as e:
            print(f"Lỗi khi đọc email: {str(e)}")
            return None

    def process_emails(self, max_emails=10):
        try:
            results = self.service.users().messages().list(
                userId='me',
                labelIds=['INBOX'],
                maxResults=max_emails
            ).execute()

            messages = results.get('messages', [])

            if not messages:
                return "Không tìm thấy email nào."

            report = f"""BÁO CÁO PHÂN TÍCH EMAIL
Thời gian tạo: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Số lượng email phân tích: {min(len(messages), max_emails)}
{'='*80}\n"""

            for i, message in enumerate(messages, 1):
                print(f"Đang xử lý email {i}/{len(messages)}...")
                email_data = self.get_email_content(message['id'])
                if email_data:
                    report += self.format_email_content(email_data, i)

            # Thêm thống kê tổng quan
            total_stats = "\nTHỐNG KÊ TỔNG QUAN:\n" + "="*30 + "\n"
            total_stats += f"Tổng số email đã phân tích: {len(messages)}\n"
            total_stats += f"Thời gian hoàn thành: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"

            return report + total_stats

        except Exception as e:
            return f"Lỗi khi xử lý email: {str(e)}"


def main():
    try:
        summarizer = GmailSummarizer()
        result = summarizer.process_emails(10)  # Phân tích 10 email

        # Lưu kết quả
        with open('email_analysis.txt', 'w', encoding='utf-8') as f:
            f.write(result)

        print("\nĐã tạo báo cáo phân tích thành công!")
        print(f"Xem kết quả trong file: email_analysis.txt")

    except Exception as e:
        print(f"Lỗi: {str(e)}")


if __name__ == "__main__":
    main()
