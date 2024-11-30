import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, font
import os
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from bs4 import BeautifulSoup
import threading


class GmailAnalyzer:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
        self.service = None

    def connect(self):
        try:
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

            self.service = build('gmail', 'v1', credentials=creds)
            return True
        except Exception as e:
            print(f"Lỗi kết nối: {str(e)}")
            return False

    def get_emails(self, max_results=10):
        try:
            results = self.service.users().messages().list(
                userId='me',
                labelIds=['INBOX'],
                maxResults=max_results
            ).execute()
            return results.get('messages', [])
        except Exception as e:
            print(f"Lỗi khi lấy email: {str(e)}")
            return []

    def decode_email_content(self, payload):
        if 'parts' in payload:
            parts = []
            for part in payload['parts']:
                if part['mimeType'] in ['text/plain', 'text/html']:
                    if 'data' in part['body']:
                        text = base64.urlsafe_b64decode(
                            part['body']['data']).decode('utf-8', 'ignore')
                        if part['mimeType'] == 'text/html':
                            text = BeautifulSoup(
                                text, 'html.parser').get_text(separator='\n')
                        parts.append(text)
            return '\n'.join(parts)
        elif 'body' in payload and 'data' in payload['body']:
            text = base64.urlsafe_b64decode(
                payload['body']['data']).decode('utf-8', 'ignore')
            if payload['mimeType'] == 'text/html':
                return BeautifulSoup(text, 'html.parser').get_text(separator='\n')
            return text
        return ""

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

            content = self.decode_email_content(message['payload'])

            return {
                'subject': subject,
                'from': sender,
                'date': date,
                'body': content
            }
        except Exception as e:
            print(f"Lỗi khi đọc email: {str(e)}")
            return None


class EmailAnalyzerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Gmail Analyzer")
        self.root.geometry("1300x800")

        # Thiết lập style
        self.setup_styles()

        self.analyzer = GmailAnalyzer()
        self.setup_gui()

    def setup_styles(self):
        self.style = ttk.Style()

        # Bảng màu cập nhật với text đen cho button
        COLORS = {
            'primary': '#1976D2',        # Xanh đậm cho nền button
            'primary_light': '#1E88E5',   # Xanh nhạt khi hover
            'disabled': '#90CAF9',        # Xanh nhạt khi disable
            'text_button': '#000000',     # Text đen cho button
            'text_normal': '#212121',     # Text đen cho nội dung thường
            'text_white': '#FFFFFF',      # Text trắng khi cần
            'background': '#F5F5F5'       # Nền xám nhạt
        }

        # Style cho Button - với text màu đen
        self.style.configure(
            'Accent.TButton',
            padding=(15, 8),
            font=('Segoe UI', 11, 'bold'),  # Tăng size font lên 11
            background=COLORS['primary'],
            foreground=COLORS['text_button'],  # Sử dụng text đen
            relief='raised',
            borderwidth=2
        )

        # Hover effect và pressed effect - giữ text đen
        self.style.map('Accent.TButton',
                       background=[
                           ('active', COLORS['primary_light']),
                           ('disabled', COLORS['disabled'])
                       ],
                       foreground=[
                           # Text vẫn đen khi hover
                           ('active', COLORS['text_button']),
                           # Text vẫn đen khi disable
                           ('disabled', COLORS['text_button'])
                       ],
                       relief=[('pressed', 'sunken')]
                       )

        # Các style khác giữ nguyên
        self.style.configure(
            'Custom.TFrame',
            background=COLORS['background']
        )

        self.style.configure(
            'Custom.TLabel',
            font=('Segoe UI', 10),
            background=COLORS['background'],
            foreground=COLORS['text_normal']  # Sử dụng text đen cho label
        )

        self.style.configure(
            'Custom.Treeview',
            font=('Segoe UI', 10),
            rowheight=30,
            foreground=COLORS['text_normal']  # Text đen cho treeview
        )

        self.style.configure(
            'Custom.Treeview.Heading',
            font=('Segoe UI', 10, 'bold'),
            foreground=COLORS['text_normal']  # Text đen cho heading
        )

    def setup_gui(self):
        # Main container
        main_frame = ttk.Frame(self.root, style='Custom.TFrame', padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Gmail Analyzer",
            font=('Segoe UI', 20, 'bold'),
            style='Custom.TLabel'
        )
        title_label.pack(pady=(0, 20))

        # Controls frame
        controls_frame = ttk.Frame(main_frame, style='Custom.TFrame')
        controls_frame.pack(fill=tk.X, pady=(0, 10))

        # Nút kết nối - sử dụng style Accent
        self.connect_btn = ttk.Button(
            controls_frame,
            text="Kết nối Gmail",
            style='Accent.TButton',
            command=self.connect_gmail,
            width=15
        )
        self.connect_btn.pack(side=tk.LEFT, padx=10)

        # Spinbox số email
        ttk.Label(
            controls_frame,
            text="Số email:",
            style='Custom.TLabel'
        ).pack(side=tk.LEFT, padx=5)

        self.email_count = ttk.Spinbox(
            controls_frame,
            from_=1,
            to=50,
            width=5,
            font=('Segoe UI', 10)
        )
        self.email_count.set(10)
        self.email_count.pack(side=tk.LEFT, padx=5)

        # Nút phân tích - sử dụng style Accent
        self.analyze_btn = ttk.Button(
            controls_frame,
            text="Phân tích",
            style='Accent.TButton',
            command=self.start_analysis,
            width=15
        )
        self.analyze_btn.pack(side=tk.LEFT, padx=10)
        self.analyze_btn['state'] = 'disabled'

        # Progress bar
        self.progress = ttk.Progressbar(
            controls_frame,
            mode='determinate',
            length=300
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15)

        # Content area
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, pady=10)

        # Email list frame
        list_frame = ttk.LabelFrame(paned, text="Danh sách email", padding=5)
        paned.add(list_frame, weight=1)

        # Email list
        self.email_list = ttk.Treeview(
            list_frame,
            columns=('From', 'Subject', 'Date'),
            show='headings',
            style='Custom.Treeview'
        )

        self.email_list.heading('From', text='Từ')
        self.email_list.heading('Subject', text='Tiêu đề')
        self.email_list.heading('Date', text='Ngày')

        self.email_list.column('From', width=200)
        self.email_list.column('Subject', width=300)
        self.email_list.column('Date', width=150)

        # Scrollbar cho email list
        list_scroll = ttk.Scrollbar(
            list_frame, orient=tk.VERTICAL, command=self.email_list.yview)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.email_list.configure(yscrollcommand=list_scroll.set)
        self.email_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.email_list.bind('<<TreeviewSelect>>', self.on_select_email)

        # Email content frame
        content_frame = ttk.LabelFrame(paned, text="Nội dung email", padding=5)
        paned.add(content_frame, weight=2)

        # Email info
        self.email_info = ttk.Label(
            content_frame,
            text="",
            wraplength=600,
            style='Custom.TLabel'
        )
        self.email_info.pack(fill=tk.X, padx=5, pady=5)

        # Email content
        self.content_text = scrolledtext.ScrolledText(
            content_frame,
            wrap=tk.WORD,
            font=('Segoe UI', 10),
            padx=5,
            pady=5
        )
        self.content_text.pack(fill=tk.BOTH, expand=True)

        # Status bar
        self.status_var = tk.StringVar(value="Chưa kết nối")
        self.status_bar = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            style='Custom.TLabel',
            relief=tk.SUNKEN,
            padding=5
        )
        self.status_bar.pack(fill=tk.X, pady=(10, 0))

    def connect_gmail(self):
        self.status_var.set("Đang kết nối...")
        self.connect_btn['state'] = 'disabled'

        def connect_thread():
            if self.analyzer.connect():
                self.root.after(0, self.connection_success)
            else:
                self.root.after(0, self.connection_failed)

        threading.Thread(target=connect_thread).start()

    def connection_success(self):
        self.status_var.set("Đã kết nối thành công")
        self.analyze_btn['state'] = 'normal'
        messagebox.showinfo("Kết nối", "Kết nối Gmail thành công!")

    def connection_failed(self):
        self.status_var.set("Kết nối thất bại")
        self.connect_btn['state'] = 'normal'
        messagebox.showerror("Lỗi", "Không thể kết nối với Gmail!")

    def start_analysis(self):
        self.analyze_btn['state'] = 'disabled'
        self.email_list.delete(*self.email_list.get_children())
        self.content_text.delete('1.0', tk.END)
        self.progress['value'] = 0
        count = int(self.email_count.get())

        def analyze_thread():
            messages = self.analyzer.get_emails(count)
            total = len(messages)

            for i, msg in enumerate(messages, 1):
                email_data = self.analyzer.get_email_content(msg['id'])
                if email_data:
                    self.root.after(0, self.add_email_to_list,
                                    email_data, msg['id'])
                progress = int((i / total) * 100)
                self.root.after(0, self.update_progress, progress)

            self.root.after(0, self.analysis_complete)

        threading.Thread(target=analyze_thread).start()

    def add_email_to_list(self, email_data, msg_id):
        self.email_list.insert('', 'end', iid=msg_id, values=(
            email_data['from'],
            email_data['subject'],
            email_data['date']
        ))

    def update_progress(self, value):
        self.progress['value'] = value
        self.status_var.set(f"Đang xử lý... {value}%")

    def analysis_complete(self):
        self.analyze_btn['state'] = 'normal'
        self.status_var.set("Phân tích hoàn tất")
        messagebox.showinfo("Hoàn thành", "Đã phân tích xong email!")

    def on_select_email(self, event):
        selection = self.email_list.selection()
        if selection:
            msg_id = selection[0]
            email_data = self.analyzer.get_email_content(msg_id)
            if email_data:
                info_text = f"""Từ: {email_data['from']}
Tiêu đề: {email_data['subject']}
Ngày: {email_data['date']}"""

                self.email_info.config(text=info_text)
                self.content_text.delete('1.0', tk.END)
                self.content_text.insert('1.0', email_data['body'])

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = EmailAnalyzerGUI()
    app.run()
