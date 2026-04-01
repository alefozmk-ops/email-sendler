import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging
from datetime import datetime

class EmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Рассылка запросов на акты сверки")
        self.setup_logging()
        self.create_widgets()
    
    def setup_logging(self):
        logging.basicConfig(
            filename=f'email_sender_{datetime.now():%Y%m%d}.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
    
    def create_widgets(self):
        # Настройки SMTP
        tk.Label(self.root, text="SMTP сервер:").grid(row=0, column=0)
        self.smtp_server = tk.Entry(self.root, width=30)
        self.smtp_server.grid(row=0, column=1)
        self.smtp_server.insert(0, "smtp.yandex.ru")
        
        tk.Label(self.root, text="Порт:").grid(row=1, column=0)
        self.smtp_port = tk.Entry(self.root, width=30)
        self.smtp_port.grid(row=1, column=1)
        self.smtp_port.insert(0, "587")
        
        tk.Label(self.root, text="Email отправителя:").grid(row=2, column=0)
        self.sender_email = tk.Entry(self.root, width=30)
        self.sender_email.grid(row=2, column=1)
        
        tk.Label(self.root, text="Пароль:").grid(row=3, column=0)
        self.sender_password = tk.Entry(self.root, width=30, show="*")
        self.sender_password.grid(row=3, column=1)
        
        # Файл CSV
        tk.Label(self.root, text="CSV файл:").grid(row=4, column=0)
        self.csv_file = tk.Entry(self.root, width=30)
        self.csv_file.grid(row=4, column=1)
        tk.Button(self.root, text="Выбрать", command=self.select_csv).grid(row=4, column=2)
        
        # Кнопка отправки
        self.send_btn = tk.Button(self.root, text="Начать рассылку", 
                                  command=self.send_emails)
        self.send_btn.grid(row=5, column=1, pady=20)
        
        # Прогресс-бар
        self.progress = ttk.Progressbar(self.root, length=300, mode='determinate')
        self.progress.grid(row=6, column=1, pady=10)
        
        # Статус
        self.status = tk.Label(self.root, text="Готов к отправке")
        self.status.grid(row=7, column=1)
    
    def select_csv(self):
        filename = filedialog.askopenfilename(
            title="Выберите CSV файл",
            filetypes=[("CSV files", "*.csv")]
        )
        if filename:
            self.csv_file.delete(0, tk.END)
            self.csv_file.insert(0, filename)
    
    def send_emails(self):
        try:
            # Читаем CSV
            df = pd.read_csv(self.csv_file.get())
            
            if 'email' not in df.columns:
                messagebox.showerror("Ошибка", 
                    "В CSV файле нет колонки 'email'")
                return
            
            # Настраиваем SMTP
            server = smtplib.SMTP(self.smtp_server.get(), 
                                  int(self.smtp_port.get()))
            server.starttls()
            server.login(self.sender_email.get(), 
                        self.sender_password.get())
            
            self.progress['maximum'] = len(df)
            
            # Отправляем письма
            for idx, row in df.iterrows():
                try:
                    msg = MIMEMultipart()
                    msg['From'] = self.sender_email.get()
                    msg['To'] = row['email']
                    msg['Subject'] = "Запрос акта сверки"
                    
                    body = f"""
                    Здравствуйте, {row.get('name', 'уважаемый контрагент')}!
                    
                    Просим вас выслать акт сверки за прошедший квартал.
                    
                    С уважением,
                    Бухгалтерия
                    """
                    
                    msg.attach(MIMEText(body, 'plain', 'utf-8'))
                    
                    server.send_message(msg)
                    logging.info(f"Письмо отправлено: {row['email']}")
                    
                    self.progress['value'] = idx + 1
                    self.status.config(text=f"Отправлено: {idx+1} из {len(df)}")
                    self.root.update()
                    
                except Exception as e:
                    logging.error(f"Ошибка при отправке {row['email']}: {e}")
                    self.status.config(text=f"Ошибка: {row['email']}")
            
            server.quit()
            messagebox.showinfo("Готово", 
                f"Рассылка завершена!\nОтправлено: {len(df)} писем\n"
                "Детали в файле email_sender.log")
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            logging.error(f"Критическая ошибка: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSender(root)
    root.mainloop()