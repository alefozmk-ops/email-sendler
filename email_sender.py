import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging
from datetime import datetime
import yaml
import os

class EmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Рассылка запросов на акты сверки")
        self.config_file = 'email_config.yaml'
        self.setup_logging()
        self.create_widgets()
        self.load_config()
    
    def setup_logging(self):
        logging.basicConfig(
            filename=f'email_sender_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )
    
    def create_widgets(self):
        # Основная рамка
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройки SMTP
        ttk.Label(main_frame, text="SMTP настройки", font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(main_frame, text="SMTP сервер:").grid(row=1, column=0, sticky=tk.W)
        self.smtp_server = ttk.Entry(main_frame, width=40)
        self.smtp_server.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Label(main_frame, text="Порт:").grid(row=2, column=0, sticky=tk.W)
        self.smtp_port = ttk.Entry(main_frame, width=40)
        self.smtp_port.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Label(main_frame, text="Email отправителя:").grid(row=3, column=0, sticky=tk.W)
        self.sender_email = ttk.Entry(main_frame, width=40)
        self.sender_email.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Label(main_frame, text="Пароль:").grid(row=4, column=0, sticky=tk.W)
        self.sender_password = ttk.Entry(main_frame, width=40, show="*")
        self.sender_password.grid(row=4, column=1, columnspan=2, sticky=(tk.W, tk.E))
        
        # Разделитель
        ttk.Separator(main_frame, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Настройки письма
        ttk.Label(main_frame, text="Настройки письма", font=('Arial', 10, 'bold')).grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(main_frame, text="Тема письма:").grid(row=7, column=0, sticky=tk.W)
        self.subject = ttk.Entry(main_frame, width=40)
        self.subject.grid(row=7, column=1, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Label(main_frame, text="Шаблон письма (доступны переменные: {name}, {company}, {email}):").grid(row=8, column=0, columnspan=3, sticky=tk.W)
        self.email_template = tk.Text(main_frame, width=60, height=8, wrap=tk.WORD)
        self.email_template.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 10))
        
        # Скроллбар для шаблона
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.email_template.yview)
        scrollbar.grid(row=9, column=3, sticky=(tk.N, tk.S))
        self.email_template.configure(yscrollcommand=scrollbar.set)
        
        # Разделитель
        ttk.Separator(main_frame, orient='horizontal').grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Файл CSV
        ttk.Label(main_frame, text="Данные для рассылки", font=('Arial', 10, 'bold')).grid(row=11, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(main_frame, text="CSV файл:").grid(row=12, column=0, sticky=tk.W)
        self.csv_file = ttk.Entry(main_frame, width=40)
        self.csv_file.grid(row=12, column=1, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Выбрать", command=self.select_csv).grid(row=12, column=2, padx=(5, 0))
        
        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=13, column=0, columnspan=3, pady=20)
        
        ttk.Button(button_frame, text="Сохранить настройки", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Загрузить настройки", command=self.load_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Тестовая отправка", command=self.test_send).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Начать рассылку", command=self.send_emails).pack(side=tk.LEFT, padx=5)
        
        # Прогресс-бар
        self.progress = ttk.Progressbar(main_frame, length=500, mode='determinate')
        self.progress.grid(row=14, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        
        # Статус
        self.status = ttk.Label(main_frame, text="Готов к работе", relief=tk.SUNKEN, anchor=tk.W)
        self.status.grid(row=15, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Настройка весов для растягивания
        self.root.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def select_csv(self):
        filename = filedialog.askopenfilename(
            title="Выберите CSV файл",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_file.delete(0, tk.END)
            self.csv_file.insert(0, filename)
    
    def load_config(self):
        """Загрузка настроек из YAML файла"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = yaml.safe_load(f)
                
                if config:
                    # SMTP настройки
                    self.smtp_server.delete(0, tk.END)
                    self.smtp_server.insert(0, config.get('smtp', {}).get('server', 'smtp.yandex.ru'))
                    
                    self.smtp_port.delete(0, tk.END)
                    self.smtp_port.insert(0, str(config.get('smtp', {}).get('port', 587)))
                    
                    self.sender_email.delete(0, tk.END)
                    self.sender_email.insert(0, config.get('smtp', {}).get('email', ''))
                    
                    # Пароль загружаем, если есть
                    if config.get('smtp', {}).get('password'):
                        self.sender_password.delete(0, tk.END)
                        self.sender_password.insert(0, config.get('smtp', {}).get('password'))
                    
                    # Настройки письма
                    self.subject.delete(0, tk.END)
                    self.subject.insert(0, config.get('email', {}).get('subject', 'Запрос акта сверки'))
                    
                    self.email_template.delete(1.0, tk.END)
                    default_template = config.get('email', {}).get('template', 
                        "Здравствуйте, {name}!\n\n"
                        "Просим вас выслать акт сверки за прошедший квартал.\n\n"
                        "С уважением,\nБухгалтерия"
                    )
                    self.email_template.insert(1.0, default_template)
                    
                    self.status.config(text=f"Конфигурация загружена из {self.config_file}")
                    logging.info("Конфигурация загружена из YAML")
                    
            except Exception as e:
                logging.error(f"Ошибка загрузки YAML: {e}")
                messagebox.showerror("Ошибка", f"Не удалось загрузить конфигурацию:\n{str(e)}")
        else:
            # Если файла нет, загружаем значения по умолчанию
            self.email_template.insert(1.0, 
                "Здравствуйте, {name}!\n\n"
                "Просим вас выслать акт сверки за прошедший квартал.\n\n"
                "С уважением,\nБухгалтерия"
            )
            self.subject.insert(0, "Запрос акта сверки")
    
    def save_config(self):
        """Сохранение настроек в YAML файл"""
        config = {
            'smtp': {
                'server': self.smtp_server.get(),
                'port': int(self.smtp_port.get()) if self.smtp_port.get().isdigit() else 587,
                'email': self.sender_email.get(),
                'password': self.sender_password.get()  # В реальном проекте лучше шифровать
            },
            'email': {
                'subject': self.subject.get(),
                'template': self.email_template.get(1.0, tk.END).strip()
            }
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
            
            self.status.config(text=f"Конфигурация сохранена в {self.config_file}")
            logging.info("Конфигурация сохранена")
            messagebox.showinfo("Успех", f"Настройки сохранены в файл:\n{self.config_file}")
            
        except Exception as e:
            logging.error(f"Ошибка сохранения YAML: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить конфигурацию:\n{str(e)}")
    
    def test_send(self):
        """Тестовая отправка на свой email"""
        if not self.sender_email.get():
            messagebox.showerror("Ошибка", "Укажите email отправителя")
            return
        
        test_email = self.sender_email.get()
        if messagebox.askyesno("Тестовая отправка", 
                               f"Отправить тестовое письмо на {test_email}?"):
            try:
                self.send_single_email(test_email, "Тестовый контрагент", "Тестовая компания")
                messagebox.showinfo("Успех", f"Тестовое письмо отправлено на {test_email}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка отправки тестового письма:\n{str(e)}")
    
    def send_single_email(self, to_email, name, company):
        """Отправка одного письма"""
        msg = MIMEMultipart()
        msg['From'] = self.sender_email.get()
        msg['To'] = to_email
        msg['Subject'] = self.subject.get()
        
        # Подставляем переменные в шаблон
        body = self.email_template.get(1.0, tk.END).strip()
        body = body.format(
            name=name,
            company=company,
            email=to_email
        )
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Отправка
        if self.smtp_port.get() == '465':  # SSL
            server = smtplib.SMTP_SSL(self.smtp_server.get(), int(self.smtp_port.get()))
        else:  # TLS
            server = smtplib.SMTP(self.smtp_server.get(), int(self.smtp_port.get()))
            server.starttls()
        
        server.login(self.sender_email.get(), self.sender_password.get())
        server.send_message(msg)
        server.quit()
        
        logging.info(f"Письмо отправлено: {to_email}")
    
    def send_emails(self):
        """Массовая рассылка"""
        if not self.csv_file.get():
            messagebox.showerror("Ошибка", "Выберите CSV файл с адресами")
            return
        
        try:
            # Читаем CSV
            df = pd.read_csv(self.csv_file.get())
            
            if 'email' not in df.columns:
                messagebox.showerror("Ошибка", 
                    "В CSV файле нет колонки 'email'.\n"
                    "Доступные колонки: " + ", ".join(df.columns))
                return
            
            # Подтверждение
            if not messagebox.askyesno("Подтверждение", 
                                       f"Найдено {len(df)} контрагентов.\n"
                                       f"Начать рассылку?"):
                return
            
            # Настраиваем SMTP
            self.status.config(text="Подключение к SMTP серверу...")
            self.root.update()
            
            if self.smtp_port.get() == '465':
                server = smtplib.SMTP_SSL(self.smtp_server.get(), int(self.smtp_port.get()))
            else:
                server = smtplib.SMTP(self.smtp_server.get(), int(self.smtp_port.get()))
                server.starttls()
            
            server.login(self.sender_email.get(), self.sender_password.get())
            
            self.progress['maximum'] = len(df)
            success_count = 0
            
            # Отправляем письма
            for idx, row in df.iterrows():
                try:
                    name = row.get('name', 'уважаемый партнер')
                    company = row.get('company', '')
                    
                    self.send_single_email(row['email'], name, company)
                    success_count += 1
                    
                    self.progress['value'] = idx + 1
                    self.status.config(text=f"Отправлено: {success_count} из {len(df)}")
                    self.root.update()
                    
                except Exception as e:
                    logging.error(f"Ошибка при отправке {row['email']}: {e}")
                    self.status.config(text=f"Ошибка: {row['email']} - {str(e)[:50]}")
                    self.root.update()
            
            server.quit()
            
            messagebox.showinfo("Готово", 
                f"Рассылка завершена!\n"
                f"Успешно: {success_count} из {len(df)}\n"
                f"Детали в файле лога")
            
            self.status.config(text=f"Рассылка завершена. Успешно: {success_count} из {len(df)}")
            logging.info(f"Рассылка завершена. Успешно: {success_count} из {len(df)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            logging.error(f"Критическая ошибка: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSender(root)
    root.mainloop()