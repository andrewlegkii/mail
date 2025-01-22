import pandas as pd
import win32com.client as win32
import logging
from tkinter import filedialog, Tk, Entry, Button, Label, messagebox


def setup_logging():
    # Настройка логирования
    logging.basicConfig(
        filename="email_send_log.txt",
        level=logging.INFO,
        format="%(asctime)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def send_emails_with_cc_and_logging(
    excel_file, sheet_name, sender_email, send_to_all=True, company_number=None
):
    # Чтение данных из Excel
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Проверка наличия необходимых столбцов
    required_columns = [
        "Номер",
        "Номер претензии",
        "Компания",
        "Инвойс",
        "Дата претензии",
        "Задолженность",
        "E-mail",
        "Copy1",
        "Copy2",
        "Copy3",
    ]
    if not all(col in df.columns for col in required_columns):
        raise ValueError(f"В таблице должны быть столбцы: {', '.join(required_columns)}")

    # Фильтрация данных: всем или конкретной компании
    if not send_to_all:
        if company_number is None:
            raise ValueError("Для отправки конкретной компании укажите её номер.")
        company_name = df.loc[df["Номер"] == company_number, "Компания"].iloc[0]
        df = df[df["Компания"] == company_name]

    # Группировка данных по e-mail
    grouped = df.groupby("E-mail")

    # Настройка Outlook
    outlook = win32.Dispatch("outlook.application")
    namespace = outlook.GetNamespace("MAPI")

    # Поиск учетной записи отправителя
    account = None
    for acc in namespace.Accounts:
        if acc.SmtpAddress.lower() == sender_email.lower():
            account = acc
            break

    if not account:
        raise ValueError(f"Учетная запись {sender_email} не найдена в Outlook.")

    for email, group in grouped:
        company = group["Компания"].iloc[0]
        debt_details = ""

        # Сбор деталей задолженностей
        for _, row in group.iterrows():
            claim_number = row["Номер претензии"]
            invoice = row["Инвойс"]
            transport_date = row["Дата претензии"]
            debt = row["Задолженность"]

            debt_details += (
                f"- Номер претензии: {claim_number}, Инвойс: {invoice}, "
                f"Дата претензии: {transport_date.strftime('%d.%m.%Y') if pd.notnull(transport_date) else 'не указана'}, "
                f"Сумма: {debt} руб.\n"
            )

        # Сбор адресов для копии
        cc_list = group[["Copy1", "Copy2", "Copy3"]].values.flatten()
        cc_list = [cc for cc in cc_list if pd.notnull(cc)]
        cc_addresses = "; ".join(cc_list)

        # Генерация текста сообщения
        subject = f"Напоминание о задолженности компании {company}"
        body = f"""
        Уважаемые коллеги,

        Напоминаем, что у компании {company} есть задолженность:
        {debt_details}

        Просим произвести оплату в ближайшее время.

        С уважением,
        Ваша компания
        """

        # Создание и отправка письма
        try:
            mail = outlook.CreateItem(0)
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Привязка отправителя
            mail.To = email
            mail.CC = cc_addresses
            mail.Subject = subject
            mail.Body = body
            mail.Send()

            # Логирование успешной отправки
            logging.info(f"Сообщение отправлено: Компания: {company}, E-mail: {email}, Копия: {cc_addresses}")
            print(f"Сообщение отправлено: {company} ({email}), Копия: {cc_addresses}")
        except Exception as e:
            logging.error(f"Ошибка при отправке сообщения для {company} ({email}): {e}")
            print(f"Ошибка при отправке сообщения для {company} ({email}): {e}")


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Отправка сообщений по компаниям")

        # Инициализация переменных
        self.excel_file = None
        self.sheet_name = None
        self.sender_email = None

        # Элементы интерфейса
        self.create_widgets()

    def create_widgets(self):
        # Кнопка для выбора файла
        self.select_file_button = Button(self.root, text="Выбрать файл", command=self.select_file)
        self.select_file_button.grid(row=0, column=0, padx=10, pady=10)

        # Строка для ввода имени листа
        self.sheet_name_label = Label(self.root, text="Введите название листа:")
        self.sheet_name_label.grid(row=1, column=0, padx=10, pady=10)
        self.sheet_name_entry = Entry(self.root)
        self.sheet_name_entry.grid(row=1, column=1, padx=10, pady=10)

        # Строка для ввода почты отправителя
        self.sender_email_label = Label(self.root, text="Введите почту отправителя:")
        self.sender_email_label.grid(row=2, column=0, padx=10, pady=10)
        self.sender_email_entry = Entry(self.root)
        self.sender_email_entry.grid(row=2, column=1, padx=10, pady=10)

        # Кнопка для отправки писем
        self.send_button = Button(self.root, text="Отправить сообщения", command=self.send_emails)
        self.send_button.grid(row=3, column=0, columnspan=2, pady=10)

    def select_file(self):
        """Открытие диалогового окна для выбора файла"""
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if self.excel_file:
            messagebox.showinfo("Информация", f"Выбран файл: {self.excel_file}")

    def send_emails(self):
        """Отправка сообщений с выбором данных"""
        self.sheet_name = self.sheet_name_entry.get()
        self.sender_email = self.sender_email_entry.get()

        if not self.excel_file or not self.sheet_name or not self.sender_email:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл, введите название листа и почту отправителя")
            return

        try:
            # Настройка логирования
            setup_logging()

            # Вызов функции отправки писем с логированием
            send_emails_with_cc_and_logging(
                self.excel_file, self.sheet_name, self.sender_email, send_to_all=True
            )
            messagebox.showinfo("Успех", "Сообщения успешно отправлены!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при отправке сообщений: {e}")


if __name__ == "__main__":
    root = Tk()
    app = EmailSenderApp(root)
    root.mainloop()
