import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox


class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Sender")

        self.excel_file = ""
        self.selected_sheet = ""
        self.companies = []

        self.setup_ui()

    def setup_ui(self):
        # Поле для выбора Excel файла
        tk.Label(self.root, text="Выберите файл:").grid(row=0, column=0, sticky="e")
        self.select_file_button = tk.Button(self.root, text="Выбрать файл", command=self.select_file)
        self.select_file_button.grid(row=0, column=1, pady=5)

        # Место для отображения выбранного файла
        self.selected_file_label = tk.Label(self.root, text="Файл не выбран")
        self.selected_file_label.grid(row=1, column=0, columnspan=2, pady=5)

        # Выпадающий список для выбора листа
        tk.Label(self.root, text="Выберите лист:").grid(row=2, column=0, sticky="e")
        self.sheet_dropdown = tk.OptionMenu(self.root, "", [])
        self.sheet_dropdown.grid(row=2, column=1, pady=5)

        # Кнопка для выгрузки компаний с выбранного листа
        self.load_companies_button = tk.Button(self.root, text="Выгрузить компании", command=self.load_companies)
        self.load_companies_button.grid(row=3, column=0, columnspan=2, pady=5)

        # Строка для ввода почты отправителя
        tk.Label(self.root, text="Почта отправителя:").grid(row=4, column=0, sticky="e")
        self.sender_email_entry = tk.Entry(self.root, width=50)
        self.sender_email_entry.grid(row=4, column=1, pady=5)

        # Кнопка отправки писем
        self.send_button = tk.Button(self.root, text="Отправить письма", command=self.send_emails)
        self.send_button.grid(row=5, column=0, columnspan=2, pady=5)

    def select_file(self):
        # Окно выбора файла
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.excel_file:
            self.selected_file_label.config(text=f"Выбран файл: {self.excel_file}")
            self.load_sheets()

    def load_sheets(self):
        # Загружаем список листов из файла
        try:
            df = pd.ExcelFile(self.excel_file)
            sheet_names = df.sheet_names
            self.sheet_dropdown['menu'].delete(0, 'end')
            for sheet in sheet_names:
                self.sheet_dropdown['menu'].add_command(label=sheet, command=tk._setit(self.selected_sheet, sheet))
            self.selected_sheet = sheet_names[0]  # Выбираем первый лист по умолчанию
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить листы: {e}")

    def load_companies(self):
        if not self.excel_file or not self.selected_sheet:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл и лист")
            return

        try:
            # Загружаем данные с выбранного листа
            df = pd.read_excel(self.excel_file, sheet_name=self.selected_sheet)
            if "Компания" not in df.columns:
                messagebox.showerror("Ошибка", "В выбранном листе нет столбца 'Компания'")
                return

            self.companies = df["Компания"].unique()

            # Отображаем компании
            self.companies_label = tk.Label(self.root, text=f"Компании: {', '.join(self.companies)}")
            self.companies_label.grid(row=3, column=0, columnspan=2, pady=5)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при выгрузке компаний: {e}")

    def send_emails(self):
        sender_email = self.sender_email_entry.get()
        if not sender_email:
            messagebox.showerror("Ошибка", "Пожалуйста, введите почту отправителя")
            return

        if not self.companies:
            messagebox.showerror("Ошибка", "Не выбраны компании")
            return

        try:
            send_emails_with_cc(self.excel_file, self.selected_sheet, sender_email, self.companies)
            messagebox.showinfo("Успех", "Письма успешно отправлены")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при отправке писем: {e}")


def send_emails_with_cc(excel_file, sheet_name, sender_email, selected_companies):
    # Загружаем данные из выбранного листа
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Фильтруем данные по выбранным компаниям
    df_filtered = df[df["Компания"].isin(selected_companies)]

    # Группируем по email
    grouped = df_filtered.groupby("E-mail")

    outlook = win32.Dispatch("outlook.application")
    namespace = outlook.GetNamespace("MAPI")

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

        cc_list = group[["Copy1", "Copy2", "Copy3"]].values.flatten()
        cc_list = [cc for cc in cc_list if pd.notnull(cc)]
        cc_addresses = "; ".join(cc_list)

        subject = f"Напоминание о задолженности компании {company}"
        body = f"""
        Уважаемые коллеги,

        Напоминаем, что у компании {company} есть задолженность:
        {debt_details}

        Просим произвести оплату в ближайшее время.

        С уважением,
        Ваша компания
        """

        try:
            mail = outlook.CreateItem(0)
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Привязка отправителя
            mail.To = email
            mail.CC = cc_addresses
            mail.Subject = subject
            mail.Body = body
            mail.Send()
        except Exception as e:
            print(f"Ошибка при отправке сообщения для {company} ({email}): {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()
