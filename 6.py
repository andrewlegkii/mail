import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import win32com.client as win32
from datetime import datetime

class DebtNotifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Debt Notifier App")
        self.sort_ascending = True  # Флаг сортировки по алфавиту

        # Элементы интерфейса
        tk.Label(root, text="Excel File Path:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.file_entry = tk.Entry(root, width=50)
        self.file_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=10, pady=5)

        tk.Label(root, text="Sheet Name:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.sheet_entry = tk.Entry(root, width=20)
        self.sheet_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # Кнопки управления
        buttons_frame = tk.Frame(root)
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        tk.Button(buttons_frame, text="Load Companies", command=self.load_companies).pack(side="left", padx=5)
        self.sort_btn = tk.Button(buttons_frame, text="Sort A-Z", command=self.toggle_sort)
        self.sort_btn.pack(side="left", padx=5)

        # Прокручиваемый фрейм
        self.scroll_canvas = tk.Canvas(root, width=600, height=300)
        self.scroll_canvas.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.scrollbar = tk.Scrollbar(root, orient="vertical", command=self.scroll_canvas.yview)
        self.scrollbar.grid(row=3, column=3, sticky="ns")
        
        self.scroll_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.inner_frame = tk.Frame(self.scroll_canvas)
        self.scroll_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", lambda e: self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all")))

        # Выбор аккаунта
        acc_frame = tk.Frame(root)
        acc_frame.grid(row=4, column=0, columnspan=3, sticky="w")
        
        tk.Label(acc_frame, text="Choose Email Account:").pack(side="left", padx=10, pady=5)
        self.account_combo = ttk.Combobox(acc_frame, width=40)
        self.account_combo.pack(side="left", padx=5)
        tk.Button(acc_frame, text="Load Accounts", command=self.load_accounts).pack(side="left", padx=5)

        # Кнопка отправки
        tk.Button(root, text="Send Emails", command=self.send_emails).grid(row=5, column=0, columnspan=3, pady=10)

        # Загружаем аккаунты при старте
        self.load_accounts()

    def toggle_sort(self):
        """Переключение порядка сортировки"""
        self.sort_ascending = not self.sort_ascending
        self.sort_btn.config(text="Sort Z-A" if self.sort_ascending else "Sort A-Z")
        self.load_companies()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def load_companies(self):
        file_path = self.file_entry.get()
        sheet_name = self.sheet_entry.get()

        if not file_path or not sheet_name:
            messagebox.showerror("Error", "Please provide both file path and sheet name.")
            return

        try:
            self.df = pd.read_excel(file_path, sheet_name=sheet_name)
            required_columns = {"Компания", "E-mail"}
            if not required_columns.issubset(self.df.columns):
                messagebox.showerror("Error", "The Excel file must contain 'Компания' and 'E-mail' columns.")
                return

            self.df["Компания"] = self.df["Компания"].astype(str)

            for widget in self.inner_frame.winfo_children():
                widget.destroy()

            companies = sorted(self.df["Компания"].unique(), reverse=not self.sort_ascending)
            
            self.check_vars = {}
            for company in companies:
                var = tk.BooleanVar()
                cb = tk.Checkbutton(self.inner_frame, text=company, variable=var)
                cb.pack(anchor="w")
                self.check_vars[company] = var

            messagebox.showinfo("Success", "Companies loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load companies: {e}")

    def load_accounts(self):
        """Загрузка доступных почтовых аккаунтов из Outlook"""
        try:
            outlook = win32.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            accounts = [acc.SmtpAddress for acc in namespace.Accounts if acc.SmtpAddress]
            self.account_combo["values"] = accounts
            if accounts:
                self.account_combo.set(accounts[0])
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not load email accounts: {e}")

    def send_emails(self):
        selected_companies = [company for company, var in self.check_vars.items() if var.get()]
        if not selected_companies:
            messagebox.showerror("Error", "No companies selected.")
            return

        account_email = self.account_combo.get()
        if not account_email:
            messagebox.showerror("Error", "Please select an email account.")
            return

        try:
            outlook = win32.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            account = None
            for acc in namespace.Accounts:
                if acc.SmtpAddress.lower() == account_email.lower():
                    account = acc
                    break

            if not account:
                messagebox.showerror("Error", f"Account with email {account_email} not found.")
                return

            log_file = "email_log.txt"
            with open(log_file, "a", encoding="utf-8") as log:
                for company in selected_companies:
                    company_data = self.df[self.df["Компания"] == company]
                    if company_data.empty:
                        continue

                    email = company_data["E-mail"].iloc[0]

                    cc_emails = []
                    for col in ["Copy1", "Copy2", "Copy3"]:
                        if col in company_data.columns and pd.notna(company_data[col].iloc[0]):
                            cc_emails.append(company_data[col].iloc[0])

                    cc_emails_str = "; ".join(cc_emails) if cc_emails else None

                    table_data = company_data[["Номер претензии", "Компания", "Инвойс", "Дата претензии", "Задолженность"]]
                    table_html = table_data.to_html(index=False, justify="center", border=1)

                    subject = f"Напоминание о задолженности по претензиям ({company})"

                    body = (f"Уважаемый партнер,<br><br>У вас имеется задолженность:<br><br>" +
                            table_html +
                            "<br><br>Просьба оплатить в ближайшее время.<br><br>С уважением, Nestle.")

                    mail = outlook.CreateItem(0)
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                    mail.To = email
                    if cc_emails_str:
                        mail.CC = cc_emails_str
                    mail.Subject = subject
                    mail.HTMLBody = body
                    mail.Send()

                    log.write(f"{datetime.now()} - Email sent to {company} ({email}) with Debt Table\n")
                    if cc_emails_str:
                        log.write(f"CC: {cc_emails_str}\n")

                messagebox.showinfo("Success", "Emails sent successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DebtNotifierApp(root)
    root.mainloop()
