import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import win32com.client as win32
from datetime import datetime

class DebtNotifierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Debt Notifier App")

        # UI Elements
        tk.Label(root, text="Excel File Path:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.file_entry = tk.Entry(root, width=50)
        self.file_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=10, pady=5)

        tk.Label(root, text="Sheet Name:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.sheet_entry = tk.Entry(root, width=20)
        self.sheet_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        tk.Button(root, text="Load Companies", command=self.load_companies).grid(row=2, column=0, columnspan=3, pady=10)

        self.company_frame = tk.Frame(root)
        self.company_frame.grid(row=3, column=0, columnspan=3, pady=10)

        tk.Button(root, text="Send Emails", command=self.send_emails).grid(row=4, column=0, columnspan=3, pady=10)

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
            if "Компания" not in self.df.columns or "E-mail" not in self.df.columns:
                messagebox.showerror("Error", "The Excel file must contain 'Компания' and 'E-mail' columns.")
                return

            # Приведение значений столбца "Компания" к строковому типу
            self.df["Компания"] = self.df["Компания"].astype(str)

            # Очистка предыдущих виджетов
            for widget in self.company_frame.winfo_children():
                widget.destroy()

            self.check_vars = {}
            for company in sorted(self.df["Компания"].unique()):
                var = tk.BooleanVar()
                cb = tk.Checkbutton(self.company_frame, text=company, variable=var)
                cb.pack(anchor="w")
                self.check_vars[company] = var

            messagebox.showinfo("Success", "Companies loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load companies: {e}")

    def send_emails(self):
        selected_companies = [company for company, var in self.check_vars.items() if var.get()]
        if not selected_companies:
            messagebox.showerror("Error", "No companies selected.")
            return

        try:
            outlook = win32.Dispatch("Outlook.Application")
            account = outlook.GetNamespace("MAPI").Accounts[0]

            log_file = "email_log.txt"
            with open(log_file, "a", encoding="utf-8") as log:
                for company in selected_companies:
                    company_data = self.df[self.df["Компания"] == company]
                    if company_data.empty:
                        continue

                    email = company_data["E-mail"].iloc[0]
                    debts = []
                    for _, row in company_data.iterrows():
                        debt_info = (f"Номер претензии: {row['Номер претензии']}, Инвойс: {row['Инвойс']}, "
                                     f"Дата претензии: {row['Дата претензии']}, Задолженность: {row['Задолженность']}")
                        debts.append(debt_info)

                    body = (f"Уважаемый партнер,\n\nУ вас имеется задолженность:\n\n" + "\n".join(debts) +
                            "\n\nПросьба оплатить в ближайшее время.\n\nС уважением, ваша компания.")

                    mail = outlook.CreateItem(0)
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Привязка отправителя
                    mail.To = email
                    mail.Subject = "Напоминание о задолженности"
                    mail.Body = body
                    mail.Send()

                    # Логирование
                    log.write(f"{datetime.now()} - Email sent to {company} ({email})\n")

                messagebox.showinfo("Success", "Emails sent successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DebtNotifierApp(root)
    root.mainloop()
