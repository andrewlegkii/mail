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

        # Scrollable Frame
        self.scroll_canvas = tk.Canvas(root, width=600, height=300)
        self.scroll_canvas.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.scrollbar = tk.Scrollbar(root, orient="vertical", command=self.scroll_canvas.yview)
        self.scrollbar.grid(row=3, column=3, sticky="ns")
        
        self.scroll_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.inner_frame = tk.Frame(self.scroll_canvas)
        self.scroll_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        self.inner_frame.bind("<Configure>", lambda e: self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all")))

        tk.Label(root, text="Choose Email Account:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.account_entry = tk.Entry(root, width=50)
        self.account_entry.grid(row=4, column=1, padx=10, pady=5, sticky="w")

        tk.Button(root, text="Send Emails", command=self.send_emails).grid(row=5, column=0, columnspan=3, pady=10)

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

            self.df["Компания"] = self.df["Компания"].astype(str)

            for widget in self.inner_frame.winfo_children():
                widget.destroy()

            self.check_vars = {}
            for company in sorted(self.df["Компания"].unique()):
                var = tk.BooleanVar()
                cb = tk.Checkbutton(self.inner_frame, text=company, variable=var)
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

        account_email = self.account_entry.get()
        if not account_email:
            messagebox.showerror("Error", "Please provide an email account.")
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

                    # Create a table with the desired columns
                    table_data = company_data[["Номер претензии", "Компания", "Инвойс", "Дата претензии", "Задолженность"]]
                    table_html = table_data.to_html(index=False, justify="center", border=1)

                    # Create subject
                    subject = f"Напоминание о задолженности по претензиям ({company})"

                    body = (f"Уважаемый партнер,<br><br>У вас имеется задолженность:<br><br>" +
                            table_html +
                            "<br><br>Просьба оплатить в ближайшее время.<br><br>С уважением, ваша компания.")

                    mail = outlook.CreateItem(0)
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                    mail.To = email
                    mail.Subject = subject  # Use dynamic subject
                    mail.HTMLBody = body  # Use HTML body for table
                    mail.Send()

                    log.write(f"{datetime.now()} - Email sent to {company} ({email}) with Debt Table\n")

                messagebox.showinfo("Success", "Emails sent successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DebtNotifierApp(root)
    root.mainloop()
