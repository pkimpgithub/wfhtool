import pandas as pd
from tkinter import Tk, Label, Button, filedialog, messagebox, Listbox, MULTIPLE, END
from tkcalendar import Calendar
from datetime import datetime

class MailLogReporter:
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Log Report Generator")
        self.file_path = ""

        Label(root, text="1. Select Mail Log Excel File").pack(pady=5)
        Button(root, text="Browse Excel File", command=self.load_file).pack(pady=5)

        Label(root, text="2. Select Dates").pack(pady=5)
        self.calendar = Calendar(root, selectmode='day')
        self.calendar.pack(pady=5)

        Button(root, text="Add Selected Date", command=self.add_date).pack(pady=5)
        self.date_listbox = Listbox(root, selectmode=MULTIPLE, width=30, height=6)
        self.date_listbox.pack(pady=5)

        Button(root, text="Remove Selected Dates", command=self.remove_dates).pack(pady=5)

        Button(root, text="3. Generate Report", command=self.generate_report).pack(pady=10)

    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            messagebox.showinfo("File Selected", f"Selected file:\n{self.file_path}")

    def add_date(self):
        date = self.calendar.get_date()
        if date not in self.date_listbox.get(0, END):
            self.date_listbox.insert(END, date)

    def remove_dates(self):
        selected = list(self.date_listbox.curselection())
        for index in reversed(selected):
            self.date_listbox.delete(index)

    def generate_report(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select an Excel file first.")
            return

        try:
            df = pd.read_excel(self.file_path, sheet_name=0)
            
            # Check required columns
            required_columns = ['Timestamp', 'Sender Email']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", f"Missing columns. Required: {required_columns}")
                return

            # Clean and convert
            df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce')
            df = df.dropna(subset=['Timestamp'])
            df['Date'] = df['Timestamp'].dt.date.astype(str)

            selected_dates = self.date_listbox.get(0, END)
            if not selected_dates:
                messagebox.showerror("Error", "Please select at least one date.")
                return

            # Filter and group
            filtered_df = df[df['Date'].isin(selected_dates)]
            if filtered_df.empty:
                messagebox.showinfo("No Data", "No emails sent on the selected dates.")
                return

            report = (
                filtered_df
                .groupby(['Date', 'Sender Email'])
                .agg(Emails_Sent=('Sender Email', 'count'))
                .reset_index()
            )

            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if output_path:
                report.to_excel(output_path, index=False)
                messagebox.showinfo("Success", f"Report saved to:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    root = Tk()
    app = MailLogReporter(root)
    root.mainloop()
