import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from tkcalendar import Calendar
from datetime import datetime

class MailReportTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Report Tool")
        self.df = None
        self.selected_dates = set()
        self.build_gui()

    def build_gui(self):
        # File upload
        self.upload_btn = tk.Button(self.root, text="Upload CSV", command=self.load_csv)
        self.upload_btn.pack(pady=5)

        # Sender filter
        tk.Label(self.root, text="Filter by Sender Email:").pack()
        self.sender_entry = ttk.Combobox(self.root)
        self.sender_entry.pack(pady=5)

        # Calendar for multi-date selection
        tk.Label(self.root, text="Select Dates:").pack()
        self.cal = Calendar(self.root, selectmode='day')
        self.cal.pack(pady=5)

        tk.Button(self.root, text="Add Date", command=self.add_date).pack(pady=2)
        self.date_listbox = tk.Listbox(self.root, height=5)
        self.date_listbox.pack(pady=5)

        tk.Button(self.root, text="Remove Selected Date", command=self.remove_date).pack(pady=2)

        # Generate Report
        self.report_btn = tk.Button(self.root, text="Generate Report", command=self.generate_report)
        self.report_btn.pack(pady=5)

        # Treeview
        self.tree = ttk.Treeview(self.root, columns=("Sender", "Date", "Count"), show='headings')
        self.tree.heading("Sender", text="Sender")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Count", text="Mail Count")
        self.tree.pack(pady=10, fill=tk.BOTH, expand=True)

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        try:
            self.df = pd.read_csv(file_path, encoding='ISO-8859-1', quotechar='\"', on_bad_lines='skip')
            self.df.columns = [col.strip().lower() for col in self.df.columns]

            if 'date_time_utc' in self.df.columns:
                self.df['date_time_utc'] = pd.to_datetime(self.df['date_time_utc'], errors='coerce')
                self.df['date'] = self.df['date_time_utc'].dt.date
            else:
                messagebox.showerror("Error", "Column 'date_time_utc' not found in the file.")
                return

            senders = sorted(self.df['sender_address'].dropna().unique())
            self.sender_entry['values'] = senders
            messagebox.showinfo("Success", "CSV Loaded Successfully")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    def add_date(self):
        selected = self.cal.get_date()
        try:
            formatted_date = datetime.strptime(selected, "%m/%d/%y").date()
        except:
            formatted_date = datetime.strptime(selected, "%m/%d/%Y").date()

        if formatted_date not in self.selected_dates:
            self.selected_dates.add(formatted_date)
            self.date_listbox.insert(tk.END, formatted_date)

    def remove_date(self):
        selected_indices = self.date_listbox.curselection()
        for index in reversed(selected_indices):
            date_value = self.date_listbox.get(index)
            self.selected_dates.discard(datetime.strptime(date_value, "%Y-%m-%d").date())
            self.date_listbox.delete(index)

    def generate_report(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please upload a CSV file first.")
            return

        filtered_df = self.df[self.df['event_id'] == 'AGENTINFO']

        if self.selected_dates:
            filtered_df = filtered_df[filtered_df['date'].isin(self.selected_dates)]

        sender = self.sender_entry.get()
        if sender:
            filtered_df = filtered_df[filtered_df['sender_address'] == sender]

        report = filtered_df.groupby(['sender_address', 'date']).size().reset_index(name='mail_count')

        for row in self.tree.get_children():
            self.tree.delete(row)

        for _, row in report.iterrows():
            self.tree.insert("", tk.END, values=(row['sender_address'], row['date'], row['mail_count']))

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("700x600")
    app = MailReportTool(root)
    root.mainloop()
