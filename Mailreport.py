import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os

class EmailReportTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Report Generator")
        self.file_path = ""
        self.df = pd.DataFrame()

        # UI Elements
        self.upload_btn = tk.Button(root, text="Import CSV", command=self.load_csv)
        self.upload_btn.pack(pady=5)

        self.email_label = tk.Label(root, text="Enter Email Addresses, comma or newline separated:")
        self.email_label.pack()
        self.email_text = tk.Text(root, height=5, width=50)
        self.email_text.pack(pady=5)

        self.dates_label = tk.Label(root, text="Enter Dates (YYYY-MM-DD), comma or newline separated:")
        self.dates_label.pack()
        self.dates_text = tk.Text(root, height=5, width=50)
        self.dates_text.pack(pady=5)

        self.generate_btn = tk.Button(root, text="Generate Report", command=self.generate_report)
        self.generate_btn.pack(pady=5)

        self.tree = ttk.Treeview(root)
        self.tree.pack(pady=10, fill='both', expand=True)

    def load_csv(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if self.file_path:
            try:
                self.df = pd.read_csv(self.file_path, encoding='latin1')
                messagebox.showinfo("Success", "CSV file loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {e}")

    def generate_report(self):
        if self.df.empty:
            messagebox.showwarning("No Data", "Please load a CSV file first.")
            return

        email_input = self.email_text.get("1.0", tk.END).strip().replace("\n", ",")
        email_list = [e.strip() for e in email_input.split(",") if e.strip()]

        date_input = self.dates_text.get("1.0", tk.END).strip().replace("\n", ",")
        date_list = [d.strip() for d in date_input.split(",") if d.strip()]

        try:
            date_list = [datetime.strptime(d, "%Y-%m-%d").date() for d in date_list]
        except ValueError:
            messagebox.showerror("Invalid Date Format", "Dates must be in YYYY-MM-DD format.")
            return

        # Preprocess
        self.df['date'] = pd.to_datetime(self.df['date_time_utc'], errors='coerce').dt.date
        filtered_df = self.df[
            (self.df['date'].isin(date_list)) &
            (self.df['sender_address'].isin(email_list)) &
            (~self.df['message_subject'].str.contains("Automatic reply", case=False, na=False))
        ]

        result = filtered_df.groupby(['sender_address', 'date']).size().reset_index(name='email_count')

        # Show in Treeview
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(result.columns)
        self.tree["show"] = "headings"
        for col in result.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        for _, row in result.iterrows():
            self.tree.insert("", "end", values=list(row))

        # Export to CSV
        export_path = os.path.join(os.path.dirname(self.file_path), "email_report.csv")
        result.to_csv(export_path, index=False)
        messagebox.showinfo("Report Generated", f"Report saved to {export_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailReportTool(root)
    root.mainloop()
