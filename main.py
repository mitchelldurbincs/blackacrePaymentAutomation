import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from datetime import datetime

# Constants
TITLE = "Data Processing GUI"
BACKGROUND = "darkblue"
FOREGROUND = "white"
CODE_SHEET_NAME = "Category Codes"
FINAL_REPORT_NAME = "final_report.xlsx"

class DataProcessingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(TITLE)
        self.setup_ui()

    def setup_ui(self):
        self.setup_style()
        self.create_file_inputs()
        self.create_date_inputs()
        self.create_status_and_progress()
        self.create_process_button()

    def setup_style(self):
        style = ttk.Style(self.root)
        style.theme_use('clam')
        date_style = ttk.Style()
        date_style.configure('my.DateEntry', 
                             fieldbackground=BACKGROUND, 
                             background=BACKGROUND, 
                             foreground=FOREGROUND, 
                             arrowcolor=FOREGROUND)

    def create_file_inputs(self):
        for row, (label_text, attr_name) in enumerate([
            ("Stripe CSV:", "stripe_entry"),
            ("Other CSV:", "other_entry")
        ]):
            tk.Label(self.root, text=label_text).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            entry = tk.Entry(self.root, width=50)
            entry.grid(row=row, column=1, padx=5, pady=5)
            setattr(self, attr_name, entry)
            tk.Button(self.root, text="Browse", command=lambda e=entry: self.browse_file(e)).grid(row=row, column=2, padx=5, pady=5)

    def create_date_inputs(self):
        for row, (label_text, attr_name) in enumerate([
            ("Start Date:", "start_date_entry"),
            ("End Date:", "end_date_entry")
        ], start=2):
            tk.Label(self.root, text=label_text).grid(row=row, column=0, sticky="e", padx=5, pady=5)
            date_entry = DateEntry(self.root, width=12, style='my.DateEntry', 
                                   selectbackground=BACKGROUND, 
                                   selectforeground=FOREGROUND,
                                   normalbackground=BACKGROUND, 
                                   normalforeground=FOREGROUND,
                                   background=BACKGROUND, 
                                   foreground=FOREGROUND,
                                   bordercolor=BACKGROUND, 
                                   headersbackground=BACKGROUND,
                                   headersforeground=FOREGROUND)
            date_entry.grid(row=row, column=1, sticky="w", padx=5, pady=5)
            setattr(self, attr_name, date_entry)

    def create_status_and_progress(self):
        self.status_label = tk.Label(self.root, text="")
        self.status_label.grid(row=4, column=1, pady=5)

        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.grid(row=5, column=1, pady=5)

    def create_process_button(self):
        tk.Button(self.root, text="Process Data", command=self.process_data).grid(row=6, column=1, pady=20)

    def browse_file(self, entry_widget):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filename)

    def safe_read_file(self, file_path, file_type, sheet_name=None):
        try:
            if file_type == 'csv':
                return pd.read_csv(file_path)
            elif file_type == 'excel':
                return pd.read_excel(file_path, sheet_name=sheet_name)
        except FileNotFoundError:
            raise ValueError(f"The file {file_path} was not found. Please check the file path.")
        except pd.errors.EmptyDataError:
            raise ValueError(f"The file {file_path} is empty. Please check the file contents.")
        except pd.errors.ParserError:
            raise ValueError(f"Unable to parse {file_path}. Please ensure it's a valid {file_type.upper()} file.")
        except Exception as e:
            raise ValueError(f"An error occurred while reading {file_path}: {str(e)}")

    def process_data(self):
        stripe_file = self.stripe_entry.get()
        other_file = self.other_entry.get()
        start_date = pd.to_datetime(self.start_date_entry.get_date())
        end_date = pd.to_datetime(self.end_date_entry.get_date())

        if not stripe_file or not other_file:
            messagebox.showerror("Error", "Please select both Stripe and Other CSV files.")
            return

        try:
            self.update_status("Loading data...")
            stripe_df = safe_read_file(stripe_file, 'csv')
            other_df = safe_read_file(other_file, 'csv')
            codes_df = safe_read_file('reportLayoutAndCodes.xlsx', 'excel', CODE_SHEET_NAME)

            self.update_status("Cleaning and processing data...")
            stripe_df_cleaned = self.clean_stripe_data(stripe_df, start_date, end_date)
            
            codes_df['Program'] = codes_df['Program'].str.rstrip()
            category_codes = codes_df.groupby('Code')['Program'].apply(list).to_dict()
            categories = codes_df.groupby('Category')['Program'].apply(list).to_dict()
            
            category_codes = {k: [x.replace("\xa0", " ") for x in v] for k, v in category_codes.items()}
            
            other_df['Amount'] = other_df['Amount'].apply(self.clean_dollar_amount)

            rows = self.process_rows(stripe_df_cleaned, other_df, category_codes, categories)
            
            final_df = pd.DataFrame(rows).sort_values('Session Date')
            final_df.to_excel(FINAL_REPORT_NAME, index=False)
            messagebox.showinfo("Success", f"Processing complete. Final report saved as '{FINAL_REPORT_NAME}'")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    @staticmethod
    def get_category_or_code(program_name, dictionary):
        return next((key for key, values in dictionary.items() if program_name in values), None)

    @staticmethod
    def clean_dollar_amount(amount):
        return float(amount.replace('$', '').replace(',', '').strip())

    @staticmethod
    def clean_stripe_data(stripe_df, start_date, end_date):
        stripe_df_cleaned = stripe_df.dropna(subset=['id'])
        stripe_df_cleaned = stripe_df_cleaned[~((stripe_df_cleaned['Captured'] == 'FALSE') | (stripe_df_cleaned['Status'] == 'Failed'))]
        stripe_df_cleaned['Created date (UTC)'] = pd.to_datetime(stripe_df_cleaned['Created date (UTC)'])
        return stripe_df_cleaned[(stripe_df_cleaned['Created date (UTC)'].dt.date >= start_date.date()) & 
                                 (stripe_df_cleaned['Created date (UTC)'].dt.date <= end_date.date())]

    def process_rows(self, stripe_df_cleaned, other_df, category_codes, categories):
        rows = []
        total_rows = len(stripe_df_cleaned)
        self.progress_bar["maximum"] = total_rows

        for index, row in stripe_df_cleaned.iterrows():
            stripe_id = row['id']
            stripe_amount = row['Amount']
            stripe_fee = row['Fee']
            amount_after_fees = stripe_amount - stripe_fee
            
            other_records = other_df[other_df['Payment Ref'] == stripe_id].reset_index(drop=True)
            
            program_name, session_date = self.get_program_info(other_records)
            
            category_code = self.get_category_or_code(program_name, category_codes)
            category = self.get_category_or_code(program_name, categories)

            rows.append({
                'Session Date': session_date,
                'Category': category,
                'Program': program_name,
                'Category Code': category_code,
                'Amount': stripe_amount,
                'Fees': stripe_fee,
                'Amount after Fees': amount_after_fees,
                'Payment Ref': stripe_id
            })
            
            self.update_progress(index)
        
        return rows

    @staticmethod
    def get_program_info(other_records):
        unique_programs = other_records['Program'].unique()
        if len(unique_programs) == 2:
            program_name = other_records[other_records['Program'] != "Payment (Thank you)"]['Program'].values[0].rstrip()
        elif len(unique_programs) > 2:
            program_name = 'More than one unique program'
        else:
            program_name = None
        
        non_payment_thank_you_records = other_records[other_records['Program'].str.rstrip() != "Payment (Thank you)"]
        if not non_payment_thank_you_records.empty:
            program_name = non_payment_thank_you_records['Program'].values[0].rstrip()
            session_date = non_payment_thank_you_records['Session Date'].values[0].rstrip()
        else:
            session_date = None
        
        return program_name, session_date

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update()

    def update_progress(self, value):
        self.progress_bar["value"] = value + 1
        self.root.update()

if __name__ == "__main__":
    root = tk.Tk()
    app = DataProcessingGUI(root)
    root.mainloop()