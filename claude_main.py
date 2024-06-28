import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from datetime import datetime

def get_category_or_code(program_name, dictionary):
    return next((key for key, values in dictionary.items() if program_name in values), None)

def clean_dollar_amount(amount):
    return float(amount.replace('$', '').replace(',', '').strip())

def browse_file(entry_widget):
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)

def process_data():
    stripe_file = stripe_entry.get()
    other_file = other_entry.get()
    start_date = pd.to_datetime(start_date_entry.get_date())
    end_date = pd.to_datetime(end_date_entry.get_date())

    if not stripe_file or not other_file:
        messagebox.showerror("Error", "Please select both Stripe and Other CSV files.")
        return

    try:
        update_status("Loading data...")
        codes_df = pd.read_excel('reportLayoutAndCodes.xlsx', sheet_name='Category Codes')
        stripe_df = pd.read_csv(stripe_file)
        other_df = pd.read_csv(other_file)

        update_status("Cleaning and processing data...")
        stripe_df_cleaned = clean_stripe_data(stripe_df, start_date, end_date)
        
        codes_df['Program'] = codes_df['Program'].str.rstrip()
        category_codes = codes_df.groupby('Code')['Program'].apply(list).to_dict()
        categories = codes_df.groupby('Category')['Program'].apply(list).to_dict()
        
        category_codes = {k: [x.replace("\xa0", " ") for x in v] for k, v in category_codes.items()}
        
        other_df['Amount'] = other_df['Amount'].apply(clean_dollar_amount)

        rows = process_rows(stripe_df_cleaned, other_df, category_codes, categories)
        
        final_df = pd.DataFrame(rows).sort_values('Session Date')
        final_df.to_excel("final_report.xlsx", index=False)
        messagebox.showinfo("Success", "Processing complete. Final report saved as 'final_report.xlsx'")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def clean_stripe_data(stripe_df, start_date, end_date):
    stripe_df_cleaned = stripe_df.dropna(subset=['id'])
    stripe_df_cleaned = stripe_df_cleaned[~((stripe_df_cleaned['Captured'] == 'FALSE') | (stripe_df_cleaned['Status'] == 'Failed'))]
    stripe_df_cleaned['Created date (UTC)'] = pd.to_datetime(stripe_df_cleaned['Created date (UTC)'])
    return stripe_df_cleaned[(stripe_df_cleaned['Created date (UTC)'].dt.date >= start_date.date()) & 
                             (stripe_df_cleaned['Created date (UTC)'].dt.date <= end_date.date())]

def process_rows(stripe_df_cleaned, other_df, category_codes, categories):
    rows = []
    total_rows = len(stripe_df_cleaned)
    progress_bar["maximum"] = total_rows

    for index, row in stripe_df_cleaned.iterrows():
        stripe_id = row['id']
        stripe_amount = row['Amount']
        stripe_amount_refunded = row['Amount Refunded']
        stripe_fee = row['Fee']
        amount_after_fees = stripe_amount - stripe_fee
        
        other_records = other_df[other_df['Payment Ref'] == stripe_id].reset_index(drop=True)
        
        program_name, session_date = get_program_info(other_records)
        
        category_code = get_category_or_code(program_name, category_codes)
        category = get_category_or_code(program_name, categories)

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
        
        update_progress(index)
    
    return rows

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

def update_status(message):
    status_label.config(text=message)
    root.update()

def update_progress(value):
    progress_bar["value"] = value + 1
    root.update()

# GUI setup
root = tk.Tk()
root.title("Data Processing GUI")

# Configure a dark theme style
style = ttk.Style(root)
style.theme_use('clam')  # or 'alt', 'default', 'classic' depending on your system

# File selection
for row, (label_text, entry_var) in enumerate([("Stripe CSV:", "stripe_entry"), ("Other CSV:", "other_entry")]):
    tk.Label(root, text=label_text).grid(row=row, column=0, sticky="e", padx=5, pady=5)
    globals()[entry_var] = tk.Entry(root, width=50)
    globals()[entry_var].grid(row=row, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda e=globals()[entry_var]: browse_file(e)).grid(row=row, column=2, padx=5, pady=5)

# Date range selection with improved visibility
date_style = ttk.Style()
date_style.configure('my.DateEntry', 
                     fieldbackground='darkblue', 
                     background='darkblue', 
                     foreground='white', 
                     arrowcolor='white')

for row, (label_text, entry_var) in enumerate([("Start Date:", "start_date_entry"), ("End Date:", "end_date_entry")], start=2):
    tk.Label(root, text=label_text).grid(row=row, column=0, sticky="e", padx=5, pady=5)
    globals()[entry_var] = DateEntry(root, width=12, style='my.DateEntry', 
                                     selectbackground='darkblue', 
                                     selectforeground='white',
                                     normalbackground='darkblue', 
                                     normalforeground='white',
                                     background='darkblue', 
                                     foreground='white',
                                     bordercolor='darkblue', 
                                     headersbackground='darkblue',
                                     headersforeground='white')
    globals()[entry_var].grid(row=row, column=1, sticky="w", padx=5, pady=5)

# Status and progress
status_label = tk.Label(root, text="")
status_label.grid(row=4, column=1, pady=5)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=5, column=1, pady=5)

# Process button
tk.Button(root, text="Process Data", command=process_data).grid(row=6, column=1, pady=20)

root.mainloop()