import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from datetime import datetime

def browse_file(entry_widget):
    filetypes = [("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
    filename = filedialog.askopenfilename(filetypes=filetypes)
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)

# Helper functions
def get_category_or_code(program_name, dictionary):
    for key in dictionary.keys():
        if program_name in dictionary[key]:
            return key
    return None

def clean_dollar_amount(amount):
    return float(amount.replace('$', '').replace(',', '').strip())

def process_data():
    stripe_file = stripe_entry.get()
    other_file = other_entry.get()
    category_codes_file = category_codes_entry.get()
    start_date = pd.to_datetime(start_date_entry.get_date())
    end_date = pd.to_datetime(end_date_entry.get_date())

    if not all([stripe_file, other_file, category_codes_file]):
        messagebox.showerror("Error", "Please select all required files.")
        return

    try:
        update_status("Loading data...")
        codes_df = pd.read_excel(category_codes_file)
        stripe_df = pd.read_csv(stripe_file)
        other_df = pd.read_csv(other_file)

        # Update status
        status_label.config(text="Cleaning and processing data...")
        root.update()

        # Clean the stripe_df
        stripe_df_cleaned = stripe_df.dropna(subset=['id'])
        stripe_df_cleaned = stripe_df_cleaned[~((stripe_df_cleaned['Captured'] == 'FALSE') | (stripe_df_cleaned['Status'] == 'Failed'))]

        # Convert dates
        stripe_df_cleaned['Created date (UTC)'] = pd.to_datetime(stripe_df_cleaned['Created date (UTC)'])
        other_df['Payment Date'] = pd.to_datetime(other_df['Payment Date'])

        # Filter by date range
        stripe_df_cleaned = stripe_df_cleaned[(stripe_df_cleaned['Created date (UTC)'].dt.date >= start_date.date()) & 
                                      (stripe_df_cleaned['Created date (UTC)'].dt.date <= end_date.date())]

        
        # remove trailing spaces from programs in codes_df
        codes_df['Program'] = codes_df['Program'].apply(lambda x: x.rstrip())

        # Get a dictionary of the codes:program from the A:B columns of codes_df
        category_codes = codes_df.groupby('Code')['Program'].apply(list).to_dict()

        categories = codes_df.groupby('Category')['Program'].apply(list).to_dict()

        # remove "\xa0" from the program names in category codes
        for key in category_codes.keys():
            category_codes[key] = [x.replace("\xa0", " ") for x in category_codes[key]]


        # Clean the Amount column in other_df
        other_df['Amount'] = other_df['Amount'].apply(clean_dollar_amount)

        # Update progress bar
        total_rows = len(stripe_df_cleaned)
        progress_bar["maximum"] = total_rows

        rows = []
        # Iterate through each row in stripe_df_cleaned
        for index, row in stripe_df_cleaned.head(1000).iterrows():
            stripe_id = row['id']
            stripe_amount = row['Amount']
            stripe_amount_refunded = row['Amount Refunded']
            stripe_fee = row['Fee']
            amount_after_fees = stripe_amount - stripe_fee
            
            # Get the rows with the same id in other_df
            other_records = other_df[other_df['Payment Ref'] == stripe_id].reset_index(drop=True)
            other_amounts = other_records['Amount']

            # Get the unique program names from other_records
            unique_programs = other_records['Program'].unique()

            # if the len of unique programs is two
            if len(unique_programs) == 2:
                # get the program name that is not "Payment (Thank you)"
                program_name = other_records[other_records['Program'] != "Payment (Thank you)"]['Program'].values[0].rstrip()
            elif len(unique_programs) > 2:
                program_name = 'More than one unique program'
                
            else:
                print("Unique programs less than 2")
                print("Unique Programs:", unique_programs)
                print("Other records: ", other_records)

            #print(other_records)
            # Get the index of "Payment (Thank you)" within other_records
            check_payment_thank_you = other_records['Program'].str.rstrip() == "Payment (Thank you)"
            if check_payment_thank_you.any():
                payment_thank_you_index = check_payment_thank_you.idxmax()
            else:
                payment_thank_you_index = None
                print("Payment (Thank you) not found")
                print("Stripe id", stripe_id)
            #print("Payment Thank You Index in other_records:", payment_thank_you_index)
            
            # Exclude the "Payment (Thank you)" row to get the program name and session date
            non_payment_thank_you_records = other_records[other_records['Program'].str.rstrip() != "Payment (Thank you)"]

            if not non_payment_thank_you_records.empty:
                program_name = non_payment_thank_you_records['Program'].values[0].rstrip()
                #print('Program Name:"', program_name + '"')
                session_date = non_payment_thank_you_records['Session Date'].values[0].rstrip()
                #print("Session Date:", session_date)
            else:
                program_name = None
                session_date = None

            # Get the category code
            category_code = get_category_or_code(program_name, category_codes)
            category = get_category_or_code(program_name, categories)

            print("Program:", program_name, "Category:", category)

            # Make sure the amounts add to zero
            if stripe_amount_refunded == 0:
                sum_of_amounts = other_records['Amount'].astype(float).sum()
                if sum_of_amounts == 0:
                    print("Amounts add up")
                    print("Amounts do not add up")
                else: 
                    print("Amounts do not add up")
            else:
                print('REFUND')
            
            # put into final_df 
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
            # Update progress bar
            progress_bar["value"] = index + 1
            root.update()
        
        status_label.config(text="Processing complete!")
        root.update()

        final_df = pd.DataFrame(rows)

        # sort by the session date
        final_df = final_df.sort_values('Session Date')

        # Save the final report
        final_df.to_excel("final_report2.xlsx", index=False)
        messagebox.showinfo("Success", "Processing complete. Final report saved as 'final_report.xlsx'")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# GUI setup
root = tk.Tk()
root.title("Data Processing GUI")

# Configure a dark theme style
style = ttk.Style(root)
style.theme_use('clam')

# File selection
file_entries = [
    ("Stripe CSV:", "stripe_entry"),
    ("Other CSV:", "other_entry"),
    ("Category Codes Excel file:", "category_codes_entry")
]

for row, (label_text, entry_var) in enumerate(file_entries):
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

for row, (label_text, entry_var) in enumerate([("Start Date:", "start_date_entry"), ("End Date:", "end_date_entry")], start=len(file_entries)):
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
status_label.grid(row=len(file_entries)+2, column=1, pady=5)

# Add this line for the progress label
tk.Label(root, text="Progress:").grid(row=len(file_entries)+3, column=1, pady=(10,0), sticky="sw")

# Update the row number for the progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=len(file_entries)+4, column=1, pady=(0,5))

# Update the row number for the Process button
tk.Button(root, text="Process Data", command=process_data).grid(row=len(file_entries)+5, column=1, pady=20)

root.mainloop()

# TODO: Attempt to fix calendar issue?
# TODO: Test on Windows
# TODO: Refunds
# TODO: Put the results on the same excel sheet as the category codes
# TODO: Add a button to add the category codes 
