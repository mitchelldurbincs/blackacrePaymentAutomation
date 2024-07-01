import pandas as pd

# Load the data
codes_df = pd.read_excel('reportLayoutAndCodes.xlsx', sheet_name='Category Codes')
stripe_df = pd.read_csv("stripe.csv")
other_df = pd.read_csv("other.csv")

# Clean the stripe_df
stripe_df_cleaned = stripe_df.dropna(subset=['id'])
stripe_df_cleaned = stripe_df_cleaned[~((stripe_df_cleaned['Captured'] == 'FALSE') | (stripe_df_cleaned['Status'] == 'Failed'))]

# get the range of dates in stripe_df "Session Date" column
stripe_df_cleaned['Created date (UTC)'] = pd.to_datetime(stripe_df_cleaned['Created date (UTC)'])

# get the range of dates in other_df "Session Date" column
other_df['Payment Date'] = pd.to_datetime(other_df['Payment Date'])

# configure the date range for stripe_df_cleaned to be the min max from other_df
stripe_df_cleaned = stripe_df_cleaned[(stripe_df_cleaned['Created date (UTC)'] >= other_df['Payment Date'].min()) & (stripe_df_cleaned['Created date (UTC)'] <= other_df['Payment Date'].max())]
stripe_df_cleaned['Created date (UTC)'] = pd.to_datetime(stripe_df_cleaned['Created date (UTC)'])

# remove trailing spaces from programs in codes_df
codes_df['Program'] = codes_df['Program'].apply(lambda x: x.rstrip())

# Get a dictionary of the codes:program from the A:B columns of codes_df
category_codes = codes_df.groupby('Code')['Program'].apply(list).to_dict()

categories = codes_df.groupby('Category')['Program'].apply(list).to_dict()

# remove "\xa0" from the program names in category codes
for key in category_codes.keys():
    category_codes[key] = [x.replace("\xa0", " ") for x in category_codes[key]]

# Helper functions
def get_category_or_code(program_name, dictionary):
    for key in dictionary.keys():
        if program_name in dictionary[key]:
            return key
    return None

def clean_dollar_amount(amount):
    return float(amount.replace('$', '').replace(',', '').strip())

# Clean the Amount column in other_df
other_df['Amount'] = other_df['Amount'].apply(clean_dollar_amount)

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
    
final_df = pd.DataFrame(rows)

# sort by the session date
final_df = final_df.sort_values('Session Date')
# convert the final_df to xlsx file named "final_report.xlsx"
final_df.to_excel("final_report.xlsx", index=False)