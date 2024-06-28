import pandas as pd

# Load the data
codes_df = pd.read_excel('reportLayoutAndCodes.xlsx', sheet_name='Category Codes')
stripe_df = pd.read_csv("stripe.csv")
other_df = pd.read_csv("other.csv")

# Clean the stripe_df
stripe_df_cleaned = stripe_df.dropna(subset=['id'])
stripe_df_cleaned = stripe_df_cleaned[~((stripe_df_cleaned['Captured'] == 'FALSE') | (stripe_df_cleaned['Status'] == 'Failed'))]

# Get a dictionary of the codes:program from the A:B columns of codes_df
category_codes = codes_df.groupby('Code')['Program'].apply(list).to_dict()

# remove "\xa0" from the program names in category codes
for key in category_codes.keys():
    category_codes[key] = [x.replace("\xa0", " ") for x in category_codes[key]]

# Helper functions
def get_category_code(program_name):
    for key in category_codes.keys():
        if program_name in category_codes[key]:
            return key
    return None

def clean_dollar_amount(amount):
    return float(amount.replace('$', '').replace(',', '').strip())

# Clean the Amount column in other_df
other_df['Amount'] = other_df['Amount'].apply(clean_dollar_amount)

# print number of unique payment refs in stripe_df_cleaned
print("Number of unique payment refs in stripe_df_cleaned:", stripe_df_cleaned['id'].nunique())

rows = []
# Iterate through each row in stripe_df_cleaned
for index, row in stripe_df_cleaned.head(1000).iterrows():
    stripe_id = row['id']
    stripe_amount = row['Amount']
    stripe_amount_refunded = row['Amount Refunded']
    stripe_fee = row['Fee']
    amount_after_fees = stripe_amount - stripe_fee
    
    # Get the rows with the same id in other_df
    other_records = other_df[other_df['Payment Ref'] == stripe_id]
    other_amounts = other_records['Amount']

    # Check if no refunds
    if stripe_amount_refunded == 0: 
        # Initialize a flag to check for "Payment (Thank you)"
        payment_thank_you_index = None
        unique_programs = set()
        # Iterate through other_records
        for idx, record in other_records.iterrows():
            program = record['Program']
            if program == "Payment (Thank you)":
                payment_thank_you_index = idx
            else:
                unique_programs.add(program)
        
        # Determine if there are multiple unique programs
        more_than_one_unique_program = len(unique_programs) > 1
        if payment_thank_you_index is not None:
            if not more_than_one_unique_program:
                # get the program from the index that is NOT Payment Thank you
                other_idx = 0 if payment_thank_you_index == 1 else 1
                program = other_records.iloc[other_idx]['Program']
                category = get_category_code(program)

                # make sure all amounts add to 0
                check = other_records['Amount'].astype(float).sum()
                if check != 0:
                    print("error: amount did not add to 0")
                if category is not None:
                    rows.append({
                        'Session Date': other_records.iloc[1]['Session Date'],
                        'Category': category,
                        'Program': program,
                        'Amount': stripe_amount,
                        'Fees': stripe_fee,
                        'Amount after Fees': amount_after_fees,
                        'Payment Ref': stripe_id
                    })
            # more than one unique program
            else:
                program = "More than one unique program"
                rows.append({
                        'Session Date': other_records.iloc[1]['Session Date'],
                        'Category': get_category_code(program),
                        'Program': program,
                        'Amount': stripe_amount,
                        'Fees': stripe_fee,
                        'Amount after Fees': amount_after_fees,
                        'Payment Ref': stripe_id
                    })

        
        else:
            print("error: no payment thank you")
            print(unique_programs)
    else: 
        # will finish this later
        continue 

final_df = pd.DataFrame(rows)

# sort by the session date
final_df = final_df.sort_values('Session Date')
# convert the final_df to xlsx file named "final_report.xlsx"
final_df.to_excel("final_report.xlsx", index=False)