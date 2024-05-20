import pandas as pd

# Load your spreadsheet
input_file = 'your_spreadsheet.xlsx'
output_file = 'msn_output.xlsx'

# Specify the sheet names you want to filter
sheets_to_filter = ['Sheet1', 'Sheet3', 'Sheet5']  # replace with your sheet names

# Create a dictionary to hold the filtered data for each sheet
filtered_data = {sheet: [] for sheet in sheets_to_filter}

# Open the workbook for reading
workbook = pd.ExcelFile(input_file)

# Iterate over each sheet in the input workbook
for sheet_name in workbook.sheet_names:
    # Read the data from the current sheet
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    if sheet_name in sheets_to_filter:
        # Ensure the relevant columns are converted to integers for comparison
        df.iloc[:, 1] = pd.to_numeric(df.iloc[:, 1], errors='coerce').fillna(0).astype(int)
        df.iloc[:, 2] = pd.to_numeric(df.iloc[:, 2], errors='coerce').fillna(0).astype(int)

        # Apply the filter conditions with "greater than or equal to" condition
        filtered_df = df[(df.iloc[:, 1] >= 2) | (df.iloc[:, 2] == 0)]

        # Append the filtered data to the dictionary
        filtered_data[sheet_name].append(filtered_df)

# Open the workbook for writing
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Iterate over each sheet in the original workbook
    for sheet_name in workbook.sheet_names:
        # Read the original data from the current sheet
        df = pd.read_excel(input_file, sheet_name=sheet_name)

        if sheet_name in sheets_to_filter:
            # Concatenate the filtered data frames
            concatenated_filtered_df = pd.concat(filtered_data[sheet_name], ignore_index=True)

            # Concatenate the original and filtered data
            df = pd.concat([df, concatenated_filtered_df], ignore_index=True)

        # Write the data to the new workbook
        df.to_excel(writer, sheet_name=sheet_name, index=False)
