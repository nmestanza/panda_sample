import pandas as pd

# Load your spreadsheet
input_file = 'your_spreadsheet.xlsx'
output_file = 'msn_output.xlsx'

# Specify the sheet names you want to filter
sheets_to_filter = ['Sheet1', 'Sheet3', 'Sheet5']  # replace with your sheet names

# Create a list to hold the filtered data from all sheets
all_filtered_data = []

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

        # Append the filtered data to the list
        all_filtered_data.append(filtered_df)

# Concatenate all filtered data frames into one
concatenated_filtered_data = pd.concat(all_filtered_data, ignore_index=True)

# Open the workbook for writing
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Write the concatenated filtered data to a new sheet
    concatenated_filtered_data.to_excel(writer, sheet_name='FilteredData', index=False)

# Optional: print the concatenated filtered data to the console
print("Concatenated filtered data:")
print(concatenated_filtered_data)
