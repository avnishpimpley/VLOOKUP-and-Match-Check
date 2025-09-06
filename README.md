# VLOOKUP-and-Match-Check
Loading data from Google Drive into a Google Colab environment.
Generating a summary sheet with the unique counts of values from a specified column.
Performing a VLOOKUP-style operation to merge data from two separate Excel files based on a common key.
Comparing two columns to identify and mark rows as 'Match' or 'Mismatch'.

#Reading Files
import pandas as pd
file_path1 = '/content/WS1.xlsx'
WS1 = pd.read_excel(file_path1)
WS1

file_path2 = '/content/WS2.xlsx'
WS2 = pd.read_excel(file_path2)
WS2

#Creating a new sheet and Counting occurence of unique value
WS2 = pd.read_excel(file_path2, "Sheet1" == 0)
count_summary = WS2['A/C Reference'].value_counts().reset_index()
count_summary.columns = ['Unique Value', 'Count']
with pd.ExcelWriter(file_path2, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        count_summary.to_excel(writer, sheet_name='Count Summary', index=False)

print(f"Success! A new sheet named 'Count Summary' has been added to {file_path2}")
print("The original data in the first sheet remains untouched.")

#VLOOKUP values and Create a new sheet
file1_sheet_name = 'Sheet1'
file2_sheet_name = 'Count Summary'
lookup_key_column_file2 = 'Unique Value' #Unique Value
lookup_key_column_file1 = 'A/C Reference' #A/C Reference
data_to_pull_column = 'Days for Settlement' #Days for Settlement

WS1_S1 = pd.read_excel(file_path1, sheet_name=file1_sheet_name)
WS2_S2 = pd.read_excel(file_path2, sheet_name=file2_sheet_name)

WS1_indexed = WS1_S1.set_index(lookup_key_column_file1)
WS2_indexed = WS2_S2.set_index(lookup_key_column_file2)

joined_df = WS2_indexed.join(WS1_indexed[[data_to_pull_column]], how='left')
joined_df = joined_df.reset_index()
joined_df

with pd.ExcelWriter(file_path2, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        joined_df.to_excel(writer, sheet_name='VLOOKUP Result', index=False)

#Understanding Match Mismatch Between Values
import numpy as np
sheet_name_to_read = 'VLOOKUP Result'

column_B = 'Count'
column_C = 'Days for Settlement'

df = pd.read_excel(file_path2, sheet_name=sheet_name_to_read)
df['Match or Mismatch'] = np.where(df[column_B] == df[column_C], 'Match', 'Mismatch')
df

with pd.ExcelWriter(file_path2, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Comparison Result', index=False)
