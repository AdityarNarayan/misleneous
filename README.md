
import os
import pandas as pd

folder_path = 'your_folder_path'  # Path to the folder containing Excel files

# Create an empty list to store the DataFrames
dfs = []

# Loop through the files in the folder
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):  # Check if file is Excel
        file_path = os.path.join(folder_path, file_name)
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Append the DataFrame to the list
        dfs.append(df)

# Concatenate all DataFrames into a single DataFrame
combined_df = pd.concat(dfs, ignore_index=True)

# Print the combined DataFrame
print(combined_df)
