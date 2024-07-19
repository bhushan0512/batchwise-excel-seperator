import pandas as pd
import os

# Load the data from the Excel file
file_path = './pending.xlsx'
data = pd.read_excel(file_path, sheet_name='Worksheet')

# Get unique courses and batches
courses = data['Course'].unique()
batches = data['Batch'].unique()

# Create output directory if it doesn't exist
output_dir = 'output_files'
os.makedirs(output_dir, exist_ok=True)

# Separate data based on Course and Batch and save to separate files
for course in courses:
    for batch in batches:
        subset = data[(data['Course'] == course) & (data['Batch'] == batch)]
        if not subset.empty:
            output_file = f"{output_dir}/{course}_Batch_{batch}.xlsx"
            subset.to_excel(output_file, index=False)
            print(f"Saved {output_file}")

print("Files have been successfully separated and saved.")
