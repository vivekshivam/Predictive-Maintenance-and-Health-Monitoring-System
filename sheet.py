import pandas as pd

# Step 1: Read the original Excel sheet into a DataFrame
input_file = 'classified_data.xlsx'
df = pd.read_excel(input_file)

# Step 2: Group the DataFrame by the desired category
category_column = 'Category'  # Replace with the column name you want to group by

# Create a new Excel writer object
output_file = 'classified_data_with_sheets.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# Step 3: Write each group to a separate sheet in a new Excel file
for category, group in df.groupby(category_column):
    # Remove characters from category names that are not allowed in Excel sheet names
    safe_category = ''.join(c if c.isalnum() else '_' for c in category)
    group.to_excel(writer, sheet_name=safe_category, index=False)

# Save the new Excel file
writer.close()

print(f'Successfully created {output_file} with separate sheets for each category.')
