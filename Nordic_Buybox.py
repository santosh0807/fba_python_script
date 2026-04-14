import pandas as pd
import os

# Input file path
input_file_path = "/home/santosh/Documents/Python/Scripts/FBA/FBA Input/Nordic BuyBox 03 Feb 2026.xlsx"

# Output directory
output_path = "/home/santosh/Documents/Python/Scripts/FBA/FBA Report/"

# Create output directory if it doesn't exist
os.makedirs(output_path, exist_ok=True)

# Output XLSX file
output_file = os.path.join(output_path, "Nordic_Competitor_BuyBox_Expanded_Report.xlsx")

# Check if input file exists
if not os.path.exists(input_file_path):
    print(f"Error: Input file not found at {input_file_path}")
    print("Please verify the file path and name.")
    exit()

# Load the Excel file
df = pd.read_excel(input_file_path, sheet_name="Sheet1")

# Clean column names
df.columns = df.columns.str.strip()

# City columns (seller) and corresponding quantity columns
city_columns = ['Mumbai', 'Delhi', 'Hyderabad', 'Bangalore', 'Chennai', 'Lucknow', 'kolkata']
qty_columns = {
    'Mumbai': 'MH',
    'Delhi': 'DE',
    'Hyderabad': 'HY',
    'Bangalore': 'BE',
    'Chennai': 'TN',
    'Lucknow': 'UP',
    'kolkata': 'KK'
}

# Filter rows where REG BuyBox is NOT 'Yes'
filtered_df = df[df['REG BuyBox'] != 'Yes'].copy()

# Handle NaN values safely
filtered_df.fillna('', inplace=True)

results = []

for _, row in filtered_df.iterrows():
    mp_sku = row['MP-sku']
    title = row['Title']
    asin = row['ASIN']
    rack_stock = row['Rack Stock']
    fba_stock = row['FBA Stock']
    
    for col in city_columns:
        seller = str(row[col]).strip()
        # Only include if it's a competitor (not Global Mart India, not '0', not empty)
        if seller and seller != '0' and seller != 'Global Mart India':
            qty = row[qty_columns[col]] if qty_columns[col] in row else ''
            city_name = 'Kolkata' if col == 'kolkata' else col.capitalize()
            results.append({
                'MP-sku': mp_sku,
                'Title': title,
                'ASIN': asin,
                'Rack Stock': rack_stock,
                'FBA Stock': fba_stock,
                'City': city_name,
                'Competitor Seller': seller,
                'Quantity': qty
            })

# Create result DataFrame
if results:
    result_df = pd.DataFrame(results)
    
    # Sort for better readability
    result_df.sort_values(['MP-sku', 'City'], inplace=True)
    
    # Print to console
    print("\nCompetitor BuyBox details (one row per state/seller):\n")
    print(result_df.to_string(index=False))
    
    # Save to Excel (.xlsx) - CORRECT METHOD
    result_df.to_excel(output_file, index=False, engine='openpyxl')
    
    print(f"\nExpanded report successfully saved as XLSX to:\n{output_file}")
else:
    print("No competitor BuyBox instances found where REG BuyBox != 'Yes'.")
    print("Nothing to save.")