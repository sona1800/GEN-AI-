import pandas as pd

# Load Dummy_dataset.csv
df = pd.read_csv('Dummy_dataset.csv')

# Question 4: Identify and fill missing values with the column average
print("Question 4: Fill missing values with column average")
print("Missing values before:")
print(df.isnull().sum())

# Fill missing values with the mean of each column
df_filled = df.fillna(df.mean(numeric_only=True))
print("\nMissing values after:")
print(df_filled.isnull().sum())

# Question 5: Identify and remove duplicate rows
print("\n\nQuestion 5: Remove duplicate rows")
print(f"Number of rows before: {len(df_filled)}")
df_cleaned = df_filled.drop_duplicates()
print(f"Number of rows after: {len(df_cleaned)}")

# Question 8: Create a new calculated column 'Total Sales' (Quantity * Unit Price)
print("\n\nQuestion 8: Create 'Total Sales' column")
if 'Quantity' in df_cleaned.columns and 'Unit Price' in df_cleaned.columns:
    df_cleaned['Total Sales'] = df_cleaned['Quantity'] * df_cleaned['Unit Price']
    print("Total Sales column created")
    print(df_cleaned[['Quantity', 'Unit Price', 'Total Sales']].head())

# Question 7: Create a pivot table showing total sales by region and product
print("\n\nQuestion 7: Pivot table - Total sales by Region and Product")
if 'Region' in df_cleaned.columns and 'Product' in df_cleaned.columns and 'Total Sales' in df_cleaned.columns:
    pivot_table = df_cleaned.pivot_table(values='Total Sales', index='Region', columns='Product', aggfunc='sum')
    print(pivot_table)
else:
    print("Required columns not found for pivot table")
