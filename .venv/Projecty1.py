import pandas as pd

# Define the file path
file_path = 'abc_manufacturing_data (1) (1).xlsx.xlsx'

# Load data from the provided Excel file
customer_data = pd.read_excel(file_path, sheet_name='Customer')
product_group_data = pd.read_excel(file_path, sheet_name='ProductGroup')
product_detail_data = pd.read_excel(file_path, sheet_name='ProductDetail')
sale_data = pd.read_excel(file_path, sheet_name='Sale')

# Display the first few rows of each dataframe to understand the structure and content
print("Customer Data:")
print(customer_data.head())

print("\nProduct Group Data:")
print(product_group_data.head())

print("\nProduct Detail Data:")
print(product_detail_data.head())

print("\nSale Data:")
print(sale_data.head())

# Handling Missing Values
print("\nMissing Values Before Cleaning:")
print(customer_data.isnull().sum())
print(product_group_data.isnull().sum())
print(product_detail_data.isnull().sum())
print(sale_data.isnull().sum())

# Drop rows with missing values
customer_data = customer_data.dropna()
product_group_data = product_group_data.dropna()
product_detail_data = product_detail_data.dropna()

# Fill missing values in 'Quantity' with mean
sale_data['Quantity'] = sale_data['Quantity'].fillna(sale_data['Quantity'].mean())

# Correcting Data Types
sale_data['SaleDate'] = pd.to_datetime(sale_data['SaleDate'])
customer_data['CustomerID'] = customer_data['CustomerID'].astype(int)
product_detail_data['ProductID'] = product_detail_data['ProductID'].astype(int)
sale_data['CustomerID'] = sale_data['CustomerID'].astype(int)
sale_data['ProductID'] = sale_data['ProductID'].astype(int)

# Removing Duplicates
customer_data = customer_data.drop_duplicates()
product_group_data = product_group_data.drop_duplicates()
product_detail_data = product_detail_data.drop_duplicates()
sale_data = sale_data.drop_duplicates()

# Standardizing Data Formats
customer_data['FirstName'] = customer_data['FirstName'].str.strip().str.lower()
customer_data['LastName'] = customer_data['LastName'].str.strip().str.lower()

# Merging DataFrames
combined_data = pd.merge(sale_data, product_detail_data, on='ProductID')
combined_data = pd.merge(combined_data, customer_data, on='CustomerID')

# Display the combined data
print("\nCombined Data:")
print(combined_data.head())

# Saving Cleaned Data
customer_data.to_csv('cleaned_customer_data.csv', index=False)
product_group_data.to_csv('cleaned_product_group_data.csv', index=False)
product_detail_data.to_csv('cleaned_product_detail_data.csv', index=False)
sale_data.to_csv('cleaned_sale_data.csv', index=False)
combined_data.to_csv('combined_data.csv', index=False)

print("\nData cleaning and preprocessing completed successfully.")