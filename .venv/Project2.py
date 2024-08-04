import pandas as pd
import matplotlib.pyplot as plt
# Define the file path
file_path = 'abc_manufacturing_data (1) (1).xlsx.xlsx'

# Load data from the provided Excel file
customer_data = pd.read_excel(file_path, sheet_name='Customer')
product_group_data = pd.read_excel(file_path, sheet_name='ProductGroup')
product_detail_data = pd.read_excel(file_path, sheet_name='ProductDetail')
sale_data = pd.read_excel(file_path, sheet_name='Sale')

# Merge data to create a comprehensive dataset
combined_data = pd.merge(sale_data, product_detail_data, on='ProductID')
combined_data = pd.merge(combined_data, customer_data, on='CustomerID')

# Sales Over Time
plt.figure(figsize=(10, 6))
combined_data['SaleDate'] = pd.to_datetime(combined_data['SaleDate'])
sales_over_time = combined_data.groupby(combined_data['SaleDate'].dt.to_period('M'))['TotalPrice'].sum()
sales_over_time.plot(kind='line', marker='o', linestyle='-', color='b')
plt.title('Total Sales Over Time')
plt.xlabel('Date')
plt.ylabel('Total Sales')
plt.grid(True)
plt.show()

# Merge combined_data with product_group_data to get GroupName
combined_data_with_groups = pd.merge(combined_data, product_group_data, left_on='ProductGroupID', right_on='ProductGroupID')

# Sales by Product Group
plt.figure(figsize=(12, 8))
sales_by_group = combined_data_with_groups.groupby('GroupName')['TotalPrice'].sum().sort_values()
sales_by_group.plot(kind='barh', color='skyblue')
plt.title('Total Sales by Product Group')
plt.xlabel('Total Sales')
plt.ylabel('Product Group')
plt.grid(axis='x')
plt.show()

# Top 10 Products by Sales
plt.figure(figsize=(12, 8))
top_products = combined_data.groupby('ProductName')['TotalPrice'].sum().nlargest(10)
top_products.plot(kind='bar', color='orange')
plt.title('Top 10 Products by Sales')
plt.xlabel('Product Name')
plt.ylabel('Total Sales')
plt.xticks(rotation=45)
plt.grid(axis='y')
plt.show()

# Sales Distribution by Country
plt.figure(figsize=(14, 10))
sales_by_country = combined_data.groupby('Country')['TotalPrice'].sum().sort_values(ascending=False).head(10)
sales_by_country.plot(kind='pie', autopct='%1.1f%%', startangle=140, colors=plt.cm.Paired.colors)
plt.title('Sales Distribution by Country (Top 10)')
plt.ylabel('')
plt.show()