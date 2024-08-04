import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
import matplotlib.pyplot as plt

# Load data
file_path = 'abc_manufacturing_data (1) (1).xlsx.xlsx'
sale_data = pd.read_excel(file_path, sheet_name='Sale')

# Convert SaleDate to datetime and add a 'Month' column
sale_data['SaleDate'] = pd.to_datetime(sale_data['SaleDate'])
sale_data['Month'] = sale_data['SaleDate'].dt.to_period('M').astype(str)

# Aggregate sales by month
monthly_sales = sale_data.groupby('Month')['TotalPrice'].sum().reset_index()
monthly_sales['Month'] = pd.to_datetime(monthly_sales['Month'])

# Create input (X) and target (y) variables
X = monthly_sales[['Month']].apply(lambda x: x.map(pd.Timestamp.toordinal))
y = monthly_sales['TotalPrice']

# Split the data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Train the linear regression model
model = LinearRegression()
model.fit(X_train, y_train)

# Make predictions and evaluate the model
y_pred = model.predict(X_test)

# Evaluate the model
mse = mean_squared_error(y_test, y_pred)
r2 = r2_score(y_test, y_pred)

print(f'Mean Squared Error: {mse}')
print(f'R^2 Score: {r2}')

# Predict future sales
# Create data for future months
future_months = pd.date_range(start=monthly_sales['Month'].max(), periods=12, freq='ME').to_period('M').astype(str)
future_X = pd.DataFrame(future_months, columns=['Month'])
future_X['Month'] = pd.to_datetime(future_X['Month']).map(pd.Timestamp.toordinal)

# Predict future sales
future_sales_pred = model.predict(future_X)

# Display the results
future_predictions = pd.DataFrame({'Month': future_months, 'PredictedSales': future_sales_pred})
print(future_predictions)

# Plot the sales predictions
plt.figure(figsize=(10, 6))
plt.plot(monthly_sales['Month'], monthly_sales['TotalPrice'], label='Actual Sales')
plt.plot(pd.to_datetime(future_predictions['Month']), future_predictions['PredictedSales'], label='Predicted Sales', linestyle='--')
plt.xlabel('Month')
plt.ylabel('Total Sales')
plt.title('Sales Prediction for the Next 12 Months')
plt.legend()
plt.grid(True)
plt.show()