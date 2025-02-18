import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt

# Sample data
data = {
    'Date': pd.date_range(start='1/1/2023', periods=10, freq='D'),
    'Sales': [100, 150, 200, 250, 180, 300, 350, 400, 450, 500]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Generate summary statistics
total_sales = df['Sales'].sum()
average_sales = df['Sales'].mean()

# Plot the data
plt.figure(figsize=(10, 5))
plt.plot(df['Date'], df['Sales'], marker='o')
plt.title('Sales Over Time')
plt.xlabel('Date')
plt.ylabel('Sales')
plt.grid(True)
plt.savefig('sales_chart.png')

# Create a report in Excel
with pd.ExcelWriter('sales_report.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sales Data', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sales Data']
    
    # Add a chart
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'categories': ['Sales Data', 1, 0, len(df), 0],
        'values':     ['Sales Data', 1, 1, len(df), 1],
        'line':       {'color': 'blue'},
    })
    chart.set_title({'name': 'Sales Over Time'})
    chart.set_x_axis({'name': 'Date'})
    chart.set_y_axis({'name': 'Sales'})
    worksheet.insert_chart('D2', chart)

    # Add summary statistics
    worksheet.write('G2', 'Total Sales')
    worksheet.write('G3', total_sales)
    worksheet.write('G5', 'Average Sales')
    worksheet.write('G6', average_sales)

print("Report generated and saved as 'sales_report.xlsx'")
