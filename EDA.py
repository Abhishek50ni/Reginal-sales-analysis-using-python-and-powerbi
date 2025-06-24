""" Problem Statement:  
"""
import pandas as pd
import numpy as np
import seaborn as sb
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter

sheets=pd.read_excel('Regional Sales Dataset.xlsx',sheet_name=None)
#print(sheets)

# Extracting the sheet names and  dataframes
df_sales=sheets['Sales Orders']
df_customers=sheets['Customers']
df_products=sheets['Products']
df_regions=sheets['Regions']   
df_states_regions=sheets['State Regions']
df_budgets=sheets['2017 Budgets']

#print(df_states_regions.head(5))
#print(df_sales.shape)

#checking for null values
""" print(df_sales.isnull().sum())
print(df_customers.isnull().sum())
print(df_products.isnull().sum())   
print(df_regions.isnull().sum())
print(df_states_regions.isnull().sum()) """ 

# Merging sales with customers 
df_merge_scp = df_sales.merge(
    df_customers,
    how='left',
    left_on='Customer Name Index',
    right_on='Customer Index'
)
#print(df_merge_scp.head(5))

#merging sales with products
df_merge_scp = df_merge_scp.merge(
    df_products,
    how='left',
    left_on='Product Description Index',
    right_on='Index'
)
#print(df_merge_scp.head(5))

#merging sales with regions
df_merge_scp = df_merge_scp.merge(
    df_regions,
    how='left',
    left_on='Delivery Region Index',
    right_on='id' 
)
#print(df_merge_scp.head(5))

#merging sales with states_regions
df_merge_scp= df_merge_scp.merge(
    df_states_regions,
    how='left',
    left_on='state_code',
    right_on='State Code'  
)
#print(df_merge_scp.head(5))

#merging with budget
df_merge_scp= df_merge_scp.merge(
    df_budgets,
    how='left',
    left_on='Product Name',
    right_on='Product Name'
)

# dropping the repeated columns after viewing csv
cols_to_drop=['Customer Index','Index','id','State Code','State']
df_merge_scp=df_merge_scp.drop(columns=cols_to_drop, errors='ignore')	
#print(df_merge_scp.head(5))

#converting to lower case
df_merge_scp.columns=df_merge_scp.columns.str.lower()
#checking the column values
#print(df_merge_scp.columns.tolist())
df_merge_scp= df_merge_scp.rename(columns={
    '2017 budgets':'budget'
})
#print(df_merge_scp.columns.tolist())


#removing budget data from non 2017 years
# Convert 'OrderDate' to datetime (if not already)
df_merge_scp['orderdate'] = pd.to_datetime(df_merge_scp['orderdate'])
# Empty 'budget' for non-2017 rows
df_merge_scp.loc[df_merge_scp['orderdate'].dt.year != 2017, 'budget'] = None  # or np.nan

df_merge_scp['total_cost']= df_merge_scp['order quantity'] * df_merge_scp['total unit cost']
df_merge_scp['profit']= df_merge_scp['line total'] - df_merge_scp['total_cost']
df_merge_scp['profit_margin']= df_merge_scp['profit'] *100 / df_merge_scp['line total']

#exporting dataframe to csv to view clearly
df_merge_scp.to_csv('merged_data.csv', index=False)




""" Creating a line chart to visualize monthly sales performance
This chart will show total sales, profit, and order volume over time."""

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter

# 1. Prepare the data
# Convert OrderDate to datetime and filter out Jan/Feb 2018
df_merge_scp['orderdate'] = pd.to_datetime(df_merge_scp['orderdate'])
df_filtered = df_merge_scp[~((df_merge_scp['orderdate'].dt.year == 2018)) & 
                          (df_merge_scp['orderdate'].dt.month.isin([1, 2]))]

# Extract month/year period
df_filtered['Month'] = df_filtered['orderdate'].dt.to_period('M')

# Aggregate monthly metrics
monthly_data = df_filtered.groupby('Month').agg({
    'line total': 'sum',    # Total sales
    'profit': 'sum',        # Total profit
    'ordernumber': 'count'  # Order volume
}).reset_index()

# Convert Month to datetime for plotting
monthly_data['Month'] = monthly_data['Month'].dt.to_timestamp()

# 2. Create the line chart
plt.figure(figsize=(12, 6))

# Plot Total Sales (blue line)
plt.plot(monthly_data['Month'], monthly_data['line total'], 
         label='Total Sales ($)', marker='o', color='#1f77b4', linewidth=2)

# Plot Total Profit (green line)
plt.plot(monthly_data['Month'], monthly_data['profit'], 
         label='Profit ($)', marker='s', color='#2ca02c', linewidth=2)

# Format x-axis as short dates (e.g., "Jan-2023")
date_format = DateFormatter("%b-%Y")  # Short month-year format
plt.gca().xaxis.set_major_formatter(date_format)
plt.gcf().autofmt_xdate()  # Auto-rotate dates

# Customize the chart
plt.title('Monthly Sales Performance (Excluding Jan/Feb 2018)', fontsize=16, pad=20)
plt.xlabel('Month', fontsize=12, labelpad=10)
plt.ylabel('Amount ($)', fontsize=12, labelpad=10)
plt.legend(fontsize=12, loc='upper left')
plt.grid(True, linestyle='--', alpha=0.3)

# Optional: Add order volume as bars (secondary y-axis)
ax2 = plt.gca().twinx()
ax2.bar(monthly_data['Month'], monthly_data['ordernumber'], 
        color='orange', alpha=0.3, width=20, label='Orders')
ax2.set_ylabel('Order Volume', fontsize=12)
ax2.legend(loc='upper right')

# Adjust layout and display
plt.tight_layout()
plt.show()




"""Monthly Sales Trend (All Years Combined) - Excluding Jan/Feb 2018"""

import pandas as pd
import matplotlib.pyplot as plt

# 1. Prepare the data
# Convert orderdate to datetime and filter out Jan/Feb 2018
df_merge_scp['orderdate'] = pd.to_datetime(df_merge_scp['orderdate'])
df_filtered = df_merge_scp[~((df_merge_scp['orderdate'].dt.year == 2018) & 
                           (df_merge_scp['orderdate'].dt.month.isin([1, 2])))]

# Extract month name from filtered data
df_filtered['Month'] = df_filtered['orderdate'].dt.month_name()

# Aggregate sales by month (combining all years)
monthly_sales = df_filtered.groupby('Month').agg({
    'line total': 'sum'  # Total sales across all years
}).reset_index()

# Order months chronologically (not alphabetically)
month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
               'July', 'August', 'September', 'October', 'November', 'December']
monthly_sales['Month'] = pd.Categorical(monthly_sales['Month'], categories=month_order, ordered=True)
monthly_sales = monthly_sales.sort_values('Month')

# 2. Create the line chart
plt.figure(figsize=(12, 6))

# Plot sales by month
plt.plot(monthly_sales['Month'], monthly_sales['line total'], 
         marker='o', color='#1f77b4', linewidth=2)

# Customize the chart
plt.title('Monthly Sales Trend (Excluding Jan/Feb 2018)', fontsize=16)
plt.xlabel('Month', fontsize=12)
plt.ylabel('Total Sales ($)', fontsize=12)
plt.xticks(rotation=45)  # Rotate month names for readability
plt.grid(True, linestyle='--', alpha=0.3)

# Display the chart
plt.tight_layout()
plt.show()




""" top 10 products by sales revenue """

import pandas as pd
import matplotlib.pyplot as plt

# Assuming df_merge_scp contains your sales data with 'product name' and 'line total' columns
# 1. Prepare the data - aggregate revenue by product
product_revenue = df_merge_scp.groupby('product name')['line total'].sum().reset_index()

# 2. Sort and get top 10 products
top_products = product_revenue.sort_values('line total', ascending=False).head(10)

# 3. Create the bar chart
plt.figure(figsize=(12, 6))
bars = plt.barh(top_products['product name'], top_products['line total'], 
                color='skyblue', edgecolor='navy')

# 4. Customize the chart
plt.title('Top 10 Products by Revenue', fontsize=16, pad=20)
plt.xlabel('Total Revenue ($)', fontsize=12)
plt.ylabel('Product Name', fontsize=12)
plt.gca().invert_yaxis()  # Show highest revenue at top
plt.grid(axis='x', linestyle='--', alpha=0.7)

# 5. Add value labels on bars
for bar in bars:
    width = bar.get_width()
    plt.text(width, bar.get_y() + bar.get_height()/2, 
             f'${width:,.0f}', 
             ha='left', va='center', fontsize=10)

plt.tight_layout()
plt.show()


""" bottom 10 products by sales revenue """
import pandas as pd
import matplotlib.pyplot as plt

# 1. Prepare the data - aggregate revenue by product
product_revenue = df_merge_scp.groupby('product name')['line total'].sum().reset_index()

# 2. Get bottom 10 products
bottom_10 = product_revenue.sort_values('line total').head(10)

# 3. Create the bar chart
plt.figure(figsize=(12, 6))
bars = plt.barh(bottom_10['product name'], bottom_10['line total'],
                color='#ff9999', edgecolor='#ff4444', height=0.8)

# 4. Customize the chart
plt.title('Bottom 10 Products by Revenue', fontsize=16, pad=20)
plt.xlabel('Total Revenue ($)', fontsize=12)
plt.ylabel('Product Name', fontsize=12)
plt.gca().invert_yaxis()  # Show lowest revenue at top
plt.grid(axis='x', linestyle='--', alpha=0.5)

# 5. Add value labels on bars
for bar in bars:
    width = bar.get_width()
    plt.text(width, bar.get_y() + bar.get_height()/2, 
             f'${width:,.0f}', 
             ha='left', va='center', fontsize=10)

# 6. Add analysis annotation
plt.annotate('These products need performance review\nor discontinuation consideration',
             xy=(0.5, 0.1), xycoords='axes fraction',
             ha='center', fontsize=12, bbox=dict(boxstyle='round', fc='white'))

plt.tight_layout()
plt.show()


""" sales by channel"""

import pandas as pd
import matplotlib.pyplot as plt

# 1. Aggregate sales by channel
channel_data = df_merge_scp.groupby('channel')['line total'].sum().sort_values(ascending=False)

# 2. Create the pie chart
plt.figure(figsize=(10, 8))

# Custom color palette
colors = plt.cm.Paired.colors[:len(channel_data)]

# Plot pie chart with percentage labels
patches, texts, autotexts = plt.pie(
    channel_data,
    labels=channel_data.index,
    colors=colors,
    autopct='%1.1f%%',
    startangle=90,
    pctdistance=0.8,
    wedgeprops={'edgecolor': 'white', 'linewidth': 1}
)

# 3. Customize the chart
plt.title('Sales Distribution by Channel', fontsize=16, pad=20)
plt.axis('equal')  # Perfect circle

# Improve label readability
for text in texts + autotexts:
    text.set_fontsize(10)
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_weight('bold')

# 4. Add legend with sales values
legend_labels = [f"{label} (${value:,.0f})" 
                for label, value in zip(channel_data.index, channel_data)]
plt.legend(
    patches,
    legend_labels,
    title="Channel Sales",
    loc="center left",
    bbox_to_anchor=(1, 0.5),
    fontsize=10
)

plt.tight_layout()
plt.show()



""" State Revenue Performance Analysis - Top and Bottom 10 States """

# 1. Prepare the data - aggregate revenue by state
state_revenue = df_merge_scp.groupby('state')['line total'].sum().reset_index()

# Get top and bottom 10 states
top_10_states = state_revenue.sort_values('line total', ascending=False).head(10)
bottom_10_states = state_revenue.sort_values('line total').head(10)

# Create figure with subplots
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
fig.suptitle('State Revenue Performance Analysis', fontsize=16, y=1.02)

# 2. Top 10 States Chart
bars_top = ax1.barh(top_10_states['state'], top_10_states['line total'],
                   color='#2ecc71', edgecolor='#27ae60')
ax1.set_title('Top 10 States by Revenue', fontsize=14, pad=15)
ax1.set_xlabel('Total Revenue ($)', fontsize=12)
ax1.grid(axis='x', linestyle='--', alpha=0.5)
ax1.invert_yaxis()  # Highest revenue at top

# Add value labels for top states
for bar in bars_top:
    width = bar.get_width()
    ax1.text(width, bar.get_y() + bar.get_height()/2,
            f'${width:,.0f}',
            ha='left', va='center', fontsize=10)

# 3. Bottom 10 States Chart
bars_bottom = ax2.barh(bottom_10_states['state'], bottom_10_states['line total'],
                      color='#e74c3c', edgecolor='#c0392b')
ax2.set_title('Bottom 10 States by Revenue', fontsize=14, pad=15)
ax2.set_xlabel('Total Revenue ($)', fontsize=12)
ax2.grid(axis='x', linestyle='--', alpha=0.5)
ax2.invert_yaxis()  # Lowest revenue at top

# Add value labels for bottom states
for bar in bars_bottom:
    width = bar.get_width()
    ax2.text(width, bar.get_y() + bar.get_height()/2,
            f'${width:,.0f}',
            ha='left', va='center', fontsize=10)

# 4. Add analysis annotations
ax1.annotate('High Performing Regions',
            xy=(0.5, 1.05), xycoords='axes fraction',
            ha='center', fontsize=12, color='#27ae60')

ax2.annotate('Areas Needing Improvement',
            xy=(0.5, 1.05), xycoords='axes fraction',
            ha='center', fontsize=12, color='#c0392b')

plt.tight_layout()
plt.show()


"""  top and bottom 10 customers by revenue """


import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap

# Load your data (replace with your actual dataframe)
# df = pd.read_csv('your_sales_data.csv')

# 1. Prepare the data
customer_revenue = df_merge_scp.groupby('customer names')['line total'].sum().reset_index()

# Get top and bottom 10 customers
top_10 = customer_revenue.nlargest(10, 'line total')
bottom_10 = customer_revenue.nsmallest(10, 'line total')

# Create gradient color maps
top_colors = LinearSegmentedColormap.from_list('top', ['#4CAF50', '#2E7D32'])  # Green gradient
bottom_colors = LinearSegmentedColormap.from_list('bottom', ['#FF9800', '#F44336'])  # Orange-Red gradient

# 2. Create the visualization
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
fig.suptitle('Customer Revenue Analysis', fontsize=16, y=1.05)

# Top 10 Customers Plot
top_colors = top_colors(np.linspace(0, 1, len(top_10)))
bars_top = ax1.barh(top_10['customer names'], top_10['line total'], color=top_colors)
ax1.set_title('Top 10 Customers by Revenue', fontsize=14, pad=15)
ax1.set_xlabel('Total Revenue ($)', fontsize=12)
ax1.grid(axis='x', linestyle='--', alpha=0.3)

# Bottom 10 Customers Plot
bottom_colors = bottom_colors(np.linspace(0, 1, len(bottom_10)))
bars_bottom = ax2.barh(bottom_10['customer names'], bottom_10['line total'], color=bottom_colors)
ax2.set_title('Bottom 10 Customers by Revenue', fontsize=14, pad=15)
ax2.set_xlabel('Total Revenue ($)', fontsize=12)
ax2.grid(axis='x', linestyle='--', alpha=0.3)

# Add value labels
def add_value_labels(ax, bars):
    for bar in bars:
        width = bar.get_width()
        ax.text(width, bar.get_y() + bar.get_height()/2,
                f'${width:,.0f}',
                ha='left', va='center', fontsize=10)

add_value_labels(ax1, bars_top)
add_value_labels(ax2, bars_bottom)

# Invert y-axis for better readability
ax1.invert_yaxis()
ax2.invert_yaxis()

# Add helpful annotations
fig.text(0.5, 0.01, 
         'Analysis Tip: Focus on retaining top customers and improving engagement with bottom customers',
         ha='center', fontsize=12, style='italic')

plt.tight_layout()
plt.show()