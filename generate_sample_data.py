"""
Generate a sample multi-sheet Excel file for testing the dashboard app.
Run: python generate_sample_data.py
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random

random.seed(42)
np.random.seed(42)

months = pd.date_range(start='2023-01-01', periods=12, freq='MS')
months_str = [m.strftime('%b %Y') for m in months]

# ── Sheet 1: Sales by Region ───────────────────────────────────────────────────
regions = ['North', 'South', 'East', 'West', 'Central']
sales_data = []
for region in regions:
    for month in months_str:
        sales_data.append({
            'Region': region,
            'Month': month,
            'Revenue': round(random.uniform(50000, 200000), 2),
            'Units_Sold': random.randint(100, 800),
            'Target': round(random.uniform(80000, 180000), 2),
        })
df_sales = pd.DataFrame(sales_data)
df_sales['Achievement_%'] = round(df_sales['Revenue'] / df_sales['Target'] * 100, 1)

# ── Sheet 2: Product Performance ──────────────────────────────────────────────
products = ['Product A', 'Product B', 'Product C', 'Product D', 'Product E']
product_data = []
for product in products:
    product_data.append({
        'Product': product,
        'Category': random.choice(['Electronics', 'Apparel', 'Home', 'Sports']),
        'Q1_Revenue': round(random.uniform(30000, 150000), 2),
        'Q2_Revenue': round(random.uniform(30000, 150000), 2),
        'Q3_Revenue': round(random.uniform(30000, 150000), 2),
        'Q4_Revenue': round(random.uniform(30000, 150000), 2),
        'Units_Sold': random.randint(200, 2000),
        'Return_Rate_%': round(random.uniform(1, 8), 1),
        'Profit_Margin_%': round(random.uniform(15, 45), 1),
    })
df_products = pd.DataFrame(product_data)

# ── Sheet 3: Customer Segments ────────────────────────────────────────────────
segments = ['Enterprise', 'SMB', 'Startup', 'Consumer', 'Government']
customer_data = []
for segment in segments:
    for month in months_str[:6]:
        customer_data.append({
            'Segment': segment,
            'Month': month,
            'New_Customers': random.randint(10, 120),
            'Churned_Customers': random.randint(2, 30),
            'Avg_Order_Value': round(random.uniform(200, 5000), 2),
            'NPS_Score': round(random.uniform(20, 80), 1),
            'Support_Tickets': random.randint(5, 60),
        })
df_customers = pd.DataFrame(customer_data)

# ── Sheet 4: Marketing Spend ──────────────────────────────────────────────────
channels = ['Google Ads', 'Social Media', 'Email', 'Events', 'SEO', 'Referral']
marketing_data = []
for channel in channels:
    for month in months_str:
        spend = round(random.uniform(5000, 50000), 2)
        leads = random.randint(50, 500)
        marketing_data.append({
            'Channel': channel,
            'Month': month,
            'Spend_USD': spend,
            'Leads_Generated': leads,
            'Conversions': random.randint(5, int(leads * 0.3)),
            'Cost_Per_Lead': round(spend / leads, 2),
        })
df_marketing = pd.DataFrame(marketing_data)
df_marketing['Conversion_Rate_%'] = round(df_marketing['Conversions'] / df_marketing['Leads_Generated'] * 100, 1)

# ── Save ───────────────────────────────────────────────────────────────────────
output_path = 'sample_business_data.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_sales.to_excel(writer, sheet_name='Sales_by_Region', index=False)
    df_products.to_excel(writer, sheet_name='Product_Performance', index=False)
    df_customers.to_excel(writer, sheet_name='Customer_Segments', index=False)
    df_marketing.to_excel(writer, sheet_name='Marketing_Spend', index=False)

print(f"✅ Sample file saved: {output_path}")
print(f"   Sheets: Sales_by_Region ({len(df_sales)} rows), "
      f"Product_Performance ({len(df_products)} rows), "
      f"Customer_Segments ({len(df_customers)} rows), "
      f"Marketing_Spend ({len(df_marketing)} rows)")
