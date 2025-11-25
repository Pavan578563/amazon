# ==============================
# Amazon Sales Performance Analysis Report
# Author: Pavan Kalyan (Internship Project)
# Design: Amazon-style (Orange, Black, Gray)
# ==============================

# --- 1. Import Libraries ---
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from datetime import datetime
import matplotlib.ticker as mticker
import os

# --- 2. Load Data ---
df = pd.read_csv("Amazon Sale Report.csv", encoding='unicode_escape')
df.columns = df.columns.str.strip()  # remove spaces
df.columns = df.columns.str.title()  # normalize case

# --- 3. Identify Key Columns Automatically ---
amount_col = next((c for c in df.columns if 'amount' in c.lower()), None)
date_col = next((c for c in df.columns if 'date' in c.lower()), None)
category_col = next((c for c in df.columns if 'category' in c.lower()), None)
fulfil_col = next((c for c in df.columns if 'fulfil' in c.lower()), None)
state_col = next((c for c in df.columns if 'state' in c.lower()), None)

# --- 4. Clean Data ---
df = df.dropna(subset=[amount_col])
df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')
df = df.dropna(subset=[date_col])
df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
df = df[df[amount_col] > 0]

# --- 5. Key Metrics ---
total_orders = len(df)
total_revenue = df[amount_col].sum()
avg_order_value = df[amount_col].mean()
date_min, date_max = df[date_col].min(), df[date_col].max()

# --- 6. Amazon Color Theme ---
sns.set_style("whitegrid")
amazon_orange = "#FF9900"
amazon_gray = "#232F3E"
amazon_black = "#0F1111"

def save_chart(fig, name):
    path = f"{name}.png"
    fig.savefig(path, bbox_inches='tight', dpi=300)
    plt.close(fig)
    return path

# --- 7. Charts ---

# Monthly Revenue Trend
df['Month'] = df[date_col].dt.to_period('M')
monthly_revenue = df.groupby('Month')[amount_col].sum()
fig, ax = plt.subplots(figsize=(7,4))
monthly_revenue.plot(kind='bar', color=amazon_orange, ax=ax)
ax.set_title("Monthly Revenue Trend", fontsize=12, color=amazon_black, weight='bold')
ax.set_xlabel("Month")
ax.set_ylabel("Revenue (₹)")
ax.yaxis.set_major_formatter(mticker.StrMethodFormatter('{x:,.0f}'))
plt.xticks(rotation=45)
monthly_chart = save_chart(fig, "monthly_revenue")

# Top 10 Product Categories
if category_col:
    top_cats = df.groupby(category_col)[amount_col].sum().sort_values(ascending=False).head(10)
    fig, ax = plt.subplots(figsize=(6,4))
    sns.barplot(x=top_cats.values, y=top_cats.index, ax=ax, palette=[amazon_orange]*10)
    ax.set_title("Top 10 Product Categories", fontsize=12, color=amazon_black, weight='bold')
    ax.set_xlabel("Revenue (₹)")
    topcat_chart = save_chart(fig, "top_categories")
else:
    topcat_chart = None

# Fulfillment Method Analysis
if fulfil_col:
    fulfill = df.groupby(fulfil_col)[amount_col].sum()
    fig, ax = plt.subplots(figsize=(5,5))
    ax.pie(fulfill, labels=fulfill.index, autopct='%1.1f%%', colors=[amazon_orange, amazon_gray, amazon_black])
    ax.set_title("Fulfillment Method Distribution", fontsize=12, weight='bold')
    fulfill_chart = save_chart(fig, "fulfillment")
else:
    fulfill_chart = None

# Top States by Revenue
if state_col:
    top_states = df.groupby(state_col)[amount_col].sum().sort_values(ascending=False).head(10)
    fig, ax = plt.subplots(figsize=(6,4))
    sns.barplot(x=top_states.values, y=top_states.index, ax=ax, palette=[amazon_orange]*10)
    ax.set_title("Top 10 States by Revenue", fontsize=12, color=amazon_black, weight='bold')
    ax.set_xlabel("Revenue (₹)")
    topstates_chart = save_chart(fig, "top_states")
else:
    topstates_chart = None

# --- 8. Generate Professional PDF Report ---
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='TitleStyle', fontSize=20, textColor=amazon_orange, spaceAfter=20, alignment=1, leading=24))
styles.add(ParagraphStyle(name='Heading', fontSize=14, textColor=amazon_black, spaceAfter=10, leading=18))
styles.add(ParagraphStyle(name='NormalText', fontSize=11, leading=16))

pdf = SimpleDocTemplate("Amazon_Sales_Performance_Report.pdf", pagesize=A4)
story = []

# Cover Page
story.append(Paragraph("Amazon Sales Performance Analysis Report", styles['TitleStyle']))
story.append(Spacer(1, 12))
story.append(Paragraph(f"Prepared by <b>Pavan Kalyan</b>", styles['NormalText']))
story.append(Paragraph("Data Analytics Internship Project", styles['NormalText']))
story.append(Paragraph(f"Date: {datetime.now().strftime('%B %d, %Y')}", styles['NormalText']))
story.append(Spacer(1, 20))

# Deliverables Section
deliverables = """
<b>Deliverables:</b><br/>
1. Comprehensive analysis report summarizing key findings, insights, and recommendations.<br/>
2. Visualizations (charts, graphs) illustrating various aspects of the data analysis.<br/>
3. Insights on product preferences, customer behaviour, and geographical sales distribution.<br/>
4. Recommendations for improving sales strategies, inventory management, and customer service.<br/><br/>

<b>Expected Outcome:</b><br/>
By conducting a thorough analysis of the Amazon sales report, the goal is to gain valuable insights that can be leveraged to optimize business operations, enhance customer experience, and drive revenue growth. The analysis provides actionable recommendations tailored to the specific needs and challenges of the business.
"""
story.append(Paragraph(deliverables, styles['NormalText']))
story.append(Spacer(1, 16))

# Executive Summary
summary = f"""
<b>Executive Summary:</b><br/><br/>
This report analyzes Amazon's sales data from <b>{date_min.strftime('%d %b %Y')}</b> to <b>{date_max.strftime('%d %b %Y')}</b>. 
It explores trends in revenue, product categories, fulfillment efficiency, and geographical performance. 
Insights derived from this analysis aim to help optimize sales strategies, improve customer experience, and 
drive sustained revenue growth.
"""
story.append(Paragraph(summary, styles['NormalText']))
story.append(Spacer(1, 16))

# Key Metrics Table
metrics_data = [
    ["Total Orders", f"{total_orders:,.0f}"],
    ["Total Revenue (₹)", f"{total_revenue:,.0f}"],
    ["Average Order Value (₹)", f"{avg_order_value:,.2f}"],
    ["Date Range", f"{date_min.strftime('%d-%b-%Y')} to {date_max.strftime('%d-%b-%Y')}"]
]
table = Table(metrics_data, hAlign='LEFT', colWidths=[200, 200])
table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), colors.HexColor(amazon_gray)),
    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
    ('ALIGN', (0,0), (-1,-1), 'LEFT'),
    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
    ('BACKGROUND', (0,1), (-1,-1), colors.whitesmoke),
    ('BOX', (0,0), (-1,-1), 0.25, colors.gray),
    ('INNERGRID', (0,0), (-1,-1), 0.25, colors.gray),
]))
story.append(table)
story.append(Spacer(1, 20))

# Add Charts
def add_chart_to_pdf(image_path, caption):
    if image_path:
        story.append(Image(image_path, width=450, height=250))
        story.append(Spacer(1, 6))
        story.append(Paragraph(caption, styles['NormalText']))
        story.append(Spacer(1, 16))

story.append(Paragraph("<b>Visual Analysis</b>", styles['Heading']))
add_chart_to_pdf(monthly_chart, "Monthly revenue shows seasonal performance patterns.")
add_chart_to_pdf(topcat_chart, "Top-performing product categories contributing to total revenue.")
add_chart_to_pdf(fulfill_chart, "Fulfillment method distribution and its impact on delivery efficiency.")
add_chart_to_pdf(topstates_chart, "Top 10 states driving majority of sales revenue.")

# Insights & Recommendations
insights = """
<b>Key Insights:</b><br/>
- Sales show strong seasonal peaks, aligning with major promotional events.<br/>
- Few categories dominate sales, suggesting focused marketing yields high returns.<br/>
- Fulfillment methods like FBA (Fulfilled by Amazon) improve delivery performance.<br/>
- Certain states exhibit strong demand concentration.<br/><br/>

<b>Recommendations:</b><br/>
1. Focus marketing budgets on top regions and categories.<br/>
2. Expand inventory for high-performing products.<br/>
3. Streamline logistics in low-performing states to reduce costs.<br/>
4. Use predictive analytics to anticipate demand during festival seasons.<br/>
"""
story.append(Paragraph(insights, styles['NormalText']))
story.append(Spacer(1, 20))

# Conclusion
conclusion = """
<b>Conclusion:</b><br/><br/>
This project successfully analyzed key patterns in Amazon’s sales data, revealing significant insights into 
revenue trends, customer preferences, and regional performance. The findings support data-driven 
strategic decisions aimed at improving efficiency and profitability.<br/><br/>

By implementing the proposed recommendations, Amazon can strengthen its competitive advantage, 
enhance customer satisfaction, and ensure sustainable business growth. This internship project also 
demonstrates the practical application of data analytics to solve real-world business challenges.
"""
story.append(Paragraph(conclusion, styles['NormalText']))

# Build PDF
pdf.build(story)

print("✅ Professional PDF report generated successfully: Amazon_Sales_Performance_Report.pdf")
