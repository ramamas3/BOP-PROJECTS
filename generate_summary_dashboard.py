import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json

# File configuration
file_path = r'C:\Users\RAMAMANIS\OneDrive - IBM\Desktop\python\BOB\Stocking Strategy Tracking.xlsx'
sheet_name = "Stocking stratergy by PN -New"

# Load data
print("Loading data...")
df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

# Clean column names
df.columns = df.columns.str.strip()

# Required columns
value_col = 'TOTAL INVESTMENT AS Per GPP($M)'
df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0)

# Generate summary statistics
print("Generating summary statistics...")
total_records = len(df)
total_investment = df[value_col].sum()
avg_investment = df[value_col].mean()
max_investment = df[value_col].max()

# Create aggregations for charts (using correct column names)
owner_data = df.groupby('OWNER')[value_col].sum().sort_values(ascending=False).reset_index()
install_data = df.groupby('Install Check')[value_col].sum().sort_values(ascending=False).reset_index()
usage_data = df.groupby('Usage Excl transfer')[value_col].sum().sort_values(ascending=False).reset_index()
dtbh_data = df.groupby('DTBH Check')[value_col].sum().sort_values(ascending=False).reset_index()
reserve_data = df.groupby('Reserve check -  3 QTR')[value_col].sum().sort_values(ascending=False).reset_index()

# Get top 10 records by investment
top_10_records = df.nlargest(10, value_col)[value_col].reset_index(drop=True)

# Create subplots
print("Creating visualizations...")
fig = make_subplots(
    rows=3, cols=2,
    subplot_titles=(
        'Investment by Owner',
        'Investment by Install Check',
        'Investment by Usage Excl Transfer',
        'Investment by DTBH Check',
        'Investment by Reserve Check - 3 QTR',
        'Top 10 Records by Investment'
    ),
    specs=[
        [{'type': 'pie'}, {'type': 'bar'}],
        [{'type': 'bar'}, {'type': 'bar'}],
        [{'type': 'bar'}, {'type': 'bar'}]
    ],
    vertical_spacing=0.12,
    horizontal_spacing=0.15
)

# 1. Owner Pie Chart
fig.add_trace(
    go.Pie(
        labels=owner_data['OWNER'],
        values=owner_data[value_col],
        name='Owner',
        textposition='inside',
        textinfo='label+percent'
    ),
    row=1, col=1
)

# 2. Install Check Bar Chart
fig.add_trace(
    go.Bar(
        x=install_data['Install Check'],
        y=install_data[value_col],
        name='Install Check',
        text=[f"${val:.2f}M" for val in install_data[value_col]],
        textposition='auto'
    ),
    row=1, col=2
)

# 3. Usage Bar Chart
fig.add_trace(
    go.Bar(
        x=usage_data['Usage Excl transfer'],
        y=usage_data[value_col],
        name='Usage',
        text=[f"${val:.2f}M" for val in usage_data[value_col]],
        textposition='auto'
    ),
    row=2, col=1
)

# 4. DTBH Bar Chart
fig.add_trace(
    go.Bar(
        x=dtbh_data['DTBH Check'],
        y=dtbh_data[value_col],
        name='DTBH',
        text=[f"${val:.2f}M" for val in dtbh_data[value_col]],
        textposition='auto'
    ),
    row=2, col=2
)

# 5. Reserve Check Bar Chart
fig.add_trace(
    go.Bar(
        x=reserve_data['Reserve check -  3 QTR'],
        y=reserve_data[value_col],
        name='Reserve',
        text=[f"${val:.2f}M" for val in reserve_data[value_col]],
        textposition='auto'
    ),
    row=3, col=1
)

# 6. Top 10 Records Bar Chart
fig.add_trace(
    go.Bar(
        x=list(range(len(top_10_records))),
        y=top_10_records,
        name='Top 10',
        text=[f"${val:.2f}M" for val in top_10_records],
        textposition='auto'
    ),
    row=3, col=2
)

# Update layout
fig.update_layout(
    title={
        'text': '',
        'font': {'size': 18}
    },
    height=1200,
    showlegend=False,
    margin=dict(t=30, b=40, l=60, r=60)
)

# Update all annotations (subplot titles)
for annotation in fig.layout.annotations:
    annotation.font.size = 14

# Generate HTML content
print("Generating HTML dashboard...")
html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stocking Strategy Summary Dashboard</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        }}
        h1 {{
            text-align: center;
            color: #667eea;
            margin-bottom: 10px;
            font-size: 2.5em;
        }}
        .subtitle {{
            text-align: center;
            color: #666;
            margin-bottom: 30px;
            font-size: 1.1em;
        }}
        .metrics {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .metric-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }}
        .metric-value {{
            font-size: 2em;
            font-weight: bold;
            margin: 10px 0;
        }}
        .metric-label {{
            font-size: 0.9em;
            opacity: 0.9;
        }}
        .chart-container {{
            margin: 30px 0;
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
        }}
        .summary-section {{
            background: #f8f9fa;
            padding: 25px;
            border-radius: 10px;
            margin-bottom: 30px;
        }}
        .summary-title {{
            color: #667eea;
            font-size: 1.3em;
            margin-bottom: 15px;
            font-weight: 600;
        }}
        .summary-table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 5px;
            overflow: hidden;
        }}
        .summary-table th {{
            background: #667eea;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }}
        .summary-table td {{
            padding: 10px 12px;
            border-bottom: 1px solid #e0e0e0;
        }}
        .summary-table tr:hover {{
            background: #f5f5f5;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 2px solid #e0e0e0;
            color: #666;
        }}
        .github-link {{
            display: inline-block;
            margin-top: 10px;
            padding: 10px 20px;
            background: #24292e;
            color: white;
            text-decoration: none;
            border-radius: 5px;
            transition: background 0.3s;
        }}
        .github-link:hover {{
            background: #0366d6;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Stocking Strategy Summary Dashboard</h1>
        <p class="subtitle">Investment Analysis & KPI Summary | Data Excludes Q1'25</p>
        
        <div class="metrics">
            <div class="metric-card">
                <div class="metric-label">Total Records</div>
                <div class="metric-value">{total_records:,}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Investment</div>
                <div class="metric-value">${total_investment:.2f}M</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Average Investment</div>
                <div class="metric-value">${avg_investment:.2f}M</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Max Investment</div>
                <div class="metric-value">${max_investment:.2f}M</div>
            </div>
        </div>
        
        <div class="chart-container">
            <div id="mainChart"></div>
        </div>
        
        <div class="summary-section">
            <div class="summary-title">📈 Investment Summary by Category</div>
            
            <h3 style="color: #667eea; margin-top: 20px;">Investment by Owner</h3>
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>Owner</th>
                        <th>Investment ($M)</th>
                        <th>Percentage</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join([f'<tr><td>{row["OWNER"]}</td><td>${row[value_col]:.2f}</td><td>{(row[value_col]/total_investment*100):.1f}%</td></tr>' for _, row in owner_data.iterrows()])}
                </tbody>
            </table>
            
            <h3 style="color: #667eea; margin-top: 30px;">Investment by Install Check</h3>
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>Install Check</th>
                        <th>Investment ($M)</th>
                        <th>Percentage</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join([f'<tr><td>{row["Install Check"]}</td><td>${row[value_col]:.2f}</td><td>{(row[value_col]/total_investment*100):.1f}%</td></tr>' for _, row in install_data.iterrows()])}
                </tbody>
            </table>
            
            <h3 style="color: #667eea; margin-top: 30px;">Investment by Reserve Check - 3 QTR</h3>
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>Reserve Status</th>
                        <th>Investment ($M)</th>
                        <th>Percentage</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join([f'<tr><td>{row["Reserve Check - 3 years"]}</td><td>${row[value_col]:.2f}</td><td>{(row[value_col]/total_investment*100):.1f}%</td></tr>' for _, row in reserve_data.iterrows()])}
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>Generated from: {sheet_name}</p>
            <p>Data Source: Stocking Strategy Tracking.xlsx</p>
            <a href="https://github.com/Ramamas3/BOP-PROJECTS" class="github-link" target="_blank">
                🔗 View on GitHub: Ramamas3/BOP-PROJECTS
            </a>
            <p style="margin-top: 20px; font-size: 0.9em;">
                Dashboard created with Python, Plotly, and ❤️
            </p>
        </div>
    </div>
    
    <script>
        const chartData = {fig.to_json()};
        Plotly.newPlot('mainChart', chartData.data, chartData.layout, {{responsive: true}});
    </script>
</body>
</html>
"""

# Save HTML file
output_file = 'stocking_summary_dashboard.html'
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\nSummary dashboard generated successfully!")
print(f"Output file: {output_file}")
print(f"\nSummary Statistics:")
print(f"   - Total Records: {total_records:,}")
print(f"   - Total Investment: ${total_investment:.2f}M")
print(f"   - Average Investment: ${avg_investment:.2f}M")
print(f"   - Max Investment: ${max_investment:.2f}M")
print(f"\nTop 3 Owners by Investment:")
for i, row in owner_data.head(3).iterrows():
    print(f"   {i+1}. {row['OWNER']}: ${row[value_col]:.2f}M")

# Made with Bob
