import pandas as pd
import json
from pathlib import Path

# File configuration
file_path = r'C:\Users\RAMAMANIS\OneDrive - IBM\Desktop\python\BOB\Stocking Strategy Tracking.xlsx'
sheet_name = "Stocking stratergy by PN -New"

print("Loading data...")
df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

# Clean column names
df.columns = df.columns.str.strip()

# Rename OWNER to Owner for better display
df.rename(columns={'OWNER': 'Owner'}, inplace=True)

# Required columns
value_col = 'TOTAL INVESTMENT AS Per GPP($M)'
df[value_col] = pd.to_numeric(df[value_col], errors='coerce').fillna(0)

print("Generating summary statistics...")
# Summary statistics
total_records = len(df)
total_investment = float(df[value_col].sum())
avg_investment = float(df[value_col].mean())
max_investment = float(df[value_col].max())

# Create aggregations
owner_data = df.groupby('Owner')[value_col].sum().sort_values(ascending=False).reset_index()
install_data = df.groupby('Install Check')[value_col].sum().sort_values(ascending=False).reset_index()
usage_data = df.groupby('Usage Excl Transfer')[value_col].sum().sort_values(ascending=False).reset_index()
dtbh_data = df.groupby('DTBH Check')[value_col].sum().sort_values(ascending=False).reset_index()
reserve_data = df.groupby('Reserve check -  3 QTR')[value_col].sum().sort_values(ascending=False).reset_index()

# Convert to JSON-serializable format
dashboard_data = {
    'summary': {
        'total_records': total_records,
        'total_investment': round(total_investment, 2),
        'avg_investment': round(avg_investment, 2),
        'max_investment': round(max_investment, 2)
    },
    'owner': owner_data.to_dict('records'),
    'install': install_data.to_dict('records'),
    'usage': usage_data.to_dict('records'),
    'dtbh': dtbh_data.to_dict('records'),
    'reserve': reserve_data.to_dict('records')
}

# Create docs directory if it doesn't exist
docs_dir = Path('docs')
docs_dir.mkdir(exist_ok=True)

# Save JSON data
json_file = docs_dir / 'dashboard_data.json'
with open(json_file, 'w') as f:
    json.dump(dashboard_data, f, indent=2)

print(f"JSON data saved to: {json_file}")

# Create HTML dashboard
html_content = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Stocking Strategy Dashboard - BOP Projects</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
            padding: 20px;
            min-height: 100vh;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        }
        
        header {
            text-align: center;
            margin-bottom: 30px;
        }
        
        h1 {
            color: #667eea;
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        
        .subtitle {
            color: #666;
            font-size: 1.1em;
            margin-bottom: 10px;
        }
        
        .github-badge {
            display: inline-block;
            padding: 8px 16px;
            background: #24292e;
            color: white;
            text-decoration: none;
            border-radius: 5px;
            font-size: 0.9em;
            transition: background 0.3s;
        }
        
        .github-badge:hover {
            background: #0366d6;
        }
        
        .metrics {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s;
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
        }
        
        .metric-value {
            font-size: 2em;
            font-weight: bold;
            margin: 10px 0;
        }
        
        .metric-label {
            font-size: 0.9em;
            opacity: 0.9;
        }
        
        .charts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 30px;
            margin: 30px 0;
        }
        
        .chart-container {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .chart-title {
            color: #667eea;
            font-size: 1.2em;
            font-weight: 600;
            margin-bottom: 15px;
            text-align: center;
        }
        
        .summary-section {
            background: #f8f9fa;
            padding: 25px;
            border-radius: 10px;
            margin: 30px 0;
        }
        
        .summary-title {
            color: #667eea;
            font-size: 1.3em;
            margin-bottom: 20px;
            font-weight: 600;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 5px;
            overflow: hidden;
            margin: 20px 0;
        }
        
        .data-table th {
            background: #667eea;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }
        
        .data-table td {
            padding: 10px 12px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        .data-table tr:hover {
            background: #f5f5f5;
        }
        
        footer {
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 2px solid #e0e0e0;
            color: #666;
        }
        
        .loading {
            text-align: center;
            padding: 40px;
            color: #667eea;
            font-size: 1.2em;
        }
        
        .error {
            background: #fee;
            border: 1px solid #fcc;
            color: #c33;
            padding: 20px;
            border-radius: 5px;
            margin: 20px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 Stocking Strategy Dashboard</h1>
            <p class="subtitle">Investment Analysis & KPI Summary | Data Excludes Q1'25</p>
            <a href="https://github.com/Ramamas3/BOP-PROJECTS" class="github-badge" target="_blank">
                🔗 View on GitHub: Ramamas3/BOP-PROJECTS
            </a>
        </header>
        
        <div id="loading" class="loading">Loading dashboard data...</div>
        <div id="error" class="error" style="display: none;"></div>
        
        <div id="dashboard" style="display: none;">
            <div class="metrics" id="metrics"></div>
            
        </div>
        
        <footer>
            <p>Generated from: Stocking stratergy by PN -New</p>
            <p>Data Source: Stocking Strategy Tracking.xlsx</p>
            <p style="margin-top: 20px; font-size: 0.9em;">
                Dashboard created with Python, Plotly, and ❤️
            </p>
        </footer>
    </div>
    
    <script>
        // Load dashboard data
        fetch('dashboard_data.json')
            .then(response => {
                if (!response.ok) {
                    throw new Error('Failed to load dashboard data');
                }
                return response.json();
            })
            .then(data => {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('dashboard').style.display = 'block';
                renderDashboard(data);
            })
            .catch(error => {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('error').style.display = 'block';
                document.getElementById('error').textContent = 'Error loading dashboard: ' + error.message;
            });
        
        function renderDashboard(data) {
            // Render metrics
            const metricsHtml = `
                <div class="metric-card">
                    <div class="metric-label">Total Records</div>
                    <div class="metric-value">${data.summary.total_records.toLocaleString()}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Total Investment</div>
                    <div class="metric-value">$${data.summary.total_investment.toFixed(2)}M</div>
                </div>
            `;
            document.getElementById('metrics').innerHTML = metricsHtml;
            
            // Charts and tables removed as per user request
        }
    </script>
</body>
</html>
'''

# Save HTML file
html_file = docs_dir / 'index.html'
with open(html_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"HTML dashboard saved to: {html_file}")
print("\nNext steps to share your dashboard:")
print("1. Commit and push the 'docs' folder to your GitHub repository")
print("2. Go to GitHub repository settings")
print("3. Enable GitHub Pages and select 'docs' folder as source")
print("4. Your dashboard will be available at:")
print("   https://ramamas3.github.io/BOP-PROJECTS/")
print("\nDashboard created successfully!")

# Made with Bob
