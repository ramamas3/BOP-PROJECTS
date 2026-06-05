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

# Rename OWNER to Owner for better display
df.rename(columns={'OWNER': 'Owner'}, inplace=True)

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
owner_data = df.groupby('Owner')[value_col].sum().sort_values(ascending=False).reset_index()
install_data = df.groupby('Install Check')[value_col].sum().sort_values(ascending=False).reset_index()
usage_data = df.groupby('Usage Excl Transfer')[value_col].sum().sort_values(ascending=False).reset_index()
dtbh_data = df.groupby('DTBH Check')[value_col].sum().sort_values(ascending=False).reset_index()
reserve_data = df.groupby('Reserve check -  3 QTR')[value_col].sum().sort_values(ascending=False).reset_index()

# Get top 10 records by investment
top_10_records = df.nlargest(10, value_col)[value_col].reset_index(drop=True)

# Create pivot table data for Owner and Work stream
print("Creating pivot table data...")
pivot_df = df.pivot_table(
    values=value_col,
    index=['Owner', 'Work stream'],
    aggfunc='sum',
    margins=False
).reset_index()
pivot_df = pivot_df.sort_values(['Owner', value_col], ascending=[True, False])

# Prepare full dataset for filtering
print("Preparing data for filters...")
filter_columns = ['Owner', 'Work stream', 'DTBH Check', 'Usage Excl Transfer',
                  'Reserve check -  3 QTR', 'Install Check', 'POLICY CHECK- 31 May-2026', value_col]
filter_data = df[filter_columns].to_dict('records')
filter_data_json = json.dumps(filter_data)

# Create chart data for Owner, Reserve check, and Investment
print("Creating chart data...")
chart_data = df.groupby(['Owner', 'Reserve check -  3 QTR'])[value_col].sum().reset_index()
chart_data_json = json.dumps(chart_data.to_dict('records'))

# Charts removed as per user request
print("Skipping chart generation...")
fig = None

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
        .pivot-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            font-size: 0.85em;
        }}
        .pivot-table th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 8px 10px;
            text-align: left;
            font-weight: 600;
            font-size: 0.9em;
        }}
        .pivot-table td {{
            padding: 6px 10px;
            border-bottom: 1px solid #e0e0e0;
        }}
        .pivot-table tr:hover {{
            background: #f5f5f5;
        }}
        .subtotal-row {{
            background: #e3e8ff;
            font-weight: bold;
            border-top: 1px solid #667eea;
            border-bottom: 1px solid #667eea;
        }}
        .grand-total-row {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: bold;
            font-size: 1em;
            border-top: 2px solid #764ba2;
        }}
        .amount {{
            text-align: right;
            font-family: 'Courier New', monospace;
        }}
        .pivot-section {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 25px 0;
        }}
        .table-container {{
            width: 75%;
            max-height: 600px;
            overflow-y: auto;
            overflow-x: auto;
            margin: 0 auto;
        }}
        .pivot-title {{
            color: #667eea;
            font-size: 1.2em;
            margin-bottom: 15px;
            font-weight: 600;
        }}
        .chart-section {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 25px 0;
        }}
        .chart-container {{
            background: white;
            padding: 15px;
            border-radius: 8px;
            margin-top: 15px;
        }}
        .small-charts-grid {{
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            margin-top: 15px;
            justify-content: space-between;
        }}
        .chart-wrapper {{
            flex: 1 1 32%;
            max-width: 420px;
            min-width: 320px;
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        .chart-wrapper h3 {{
            color: #667eea;
            font-size: 0.95em;
            margin: 0 0 10px 0;
            text-align: center;
        }}
        .filters-section {{
            background: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
        }}
        .filters-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }}
        .filter-group {{
            display: flex;
            flex-direction: column;
        }}
        .filter-group label {{
            font-weight: 600;
            color: #667eea;
            margin-bottom: 5px;
            font-size: 0.9em;
        }}
        .filter-group select {{
            padding: 8px;
            border: 2px solid #e0e0e0;
            border-radius: 5px;
            font-size: 0.9em;
            background: white;
            cursor: pointer;
            transition: border-color 0.3s;
        }}
        .filter-group select:hover {{
            border-color: #667eea;
        }}
        .filter-group select:focus {{
            outline: none;
            border-color: #764ba2;
        }}
        .filter-buttons {{
            display: flex;
            gap: 10px;
            margin-top: 15px;
        }}
        .btn {{
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s;
        }}
        .btn-apply {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }}
        .btn-apply:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        }}
        .btn-reset {{
            background: #e0e0e0;
            color: #333;
        }}
        .btn-reset:hover {{
            background: #d0d0d0;
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
        </div>
        
        <div class="filters-section">
            <div class="pivot-title">🔍 Filters</div>
            <div class="filters-grid">
                <div class="filter-group">
                    <label for="filter-owner">Owner</label>
                    <select id="filter-owner">
                        <option value="">All Owners</option>
"""

# Add unique owners to dropdown
for owner in sorted(df['Owner'].dropna().unique()):
    html_content += f'                        <option value="{owner}">{owner}</option>\n'

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="filter-workstream">Work Stream</label>
                    <select id="filter-workstream">
                        <option value="">All Work Streams</option>
"""

# Add unique work streams to dropdown
for ws in sorted(df['Work stream'].dropna().unique()):
    html_content += f'                        <option value="{ws}">{ws}</option>\n'

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="filter-dtbh">DTBH Check</label>
                    <select id="filter-dtbh">
                        <option value="">All DTBH</option>
"""

# Add unique DTBH values
for dtbh in sorted(df['DTBH Check'].dropna().unique()):
    html_content += f'                        <option value="{dtbh}">{dtbh}</option>\n'

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="filter-usage">Usage Excl Transfer</label>
                    <select id="filter-usage">
                        <option value="">All Usage</option>
"""

# Add unique usage values
for usage in sorted(df['Usage Excl Transfer'].dropna().unique()):
    html_content += f'                        <option value="{usage}">{usage}</option>\n'

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="filter-reserve">Reserve Check - 3 QTR</label>
                    <select id="filter-reserve">
                        <option value="">All Reserve</option>
"""

# Add unique reserve values
for reserve in sorted(df['Reserve check -  3 QTR'].dropna().unique()):
    html_content += f'                        <option value="{reserve}">{reserve}</option>\n'

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="filter-install">Install Check</label>
                    <select id="filter-install">
                        <option value="">All Install</option>
"""

# Add unique install values
for install in sorted(df['Install Check'].dropna().unique()):
    html_content += f'                        <option value="{install}">{install}</option>\n'

html_content += """                    </select>
                </div>
                <div class="filter-group">
                    <label for="filter-policy">POLICY CHECK- 31 May-2026</label>
                    <select id="filter-policy">
                        <option value="">All Policy</option>
"""

# Add unique policy values
for policy in sorted(df['POLICY CHECK- 31 May-2026'].dropna().unique()):
    html_content += f'                        <option value="{policy}">{policy}</option>\n'

html_content += """                    </select>
                </div>
            </div>
            <div class="filter-buttons">
                <button class="btn btn-apply" onclick="applyFilters()">Apply Filters</button>
                <button class="btn btn-reset" onclick="resetFilters()">Reset Filters</button>
            </div>
        </div>
        
        <div class="pivot-section">
            <div class="pivot-title">📊 Investment Pivot Table - Owner & Work Stream</div>
            <div class="table-container">
                <table class="pivot-table">
                    <thead>
                        <tr>
                            <th>Owner</th>
                            <th>Work Stream</th>
                            <th style="text-align: right;">Investment ($M)</th>
                        </tr>
                    </thead>
                    <tbody>
"""

# Generate pivot table rows with subtotals
current_owner = None
owner_total = 0
grand_total = 0

for idx, row in pivot_df.iterrows():
    owner = row['Owner']
    work_stream = row['Work stream']
    investment = row[value_col]
    
    # If owner changed, add subtotal for previous owner
    if current_owner is not None and current_owner != owner:
        html_content += f"""
                    <tr class="subtotal-row">
                        <td><strong>{current_owner} Subtotal:</strong></td>
                        <td></td>
                        <td class="amount">${owner_total:.2f}</td>
                    </tr>
"""
        owner_total = 0
    
    # Add data row
    html_content += f"""
                    <tr>
                        <td>{owner if owner != current_owner else ''}</td>
                        <td>{work_stream}</td>
                        <td class="amount">${investment:.2f}</td>
                    </tr>
"""
    
    current_owner = owner
    owner_total += investment
    grand_total += investment

# Add last owner's subtotal
if current_owner is not None:
    html_content += f"""
                    <tr class="subtotal-row">
                        <td><strong>{current_owner} Subtotal:</strong></td>
                        <td></td>
                        <td class="amount">${owner_total:.2f}</td>
                    </tr>
"""

# Add grand total
html_content += f"""
                    <tr class="grand-total-row">
                        <td colspan="2" style="text-align: right; padding-right: 20px;">
                            <strong>GRAND TOTAL:</strong>
                        </td>
                        <td class="amount">${grand_total:.2f}</td>
                    </tr>
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="chart-section">
            <div class="pivot-title">📊 Investment Analysis Charts</div>
            <div class="small-charts-grid">
                <div class="chart-wrapper">
                    <h3>Owner vs Investment</h3>
                    <div id="ownerChart"></div>
                </div>
                <div class="chart-wrapper">
                    <h3>Install Check vs Investment</h3>
                    <div id="installChart"></div>
                </div>
                <div class="chart-wrapper">
                    <h3>Usage Excl Transfer vs Investment</h3>
                    <div id="usageChart"></div>
                </div>
                <div class="chart-wrapper">
                    <h3>Reserve Check vs Investment</h3>
                    <div id="reserveChart"></div>
                </div>
                <div class="chart-wrapper">
                    <h3>DTBH Check vs Investment</h3>
                    <div id="dtbhChart"></div>
                </div>
                <div class="chart-wrapper">
                    <h3>Policy Check vs Investment</h3>
                    <div id="policyChart"></div>
                </div>
            </div>
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
        // Full dataset for filtering
        const fullData = {filter_data_json};
        const chartData = {chart_data_json};
        const valueCol = '{value_col}';
        
        // Create initial charts
        createAllCharts(fullData);
        
        function createAllCharts(data) {{
            createOwnerChart(data);
            createInstallChart(data);
            createUsageChart(data);
            createReserveChart(data);
            createDtbhChart(data);
            createPolicyChart(data);
        }}
        
        function createOwnerChart(data) {{
            const grouped = {{}};
            data.forEach(row => {{
                const owner = row['Owner'];
                const investment = row[valueCol];
                if (!grouped[owner]) grouped[owner] = 0;
                grouped[owner] += investment;
            }});
            
            const sortedData = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
            const categories = sortedData.map(d => d[0]);
            const values = sortedData.map(d => d[1]);
            
            const trace = {{
                x: categories,
                y: values,
                type: 'bar',
                marker: {{ color: '#667eea' }},
                text: values.map(v => `$${{v.toFixed(2)}}M`),
                textposition: 'outside'
            }};
            
            const layout = {{
                margin: {{ l: 50, r: 20, t: 20, b: 80 }},
                xaxis: {{ tickangle: -45 }},
                yaxis: {{ title: 'Investment ($M)' }},
                height: 300,
                showlegend: false
            }};
            
            Plotly.newPlot('ownerChart', [trace], layout, {{responsive: true}});
        }}
        
        function createInstallChart(data) {{
            const grouped = {{}};
            data.forEach(row => {{
                const install = row['Install Check'];
                const investment = row[valueCol];
                if (!grouped[install]) grouped[install] = 0;
                grouped[install] += investment;
            }});
            
            const sortedData = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
            const categories = sortedData.map(d => d[0]);
            const values = sortedData.map(d => d[1]);
            
            const trace = {{
                x: categories,
                y: values,
                type: 'bar',
                marker: {{ color: '#764ba2' }},
                text: values.map(v => `$${{v.toFixed(2)}}M`),
                textposition: 'outside'
            }};
            
            const layout = {{
                margin: {{ l: 50, r: 20, t: 20, b: 80 }},
                xaxis: {{ tickangle: -45 }},
                yaxis: {{ title: 'Investment ($M)' }},
                height: 300,
                showlegend: false
            }};
            
            Plotly.newPlot('installChart', [trace], layout, {{responsive: true}});
        }}
        
        function createUsageChart(data) {{
            const grouped = {{}};
            data.forEach(row => {{
                const usage = row['Usage Excl Transfer'];
                const investment = row[valueCol];
                if (!grouped[usage]) grouped[usage] = 0;
                grouped[usage] += investment;
            }});
            
            const sortedData = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
            const categories = sortedData.map(d => d[0]);
            const values = sortedData.map(d => d[1]);
            
            const trace = {{
                x: categories,
                y: values,
                type: 'bar',
                marker: {{ color: '#FFA500' }},
                text: values.map(v => `$${{v.toFixed(2)}}M`),
                textposition: 'outside'
            }};
            
            const layout = {{
                margin: {{ l: 50, r: 20, t: 20, b: 80 }},
                xaxis: {{ tickangle: -45 }},
                yaxis: {{ title: 'Investment ($M)' }},
                height: 300,
                showlegend: false
            }};
            
            Plotly.newPlot('usageChart', [trace], layout, {{responsive: true}});
        }}
        
        function createReserveChart(data) {{
            const grouped = {{}};
            data.forEach(row => {{
                const reserve = row['Reserve check -  3 QTR'];
                const investment = row[valueCol];
                if (!grouped[reserve]) grouped[reserve] = 0;
                grouped[reserve] += investment;
            }});
            
            const sortedData = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
            const categories = sortedData.map(d => d[0]);
            const values = sortedData.map(d => d[1]);
            
            const colorMap = {{
                'No Reserve Impact': '#90EE90',
                'Reserve Forecast': '#FFA500',
                'Already Reserved': '#FFB6C1'
            }};
            const colors = categories.map(c => colorMap[c] || '#667eea');
            
            const trace = {{
                x: categories,
                y: values,
                type: 'bar',
                marker: {{ color: colors }},
                text: values.map(v => `$${{v.toFixed(2)}}M`),
                textposition: 'outside'
            }};
            
            const layout = {{
                margin: {{ l: 50, r: 20, t: 20, b: 80 }},
                xaxis: {{ tickangle: -45 }},
                yaxis: {{ title: 'Investment ($M)' }},
                height: 300,
                showlegend: false
            }};
            
            Plotly.newPlot('reserveChart', [trace], layout, {{responsive: true}});
        }}
        
        function createDtbhChart(data) {{
            const grouped = {{}};
            data.forEach(row => {{
                const dtbh = row['DTBH Check'];
                const investment = row[valueCol];
                if (!grouped[dtbh]) grouped[dtbh] = 0;
                grouped[dtbh] += investment;
            }});
            
            const sortedData = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
            const categories = sortedData.map(d => d[0]);
            const values = sortedData.map(d => d[1]);
            
            const trace = {{
                x: categories,
                y: values,
                type: 'bar',
                marker: {{ color: '#20B2AA' }},
                text: values.map(v => `$${{v.toFixed(2)}}M`),
                textposition: 'outside'
            }};
            
            const layout = {{
                margin: {{ l: 50, r: 20, t: 20, b: 80 }},
                xaxis: {{ tickangle: -45 }},
                yaxis: {{ title: 'Investment ($M)' }},
                height: 300,
                showlegend: false
            }};
            
            Plotly.newPlot('dtbhChart', [trace], layout, {{responsive: true}});
        }}
        
        function createPolicyChart(data) {{
            const grouped = {{}};
            data.forEach(row => {{
                const policy = row['POLICY CHECK- 31 May-2026'];
                const investment = row[valueCol];
                if (!grouped[policy]) grouped[policy] = 0;
                grouped[policy] += investment;
            }});
            
            const sortedData = Object.entries(grouped).sort((a, b) => b[1] - a[1]);
            const categories = sortedData.map(d => d[0]);
            const values = sortedData.map(d => d[1]);
            
            const trace = {{
                x: categories,
                y: values,
                type: 'bar',
                marker: {{ color: '#FF6B6B' }},
                text: values.map(v => `$${{v.toFixed(2)}}M`),
                textposition: 'outside'
            }};
            
            const layout = {{
                margin: {{ l: 50, r: 20, t: 20, b: 80 }},
                xaxis: {{ tickangle: -45 }},
                yaxis: {{ title: 'Investment ($M)' }},
                height: 300,
                showlegend: false
            }};
            
            Plotly.newPlot('policyChart', [trace], layout, {{responsive: true}});
        }}
        
        function createChart(data) {{
            // This function is kept for backward compatibility but not used
            const owners = [...new Set(data.map(d => d.Owner))];
            const reserves = [...new Set(data.map(d => d['Reserve check -  3 QTR']))];
            
            const colorMap = {{
                'No Reserve Impact': '#90EE90',
                'Reserve Forecast': '#FFA500',
                'Already Reserved': '#FFB6C1'
            }};
            
            const traces = reserves.map(reserve => {{
                const yValues = owners.map(owner => {{
                    const record = data.find(d => d.Owner === owner && d['Reserve check -  3 QTR'] === reserve);
                    return record ? record[valueCol] : 0;
                }});
                
                return {{
                    x: owners,
                    y: yValues,
                    name: reserve,
                    type: 'bar',
                    marker: {{
                        color: colorMap[reserve] || '#667eea'
                    }},
                    text: yValues.map(v => `$${{v.toFixed(2)}}`),
                    textposition: 'auto',
                    textfont: {{
                        size: 10
                    }}
                }};
            }});
            
            const layout = {{
                title: {{
                    text: 'Investment by Owner and Reserve Check - 3 QTR',
                    font: {{ size: 16, color: '#667eea' }}
                }},
                barmode: 'group',
                xaxis: {{
                    title: 'Owner',
                    tickangle: -45
                }},
                yaxis: {{
                    title: 'Investment ($M)'
                }},
                showlegend: true,
                legend: {{
                    orientation: 'h',
                    y: -0.2
                }},
                margin: {{ t: 50, b: 100, l: 60, r: 20 }},
                height: 500
            }};
            
            Plotly.newPlot('investmentChart', traces, layout, {{responsive: true}});
        }}
        
        function applyFilters() {{
            const filters = {{
                owner: document.getElementById('filter-owner').value,
                workstream: document.getElementById('filter-workstream').value,
                dtbh: document.getElementById('filter-dtbh').value,
                usage: document.getElementById('filter-usage').value,
                reserve: document.getElementById('filter-reserve').value,
                install: document.getElementById('filter-install').value,
                policy: document.getElementById('filter-policy').value
            }};
            
            // Filter the data
            let filteredData = fullData.filter(row => {{
                if (filters.owner && row['Owner'] !== filters.owner) return false;
                if (filters.workstream && row['Work stream'] !== filters.workstream) return false;
                if (filters.dtbh && row['DTBH Check'] !== filters.dtbh) return false;
                if (filters.usage && row['Usage Excl Transfer'] !== filters.usage) return false;
                if (filters.reserve && row['Reserve check -  3 QTR'] !== filters.reserve) return false;
                if (filters.install && row['Install Check'] !== filters.install) return false;
                if (filters.policy && row['POLICY CHECK- 31 May-2026'] !== filters.policy) return false;
                return true;
            }});
            
            // Rebuild pivot table
            rebuildPivotTable(filteredData);
            
            // Update metrics
            updateMetrics(filteredData);
            
            // Update all charts
            createAllCharts(filteredData);
        }}
        
        function rebuildPivotTable(data) {{
            // Group by Owner and Work stream
            const grouped = {{}};
            data.forEach(row => {{
                const owner = row['Owner'];
                const workstream = row['Work stream'];
                const investment = row[valueCol];
                
                if (!grouped[owner]) {{
                    grouped[owner] = {{}};
                }}
                if (!grouped[owner][workstream]) {{
                    grouped[owner][workstream] = 0;
                }}
                grouped[owner][workstream] += investment;
            }});
            
            // Build table HTML
            let tableHTML = '';
            let grandTotal = 0;
            
            Object.keys(grouped).sort().forEach((owner, ownerIndex) => {{
                let ownerTotal = 0;
                const workstreams = grouped[owner];
                
                Object.keys(workstreams).sort((a, b) => workstreams[b] - workstreams[a]).forEach((workstream, wsIndex) => {{
                    const investment = workstreams[workstream];
                    ownerTotal += investment;
                    grandTotal += investment;
                    
                    tableHTML += `
                        <tr>
                            <td>${{wsIndex === 0 ? owner : ''}}</td>
                            <td>${{workstream}}</td>
                            <td class="amount">$${{investment.toFixed(2)}}</td>
                        </tr>
                    `;
                }});
                
                // Add subtotal
                tableHTML += `
                    <tr class="subtotal-row">
                        <td><strong>${{owner}} Subtotal:</strong></td>
                        <td></td>
                        <td class="amount">$${{ownerTotal.toFixed(2)}}</td>
                    </tr>
                `;
            }});
            
            // Add grand total
            tableHTML += `
                <tr class="grand-total-row">
                    <td colspan="2" style="text-align: right; padding-right: 20px;">
                        <strong>GRAND TOTAL:</strong>
                    </td>
                    <td class="amount">$${{grandTotal.toFixed(2)}}</td>
                </tr>
            `;
            
            document.querySelector('.pivot-table tbody').innerHTML = tableHTML;
        }}
        
        function updateMetrics(data) {{
            const totalRecords = data.length;
            const totalInvestment = data.reduce((sum, row) => sum + row[valueCol], 0);
            
            document.querySelector('.metric-card:nth-child(1) .metric-value').textContent = totalRecords.toLocaleString();
            document.querySelector('.metric-card:nth-child(2) .metric-value').textContent = `$${{totalInvestment.toFixed(2)}}M`;
        }}
        
        function resetFilters() {{
            document.getElementById('filter-owner').value = '';
            document.getElementById('filter-workstream').value = '';
            document.getElementById('filter-dtbh').value = '';
            document.getElementById('filter-usage').value = '';
            document.getElementById('filter-reserve').value = '';
            document.getElementById('filter-install').value = '';
            document.getElementById('filter-policy').value = '';
            
            rebuildPivotTable(fullData);
            updateMetrics(fullData);
            createAllCharts(fullData);
        }}
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
print(f"\nTop 3 Owners by Investment:")
for i, row in owner_data.head(3).iterrows():
    print(f"   {i+1}. {row['Owner']}: ${row[value_col]:.2f}M")

# Made with Bob
