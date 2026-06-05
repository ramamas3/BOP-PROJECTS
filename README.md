# Stocking Strategy Dashboard

## 📊 Overview
Interactive dashboard for analyzing stocking strategy and investment tracking data from the "Stocking Strategy Tracking.xlsx" file.

## 🎯 Key Metrics Summary

### Total Statistics
- **Total Records**: 15,733
- **Total Investment**: $105.04M
- **Average Investment per Record**: $0.01M
- **Maximum Investment**: $1.06M

### Investment by Owner
| Owner | Total Investment ($M) | Count | Avg Investment ($M) |
|-------|----------------------|-------|---------------------|
| Aurelio | 47.07 | 5,568 | 0.01 |
| Diego | 24.65 | 5,551 | 0.00 |
| Mahesh | 16.01 | 1,538 | 0.01 |
| Krish | 11.69 | 1,092 | 0.01 |
| Asha | 4.13 | 1,385 | 0.00 |
| Smitha | 0.91 | 256 | 0.00 |
| Ramamani | 0.57 | 343 | 0.00 |

### Investment by Install Check
- **>500**: $48.19M
- **101 to 500**: $29.00M
- **1 to 100**: $21.34M
- **Zero Installs**: $6.51M

### Investment by Usage Excl Transfer
- **0 USAGE**: $86.42M
- **1 to 3**: $8.45M
- **>10**: $5.20M
- **4 to 10**: $4.97M

### Investment by DTBH Check
- **DOB < 2 years**: $57.21M
- **DOB>5**: $24.71M
- **DOB 2-5 years**: $23.12M

### Investment by Reserve Check (3 years)
- **No Reserve Impact**: $96.37M
- **Already Reserved**: $5.07M
- **Reserve Forecast**: $3.61M

## 📁 Files

- **stocking_dashboard.html** - Main interactive dashboard (standalone HTML file)
- **dashboard_data.json** - JSON data export for API integration
- **create_github_dashboard.py** - Python script to generate dashboard and summary

## 🚀 Viewing the Dashboard

### Option 1: Local Viewing
Simply open `stocking_dashboard.html` in any modern web browser.

### Option 2: GitHub Pages (Recommended for Sharing)

1. **Create/Access Repository**
   - Repository: `BOP-PROJECTS`
   - GitHub Username: `Ramamas3`

2. **Upload Files**
   ```bash
   git clone https://github.com/Ramamas3/BOP-PROJECTS.git
   cd BOP-PROJECTS
   # Copy stocking_dashboard.html and dashboard_data.json to the repository
   git add stocking_dashboard.html dashboard_data.json README.md
   git commit -m "Add stocking strategy dashboard"
   git push origin main
   ```

3. **Enable GitHub Pages**
   - Go to repository Settings
   - Navigate to "Pages" section
   - Under "Source", select "main" branch
   - Click "Save"

4. **Access Dashboard**
   - URL: `https://ramamas3.github.io/BOP-PROJECTS/stocking_dashboard.html`
   - Share this link with stakeholders

## 🛠️ Technical Details

### Technologies Used
- **Python**: Data processing and analysis
- **Pandas**: Excel data manipulation
- **Plotly**: Interactive visualizations
- **HTML/CSS/JavaScript**: Dashboard frontend

### Data Source
- **File**: Stocking Strategy Tracking.xlsx
- **Sheet**: "Stocking stratergy by PN -New"
- **Total Columns**: 67
- **Key Metrics Column**: TOTAL INVESTMENT AS Per GPP($M)

## 📊 Dashboard Features

1. **Summary Cards**: Quick overview of key metrics
2. **Interactive Charts**: 
   - Investment by Owner (Pie Chart)
   - Investment by Install Check (Bar Chart)
   - Investment by Usage Excl Transfer (Bar Chart)
   - Investment by DTBH Check (Bar Chart)
   - Investment by Reserve Check (Bar Chart)
   - Top 10 Records by Investment (Bar Chart)
3. **Detailed Tables**: Breakdown by various categories
4. **Responsive Design**: Works on desktop and mobile devices
5. **Embedded Data**: No external dependencies required

## 🔄 Updating the Dashboard

To regenerate the dashboard with updated data:

```bash
python create_github_dashboard.py
```

This will:
1. Read the latest Excel data
2. Generate a new summary
3. Create updated JSON data
4. Generate a fresh HTML dashboard

## 📝 Notes

- Dashboard is fully self-contained (no external API calls)
- All data is embedded in the HTML file
- Can be shared as a single file
- Works offline after initial load

## 👤 Author

**GitHub**: [Ramamas3](https://github.com/Ramamas3)  
**Repository**: [BOP-PROJECTS](https://github.com/Ramamas3/BOP-PROJECTS)

## 📅 Last Updated

Generated: 2026-06-04

---

For questions or issues, please open an issue in the GitHub repository.