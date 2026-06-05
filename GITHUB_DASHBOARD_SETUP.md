# GitHub Dashboard Setup Guide

## 📋 Prerequisites
1. Close the Excel file: `Stocking Strategy Tracking.xlsx`
2. Ensure you have Python with pandas, openpyxl, and pathlib installed

## 🚀 Steps to Create and Deploy Dashboard

### Step 1: Generate Dashboard Files
```bash
python create_github_dashboard.py
```

This will create:
- `docs/dashboard_data.json` - Data file
- `docs/index.html` - Interactive dashboard

### Step 2: Push to GitHub
```bash
git add docs/
git commit -m "Add interactive dashboard"
git push origin main
```

### Step 3: Enable GitHub Pages
1. Go to your repository: https://github.com/Ramamas3/BOP-PROJECTS
2. Click on **Settings**
3. Scroll down to **Pages** section (left sidebar)
4. Under **Source**, select:
   - Branch: `main`
   - Folder: `/docs`
5. Click **Save**

### Step 4: Access Your Dashboard
After a few minutes, your dashboard will be live at:
```
https://ramamas3.github.io/BOP-PROJECTS/
```

## 📊 Dashboard Features
- ✅ Interactive charts using Plotly
- ✅ Real-time data loading from JSON
- ✅ Responsive design
- ✅ Summary metrics cards
- ✅ Detailed breakdown tables
- ✅ Shareable link

## 🔄 Updating the Dashboard
To update the dashboard with new data:
1. Close the Excel file
2. Run: `python create_github_dashboard.py`
3. Commit and push the updated `docs/` folder
4. GitHub Pages will automatically update

## 📱 Sharing the Dashboard
Simply share this link with anyone:
```
https://ramamas3.github.io/BOP-PROJECTS/
```

No login required - anyone with the link can view the dashboard!

## 🛠️ Troubleshooting

### Permission Error
- **Problem**: `PermissionError: [Errno 13]`
- **Solution**: Close the Excel file in Excel application

### Dashboard Not Loading
- **Problem**: Dashboard shows "Loading..." forever
- **Solution**: Check that `dashboard_data.json` exists in the `docs/` folder

### GitHub Pages Not Working
- **Problem**: 404 error when accessing the URL
- **Solution**: 
  - Wait 5-10 minutes after enabling GitHub Pages
  - Verify the `docs/` folder is in the main branch
  - Check GitHub Pages settings are correct

## 📁 File Structure
```
BOB/
├── docs/
│   ├── index.html              # Dashboard HTML
│   └── dashboard_data.json     # Data file
├── create_github_dashboard.py  # Generator script
└── GITHUB_DASHBOARD_SETUP.md   # This file
```

## 🎯 Next Steps
1. Close Excel file
2. Run `python create_github_dashboard.py`
3. Follow steps above to deploy to GitHub Pages
4. Share the link with your team!