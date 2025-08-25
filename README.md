# 📊 Membership Analysis Dashboard

A comprehensive data analysis tool for analyzing membership data across different locations and time periods.

## 🚀 Features

- **Auto-load Excel files**: Place your Excel file in the root directory for automatic loading
- **Multi-sheet Excel support**: Analyze data from Stockholm, Göteborg, and other locations across different years
- **Three analysis modes**:
  - Single Location Analysis
  - Two Location Comparison (side-by-side)
  - Multi-Location Comparison (all locations overview)
- **Interactive visualizations**: 
  - Monthly membership trends
  - Membership type distribution pie charts
  - Activity heatmaps
  - Top performers analysis
  - Growth rate comparisons
- **Automated insights**: Get key metrics and trends automatically
- **Export capabilities**: Save analysis results to Excel

## 📁 Auto-Loading Excel Files

Place your Excel file in the root directory with one of these names for automatic loading:
- `membership_data.xlsx`
- `memberships.xlsx`
- `data.xlsx`
- `Masterdokument Memberships Antal.xlsx`

The dashboard will automatically detect and load the file on startup!

## 📋 Data Format

Your Excel file should contain sheets named like:
- `Stockholm 2024`
- `Stockholm 2025`
- `Göteborg 2024`
- `Göteborg 2025`

Each sheet should have:
- First column: Membership types
- Monthly columns: `Januar 2024`, `Februari 2024`, etc.
- The system automatically filters out total and percentage columns

## 🛠️ Local Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Run the dashboard:
```bash
streamlit run dashboard.py
```

## 🌐 Online Deployment

### Streamlit Community Cloud (Recommended)
1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repository
4. Deploy with main file: `streamlit_app.py`

### Vercel Deployment
1. Install Vercel CLI: `npm i -g vercel`
2. Run `vercel` in this directory
3. Follow the prompts

### Heroku Deployment
1. Create a `Procfile`:
```
web: streamlit run streamlit_app.py --server.port=$PORT --server.address=0.0.0.0
```

2. Deploy to Heroku:
```bash
git init
git add .
git commit -m "Initial commit"
heroku create your-app-name
git push heroku main
```

## 🎯 Usage

### Single Location Analysis
- Deep dive into one specific location
- All standard charts and insights
- Perfect for detailed analysis

### Two Location Comparison
- Side-by-side comparison of any 2 locations
- All charts shown in parallel
- Head-to-head winner analysis
- Performance gap calculations

### Multi-Location Overview
- Compare all locations at once
- Total memberships ranking
- Monthly performance overlay
- Growth rate comparison
- Membership type distribution across locations
- Summary statistics table

## 📊 Available Visualizations

1. **Monthly Trends**: Line charts showing membership acquisitions over time
2. **Distribution Pie Charts**: Percentage breakdown of membership types
3. **Activity Heatmaps**: Visual representation of activity by type and month
4. **Top Performers**: Bar charts of most popular membership types
5. **Growth Comparisons**: Color-coded growth rate analysis
6. **Distribution Comparisons**: Stacked bars showing membership types across locations

## 🔍 Automated Insights

The system automatically generates:
- Total membership counts
- Top performing membership types
- Best and worst performing months
- Growth trends and patterns
- Performance gaps between locations
- Year-over-year comparisons

## 📁 File Structure

```
├── dashboard.py              # Main dashboard application
├── streamlit_app.py         # Entry point for deployment
├── membership_analyzer.py   # Core analysis engine
├── requirements.txt         # Python dependencies
├── runtime.txt             # Python version for deployment
├── .streamlit/config.toml  # Streamlit configuration
└── README.md               # This file
```

## 🎨 Dashboard Features

- **Responsive Design**: Works on desktop and mobile
- **Interactive Charts**: Hover, zoom, and explore data
- **Real-time Analysis**: Instant updates when switching views
- **Export Options**: Download results and summaries
- **Clean UI**: Professional styling with clear navigation

## 🚀 Getting Started

1. Clone or download this repository
2. Place your Excel file in the root directory (it will auto-load!)
3. Run `streamlit run dashboard.py`
4. Open your browser to `http://localhost:8501`
5. Start exploring your membership data!

## 📈 Example Analysis Output

```
📊 Total memberships: 482
🏆 Top performing membership: Löpande Membership Standard (162 members)
📈 Average memberships per active type: 26.8
🗓️ Best performing month: Februari 2024 (49 new members)
📉 Lowest performing month: December 2024 (17 new members)
📈 Growing trend: 15.2% increase from first to second half
```

Enjoy analyzing your membership data! 📊✨