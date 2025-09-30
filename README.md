```markdown
# 🎫 Ticket Sales Analytics Platform

## 📊 Project Overview
A comprehensive web-based analytics platform for ticket sales data with interactive visualizations and Excel export capabilities.

## 🚀 Quick Start

### Installation
```bash
pip install -r requirements.txt
```

### Run Application
```bash
python app.py
```

### Access Dashboard
Open in browser: `http://127.0.0.1:56777`

## 📈 Features

### Static Charts (Matplotlib)
- **Pie Chart** - Revenue distribution by event categories
- **Bar Chart** - Top venues by event count
- **Horizontal Bar Chart** - Average transaction by state
- **Line Chart** - Monthly sales trends
- **Histogram** - Ticket price distribution
- **Scatter Plot** - Price vs quantity sold correlation

### Interactive Dashboards (Plotly)
- **Animated Category Sales** - Monthly revenue by categories with slider
- **Advanced Sales Dashboard** - Interactive scatter plot with state data
- **Time Series Analysis** - Dynamic sales trends with animation

### Data Export
- **Excel Reports** with automatic formatting:
  - Gradient color scales
  - Column filters
  - Frozen headers
  - Conditional formatting

## 🛠 Technologies Used
- **Backend**: Python, Flask, SQLAlchemy
- **Database**: PostgreSQL
- **Visualization**: Matplotlib, Plotly, Seaborn
- **Data Processing**: Pandas
- **Export**: OpenPyXL

## 📁 Project Structure
```
Ticket2/
├── analytics.py          # Main analytics engine
├── app.py               # Flask web server
├── requirements.txt     # Python dependencies
├── templates/          # HTML templates
│   ├── index.html
│   └── interactive_charts.html
└── README.md
```

## 🗃️ Database Schema
- **Users** - Customer information and preferences
- **Events** - Concerts, shows, and performances
- **Categories** - Event types (Music, Sports, Theater)
- **Sales** - Transaction records
- **Venues** - Event locations
- **Listings** - Ticket listings

## 🎯 Use Cases
- Sales performance analysis
- Customer behavior insights
- Revenue optimization
- Event popularity tracking
- Geographic sales distribution

## 🔧 Development
```bash
# Create virtual environment
python -m venv venv

# Activate environment
venv\Scripts\activate  # Windows
source venv/bin/activate  # Linux/Mac

# Install dependencies
pip install -r requirements.txt

# Run development server
python app.py
```

## 📊 Sample Insights
- Identify most profitable event categories
- Track seasonal sales patterns
- Analyze customer spending habits
- Optimize ticket pricing strategies
- Monitor venue performance metrics

---

**Developed for educational purposes** 🎓
```
