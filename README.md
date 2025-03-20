# Amazon Ad Campaign Analysis Dashboard

A comprehensive analytics dashboard for Amazon advertising campaigns, providing insights into campaign performance, product metrics, and optimization opportunities.

## Features

### 1. Campaign Performance Analysis

- Total spend and sales tracking
- Campaign-level performance metrics
- ROAS and ACOS analysis
- Click-through and conversion rates

### 2. Product Performance Insights

- Product-level performance metrics
- Sales and spend analysis by ASIN
- Performance categorization (Over-performing, Moderate, Under-performing)
- ROI analysis for each product

### 3. Search Term Analysis

- Search term performance tracking
- Impression share analysis
- Search term ranking insights
- Conversion optimization opportunities

### 4. Trend Analysis

- Daily performance trends
- Hourly performance patterns
- ROAS and conversion rate tracking
- Spend and sales trends

### 5. Interactive Visualizations

- Campaign Spend vs Sales comparison
- Product Performance scatter plots
- Daily Trends line charts
- Search Term Performance analysis

## Project Structure

```
amazon-ad-analysis/
├── data/                    # Data files directory
│   ├── campaign_reports/    # Campaign performance data
│   ├── search_terms/       # Search term data
│   ├── products/          # Product performance data
│   └── mappings/          # SKU and campaign mappings
├── reports/                # Generated reports and charts
├── src/                    # Source code
│   ├── app.py             # Streamlit dashboard application
│   └── analyze_campaigns.py # Core analysis functionality
├── notebooks/             # Jupyter notebooks for analysis
├── requirements.txt       # Project dependencies
└── README.md             # Project documentation
```

## How to Run the Application

### Prerequisites

1. Python 3.8 or higher installed on your system
2. Git (for cloning the repository)

### Step-by-Step Setup

1. **Clone the Repository**

```bash
git clone https://github.com/yourusername/amazon-ad-analysis.git
cd amazon-ad-analysis
```

2. **Set Up Virtual Environment**

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

3. **Install Dependencies**

```bash
pip install -r requirements.txt
```

4. **Prepare Data Files**
   Make sure you have the following files in their respective directories:

- `data/campaign_reports/Sponsored_Products_Campaign_report_-_01-02_-15-03.xlsx`
- `data/campaign_reports/Sponsored_Products_Campaign_report-_hourly_(18th_to_2nd_march).csv`
- `data/search_terms/Sponsored_Products_Search_Term_Impression_Share_report - summary.csv`
- `data/search_terms/Sponsored_Products_Search_Term_Impression_Share_report_-Daily.csv`
- `data/products/Sponsored_Products_Advertised_product_report - SUMMARY.xlsx`
- `data/mappings/MSKUS_to_SKU_Amazon.xlsx`

5. **Run the Application**

There are two ways to use the application:

#### Option 1: Run the Interactive Dashboard

```bash
streamlit run src/app.py
```

This will:

- Start the Streamlit dashboard
- Open your default web browser to `http://localhost:8501`
- Show interactive visualizations and metrics
- Allow you to generate Excel reports from the dashboard

#### Option 2: Generate Excel Report Only

```bash
python src/analyze_campaigns.py
```

This will:

- Process all data files
- Generate a comprehensive Excel report
- Save the report in the `reports` directory with timestamp

### Using the Dashboard

1. **Access the Dashboard**

   - Open your web browser
   - Navigate to `http://localhost:8501`
   - The dashboard will load automatically

2. **Navigate Through Tabs**

   - Summary: Overview of key metrics
   - Campaign Performance: Detailed campaign analysis
   - Product Analysis: Product-level insights
   - Search Term Analysis: Search term performance

3. **Generate Reports**
   - Click the "Generate Excel Report" button in the sidebar
   - Wait for the report to generate
   - Download the report when ready

### Troubleshooting

If you encounter any issues:

1. **Data Loading Errors**

   - Verify all required data files are in the correct directories
   - Check file names match exactly
   - Ensure files are not corrupted

2. **Dependency Issues**

   - Make sure all requirements are installed correctly
   - Try reinstalling requirements: `pip install -r requirements.txt --upgrade`

3. **Dashboard Not Loading**

   - Check if port 8501 is available
   - Try running with a different port: `streamlit run src/app.py --server.port 8502`

4. **Report Generation Issues**
   - Ensure Excel is not open when generating reports
   - Check write permissions in the reports directory
   - Verify sufficient disk space

## Data Requirements

The dashboard expects the following data files:

1. Campaign Reports:

   - `Sponsored_Products_Campaign_report_-_01-02_-15-03.xlsx`
   - `Sponsored_Products_Campaign_report-_hourly_(18th_to_2nd_march).csv`

2. Search Term Data:

   - `Sponsored_Products_Search_Term_Impression_Share_report - summary.csv`
   - `Sponsored_Products_Search_Term_Impression_Share_report_-Daily.csv`

3. Product Data:

   - `Sponsored_Products_Advertised_product_report - SUMMARY.xlsx`

4. Mappings:
   - `MSKUS_to_SKU_Amazon.xlsx`

## Features in Detail

### Dashboard Interface

- Interactive metrics display
- Real-time data visualization
- Customizable date ranges
- Export functionality for reports

### Analysis Components

1. **Campaign Overview**

   - Total spend and sales
   - Key performance indicators
   - Campaign-level insights

2. **Product Analysis**

   - Performance by ASIN
   - Sales and spend metrics
   - ROI analysis

3. **Search Term Analysis**

   - Impression share tracking
   - Search term performance
   - Optimization opportunities

4. **Trend Analysis**
   - Daily performance tracking
   - Hourly patterns
   - Performance trends

### Report Generation

- Comprehensive Excel reports
- Interactive charts
- Detailed metrics and insights
- Professional formatting

## Dependencies

- Python 3.8+
- pandas
- numpy
- plotly
- streamlit
- xlsxwriter
- openpyxl
- seaborn
- matplotlib

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please open an issue in the GitHub repository or contact the maintainers.

## Acknowledgments

- Amazon Advertising API
- Streamlit team for the dashboard framework
- Contributors and maintainers
