# Amazon Advertising Campaign Analysis

This project provides tools for analyzing Amazon advertising campaign performance data and generating actionable insights for optimization.

## Project Structure

```
amazon-ad-analysis/
├── data/
│   ├── mappings/           # Contains MSKU to SKU mapping files
│   ├── campaign_reports/   # Campaign performance reports
│   ├── search_terms/      # Search term performance data
│   └── products/          # Product-level performance data
├── src/                   # Source code
├── notebooks/            # Jupyter notebooks for analysis
├── reports/             # Generated reports and dashboards
└── venv/                # Python virtual environment
```

## Setup

1. Create and activate a virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Data Requirements

The following data files are required:

1. Campaign Reports:

   - `Sponsored_Products_Campaign_report_-_01-02_-15-03.xlsx`
   - `Sponsored_Products_Campaign_report-_hourly_(18th_to_2nd_march).csv`

2. Search Term Reports:

   - `Sponsored_Products_Search_Term_Impression_Share_report - summary.csv`
   - `Sponsored_Products_Search_Term_Impression_Share_report_-Daily.csv`

3. Product Reports:

   - `Sponsored_Products_Advertised_product_report - SUMMARY.xlsx`

4. Mapping Files:
   - `MSKUS_to_SKU_Amazon.xlsx` (in the mappings directory)

## Usage

1. Run the analysis script:

```bash
python src/analyze_campaigns.py
```

This will:

- Load and process all data files
- Generate a comprehensive report
- Export results to Excel for dashboard creation

2. The generated Excel file (`reports/campaign_analysis.xlsx`) contains:
   - Campaign Performance sheet
   - Product Performance sheet
   - Search Term Performance sheet
   - Daily Trends sheet

## Analysis Features

1. Campaign Performance Analysis:

   - Total spend, impressions, clicks, and sales
   - ROAS and ACOS metrics
   - Performance by SKU

2. Product-Level Analysis:

   - Performance metrics by ASIN
   - Identification of over/under-performing products
   - ROAS calculation

3. Search Term Analysis:

   - Impression share and rank
   - Click and order metrics
   - Optimization opportunities

4. Trend Analysis:
   - Daily performance metrics
   - ROAS trends
   - Spend and sales patterns

## Dashboard Creation

The exported Excel file can be used to create a dashboard using:

- Excel PivotTables and Charts
- Power BI
- Google Sheets

Recommended dashboard components:

1. Campaign Performance Overview
2. Product Performance Analysis
3. Search Term Optimization Opportunities
4. Daily Performance Trends

## Output Files

1. `campaign_analysis.xlsx`: Contains all analysis results in separate sheets
2. Console output: Summary of key metrics and insights

## Contributing

Feel free to submit issues and enhancement requests!
