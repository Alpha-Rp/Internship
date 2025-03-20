import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
from pathlib import Path

class CampaignAnalyzer:
    def __init__(self, data_dir):
        # Get the project root directory (2 levels up from src)
        self.project_root = Path(__file__).parent.parent
        self.data_dir = self.project_root / data_dir
        self.campaign_data = None
        self.hourly_data = None
        self.search_terms = None
        self.daily_search = None
        self.product_data = None
        self.sku_mapping = None
        
    def load_data(self):
        """Load all campaign data from various sources"""
        try:
            # Load SKU mapping data
            self.sku_mapping = pd.read_excel(
                self.data_dir / 'mappings' / 'MSKUS_to_SKU_Amazon.xlsx'
            )
            print("\nSKU Mapping columns:", self.sku_mapping.columns.tolist())
            
            # Rename columns in SKU mapping to match campaign data
            self.sku_mapping = self.sku_mapping.rename(columns={
                'msku': 'Campaign Name',  # This is the key column we'll use for merging
                'sku': 'SKU'
            })
            
            # Load campaign data
            self.campaign_data = pd.read_excel(
                self.data_dir / 'campaign_reports' / 'Sponsored_Products_Campaign_report_-_01-02_-15-03.xlsx'
            )
            print("\nCampaign data columns:", self.campaign_data.columns.tolist())
            
            # Load hourly data
            self.hourly_data = pd.read_csv(
                self.data_dir / 'campaign_reports' / 'Sponsored_Products_Campaign_report-_hourly_(18th_to_2nd_march).csv'
            )
            print("\nHourly data columns:", self.hourly_data.columns.tolist())
            
            # Load search term data
            self.search_terms = pd.read_csv(
                self.data_dir / 'search_terms' / 'Sponsored_Products_Search_Term_Impression_Share_report - summary.csv'
            )
            print("\nSearch terms columns:", self.search_terms.columns.tolist())
            
            self.daily_search = pd.read_csv(
                self.data_dir / 'search_terms' / 'Sponsored_Products_Search_Term_Impression_Share_report_-Daily.csv'
            )
            print("\nDaily search columns:", self.daily_search.columns.tolist())
            
            # Load product data
            self.product_data = pd.read_excel(
                self.data_dir / 'products' / 'Sponsored_Products_Advertised_product_report - SUMMARY.xlsx'
            )
            print("\nProduct data columns:", self.product_data.columns.tolist())
            
            # Now proceed with merging and data processing
            # Merge campaign data with SKU mapping
            self.campaign_data = self.campaign_data.merge(
                self.sku_mapping,
                left_on='Campaign Name',
                right_on='Campaign Name',
                how='left'
            )
            
            # Convert currency columns to numeric, removing currency symbols and commas
            currency_columns = ['Spend', '7 Day Total Sales (₹)', 'Total Return on Advertising Spend (ROAS)', 
                              'Total Advertising Cost of Sales (ACOS) ', 'Click-Thru Rate (CTR)']
            for col in currency_columns:
                if col in self.campaign_data.columns:
                    self.campaign_data[col] = pd.to_numeric(self.campaign_data[col].astype(str).str.replace('₹', '').str.replace(',', ''), errors='coerce')
            
            # Merge hourly data with SKU mapping
            self.hourly_data = self.hourly_data.merge(
                self.sku_mapping,
                left_on='Campaign Name',
                right_on='Campaign Name',
                how='left'
            )
            
            # Convert currency columns in hourly data
            currency_columns = ['Spend', '7 Day Total Sales (₹)']
            for col in currency_columns:
                if col in self.hourly_data.columns:
                    self.hourly_data[col] = pd.to_numeric(self.hourly_data[col].astype(str).str.replace('₹', '').str.replace(',', ''), errors='coerce')
            
            # Merge search terms with SKU mapping
            self.search_terms = self.search_terms.merge(
                self.sku_mapping,
                left_on='Campaign Name',
                right_on='Campaign Name',
                how='left'
            )
            
            # Convert numeric columns in search terms
            numeric_columns = ['Search Term Impression Share', 'Search Term Impression Rank']
            for col in numeric_columns:
                if col in self.search_terms.columns:
                    self.search_terms[col] = pd.to_numeric(self.search_terms[col], errors='coerce')
            
            # Merge daily search with SKU mapping
            self.daily_search = self.daily_search.merge(
                self.sku_mapping,
                left_on='Campaign Name',
                right_on='Campaign Name',
                how='left'
            )
            
            # Convert currency columns in product data
            currency_columns = ['Spend', '7 Day Total Sales (₹)']
            for col in currency_columns:
                if col in self.product_data.columns:
                    self.product_data[col] = pd.to_numeric(self.product_data[col].astype(str).str.replace('₹', '').str.replace(',', ''), errors='coerce')
            
            print("\nSuccessfully loaded all data files!")
            
        except FileNotFoundError as e:
            print(f"Error: Could not find data file: {e}")
            print("Please make sure all data files are in the correct directories:")
            print(f"Campaign reports: {self.data_dir / 'campaign_reports'}")
            print(f"Search terms: {self.data_dir / 'search_terms'}")
            print(f"Products: {self.data_dir / 'products'}")
            print(f"Mappings: {self.data_dir / 'mappings'}")
            raise
        except Exception as e:
            print(f"Error loading data: {e}")
            print("\nPlease check the column names in your data files.")
            raise

    def calculate_metrics(self):
        """Calculate key performance metrics"""
        try:
            # Calculate metrics using the correct column names from the data
            metrics = {
                'total_spend': float(self.campaign_data['Spend'].sum()),
                'total_sales': float(self.campaign_data['7 Day Total Sales (₹)'].sum()),
                'total_impressions': int(self.campaign_data['Impressions'].sum()),
                'total_clicks': int(self.campaign_data['Clicks'].sum()),
                'total_orders': int(self.campaign_data['7 Day Total Orders (#)'].sum()),
                'average_roi': float(self.campaign_data['Total Return on Advertising Spend (ROAS)'].mean()),
                'average_acos': float(self.campaign_data['Total Advertising Cost of Sales (ACOS) '].mean()),
                'average_ctr': float(self.campaign_data['Click-Thru Rate (CTR)'].mean()),
                'average_conversion_rate': float(self.campaign_data['7 Day Total Orders (#)'].sum() / self.campaign_data['Clicks'].sum())
            }
            return metrics
            
        except Exception as e:
            print(f"Error calculating metrics: {e}")
            print("Available columns in campaign data:", self.campaign_data.columns.tolist())
            raise
    
    def analyze_hourly_performance(self):
        """Analyze hourly performance trends"""
        if self.hourly_data is None:
            raise ValueError("Hourly data not loaded. Please run load_data() first.")
            
        # Convert Start Time to datetime with flexible format handling
        try:
            # First try the specific format
            self.hourly_data['Start Time'] = pd.to_datetime(self.hourly_data['Start Time'], format='%Y-%m-%d %H:%M:%S')
        except ValueError:
            try:
                # If that fails, try parsing as time only
                self.hourly_data['Start Time'] = pd.to_datetime(self.hourly_data['Start Time'], format='%H:%M')
            except ValueError:
                # If both fail, try inferring the format
                self.hourly_data['Start Time'] = pd.to_datetime(self.hourly_data['Start Time'])
        
        # Extract hour from the datetime
        self.hourly_data['hour'] = self.hourly_data['Start Time'].dt.hour
        
        # Group by hour and calculate metrics
        hourly_metrics = self.hourly_data.groupby('hour').agg({
            'Spend': 'sum',
            '7 Day Total Sales (₹)': 'sum',
            'Impressions': 'sum',
            'Clicks': 'sum',
            '7 Day Total Orders (#)': 'sum'
        }).reset_index()
        
        # Calculate hourly ROAS and conversion rate
        hourly_metrics['ROAS'] = (hourly_metrics['7 Day Total Sales (₹)'] / hourly_metrics['Spend']).round(2)
        hourly_metrics['Conversion Rate'] = (hourly_metrics['7 Day Total Orders (#)'] / hourly_metrics['Clicks']).round(3)
        
        # Sort by hour
        hourly_metrics = hourly_metrics.sort_values('hour')
        
        return hourly_metrics
    
    def analyze_search_terms(self):
        """Analyze search term performance"""
        top_terms = self.search_terms.nlargest(10, 'Search Term Impression Share')
        return top_terms
    
    def analyze_product_performance(self):
        """Analyze product-level performance"""
        product_metrics = self.product_data.groupby('Advertised ASIN').agg({
            'Spend': 'sum',
            '7 Day Total Sales (₹)': 'sum',
            'Impressions': 'sum',
            'Clicks': 'sum',
            '7 Day Total Orders (#)': 'sum'
        }).reset_index()
        
        product_metrics['roi'] = (product_metrics['7 Day Total Sales (₹)'] - product_metrics['Spend']) / product_metrics['Spend']
        product_metrics['acos'] = product_metrics['Spend'] / product_metrics['7 Day Total Sales (₹)']
        
        return product_metrics
    
    def generate_insights(self):
        """Generate insights and recommendations"""
        insights = {
            'high_impression_low_sales': [],
            'overspending': [],
            'low_conversion': [],
            'opportunities': []
        }
        
        # Calculate conversion rate for campaign data
        self.campaign_data['Conversion Rate'] = self.campaign_data['7 Day Total Orders (#)'] / self.campaign_data['Clicks']
        
        # Identify high impression, low sales campaigns
        high_impression = self.campaign_data[self.campaign_data['Impressions'] > self.campaign_data['Impressions'].mean()]
        low_sales = high_impression[high_impression['7 Day Total Sales (₹)'] < high_impression['7 Day Total Sales (₹)'].mean()]
        insights['high_impression_low_sales'] = low_sales['Campaign Name'].unique().tolist()
        
        # Identify overspending campaigns
        overspending = self.campaign_data[self.campaign_data['Total Advertising Cost of Sales (ACOS) '] > 1.0]
        insights['overspending'] = overspending['Campaign Name'].unique().tolist()
        
        # Identify low conversion campaigns
        low_conversion = self.campaign_data[self.campaign_data['Conversion Rate'] < self.campaign_data['Conversion Rate'].mean()]
        insights['low_conversion'] = low_conversion['Campaign Name'].unique().tolist()
        
        # Identify opportunities from search terms
        high_rank_low_share = self.search_terms[
            (self.search_terms['Search Term Impression Rank'] <= 3) & 
            (self.search_terms['Search Term Impression Share'] < 0.1)
        ]
        insights['opportunities'] = high_rank_low_share['Customer Search Term'].unique().tolist()
        
        return insights
    
    def analyze_campaign_performance(self):
        """Analyze overall campaign performance metrics"""
        if self.campaign_data is None:
            raise ValueError("Campaign data not loaded. Please run load_data() first.")
            
        # Calculate key metrics
        metrics = {
            'total_spend': self.campaign_data['Spend'].sum(),
            'total_impressions': self.campaign_data['Impressions'].sum(),
            'total_clicks': self.campaign_data['Clicks'].sum(),
            'total_sales': self.campaign_data['7 Day Total Sales (₹)'].sum(),
            'average_roas': self.campaign_data['Total Return on Advertising Spend (ROAS)'].mean(),
            'average_acos': self.campaign_data['Total Advertising Cost of Sales (ACOS) '].mean(),
            'average_ctr': self.campaign_data['Click-Thru Rate (CTR)'].mean()
        }
        
        # Campaign performance by SKU
        sku_performance = self.campaign_data.groupby('SKU').agg({
            'Spend': 'sum',
            'Impressions': 'sum',
            'Clicks': 'sum',
            '7 Day Total Sales (₹)': 'sum',
            'Total Return on Advertising Spend (ROAS)': 'mean',
            'Total Advertising Cost of Sales (ACOS) ': 'mean'
        }).round(2)
        
        return metrics, sku_performance

    def analyze_product_performance(self):
        """Analyze product-level performance and identify opportunities"""
        if self.product_data is None:
            raise ValueError("Product data not loaded. Please run load_data() first.")
            
        # Calculate product metrics
        product_metrics = self.product_data.groupby('Advertised ASIN').agg({
            'Spend': 'sum',
            'Impressions': 'sum',
            'Clicks': 'sum',
            '7 Day Total Sales (₹)': 'sum'
        }).round(2)
        
        # Calculate ROAS for each product
        product_metrics['ROAS'] = (product_metrics['7 Day Total Sales (₹)'] / product_metrics['Spend']).round(2)
        
        # Identify over-performing and under-performing products
        product_metrics['Performance'] = pd.cut(
            product_metrics['ROAS'],
            bins=[-np.inf, 1, 2, np.inf],
            labels=['Under-performing', 'Moderate', 'Over-performing']
        )
        
        return product_metrics

    def analyze_search_term_performance(self):
        """Analyze search term performance and identify optimization opportunities"""
        if self.search_terms is None:
            raise ValueError("Search term data not loaded. Please run load_data() first.")
            
        # Calculate search term metrics
        search_metrics = self.search_terms.groupby('Customer Search Term').agg({
            'Search Term Impression Share': 'mean',
            'Search Term Impression Rank': 'mean',
            'Clicks': 'sum',
            '7 Day Total Orders (#)': 'sum',
            '7 Day Total Sales (₹)': 'sum'
        }).round(2)
        
        # Identify optimization opportunities
        search_metrics['Opportunity'] = pd.cut(
            search_metrics['Search Term Impression Share'],
            bins=[0, 0.3, 0.7, 1],
            labels=['Low Share', 'Moderate Share', 'High Share']
        )
        
        return search_metrics

    def analyze_trends(self):
        """Analyze performance trends over time"""
        if self.hourly_data is None:
            raise ValueError("Hourly data not loaded. Please run load_data() first.")
            
        # Convert date column to datetime
        self.hourly_data['Start Date'] = pd.to_datetime(self.hourly_data['Start Date'])
        
        # Calculate daily metrics
        daily_metrics = self.hourly_data.groupby('Start Date').agg({
            'Spend': 'sum',
            'Impressions': 'sum',
            'Clicks': 'sum',
            '7 Day Total Sales (₹)': 'sum',
            '7 Day Total Orders (#)': 'sum'
        }).round(2)
        
        # Calculate daily ROAS and conversion rate
        daily_metrics['ROAS'] = (daily_metrics['7 Day Total Sales (₹)'] / daily_metrics['Spend']).round(2)
        daily_metrics['Conversion Rate'] = (daily_metrics['7 Day Total Orders (#)'] / daily_metrics['Clicks']).round(3)
        
        # Sort by date
        daily_metrics = daily_metrics.sort_values('Start Date')
        
        return daily_metrics

    def generate_report(self):
        """Generate a comprehensive report with insights and recommendations"""
        # Get all analysis results
        campaign_metrics, sku_performance = self.analyze_campaign_performance()
        product_metrics = self.analyze_product_performance()
        search_metrics = self.analyze_search_term_performance()
        daily_metrics = self.analyze_trends()
        
        # Create report dictionary
        report = {
            'overview': {
                'total_spend': campaign_metrics['total_spend'],
                'total_sales': campaign_metrics['total_sales'],
                'average_roas': campaign_metrics['average_roas'],
                'average_acos': campaign_metrics['average_acos']
            },
            'top_performing_products': product_metrics[product_metrics['Performance'] == 'Over-performing'].head(5),
            'underperforming_products': product_metrics[product_metrics['Performance'] == 'Under-performing'].head(5),
            'high_impression_keywords': search_metrics[search_metrics['Opportunity'] == 'High Share'].head(5),
            'trend_analysis': {
                'roas_trend': daily_metrics['ROAS'].mean(),
                'spend_trend': daily_metrics['Spend'].mean(),
                'sales_trend': daily_metrics['7 Day Total Sales (₹)'].mean()
            }
        }
        
        return report

    def export_to_excel(self, output_path):
        """Export analysis results to Excel for dashboard creation"""
        # Create Excel writer
        with pd.ExcelWriter(output_path) as writer:
            # Export campaign performance
            _, sku_performance = self.analyze_campaign_performance()
            sku_performance.to_excel(writer, sheet_name='Campaign Performance')
            
            # Export product performance
            product_metrics = self.analyze_product_performance()
            product_metrics.to_excel(writer, sheet_name='Product Performance')
            
            # Export search term performance
            search_metrics = self.analyze_search_term_performance()
            search_metrics.to_excel(writer, sheet_name='Search Term Performance')
            
            # Export daily trends
            daily_metrics = self.analyze_trends()
            daily_metrics.to_excel(writer, sheet_name='Daily Trends')
            
            # Export hourly performance
            hourly_metrics = self.analyze_hourly_performance()
            hourly_metrics.to_excel(writer, sheet_name='Hourly Performance')
            
            # Export insights
            insights = self.generate_insights()
            insights_df = pd.DataFrame({
                'Category': ['High Impression, Low Sales', 'Overspending', 'Low Conversion', 'Opportunities'],
                'Items': [
                    ', '.join(insights['high_impression_low_sales']),
                    ', '.join(insights['overspending']),
                    ', '.join(insights['low_conversion']),
                    ', '.join(insights['opportunities'])
                ]
            })
            insights_df.to_excel(writer, sheet_name='Insights', index=False)

def main():
    try:
        # Initialize analyzer
        analyzer = CampaignAnalyzer('data')
        
        # Load data
        analyzer.load_data()
        
        # Calculate metrics
        metrics = analyzer.calculate_metrics()
        print("\nOverall Campaign Metrics:")
        for key, value in metrics.items():
            print(f"{key}: {value:.2f}")
        
        # Generate report
        report = analyzer.generate_report()
        print("\nCampaign Analysis Report:")
        print(f"Total Spend: ₹{report['overview']['total_spend']:,.2f}")
        print(f"Total Sales: ₹{report['overview']['total_sales']:,.2f}")
        print(f"Average ROAS: {report['overview']['average_roas']:.2f}")
        print(f"Average ACOS: {report['overview']['average_acos']:.2f}")
        
        # Export to Excel for dashboard creation
        output_path = analyzer.project_root / 'reports' / 'campaign_analysis.xlsx'
        analyzer.export_to_excel(output_path)
        print(f"\nReport generated successfully and exported to: {output_path}")
        
    except Exception as e:
        print(f"Error running analysis: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main()) 