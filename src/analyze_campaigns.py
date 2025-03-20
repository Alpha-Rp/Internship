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
        
        # Reset index to make Advertised ASIN a column
        product_metrics = product_metrics.reset_index()
        
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
        
        # Clean up sales column - remove any duplicate currency symbols and convert to numeric
        search_metrics['7 Day Total Sales (₹)'] = search_metrics['7 Day Total Sales (₹)'].astype(str).str.replace('Rs.', '').str.replace('₹', '').str.replace(',', '')
        search_metrics['7 Day Total Sales (₹)'] = pd.to_numeric(search_metrics['7 Day Total Sales (₹)'], errors='coerce')
        
        # Identify optimization opportunities based on multiple factors
        search_metrics['Opportunity'] = pd.cut(
            search_metrics['Search Term Impression Share'],
            bins=[0, 0.3, 0.7, 1],
            labels=['Low Share', 'Moderate Share', 'High Share']
        )
        
        # Add opportunity details
        search_metrics['Opportunity Details'] = ''
        for idx, row in search_metrics.iterrows():
            details = []
            if row['Search Term Impression Share'] < 0.3:
                details.append('Low impression share')
            if row['Search Term Impression Rank'] > 3:
                details.append('Poor ranking')
            if row['Clicks'] > 0 and row['7 Day Total Sales (₹)'] == 0:
                details.append('No sales despite clicks')
            search_metrics.loc[idx, 'Opportunity Details'] = ' | '.join(details) if details else 'No issues'
        
        # Reset index to make Customer Search Term a column
        search_metrics = search_metrics.reset_index()
        
        return search_metrics

    def analyze_trends(self):
        """Analyze performance trends over time"""
        if self.hourly_data is None:
            raise ValueError("Hourly data not loaded. Please run load_data() first.")
            
        # Convert date column to datetime and extract only the date part
        self.hourly_data['Start Date'] = pd.to_datetime(self.hourly_data['Start Date']).dt.date
        
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
        
        # Reset index to make Start Date a column
        daily_metrics = daily_metrics.reset_index()
        
        return daily_metrics

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
            'total_orders': self.campaign_data['7 Day Total Orders (#)'].sum(),
            'average_roas': self.campaign_data['Total Return on Advertising Spend (ROAS)'].mean(),
            'average_acos': self.campaign_data['Total Advertising Cost of Sales (ACOS) '].mean(),
            'average_ctr': self.campaign_data['Click-Thru Rate (CTR)'].mean(),
            'average_conversion_rate': (self.campaign_data['7 Day Total Orders (#)'].sum() / self.campaign_data['Clicks'].sum())
        }
        
        return metrics

    def generate_report(self):
        """Generate a comprehensive report with insights and recommendations"""
        # Get all analysis results
        campaign_metrics = self.analyze_campaign_performance()
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

    def create_charts(self):
        """Create visualizations for the analysis"""
        charts = {}
        
        # 1. Campaign Performance Chart (Bar Chart)
        fig = px.bar(
            self.campaign_data,
            x='Campaign Name',
            y=['Spend', '7 Day Total Sales (₹)'],
            title="Campaign Spend vs Sales",
            barmode='group',
            labels={
                'Campaign Name': 'Campaign',
                'value': 'Amount (₹)',
                'variable': 'Metric'
            }
        )
        fig.update_layout(
            xaxis_tickangle=-45,
            height=480,
            width=720,
            showlegend=True,
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(t=50, b=50, l=50, r=50)
        )
        charts['campaign_performance'] = fig
        
        # 2. Product Performance Chart (Scatter Plot)
        product_metrics = self.analyze_product_performance()
        fig = px.scatter(
            product_metrics,
            x='Spend',
            y='7 Day Total Sales (₹)',
            size='Impressions',
            color='ROAS',
            hover_data=['Advertised ASIN', 'Clicks', '7 Day Total Orders (#)'],
            title="Product Performance Analysis",
            labels={
                'Spend': 'Total Spend (₹)',
                '7 Day Total Sales (₹)': 'Total Sales (₹)',
                'Impressions': 'Number of Impressions',
                'ROAS': 'ROAS',
                'Clicks': 'Number of Clicks',
                '7 Day Total Orders (#)': 'Number of Orders'
            }
        )
        fig.update_layout(
            height=480,
            width=720,
            showlegend=True,
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(t=50, b=50, l=50, r=50)
        )
        charts['product_performance'] = fig
        
        # 3. Daily Trends Chart (Line Chart)
        daily_metrics = self.analyze_trends()
        fig = go.Figure()
        
        # Add Spend trace
        fig.add_trace(go.Scatter(
            x=daily_metrics['Start Date'],
            y=daily_metrics['Spend'],
            name='Spend',
            line=dict(color='#1f77b4', width=2)
        ))
        
        # Add Sales trace
        fig.add_trace(go.Scatter(
            x=daily_metrics['Start Date'],
            y=daily_metrics['7 Day Total Sales (₹)'],
            name='Sales',
            line=dict(color='#2ca02c', width=2)
        ))
        
        fig.update_layout(
            title="Daily Performance Trends",
            xaxis_title="Date",
            yaxis_title="Amount (₹)",
            height=480,
            width=720,
            showlegend=True,
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(t=50, b=50, l=50, r=50)
        )
        charts['daily_trends'] = fig
        
        # 4. Search Term Performance Chart (Scatter Plot)
        search_metrics = self.analyze_search_term_performance()
        fig = px.scatter(
            search_metrics,
            x='Search Term Impression Share',
            y='7 Day Total Sales (₹)',
            size='Clicks',
            color='Search Term Impression Rank',
            hover_data=['Customer Search Term', 'Clicks', '7 Day Total Orders (#)'],
            title="Search Term Performance Analysis",
            labels={
                'Search Term Impression Share': 'Impression Share',
                '7 Day Total Sales (₹)': 'Sales (₹)',
                'Clicks': 'Number of Clicks',
                'Search Term Impression Rank': 'Rank',
                '7 Day Total Orders (#)': 'Orders'
            }
        )
        fig.update_layout(
            height=480,
            width=720,
            showlegend=True,
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(t=50, b=50, l=50, r=50)
        )
        charts['search_term_performance'] = fig
        
        return charts

    def export_to_excel(self, output_path):
        """Export analysis results to Excel for dashboard creation"""
        try:
            # Create Excel writer
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                # Get the workbook and create formats
                workbook = writer.book
                
                # Define formats
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#D9E1F2',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                currency_format = workbook.add_format({
                    'num_format': '₹#,##0.00',
                    'border': 1
                })
                
                number_format = workbook.add_format({
                    'num_format': '#,##0.00',
                    'border': 1
                })
                
                percent_format = workbook.add_format({
                    'num_format': '0.00%',
                    'border': 1
                })
                
                date_format = workbook.add_format({
                    'num_format': 'dd/mm/yyyy',
                    'border': 1
                })
                
                # Create Summary Sheet
                summary_sheet = workbook.add_worksheet('Summary')
                
                # Get campaign metrics
                metrics = self.analyze_campaign_performance()
                
                # Write summary header
                summary_sheet.merge_range('A1:D1', 'Campaign Performance Summary', header_format)
                summary_sheet.set_column('A:D', 20)
                
                # Write key metrics
                summary_sheet.write('A3', 'Metric', header_format)
                summary_sheet.write('B3', 'Value', header_format)
                
                metrics_to_show = [
                    ('Total Spend', metrics['total_spend'], currency_format),
                    ('Total Sales', metrics['total_sales'], currency_format),
                    ('Total Impressions', metrics['total_impressions'], number_format),
                    ('Total Clicks', metrics['total_clicks'], number_format),
                    ('Total Orders', metrics['total_orders'], number_format),
                    ('Average ROAS', metrics['average_roas'], number_format),
                    ('Average ACOS', metrics['average_acos'], percent_format),
                    ('Average CTR', metrics['average_ctr'], percent_format),
                    ('Conversion Rate', metrics['average_conversion_rate'], percent_format)
                ]
                
                for i, (label, value, fmt) in enumerate(metrics_to_show, start=4):
                    summary_sheet.write(f'A{i}', label, header_format)
                    summary_sheet.write(f'B{i}', value, fmt)
                
                # Add insights section
                insights = self.generate_insights()
                summary_sheet.write('A12', 'Key Insights', header_format)
                summary_sheet.write('A13', 'High Impression, Low Sales Campaigns:', header_format)
                summary_sheet.write('A14', ', '.join(insights['high_impression_low_sales']))
                summary_sheet.write('A15', 'Overspending Campaigns:', header_format)
                summary_sheet.write('A16', ', '.join(insights['overspending']))
                summary_sheet.write('A17', 'Low Conversion Campaigns:', header_format)
                summary_sheet.write('A18', ', '.join(insights['low_conversion']))
                summary_sheet.write('A19', 'Optimization Opportunities:', header_format)
                summary_sheet.write('A20', ', '.join(insights['opportunities']))
                
                # Export campaign performance
                metrics_df = pd.DataFrame([metrics])
                metrics_df.to_excel(writer, sheet_name='Campaign Performance', index=False)
                
                # Get the worksheet and apply formatting
                worksheet = writer.sheets['Campaign Performance']
                for col_num, value in enumerate(metrics_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    if 'spend' in value.lower() or 'sales' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, currency_format)
                    elif 'rate' in value.lower() or 'roas' in value.lower() or 'acos' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, percent_format)
                    else:
                        worksheet.set_column(col_num, col_num, 15, number_format)
                
                # Export product performance
                product_metrics = self.analyze_product_performance()
                product_metrics.to_excel(writer, sheet_name='Product Performance', index=False)
                worksheet = writer.sheets['Product Performance']
                
                # Apply formatting to product performance
                for col_num, value in enumerate(product_metrics.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    if 'spend' in value.lower() or 'sales' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, currency_format)
                    elif 'rate' in value.lower() or 'roas' in value.lower() or 'acos' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, percent_format)
                    else:
                        worksheet.set_column(col_num, col_num, 15, number_format)
                
                # Export search term performance
                search_metrics = self.analyze_search_term_performance()
                search_metrics.to_excel(writer, sheet_name='Search Term Performance', index=False)
                worksheet = writer.sheets['Search Term Performance']
                
                # Apply formatting to search term performance
                for col_num, value in enumerate(search_metrics.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    if 'spend' in value.lower() or 'sales' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, currency_format)
                    elif 'rate' in value.lower() or 'roas' in value.lower() or 'acos' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, percent_format)
                    else:
                        worksheet.set_column(col_num, col_num, 15, number_format)
                
                # Export daily trends
                daily_metrics = self.analyze_trends()
                daily_metrics.to_excel(writer, sheet_name='Daily Trends', index=False)
                worksheet = writer.sheets['Daily Trends']
                
                # Apply formatting to daily trends
                for col_num, value in enumerate(daily_metrics.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    if value == 'Start Date':
                        worksheet.set_column(col_num, col_num, 12, date_format)
                    elif 'spend' in value.lower() or 'sales' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, currency_format)
                    elif 'rate' in value.lower() or 'roas' in value.lower() or 'acos' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, percent_format)
                    else:
                        worksheet.set_column(col_num, col_num, 15, number_format)
                
                # Export hourly performance
                hourly_metrics = self.analyze_hourly_performance()
                hourly_metrics.to_excel(writer, sheet_name='Hourly Performance', index=False)
                worksheet = writer.sheets['Hourly Performance']
                
                # Apply formatting to hourly performance
                for col_num, value in enumerate(hourly_metrics.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    if 'spend' in value.lower() or 'sales' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, currency_format)
                    elif 'rate' in value.lower() or 'roas' in value.lower() or 'acos' in value.lower():
                        worksheet.set_column(col_num, col_num, 15, percent_format)
                    else:
                        worksheet.set_column(col_num, col_num, 15, number_format)
                
                # Export insights
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
                worksheet = writer.sheets['Insights']
                
                # Apply formatting to insights
                for col_num, value in enumerate(insights_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 50)
                
                # Create charts sheet
                chart_sheet = workbook.add_worksheet('Charts')
                
                # Campaign Performance Chart
                campaign_chart = workbook.add_chart({'type': 'column'})
                
                # Add spend series
                campaign_chart.add_series({
                    'name': 'Spend',
                    'categories': ['Campaign Performance', 1, 0, len(self.campaign_data), 0],
                    'values': ['Campaign Performance', 1, 1, len(self.campaign_data), 1],
                    'data_labels': {'value': True, 'num_format': '₹#,##0.00'}
                })
                
                # Add sales series
                campaign_chart.add_series({
                    'name': 'Sales',
                    'categories': ['Campaign Performance', 1, 0, len(self.campaign_data), 0],
                    'values': ['Campaign Performance', 1, 2, len(self.campaign_data), 2],
                    'data_labels': {'value': True, 'num_format': '₹#,##0.00'}
                })
                
                # Set chart properties
                campaign_chart.set_title({
                    'name': 'Campaign Spend vs Sales',
                    'font': {'size': 14, 'bold': True}
                })
                campaign_chart.set_x_axis({
                    'name': 'Campaign',
                    'name_font': {'size': 12, 'bold': True},
                    'num_font': {'rotation': -45}
                })
                campaign_chart.set_y_axis({
                    'name': 'Amount (₹)',
                    'name_font': {'size': 12, 'bold': True},
                    'num_format': '₹#,##0.00'
                })
                campaign_chart.set_legend({'position': 'top'})
                campaign_chart.set_size({'width': 720, 'height': 480})
                
                # Insert the chart
                chart_sheet.insert_chart('B2', campaign_chart)
                
                # Product Performance Chart
                product_chart = workbook.add_chart({'type': 'scatter'})
                
                # Add product performance series
                product_chart.add_series({
                    'name': 'Product Performance',
                    'categories': ['Product Performance', 1, 1, len(product_metrics), 1],  # Spend
                    'values': ['Product Performance', 1, 3, len(product_metrics), 3],     # Sales
                    'data_labels': {'value': True, 'num_format': '₹#,##0.00'}
                })
                
                # Set chart properties
                product_chart.set_title({
                    'name': 'Product Performance Analysis',
                    'font': {'size': 14, 'bold': True}
                })
                product_chart.set_x_axis({
                    'name': 'Spend (₹)',
                    'name_font': {'size': 12, 'bold': True},
                    'num_format': '₹#,##0.00'
                })
                product_chart.set_y_axis({
                    'name': 'Sales (₹)',
                    'name_font': {'size': 12, 'bold': True},
                    'num_format': '₹#,##0.00'
                })
                product_chart.set_legend({'position': 'top'})
                product_chart.set_size({'width': 720, 'height': 480})
                
                # Insert the chart
                chart_sheet.insert_chart('B30', product_chart)
                
                # Search Term Performance Chart
                search_chart = workbook.add_chart({'type': 'scatter'})
                
                # Add search term performance series
                search_chart.add_series({
                    'name': 'Search Term Performance',
                    'categories': ['Search Term Performance', 1, 1, len(search_metrics), 1],  # Impression Share
                    'values': ['Search Term Performance', 1, 4, len(search_metrics), 4],     # Sales
                    'data_labels': {'value': True, 'num_format': '₹#,##0.00'}
                })
                
                # Set chart properties
                search_chart.set_title({
                    'name': 'Search Term Performance Analysis',
                    'font': {'size': 14, 'bold': True}
                })
                search_chart.set_x_axis({
                    'name': 'Impression Share',
                    'name_font': {'size': 12, 'bold': True},
                    'num_format': '0.00%'
                })
                search_chart.set_y_axis({
                    'name': 'Sales (₹)',
                    'name_font': {'size': 12, 'bold': True},
                    'num_format': '₹#,##0.00'
                })
                search_chart.set_legend({'position': 'top'})
                search_chart.set_size({'width': 720, 'height': 480})
                
                # Insert the chart
                chart_sheet.insert_chart('B58', search_chart)
                
                # Daily Trends Chart
                trends_chart = workbook.add_chart({'type': 'line'})
                
                # Add ROAS series
                trends_chart.add_series({
                    'name': 'ROAS',
                    'categories': ['Daily Trends', 1, 0, len(daily_metrics), 0],
                    'values': ['Daily Trends', 1, 5, len(daily_metrics), 5],
                    'data_labels': {'value': True, 'num_format': '0.00'}
                })
                
                # Add conversion rate series
                trends_chart.add_series({
                    'name': 'Conversion Rate',
                    'categories': ['Daily Trends', 1, 0, len(daily_metrics), 0],
                    'values': ['Daily Trends', 1, 6, len(daily_metrics), 6],
                    'data_labels': {'value': True, 'num_format': '0.00%'}
                })
                
                # Set chart properties
                trends_chart.set_title({
                    'name': 'Daily Performance Trends',
                    'font': {'size': 14, 'bold': True}
                })
                trends_chart.set_x_axis({
                    'name': 'Date',
                    'name_font': {'size': 12, 'bold': True},
                    'num_format': 'dd/mm/yyyy'
                })
                trends_chart.set_y_axis({
                    'name': 'Rate',
                    'name_font': {'size': 12, 'bold': True}
                })
                trends_chart.set_legend({'position': 'top'})
                trends_chart.set_size({'width': 720, 'height': 480})
                
                # Insert the chart
                chart_sheet.insert_chart('B86', trends_chart)
                
                print(f"\nReport generated successfully and exported to: {output_path}")
                
        except Exception as e:
            print(f"\nError exporting to Excel: {str(e)}")
            print("Please make sure you have xlsxwriter installed:")
            print("pip install xlsxwriter")
            raise

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
        
        # Create reports directory if it doesn't exist
        reports_dir = analyzer.project_root / 'reports'
        reports_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = reports_dir / f'campaign_analysis_{timestamp}.xlsx'
        
        # Try to export to Excel
        try:
            analyzer.export_to_excel(output_path)
            print(f"\nReport generated successfully and exported to: {output_path}")
        except PermissionError:
            print(f"\nError: Could not write to {output_path}")
            print("Please make sure the file is not open in Excel and you have write permissions.")
            print("Try closing any open Excel files and running the script again.")
            return 1
        except Exception as e:
            print(f"\nError exporting to Excel: {e}")
            return 1
        
    except Exception as e:
        print(f"Error running analysis: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main()) 