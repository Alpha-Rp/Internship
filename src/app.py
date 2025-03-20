import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from analyze_campaigns import CampaignAnalyzer
from datetime import datetime
import os
from pathlib import Path

def create_dashboard():
    st.set_page_config(page_title="Amazon Ad Campaign Analysis", layout="wide")
    
    # Initialize analyzer
    analyzer = CampaignAnalyzer('data')
    
    try:
        # Load data
        with st.spinner('Loading data...'):
            analyzer.load_data()
        
        # Add download button to sidebar
        with st.sidebar:
            st.title("Report Generation")
            if st.button("Generate Excel Report", type="primary"):
                with st.spinner('Generating report...'):
                    # Create reports directory if it doesn't exist
                    reports_dir = Path('reports')
                    reports_dir.mkdir(parents=True, exist_ok=True)
                    
                    # Generate filename with timestamp
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    output_path = reports_dir / f'campaign_analysis_{timestamp}.xlsx'
                    
                    # Generate the report
                    analyzer.export_to_excel(output_path)
                    
                    # Read the file and create download button
                    with open(output_path, 'rb') as f:
                        st.download_button(
                            label="Download Excel Report",
                            data=f,
                            file_name=f'campaign_analysis_{timestamp}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
        
        # Add title and description
        st.title("Amazon Ad Campaign Analysis Dashboard")
        st.markdown("""
        This dashboard provides insights into your Amazon advertising campaigns, including performance metrics,
        trends, and optimization opportunities.
        """)
        
        # Create tabs for different sections
        tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Campaign Performance", "Product Analysis", "Search Term Analysis"])
        
        with tab1:
            st.header("Campaign Summary")
            
            # Get campaign metrics
            metrics = analyzer.analyze_campaign_performance()
            
            # Create three columns for key metrics
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown("### Total Spend")
                st.markdown(f"<h2 style='font-size: 24px; color: #1f77b4;'>{metrics['total_spend']:,.2f}</h2>", unsafe_allow_html=True)
                st.markdown("### Total Sales")
                st.markdown(f"<h2 style='font-size: 24px; color: #2ca02c;'>{metrics['total_sales']:,.2f}</h2>", unsafe_allow_html=True)
                st.markdown("### Total Impressions")
                st.markdown(f"<h2 style='font-size: 24px; color: #ff7f0e;'>{metrics['total_impressions']:,.0f}</h2>", unsafe_allow_html=True)
            
            with col2:
                st.markdown("### Total Clicks")
                st.markdown(f"<h2 style='font-size: 24px; color: #d62728;'>{metrics['total_clicks']:,.0f}</h2>", unsafe_allow_html=True)
                st.markdown("### Total Orders")
                st.markdown(f"<h2 style='font-size: 24px; color: #9467bd;'>{metrics['total_orders']:,.0f}</h2>", unsafe_allow_html=True)
                st.markdown("### Average ROAS")
                st.markdown(f"<h2 style='font-size: 24px; color: #8c564b;'>{metrics['average_roas']:.2f}</h2>", unsafe_allow_html=True)
            
            with col3:
                st.markdown("### Average ACOS")
                st.markdown(f"<h2 style='font-size: 24px; color: #e377c2;'>{metrics['average_acos']:.2%}</h2>", unsafe_allow_html=True)
                st.markdown("### Average CTR")
                st.markdown(f"<h2 style='font-size: 24px; color: #7f7f7f;'>{metrics['average_ctr']:.2%}</h2>", unsafe_allow_html=True)
                st.markdown("### Conversion Rate")
                st.markdown(f"<h2 style='font-size: 24px; color: #bcbd22;'>{metrics['average_conversion_rate']:.2%}</h2>", unsafe_allow_html=True)
            
            # Add insights section
            st.subheader("Key Insights")
            insights = analyzer.generate_insights()
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**High Impression, Low Sales Campaigns**")
                for campaign in insights['high_impression_low_sales']:
                    st.write(f"- {campaign}")
                
                st.markdown("**Overspending Campaigns**")
                for campaign in insights['overspending']:
                    st.write(f"- {campaign}")
            
            with col2:
                st.markdown("**Low Conversion Campaigns**")
                for campaign in insights['low_conversion']:
                    st.write(f"- {campaign}")
                
                st.markdown("**Optimization Opportunities**")
                for term in insights['opportunities']:
                    st.write(f"- {term}")
        
        with tab2:
            st.header("Campaign Performance")
            
            # Get campaign data
            campaign_data = analyzer.campaign_data
            
            # Create campaign performance chart
            fig = px.bar(
                campaign_data,
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
                height=500,
                showlegend=True
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Display campaign metrics table
            st.subheader("Campaign Metrics")
            campaign_metrics = campaign_data[[
                'Campaign Name', 'Spend', 'Impressions', 'Clicks',
                '7 Day Total Sales (₹)', '7 Day Total Orders (#)',
                'Total Return on Advertising Spend (ROAS)',
                'Total Advertising Cost of Sales (ACOS) ',
                'Click-Thru Rate (CTR)'
            ]].round(2)
            st.dataframe(campaign_metrics)
        
        with tab3:
            st.header("Product Analysis")
            
            # Get product metrics
            product_metrics = analyzer.analyze_product_performance()
            
            # Create product performance scatter plot
            fig = px.scatter(
                product_metrics,
                x='Spend',
                y='7 Day Total Sales (₹)',
                size='Impressions',
                color='Performance',
                hover_data=['Advertised ASIN', 'Clicks', 'ROAS'],
                title="Product Performance Analysis",
                labels={
                    'Spend': 'Total Spend (₹)',
                    '7 Day Total Sales (₹)': 'Total Sales (₹)',
                    'Impressions': 'Number of Impressions',
                    'Performance': 'Performance Category',
                    'ROAS': 'ROAS'
                }
            )
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)
            
            # Display product metrics table
            st.subheader("Product Metrics")
            st.dataframe(product_metrics)
        
        with tab4:
            st.header("Search Term Analysis")
            
            # Get search term metrics
            search_metrics = analyzer.analyze_search_term_performance()
            
            # Create search term performance scatter plot
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
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)
            
            # Display search term metrics table
            st.subheader("Search Term Metrics")
            st.dataframe(search_metrics)
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.error("Please make sure all data files are in the correct directories.")

if __name__ == "__main__":
    create_dashboard() 