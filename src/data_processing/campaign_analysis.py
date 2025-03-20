import pandas as pd
import numpy as np
from datetime import datetime
import os

def load_campaign_data(file_path):
    """
    Load and process campaign data from Excel file.
    
    Args:
        file_path (str): Path to the campaign data Excel file
        
    Returns:
        pd.DataFrame: Processed campaign data
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Clean column names
        df.columns = df.columns.str.strip().str.lower()
        
        # Convert date columns to datetime
        date_columns = [col for col in df.columns if 'date' in col or 'time' in col]
        for col in date_columns:
            df[col] = pd.to_datetime(df[col])
        
        # Calculate key metrics
        if 'spend' in df.columns and 'sales' in df.columns:
            df['roi'] = (df['sales'] - df['spend']) / df['spend']
            df['acos'] = df['spend'] / df['sales']
        
        if 'impressions' in df.columns and 'clicks' in df.columns:
            df['ctr'] = df['clicks'] / df['impressions']
        
        if 'clicks' in df.columns and 'orders' in df.columns:
            df['conversion_rate'] = df['orders'] / df['clicks']
        
        return df
    
    except Exception as e:
        print(f"Error loading campaign data: {str(e)}")
        return None

def analyze_campaign_performance(df):
    """
    Analyze campaign performance metrics.
    
    Args:
        df (pd.DataFrame): Campaign data DataFrame
        
    Returns:
        dict: Dictionary containing performance metrics
    """
    try:
        metrics = {
            'total_spend': df['spend'].sum(),
            'total_sales': df['sales'].sum(),
            'total_impressions': df['impressions'].sum(),
            'total_clicks': df['clicks'].sum(),
            'total_orders': df['orders'].sum(),
            'average_roi': df['roi'].mean(),
            'average_acos': df['acos'].mean(),
            'average_ctr': df['ctr'].mean(),
            'average_conversion_rate': df['conversion_rate'].mean()
        }
        
        # Campaign-level metrics
        campaign_metrics = df.groupby('campaign_name').agg({
            'spend': 'sum',
            'sales': 'sum',
            'impressions': 'sum',
            'clicks': 'sum',
            'orders': 'sum',
            'roi': 'mean',
            'acos': 'mean',
            'ctr': 'mean',
            'conversion_rate': 'mean'
        }).round(2)
        
        metrics['campaign_breakdown'] = campaign_metrics
        
        return metrics
    
    except Exception as e:
        print(f"Error analyzing campaign performance: {str(e)}")
        return None

def identify_optimization_opportunities(df):
    """
    Identify optimization opportunities in campaign data.
    
    Args:
        df (pd.DataFrame): Campaign data DataFrame
        
    Returns:
        dict: Dictionary containing optimization opportunities
    """
    try:
        opportunities = {
            'high_impression_low_sales': [],
            'overspending': [],
            'low_conversion': []
        }
        
        # High impression share but low sales
        high_impression = df[df['impressions'] > df['impressions'].mean()]
        low_sales = high_impression[high_impression['sales'] < high_impression['sales'].mean()]
        opportunities['high_impression_low_sales'] = low_sales['campaign_name'].unique().tolist()
        
        # Overspending campaigns
        overspending = df[df['acos'] > 1.0]  # ACOS > 100%
        opportunities['overspending'] = overspending['campaign_name'].unique().tolist()
        
        # Low conversion campaigns
        low_conversion = df[df['conversion_rate'] < df['conversion_rate'].mean()]
        opportunities['low_conversion'] = low_conversion['campaign_name'].unique().tolist()
        
        return opportunities
    
    except Exception as e:
        print(f"Error identifying optimization opportunities: {str(e)}")
        return None

if __name__ == "__main__":
    # Example usage
    campaign_file = os.path.join("data", "campaign_reports", "campaign_data.xlsx")
    df = load_campaign_data(campaign_file)
    
    if df is not None:
        metrics = analyze_campaign_performance(df)
        opportunities = identify_optimization_opportunities(df)
        
        print("\nCampaign Performance Metrics:")
        for key, value in metrics.items():
            if key != 'campaign_breakdown':
                print(f"{key}: {value}")
        
        print("\nOptimization Opportunities:")
        for key, value in opportunities.items():
            print(f"{key}: {value}") 