import dash
from dash import html, dcc
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import os

def create_dashboard(campaign_data, metrics, opportunities):
    """
    Create an interactive dashboard using Dash.
    
    Args:
        campaign_data (pd.DataFrame): Processed campaign data
        metrics (dict): Campaign performance metrics
        opportunities (dict): Optimization opportunities
    """
    app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
    
    # Create layout
    app.layout = dbc.Container([
        html.H1("Amazon Advertising Campaign Dashboard", className="text-center my-4"),
        
        # Key Metrics Cards
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("Total Spend"),
                        html.H2(f"${metrics['total_spend']:,.2f}")
                    ])
                ])
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("Total Sales"),
                        html.H2(f"${metrics['total_sales']:,.2f}")
                    ])
                ])
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("Average ROI"),
                        html.H2(f"{metrics['average_roi']:.2%}")
                    ])
                ])
            ], width=3),
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        html.H4("Average ACOS"),
                        html.H2(f"{metrics['average_acos']:.2%}")
                    ])
                ])
            ], width=3)
        ], className="mb-4"),
        
        # Campaign Performance Chart
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("Campaign Performance"),
                    dbc.CardBody([
                        dcc.Graph(
                            figure=px.bar(
                                campaign_data,
                                x='campaign_name',
                                y=['spend', 'sales'],
                                title="Campaign Spend vs Sales",
                                barmode='group'
                            )
                        )
                    ])
                ])
            ], width=12)
        ], className="mb-4"),
        
        # Optimization Opportunities
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("Optimization Opportunities"),
                    dbc.CardBody([
                        html.H5("High Impression, Low Sales Campaigns:"),
                        html.Ul([html.Li(campaign) for campaign in opportunities['high_impression_low_sales']]),
                        
                        html.H5("Overspending Campaigns:"),
                        html.Ul([html.Li(campaign) for campaign in opportunities['overspending']]),
                        
                        html.H5("Low Conversion Campaigns:"),
                        html.Ul([html.Li(campaign) for campaign in opportunities['low_conversion']])
                    ])
                ])
            ], width=12)
        ], className="mb-4"),
        
        # Campaign Metrics Table
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader("Campaign Metrics"),
                    dbc.CardBody([
                        html.Div([
                            html.Table([
                                html.Thead([
                                    html.Tr([html.Th(col) for col in metrics['campaign_breakdown'].columns])
                                ]),
                                html.Tbody([
                                    html.Tr([
                                        html.Td(metrics['campaign_breakdown'].iloc[i][col])
                                        for col in metrics['campaign_breakdown'].columns
                                    ]) for i in range(len(metrics['campaign_breakdown']))
                                ])
                            ], className="table table-striped")
                        ])
                    ])
                ])
            ], width=12)
        ])
    ], fluid=True)
    
    return app

if __name__ == "__main__":
    # Example usage
    from data_processing.campaign_analysis import load_campaign_data, analyze_campaign_performance, identify_optimization_opportunities
    
    campaign_file = os.path.join("data", "campaign_reports", "campaign_data.xlsx")
    df = load_campaign_data(campaign_file)
    
    if df is not None:
        metrics = analyze_campaign_performance(df)
        opportunities = identify_optimization_opportunities(df)
        
        app = create_dashboard(df, metrics, opportunities)
        app.run_server(debug=True) 