"""
Visualizations Module for PMO Pulse Dashboard

Creates interactive charts and visualizations for the dashboard.
"""

import plotly.express as px
import plotly.graph_objects as go
import plotly.figure_factory as ff
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Arcadis Brand Colors
ARCADIS_ORANGE = "#E67300"
ARCADIS_BLACK = "#000000"
ARCADIS_GREY = "#6c757d"
ARCADIS_DARK_GREY = "#646469"
ARCADIS_WHITE = "#FFFFFF"
ARCADIS_LIGHT_GREY = "#F5F5F5"
ARCADIS_TEAL = "#00A3A1"

def create_gauge_chart(value, title, color=ARCADIS_ORANGE):
    """
    Create a gauge chart for KPI visualization.
    
    Parameters:
    -----------
    value : float
        The value to display in the gauge (typically SPI or CPI)
    title : str
        The title for the gauge
    color : str
        The color for the gauge indicator
        
    Returns:
    --------
    plotly.graph_objects.Figure
        The configured gauge chart
    """
    # Determine gauge color bands based on value ranges
    if title == "SPI" or title == "CPI":
        steps = [
            {'range': [0, 0.7], 'color': "#FADBD8"},  # Red zone
            {'range': [0.7, 0.9], 'color': "#FEF9E7"},  # Yellow zone
            {'range': [0.9, 1.1], 'color': "#D5F5E3"},  # Green zone
            {'range': [1.1, 2], 'color': "#D5F5E3"}  # Green zone
        ]
    else:
        steps = [
            {'range': [0, 0.6], 'color': "#FADBD8"},
            {'range': [0.6, 0.8], 'color': "#FEF9E7"},
            {'range': [0.8, 2], 'color': "#D5F5E3"}
        ]
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        title={'text': title, 'font': {'size': 24}},
        number={'suffix': "", 'font': {'size': 26}},
        gauge={
            'axis': {'range': [0, 2], 'tickwidth': 1, 'tickcolor': "darkblue"},
            'bar': {'color': color},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': steps,
            'threshold': {
                'line': {'color': "black", 'width': 4},
                'thickness': 0.75,
                'value': 1.0
            }
        }
    ))
    
    fig.update_layout(
        height=250,
        margin=dict(l=20, r=20, t=30, b=20),
        paper_bgcolor="white",
        font={'color': "darkblue", 'family': "Arial"}
    )
    
    return fig

def create_trend_chart(data):
    """
    Create a trend chart for SPI and CPI over time.
    
    Parameters:
    -----------
    data : pandas.DataFrame
        DataFrame containing 'month', 'spi', and 'cpi' columns
        
    Returns:
    --------
    plotly.graph_objects.Figure
        The configured trend chart
    """
    fig = go.Figure()
    
    # Add SPI line
    fig.add_trace(go.Scatter(
        x=data['month_dt'] if 'month_dt' in data.columns else data['month'],
        y=data['spi'],
        mode='lines+markers',
        name='SPI',
        line=dict(color=ARCADIS_TEAL, width=3),
        marker=dict(size=8)
    ))
    
    # Add CPI line
    fig.add_trace(go.Scatter(
        x=data['month_dt'] if 'month_dt' in data.columns else data['month'],
        y=data['cpi'],
        mode='lines+markers',
        name='CPI',
        line=dict(color=ARCADIS_ORANGE, width=3),
        marker=dict(size=8)
    ))
    
    # Add threshold line at 1.0
    fig.add_shape(
        type="line",
        x0=data['month_dt'].iloc[0] if 'month_dt' in data.columns else data['month'].iloc[0],
        y0=1.0,
        x1=data['month_dt'].iloc[-1] if 'month_dt' in data.columns else data['month'].iloc[-1],
        y1=1.0,
        line=dict(color="rgba(0, 0, 0, 0.3)", width=1, dash="dash")
    )
    
    # Update layout
    fig.update_layout(
        title="SPI & CPI Trends",
        xaxis_title="Month",
        yaxis_title="Index Value",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        yaxis=dict(range=[0.7, 1.3]),
        plot_bgcolor='rgba(255, 255, 255, 1)',
        paper_bgcolor='rgba(255, 255, 255, 1)',
        margin=dict(l=20, r=20, t=50, b=50)
    )
    
    return fig

def create_portfolio_overview(projects_df, group_by='status'):
    """
    Create a portfolio overview chart grouped by a specific column.
    
    Parameters:
    -----------
    projects_df : pandas.DataFrame
        DataFrame containing project information
    group_by : str
        Column to group projects by ('status', 'sector', etc.)
        
    Returns:
    --------
    plotly.graph_objects.Figure
        The configured chart
    """
    # Group projects by the specified column and count
    grouped = projects_df.groupby(group_by).size().reset_index(name='count')
    
    # Create a color map based on the group_by column
    if group_by == 'status':
        color_map = {
            'On Track': '#2ECC71',
            'Minor Issues': '#F1C40F',
            'At Risk': '#E74C3C',
            'Delayed': '#8E44AD',
            'Completed': '#3498DB'
        }
        title = "Projects by Status"
    elif group_by == 'sector':
        # Use a colorscale for sectors
        sectors = grouped[group_by].unique()
        colors = px.colors.qualitative.Plotly[:len(sectors)]
        color_map = {sector: color for sector, color in zip(sectors, colors)}
        title = "Projects by Sector"
    else:
        color_map = None
        title = f"Projects by {group_by.replace('_', ' ').title()}"
    
    # Create the chart based on count
    fig = px.bar(
        grouped,
        x=group_by,
        y='count',
        color=group_by,
        color_discrete_map=color_map,
        title=title,
        text='count'
    )
    
    # Update layout
    fig.update_layout(
        xaxis_title=group_by.replace('_', ' ').title(),
        yaxis_title="Number of Projects",
        plot_bgcolor='rgba(255, 255, 255, 1)',
        paper_bgcolor='rgba(255, 255, 255, 1)',
        margin=dict(l=20, r=20, t=50, b=80),
        height=400
    )
    
    # Update text position
    fig.update_traces(textposition='outside')
    
    return fig

def create_risk_matrix(risks_df):
    """
    Create a risk matrix visualization.
    
    Parameters:
    -----------
    risks_df : pandas.DataFrame
        DataFrame containing risk information
        
    Returns:
    --------
    plotly.graph_objects.Figure
        The configured risk matrix
    """
    fig = go.Figure()
    
    # Add scatter plot for risks
    fig.add_trace(go.Scatter(
        x=risks_df['probability'],
        y=risks_df['impact_cost'] / risks_df['impact_cost'].max(),  # Normalize impact to 0-1
        mode='markers',
        marker=dict(
            size=15,
            color=risks_df['risk_score'],
            colorscale=[
                [0, '#3498DB'],    # Low risk
                [0.3, '#F1C40F'],  # Medium risk
                [0.6, '#E74C3C']   # High risk
            ],
            colorbar=dict(
                title="Risk Score",
                tickvals=[0, 0.3, 0.6, 1.0],
                ticktext=["Low", "Medium", "High", "Critical"]
            ),
            line=dict(width=1, color='darkgrey')
        ),
        text=risks_df['risk_name'],
        hovertemplate=
        '<b>%{text}</b><br>' +
        'Probability: %{x:.0%}<br>' +
        'Impact: $%{customdata:,.0f}<br>' +
        'Risk Score: %{marker.color:.2f}' +
        '<extra></extra>',
        customdata=risks_df['impact_cost']
    ))
    
    # Add risk zones (colored rectangles)
    # High risk zone (red)
    fig.add_shape(
        type="rect",
        x0=0.5, y0=0.5,
        x1=1.0, y1=1.0,
        fillcolor="rgba(231, 76, 60, 0.2)",
        line=dict(width=0),
        layer="below"
    )
    
    # Medium risk zone (yellow)
    fig.add_shape(
        type="rect",
        x0=0.2, y0=0.2,
        x1=0.5, y1=0.5,
        fillcolor="rgba(241, 196, 15, 0.2)",
        line=dict(width=0),
        layer="below"
    )
    
    # Medium-high risk zones
    fig.add_shape(
        type="rect",
        x0=0.2, y0=0.5,
        x1=0.5, y1=1.0,
        fillcolor="rgba(230, 126, 34, 0.2)",
        line=dict(width=0),
        layer="below"
    )
    
    fig.add_shape(
        type="rect",
        x0=0.5, y0=0.2,
        x1=1.0, y1=0.5,
        fillcolor="rgba(230, 126, 34, 0.2)",
        line=dict(width=0),
        layer="below"
    )
    
    # Low risk zone (blue)
    fig.add_shape(
        type="rect",
        x0=0, y0=0,
        x1=0.2, y1=0.2,
        fillcolor="rgba(52, 152, 219, 0.2)",
        line=dict(width=0),
        layer="below"
    )
    
    # Low-medium risk zones
    fig.add_shape(
        type="rect",
        x0=0, y0=0.2,
        x1=0.2, y1=1.0,
        fillcolor="rgba(41, 128, 185, 0.2)",
        line=dict(width=0),
        layer="below"
    )
    
    fig.add_shape(
        type="rect",
        x0=0.2, y0=0,
        x1=1.0, y1=0.2,
        fillcolor="rgba(41, 128, 185, 0.2)",
        line=dict(width=0),
        layer="below"
    )
    
    # Update layout
    fig.update_layout(
        title="Risk Matrix",
        xaxis=dict(
            title="Probability",
            tickformat=".0%",
            range=[0, 1]
        ),
        yaxis=dict(
            title="Impact (Normalized)",
            tickformat=".0%",
            range=[0, 1]
        ),
        plot_bgcolor='rgba(255, 255, 255, 1)',
        paper_bgcolor='rgba(255, 255, 255, 1)',
        margin=dict(l=20, r=20, t=50, b=50),
        height=500
    )
    
    return fig

def create_schedule_gantt(tasks_df, project_df):
    """
    Create a Gantt chart for project schedule visualization.
    
    Parameters:
    -----------
    tasks_df : pandas.DataFrame
        DataFrame containing task information
    project_df : pandas.Series
        Series containing project information
        
    Returns:
    --------
    plotly.graph_objects.Figure
        The configured Gantt chart
    """
    # Sort tasks by start date
    tasks_df = tasks_df.sort_values('planned_start')
    
    # Calculate progress for each task
    tasks_df = tasks_df.copy()
    tasks_df['progress'] = (tasks_df['earned_value'] / tasks_df['planned_cost']).clip(0, 1) * 100
    
    # Determine color based on task status
    def get_task_color(row):
        if row['progress'] >= 100:
            return '#2ECC71'  # Completed
        
        planned_end = row['planned_end']
        actual_end = row.get('actual_end', datetime.now())
        
        if actual_end > planned_end:
            return '#E74C3C'  # Delayed
        elif row['progress'] < row['expected_progress']:
            return '#F1C40F'  # Behind schedule
        else:
            return '#3498DB'  # On track
    
    tasks_df['color'] = tasks_df.apply(get_task_color, axis=1)
    
    # Create figure
    # Convert the color Series to a list to avoid the multiplication error
    colors_list = tasks_df['color'].tolist()
    
    # Prepare dataframe with required columns for Gantt
    gantt_df = tasks_df.rename(columns={
        'task_name': 'Task',
        'planned_start': 'Start',
        'planned_end': 'Finish'
    })
    
    # Create the Gantt chart
    fig = ff.create_gantt(
        gantt_df,
        colors=colors_list,
        index_col='Task',
        show_colorbar=True,
        group_tasks=True
    )
    
    # Get today's date for reference line
    today = datetime.now()
    
    # Add a vertical line for today's date
    fig.add_shape(
        type="line",
        x0=today,
        y0=0,
        x1=today,
        y1=len(tasks_df),
        line=dict(color="black", width=2, dash="dash")
    )
    
    # Add annotation for today's date
    fig.add_annotation(
        x=today,
        y=len(tasks_df),
        text="Today",
        showarrow=True,
        arrowhead=1,
        ax=0,
        ay=-40
    )
    
    # Update layout
    fig.update_layout(
        title="Project Schedule",
        xaxis=dict(
            title="Date",
            rangeselector=dict(
                buttons=list([
                    dict(count=1, label="1m", step="month", stepmode="backward"),
                    dict(count=6, label="6m", step="month", stepmode="backward"),
                    dict(count=1, label="YTD", step="year", stepmode="todate"),
                    dict(count=1, label="1y", step="year", stepmode="backward"),
                    dict(step="all")
                ])
            )
        ),
        yaxis=dict(
            title="Tasks"
        ),
        plot_bgcolor='rgba(255, 255, 255, 1)',
        paper_bgcolor='rgba(255, 255, 255, 1)',
        margin=dict(l=20, r=20, t=50, b=50),
        height=600
    )
    
    return fig
