"""
Data Processor Module for PMO Pulse Dashboard

Processes project data for analysis and visualization.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def calculate_portfolio_metrics(projects_df, tasks_df, history_df):
    """
    Calculate key portfolio metrics for dashboard display.
    
    Parameters:
    -----------
    projects_df : pandas.DataFrame
        DataFrame containing project information
    tasks_df : pandas.DataFrame
        DataFrame containing task information
    history_df : pandas.DataFrame
        DataFrame containing historical performance data
        
    Returns:
    --------
    dict
        Dictionary containing calculated metrics
    """
    # Get the latest month's data for each project
    latest_performance = history_df.groupby('project_id').apply(
        lambda x: x.sort_values('month').iloc[-1]
    ).reset_index(drop=True)
    
    # Get the previous month's data for comparison
    previous_month_performance = history_df.groupby('project_id').apply(
        lambda x: x.sort_values('month').iloc[-2] if len(x) > 1 else None
    ).reset_index(drop=True).dropna()
    
    # Calculate average SPI and CPI
    avg_spi = latest_performance['spi'].mean()
    avg_cpi = latest_performance['cpi'].mean()
    
    # Calculate change in SPI and CPI from previous month
    if not previous_month_performance.empty:
        prev_avg_spi = previous_month_performance['spi'].mean()
        prev_avg_cpi = previous_month_performance['cpi'].mean()
        spi_change = avg_spi - prev_avg_spi
        cpi_change = avg_cpi - prev_avg_cpi
    else:
        spi_change = 0
        cpi_change = 0
    
    # Count active and at-risk projects
    active_projects = len(projects_df[projects_df['status'] != 'Completed'])
    at_risk_projects = len(projects_df[projects_df['status'].isin(['At Risk', 'Delayed'])])
    
    # Calculate change in active and at-risk projects
    # For this mock data, we'll simulate a change
    active_projects_change = np.random.randint(-2, 3)
    at_risk_projects_change = np.random.randint(-2, 3)
    
    # Return metrics as a dictionary
    return {
        'avg_spi': avg_spi,
        'avg_cpi': avg_cpi,
        'spi_change': spi_change,
        'cpi_change': cpi_change,
        'active_projects': active_projects,
        'at_risk_projects': at_risk_projects,
        'active_projects_change': active_projects_change,
        'at_risk_projects_change': at_risk_projects_change
    }

def calculate_project_metrics(project_df, project_tasks, project_history):
    """
    Calculate metrics for a specific project.
    
    Parameters:
    -----------
    project_df : pandas.Series
        Series containing project information
    project_tasks : pandas.DataFrame
        DataFrame containing tasks for this project
    project_history : pandas.DataFrame
        DataFrame containing historical performance data for this project
        
    Returns:
    --------
    dict
        Dictionary containing calculated metrics
    """
    # Get latest SPI and CPI values
    if not project_history.empty:
        latest_performance = project_history.sort_values('month').iloc[-1]
        spi = latest_performance['spi']
        cpi = latest_performance['cpi']
        
        # Get previous month's values for change calculation
        if len(project_history) > 1:
            previous_performance = project_history.sort_values('month').iloc[-2]
            spi_change = spi - previous_performance['spi']
            cpi_change = cpi - previous_performance['cpi']
        else:
            spi_change = 0
            cpi_change = 0
    else:
        # Default values if no history
        spi = 1.0
        cpi = 1.0
        spi_change = 0
        cpi_change = 0
    
    # Calculate budget metrics
    budget = project_df['budget']
    actual_cost = project_tasks['actual_cost'].sum() if not project_tasks.empty else 0
    spend_percentage = (actual_cost / budget) * 100 if budget > 0 else 0
    
    # Calculate schedule metrics
    today = datetime.now().date()
    start_date = project_df['start_date'].date()
    end_date = project_df['planned_end_date'].date()
    total_days = (end_date - start_date).days if end_date > start_date else 1
    days_passed = (min(today, end_date) - start_date).days
    days_remaining = (end_date - today).days if end_date > today else 0
    
    schedule_progress = (days_passed / total_days) * 100 if total_days > 0 else 0
    
    # Return metrics as a dictionary
    return {
        'spi': spi,
        'cpi': cpi,
        'spi_change': spi_change,
        'cpi_change': cpi_change,
        'actual_cost': actual_cost,
        'spend_percentage': spend_percentage,
        'schedule_progress': schedule_progress,
        'days_remaining': days_remaining
    }

def get_projects_needing_attention(projects_df, history_df, spi_threshold=0.9, cpi_threshold=0.9):
    """
    Identify projects that need attention based on performance thresholds.
    
    Parameters:
    -----------
    projects_df : pandas.DataFrame
        DataFrame containing project information
    history_df : pandas.DataFrame
        DataFrame containing historical performance data
    spi_threshold : float
        Threshold below which SPI is considered at risk
    cpi_threshold : float
        Threshold below which CPI is considered at risk
        
    Returns:
    --------
    pandas.DataFrame
        DataFrame containing projects needing attention
    """
    # Get the latest performance data for each project
    latest_performance = history_df.groupby('project_id').apply(
        lambda x: x.sort_values('month').iloc[-1]
    ).reset_index(drop=True)
    
    # Merge with project data
    merged_data = pd.merge(
        projects_df,
        latest_performance[['project_id', 'spi', 'cpi']],
        on='project_id',
        how='left'
    )
    
    # Fill NaN values with defaults
    merged_data['spi'] = merged_data['spi'].fillna(1.0)
    merged_data['cpi'] = merged_data['cpi'].fillna(1.0)
    
    # Filter projects based on thresholds
    attention_projects = merged_data[
        (merged_data['spi'] < spi_threshold) | 
        (merged_data['cpi'] < cpi_threshold) |
        (merged_data['status'].isin(['At Risk', 'Delayed']))
    ].copy()
    
    # Add combined risk score for sorting
    attention_projects['risk_score'] = (
        (1 - attention_projects['spi']) * 0.5 + 
        (1 - attention_projects['cpi']) * 0.5
    )
    
    # Sort by risk score (highest risk first)
    attention_projects = attention_projects.sort_values('risk_score', ascending=False)
    
    return attention_projects

def prepare_forecast_data(project_df, project_history):
    """
    Prepare forecast data for a project based on historical performance.
    
    Parameters:
    -----------
    project_df : pandas.Series
        Series containing project information
    project_history : pandas.DataFrame
        DataFrame containing historical performance data for this project
        
    Returns:
    --------
    dict
        Dictionary containing forecast data
    """
    # Default values
    today = datetime.now().date()
    start_date = project_df['start_date'].date()
    planned_end_date = project_df['planned_end_date'].date()
    budget = project_df['budget']
    
    # Initialize with empty data
    historical_dates = pd.Series([])
    historical_spi = pd.Series([])
    historical_cpi = pd.Series([])
    forecast_dates = pd.Series([])
    forecast_spi = pd.Series([])
    forecast_cpi = pd.Series([])
    
    # Calculate forecast if we have historical data
    if not project_history.empty:
        # Process historical data
        project_history = project_history.sort_values('month')
        historical_months = [datetime.strptime(month, '%Y-%m').date() for month in project_history['month']]
        historical_dates = pd.Series(historical_months)
        historical_spi = project_history['spi']
        historical_cpi = project_history['cpi']
        
        # Calculate average SPI and CPI for forecasting
        # Use more recent values with higher weight
        weights = np.linspace(1, 3, len(project_history))
        weights = weights / weights.sum()
        avg_spi = np.average(project_history['spi'], weights=weights)
        avg_cpi = np.average(project_history['cpi'], weights=weights)
        
        # Generate forecast dates (3 months into the future)
        last_date = historical_months[-1]
        forecast_months = [
            last_date + timedelta(days=30 * i)
            for i in range(1, 4)
        ]
        forecast_dates = pd.Series(forecast_months)
        
        # Generate forecast SPI and CPI with slight regression to mean
        forecast_spi_values = []
        forecast_cpi_values = []
        
        for i in range(3):
            regression_factor = 0.2 * (i + 1)
            forecast_spi_values.append(
                avg_spi * (1 - regression_factor) + 1.0 * regression_factor
            )
            forecast_cpi_values.append(
                avg_cpi * (1 - regression_factor) + 1.0 * regression_factor
            )
        
        forecast_spi = pd.Series(forecast_spi_values)
        forecast_cpi = pd.Series(forecast_cpi_values)
        
        # Calculate schedule variance
        # If SPI < 1, project will take longer than planned
        total_duration = (planned_end_date - start_date).days
        forecasted_duration = total_duration / avg_spi if avg_spi > 0 else float('inf')
        schedule_variance_days = int(forecasted_duration - total_duration)
        
        # Calculate forecast end date
        forecast_end_date = start_date + timedelta(days=int(forecasted_duration))
        
        # Calculate cost forecast
        # If CPI < 1, project will cost more than budgeted
        forecast_cost = budget / avg_cpi if avg_cpi > 0 else float('inf')
    else:
        # Default values if no history
        avg_spi = 1.0
        avg_cpi = 1.0
        schedule_variance_days = 0
        forecast_end_date = planned_end_date
        forecast_cost = budget
    
    return {
        'historical_dates': historical_dates,
        'historical_spi': historical_spi,
        'historical_cpi': historical_cpi,
        'forecast_dates': forecast_dates,
        'forecast_spi': forecast_spi,
        'forecast_cpi': forecast_cpi,
        'avg_spi': avg_spi,
        'avg_cpi': avg_cpi,
        'schedule_variance_days': schedule_variance_days,
        'forecast_end_date': forecast_end_date,
        'forecast_cost': forecast_cost
    }
