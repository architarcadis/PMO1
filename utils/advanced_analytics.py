#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Advanced Analytics Module for PMO Pulse Dashboard

This module provides advanced analytics capabilities including:
- Project anomaly detection
- Resource optimization
- Budget forecast models
- Natural language insights generation

Author: Arcadis PMO Analytics Team
Version: 1.0
"""

import pandas as pd
import numpy as np
import datetime
from datetime import datetime, timedelta
from scipy import stats
from sklearn.cluster import DBSCAN
from sklearn.ensemble import IsolationForest
from sklearn.preprocessing import StandardScaler
import joblib
import os

# Define constants
MODEL_DIR = 'models'
if not os.path.exists(MODEL_DIR):
    os.makedirs(MODEL_DIR)

def detect_project_anomalies(projects_df, metrics=None, contamination=0.1):
    """
    Detects anomalous projects based on performance metrics
    
    Parameters:
    -----------
    projects_df : pandas DataFrame
        DataFrame containing project performance data
    metrics : list
        List of metric columns to consider for anomaly detection
    contamination : float
        Expected proportion of outliers (0.0 to 0.5)
        
    Returns:
    --------
    pandas DataFrame
        Original DataFrame with anomaly score and flag columns added
    """
    if projects_df.empty:
        return projects_df
    
    # Define metrics to use if not specified
    if metrics is None:
        metrics = ['spi', 'cpi', 'budget_variance_pct', 'schedule_variance_days']
    
    # Filter to only include numeric metrics that exist in the DataFrame
    valid_metrics = [col for col in metrics if col in projects_df.columns 
                    and np.issubdtype(projects_df[col].dtype, np.number)]
    
    if not valid_metrics:
        # No valid metrics found
        return projects_df
    
    # Make a copy of the DataFrame for results
    result_df = projects_df.copy()
    
    # Prepare the data (fill missing values with column means)
    X = result_df[valid_metrics].fillna(result_df[valid_metrics].mean())
    
    # Normalize the data
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    # Use Isolation Forest for anomaly detection
    model = IsolationForest(contamination=contamination, random_state=42)
    result_df['anomaly_score'] = model.fit_predict(X_scaled)
    
    # Convert predictions (-1 for outliers, 1 for inliers) to boolean flags
    result_df['is_anomaly'] = result_df['anomaly_score'] == -1
    
    return result_df

def optimize_resource_allocation(projects_df, resources_df, task_df=None):
    """
    Analyzes resource allocation and identifies optimization opportunities
    
    Parameters:
    -----------
    projects_df : pandas DataFrame
        DataFrame containing project data
    resources_df : pandas DataFrame
        DataFrame containing resource allocation data
    task_df : pandas DataFrame, optional
        DataFrame containing task-level data
        
    Returns:
    --------
    dict
        Dictionary with optimization results and recommendations
    """
    results = {
        'overallocated_resources': [],
        'underallocated_resources': [],
        'resource_conflicts': [],
        'optimization_recommendations': []
    }
    
    if projects_df.empty or resources_df.empty:
        return results
    
    # Placeholder for resource optimization logic
    # This would involve analyzing resource utilization across projects
    # and identifying conflicts, over/under allocations, etc.
    
    return results

def forecast_project_completion(project_data, tasks_data, historical_data=None):
    """
    Forecasts project completion date and final cost based on current progress
    
    Parameters:
    -----------
    project_data : pandas Series
        Data for a specific project
    tasks_data : pandas DataFrame
        Tasks associated with the project
    historical_data : pandas DataFrame, optional
        Historical performance data for better forecasting
        
    Returns:
    --------
    dict
        Dictionary with forecast results
    """
    results = {
        'forecast_completion_date': None,
        'forecast_final_cost': None,
        'schedule_variance_days': None,
        'cost_variance': None,
        'completion_confidence': None
    }
    
    if project_data is None:
        return results
    
    # Basic forecasting using Earned Value Management (EVM) principles
    if not tasks_data.empty:
        # Calculate key EVM metrics
        total_planned_value = tasks_data['planned_cost'].sum()
        total_earned_value = tasks_data['earned_value'].sum() if 'earned_value' in tasks_data.columns else 0
        total_actual_cost = tasks_data['actual_cost'].sum() if 'actual_cost' in tasks_data.columns else 0
        
        # Calculate SPI and CPI
        if total_planned_value > 0:
            spi = total_earned_value / total_planned_value
            cpi = total_earned_value / total_actual_cost if total_actual_cost > 0 else 1.0
        else:
            spi = 1.0
            cpi = 1.0
        
        # Calculate schedule and cost forecasts
        planned_duration = (project_data['planned_end'] - project_data['planned_start']).days
        elapsed_days = (datetime.now() - project_data['planned_start']).days
        
        # Estimate remaining duration adjusted by SPI
        if spi > 0:
            remaining_duration = (planned_duration - elapsed_days) / spi
        else:
            remaining_duration = planned_duration - elapsed_days  # Default if SPI is 0
            
        # Calculate forecast completion date
        forecast_completion = project_data['planned_start'] + timedelta(days=elapsed_days + remaining_duration)
        results['forecast_completion_date'] = forecast_completion
        
        # Calculate schedule variance in days
        results['schedule_variance_days'] = (forecast_completion - project_data['planned_end']).days
        
        # Calculate forecast final cost (EAC - Estimate at Completion)
        if cpi > 0:
            forecast_cost = total_actual_cost + ((project_data['budget'] - total_earned_value) / cpi)
        else:
            forecast_cost = project_data['budget'] * 1.5  # Default assumption if CPI is 0
            
        results['forecast_final_cost'] = forecast_cost
        results['cost_variance'] = forecast_cost - project_data['budget']
        
        # Calculate completion confidence based on SPI and CPI
        # Higher values of SPI and CPI indicate higher confidence
        spi_factor = min(1.0, spi)  # Cap at 1.0
        cpi_factor = min(1.0, cpi)  # Cap at 1.0
        confidence = (spi_factor * 0.5 + cpi_factor * 0.5) * 100  # Simple weighted average
        results['completion_confidence'] = int(confidence)
    
    return results

def generate_nlp_insights(project_data, portfolio_data=None, history_data=None):
    """
    Generates natural language insights about project performance
    
    Parameters:
    -----------
    project_data : pandas Series or DataFrame
        Project data to analyze
    portfolio_data : pandas DataFrame, optional
        Overall portfolio data for comparison
    history_data : pandas DataFrame, optional
        Historical data for trend analysis
        
    Returns:
    --------
    list
        List of insight statements
    """
    insights = []
    
    # Check if we have a single project or a portfolio
    is_portfolio = isinstance(project_data, pd.DataFrame) and len(project_data) > 1
    
    if is_portfolio:
        # Portfolio-level insights
        if 'spi' in project_data.columns:
            avg_spi = project_data['spi'].mean()
            below_threshold = len(project_data[project_data['spi'] < 0.9])
            
            if below_threshold > 0:
                insights.append(f"{below_threshold} projects are significantly behind schedule (SPI < 0.9) and may require intervention.")
                
            if avg_spi < 0.95:
                insights.append(f"Portfolio is generally behind schedule with average SPI of {avg_spi:.2f}. Resource allocation should be reviewed.")
            elif avg_spi > 1.05:
                insights.append(f"Portfolio is performing well on schedule with average SPI of {avg_spi:.2f}.")
        
        if 'cpi' in project_data.columns:
            avg_cpi = project_data['cpi'].mean()
            below_budget = len(project_data[project_data['cpi'] < 0.9])
            
            if below_budget > 0:
                insights.append(f"{below_budget} projects are over budget (CPI < 0.9) and require cost management attention.")
                
            if avg_cpi < 0.95:
                insights.append(f"Portfolio is generally over budget with average CPI of {avg_cpi:.2f}. Cost controls should be strengthened.")
            elif avg_cpi > 1.05:
                insights.append(f"Portfolio is performing well on budget with average CPI of {avg_cpi:.2f}.")
    else:
        # Single project insights
        if isinstance(project_data, pd.Series):
            if 'spi' in project_data:
                spi = project_data['spi']
                if spi < 0.8:
                    insights.append(f"Project is significantly behind schedule (SPI = {spi:.2f}). Immediate attention required.")
                elif spi < 0.95:
                    insights.append(f"Project is slightly behind schedule (SPI = {spi:.2f}). Monitor closely.")
                elif spi > 1.1:
                    insights.append(f"Project is ahead of schedule (SPI = {spi:.2f}).")
                    
            if 'cpi' in project_data:
                cpi = project_data['cpi']
                if cpi < 0.8:
                    insights.append(f"Project is significantly over budget (CPI = {cpi:.2f}). Budget control measures required.")
                elif cpi < 0.95:
                    insights.append(f"Project is slightly over budget (CPI = {cpi:.2f}). Monitor expenses closely.")
                elif cpi > 1.1:
                    insights.append(f"Project is under budget (CPI = {cpi:.2f}), showing good cost management.")
    
    # Add trend insights if history data is available
    if history_data is not None and not history_data.empty:
        # This would analyze trends over time and generate insights
        pass
    
    # If no insights could be generated
    if not insights:
        insights.append("No significant patterns detected in the project data.")
    
    return insights