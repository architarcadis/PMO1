#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Predictive Analytics Module for PMO Pulse Dashboard

This module provides advanced analytics capabilities including:
- Machine learning-based forecasts for project completion 
- Anomaly detection for identifying outliers in project performance
- Natural language insights generation
- Risk prediction models

Author: Arcadis PMO Analytics Team
Version: 1.0
"""

import pandas as pd
import numpy as np
import datetime
from datetime import datetime, timedelta
import joblib
import os
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor, IsolationForest
from sklearn.preprocessing import StandardScaler
import scipy.stats as stats

# Constants
MODEL_DIR = 'models'
if not os.path.exists(MODEL_DIR):
    os.makedirs(MODEL_DIR)

def train_forecast_model(historical_data, target_column='spi', features=None):
    """
    Trains a forecasting model for project metrics based on historical data
    
    Parameters:
    -----------
    historical_data : pandas.DataFrame
        Historical project performance data
    target_column : str
        The metric to forecast (e.g., 'spi', 'cpi')
    features : list
        List of feature columns to use (if None, uses all available except target)
        
    Returns:
    --------
    sklearn model
        Trained forecasting model
    """
    if historical_data.empty:
        return None
        
    # Prepare the data
    data = historical_data.copy()
    
    # Convert dates to numerical features
    if 'month' in data.columns:
        data['month_num'] = pd.to_datetime(data['month']).apply(lambda x: x.timestamp())
    
    # If no features are specified, use all columns except the target
    if features is None:
        features = [col for col in data.columns if col != target_column 
                   and col != 'project_id' and col != 'month'
                   and data[col].dtype in [np.int64, np.float64]]
    
    # Ensure we have sufficient data
    if len(data) < 5:  # Need at least 5 data points
        return None
    
    # Handle missing values
    data = data.dropna(subset=features + [target_column])
    
    if len(data) < 5:  # Re-check after dropping NAs
        return None
    
    # Train a Random Forest model
    X = data[features]
    y = data[target_column]
    
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X, y)
    
    # Save the model
    model_path = os.path.join(MODEL_DIR, f'{target_column}_forecast_model.joblib')
    joblib.dump(model, model_path)
    
    return model

def predict_project_completion(project_data, tasks_data, historical_data=None):
    """
    Predicts project completion date and final cost based on current status and trends
    
    Parameters:
    -----------
    project_data : pandas.Series
        Current project data
    tasks_data : pandas.DataFrame
        Tasks related to the project
    historical_data : pandas.DataFrame, optional
        Historical performance data for the project
        
    Returns:
    --------
    dict
        Predicted completion date, cost, and confidence metrics
    """
    results = {}
    
    # Basic calculations (EVM-based)
    planned_duration = (project_data['planned_end'] - project_data['planned_start']).days
    elapsed_days = (datetime.now() - project_data['planned_start']).days
    
    # Calculate current progress
    if not tasks_data.empty:
        total_planned_cost = tasks_data['planned_cost'].sum()
        total_earned_value = tasks_data['earned_value'].sum()
        total_actual_cost = tasks_data['actual_cost'].sum()
        
        if total_planned_cost > 0:
            progress_pct = total_earned_value / total_planned_cost
            spi = total_earned_value / total_planned_cost if total_planned_cost > 0 else 1.0
            cpi = total_earned_value / total_actual_cost if total_actual_cost > 0 else 1.0
        else:
            progress_pct = 0
            spi = 1.0
            cpi = 1.0
    else:
        progress_pct = elapsed_days / planned_duration if planned_duration > 0 else 0
        spi = 1.0
        cpi = 1.0
            
    # Cap progress at 100%
    progress_pct = min(1.0, progress_pct)
    
    # Calculate remaining duration and adjust by SPI
    remaining_duration = planned_duration - elapsed_days
    
    if spi > 0:
        adjusted_remaining = remaining_duration / spi
    else:
        adjusted_remaining = remaining_duration * 2  # Assume double time if SPI is 0
    
    adjusted_remaining = max(0, adjusted_remaining)  # Ensure it's not negative
    
    # Calculate forecast completion date
    remaining_days = int(adjusted_remaining)
    forecast_completion = project_data['planned_start'] + timedelta(days=elapsed_days + remaining_days)
    
    # Calculate schedule variance
    schedule_variance = remaining_days - remaining_duration
    
    # Calculate cost variance and EAC (Estimate at Completion)
    budget = project_data['budget']
    spent = budget * progress_pct
    
    if cpi > 0:
        eac = spent + ((budget - spent) / cpi)
    else:
        eac = budget * 2  # Assume double cost if CPI is 0
        
    cost_variance = eac - budget
    
    # Calculate confidence level (rough estimate based on variance)
    spi_confidence = min(100, int(100 * (1 - (0.5 * abs(1 - spi)))))
    cpi_confidence = min(100, int(100 * (1 - (0.5 * abs(1 - cpi)))))
    
    avg_confidence = (spi_confidence + cpi_confidence) / 2
    
    # Prepare the results
    results = {
        'forecast_completion_date': forecast_completion,
        'schedule_variance_days': schedule_variance,
        'forecast_cost': eac,
        'cost_variance': cost_variance,
        'schedule_confidence': spi_confidence,
        'cost_confidence': cpi_confidence,
        'overall_confidence': avg_confidence,
        'progress_percentage': progress_pct * 100
    }
    
    # Add ML-based predictions if historical data is available
    if historical_data is not None and len(historical_data) >= 5:
        # Advanced ML prediction could be added here
        pass
        
    return results

def detect_anomalies(project_data, method='isolation_forest', contamination=0.05):
    """
    Detects anomalies in project performance data
    
    Parameters:
    -----------
    project_data : pandas.DataFrame
        Project performance data with metrics
    method : str
        Detection method: 'isolation_forest', 'zscore', etc.
    contamination : float
        Expected ratio of outliers in the data
        
    Returns:
    --------
    pandas.DataFrame
        Original data with added anomaly flag and score
    """
    if project_data.empty:
        return project_data
    
    # Select numerical columns for anomaly detection
    num_cols = project_data.select_dtypes(include=[np.number]).columns.tolist()
    
    # Remove ID columns from features
    features = [col for col in num_cols if 'id' not in col.lower() 
               and 'date' not in col.lower() 
               and 'month' not in col.lower()]
    
    if not features:
        # No suitable features for anomaly detection
        return project_data
    
    result_df = project_data.copy()
    
    # Fill NAs to avoid errors (with column means)
    X = result_df[features].fillna(result_df[features].mean())
    
    if method == 'isolation_forest':
        # Isolation Forest method
        model = IsolationForest(contamination=contamination, random_state=42)
        result_df['anomaly_score'] = model.fit_predict(X)
        result_df['is_anomaly'] = result_df['anomaly_score'] == -1
        
    elif method == 'zscore':
        # Z-score method (statistical)
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(X)
        
        # Calculate mean Z-score across all features
        z_scores = np.mean(np.abs(scaled_data), axis=1)
        threshold = 2.5  # Z-score threshold for anomalies
        
        result_df['anomaly_score'] = z_scores
        result_df['is_anomaly'] = z_scores > threshold
    
    return result_df

def generate_natural_language_insights(project_data, portfolio_data=None, history_data=None):
    """
    Generates natural language insights about project performance
    
    Parameters:
    -----------
    project_data : pandas.DataFrame or pandas.Series
        Current project data
    portfolio_data : pandas.DataFrame, optional
        Overall portfolio data for comparison
    history_data : pandas.DataFrame, optional
        Historical project data for trend analysis
        
    Returns:
    --------
    list
        List of insight statements
    """
    insights = []
    
    if isinstance(project_data, pd.DataFrame) and len(project_data) == 0:
        return ["Insufficient data available to generate insights."]
    
    # Handle both Series (single project) and DataFrame (multiple projects)
    if isinstance(project_data, pd.Series):
        # Single project analysis
        spi = project_data.get('spi', None)
        cpi = project_data.get('cpi', None)
        status = project_data.get('status', None)
        budget = project_data.get('budget', 0)
        
        # Project schedule insights
        if spi is not None:
            if spi < 0.8:
                insights.append(f"Project is significantly behind schedule (SPI = {spi:.2f}), which may impact delivery milestones.")
            elif spi < 0.95:
                insights.append(f"Project is slightly behind schedule (SPI = {spi:.2f}). Close monitoring recommended.")
            elif spi > 1.1:
                insights.append(f"Project is ahead of schedule (SPI = {spi:.2f}), potentially allowing for early delivery.")
        
        # Project cost insights
        if cpi is not None:
            if cpi < 0.8:
                insights.append(f"Project is significantly over budget (CPI = {cpi:.2f}). Budget controls should be reviewed.")
            elif cpi < 0.95:
                insights.append(f"Project is slightly over budget (CPI = {cpi:.2f}). Cost management attention required.")
            elif cpi > 1.1:
                insights.append(f"Project is under budget (CPI = {cpi:.2f}), showing strong cost management.")
        
        # Combined insights
        if spi is not None and cpi is not None:
            if spi < 0.9 and cpi < 0.9:
                insights.append("Project is at high risk due to both schedule delays and cost overruns.")
            elif spi > 1.05 and cpi > 1.05:
                insights.append("Project is performing exceptionally well in both schedule and cost dimensions.")
    
    else:
        # Portfolio-level insights
        avg_spi = project_data['spi'].mean() if 'spi' in project_data.columns else None
        avg_cpi = project_data['cpi'].mean() if 'cpi' in project_data.columns else None
        
        # Count projects in different status categories
        if 'status' in project_data.columns:
            status_counts = project_data['status'].value_counts()
            at_risk_count = status_counts.get('At Risk', 0) + status_counts.get('Delayed', 0)
            
            if at_risk_count > 0 and len(project_data) > 0:
                insights.append(f"{at_risk_count} projects ({at_risk_count/len(project_data)*100:.1f}%) are at risk or delayed and may require intervention.")
        
        # Portfolio schedule insights
        if avg_spi is not None:
            if avg_spi < 0.9:
                insights.append(f"Portfolio is generally behind schedule (SPI = {avg_spi:.2f}). Resource allocation should be reviewed.")
            elif avg_spi > 1.05:
                insights.append(f"Portfolio is ahead of schedule (SPI = {avg_spi:.2f}), indicating effective schedule management.")
        
        # Portfolio cost insights
        if avg_cpi is not None:
            if avg_cpi < 0.9:
                insights.append(f"Portfolio is over budget (CPI = {avg_cpi:.2f}). Cost control measures should be reinforced.")
            elif avg_cpi > 1.05:
                insights.append(f"Portfolio is under budget (CPI = {avg_cpi:.2f}), demonstrating effective cost management.")
        
        # Performance distribution insights
        if 'spi' in project_data.columns:
            low_spi_count = len(project_data[project_data['spi'] < 0.8])
            if low_spi_count > 0:
                insights.append(f"{low_spi_count} projects have critical schedule performance (SPI < 0.8) requiring immediate attention.")
    
    # Add trend insights if historical data is available
    if history_data is not None and not history_data.empty:
        # Analysis could be added here
        pass
    
    # If no insights could be generated
    if not insights:
        insights.append("No significant patterns or issues identified in the current project data.")
    
    return insights