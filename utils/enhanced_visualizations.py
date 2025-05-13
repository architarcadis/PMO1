#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Enhanced Visualizations Module for PMO Pulse Dashboard

This module provides advanced and interactive visualizations including:
- Interactive Gantt charts
- Geospatial project visualizations
- Resource allocation heatmaps
- Advanced performance tracking charts
- Custom dashboard components

Author: Arcadis PMO Analytics Team
Version: 1.0
"""

import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Color constants
ARCADIS_ORANGE = "#E67300"
ARCADIS_TEAL = "#00A3A1"
ARCADIS_BLUE = "#003366"
ARCADIS_GRAY = "#646469"

# Color themes
COLOR_THEMES = {
    "Arcadis Default": {
        "primary": ARCADIS_ORANGE,
        "secondary": ARCADIS_TEAL,
        "tertiary": ARCADIS_BLUE,
        "neutral": ARCADIS_GRAY,
        "success": "#2ECC71",
        "warning": "#F1C40F",
        "danger": "#E74C3C",
        "info": "#3498DB",
        "background": "#FFFFFF",
        "text": "#333333"
    },
    "Monochrome": {
        "primary": "#2C3E50",
        "secondary": "#7F8C8D",
        "tertiary": "#BDC3C7",
        "neutral": "#95A5A6",
        "success": "#27AE60",
        "warning": "#F39C12",
        "danger": "#C0392B",
        "info": "#2980B9",
        "background": "#FFFFFF",
        "text": "#2C3E50"
    },
    "Colorful": {
        "primary": "#9B59B6",
        "secondary": "#3498DB",
        "tertiary": "#2ECC71",
        "neutral": "#95A5A6",
        "success": "#2ECC71",
        "warning": "#F1C40F",
        "danger": "#E74C3C",
        "info": "#3498DB",
        "background": "#FFFFFF",
        "text": "#2C3E50"
    },
    "High Contrast": {
        "primary": "#000000",
        "secondary": "#FFFFFF",
        "tertiary": "#FFD700",
        "neutral": "#808080",
        "success": "#008000",
        "warning": "#FFA500",
        "danger": "#FF0000",
        "info": "#0000FF",
        "background": "#FFFFFF",
        "text": "#000000"
    }
}

def get_theme_colors(theme_name="Arcadis Default"):
    """
    Returns color palette for the specified theme
    
    Parameters:
    -----------
    theme_name : str
        Name of the theme to use
        
    Returns:
    --------
    dict
        Dictionary of theme colors
    """
    return COLOR_THEMES.get(theme_name, COLOR_THEMES["Arcadis Default"])

def create_interactive_gantt(tasks_df, project_df=None, show_critical_path=True, editable=False):
    """
    Creates an interactive Gantt chart with optional critical path highlighting
    
    Parameters:
    -----------
    tasks_df : pandas DataFrame
        DataFrame containing task data
    project_df : pandas Series, optional
        Project data for context
    show_critical_path : bool
        Whether to highlight the critical path
    editable : bool
        Whether to enable interactive editing (drag to adjust dates)
        
    Returns:
    --------
    plotly.graph_objects.Figure
        Interactive Gantt chart
    """
    if tasks_df.empty:
        # Create empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No task data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        return fig
    
    # Create a copy of the dataframe to avoid modification warnings
    df = tasks_df.copy()
    
    # Determine which tasks are on the critical path (if requested)
    if show_critical_path and 'is_critical' not in df.columns:
        # For simplicity, use a heuristic: tasks with zero float are critical
        if 'total_float' in df.columns:
            df['is_critical'] = df['total_float'] <= 1  # 1 day or less float
        else:
            # If no float information, can't determine critical path
            df['is_critical'] = False
    
    # Calculate percent complete if not provided
    if 'percent_complete' not in df.columns and 'earned_value' in df.columns and 'planned_cost' in df.columns:
        df['percent_complete'] = (df['earned_value'] / df['planned_cost']).clip(0, 1) * 100
    elif 'percent_complete' not in df.columns:
        df['percent_complete'] = 0
    
    # Determine color based on task status or critical path
    if 'is_critical' in df.columns and show_critical_path:
        # Color by critical path (true/false)
        color_column = 'is_critical'
        color_map = {True: "red", False: "blue"}
    elif 'status' in df.columns:
        # Color by task status
        color_column = 'status'
        color_map = {
            'Completed': '#2ECC71',
            'In Progress': '#3498DB',
            'Not Started': '#95A5A6',
            'Behind Schedule': '#E74C3C',
            'At Risk': '#F1C40F'
        }
    else:
        # Default coloring
        color_column = None
        color_map = None
    
    # Create the Gantt chart
    if 'task_name' in df.columns and 'planned_start' in df.columns and 'planned_end' in df.columns:
        fig = px.timeline(
            df,
            x_start="planned_start",
            x_end="planned_end",
            y="task_name",
            color=color_column,
            color_discrete_map=color_map,
            hover_name="task_name",
            hover_data=["percent_complete", "planned_cost"] if "planned_cost" in df.columns else ["percent_complete"],
            labels={"task_name": "Task", "planned_start": "Start", "planned_end": "End"},
            title="Project Schedule"
        )
        
        # Add completion indicators as markers
        has_actual_dates = 'actual_start' in df.columns and 'actual_end' in df.columns
        
        if has_actual_dates:
            # Add actual start/end markers
            for i, task in df.iterrows():
                if pd.notna(task['actual_start']):
                    fig.add_scatter(
                        x=[task['actual_start']],
                        y=[task['task_name']],
                        mode='markers',
                        marker=dict(symbol='triangle-right', size=12, color='green'),
                        name='Actual Start',
                        showlegend=i==0  # Only show in legend once
                    )
                    
                if pd.notna(task['actual_end']) and task['percent_complete'] >= 100:
                    fig.add_scatter(
                        x=[task['actual_end']],
                        y=[task['task_name']],
                        mode='markers',
                        marker=dict(symbol='triangle-right', size=12, color='red'),
                        name='Actual End',
                        showlegend=i==0  # Only show in legend once
                    )
        
        # Add today's date line
        today = datetime.now()
        fig.add_vline(
            x=today, 
            line_width=2, 
            line_dash="dash", 
            line_color="black",
            annotation_text="Today",
            annotation_position="top right"
        )
        
        # Customize the layout
        fig.update_layout(
            xaxis_title="Date",
            yaxis_title="",
            yaxis={'categoryorder':'total ascending'},
            plot_bgcolor='white',
            showlegend=True,
            legend_title="Task Status" if color_column == 'status' else "Critical Path" if color_column == 'is_critical' else "",
            height=max(400, len(df) * 30),  # Dynamic height based on number of tasks
            margin=dict(l=10, r=10, t=50, b=10)
        )
        
        # Add project start/end lines if project data provided
        if project_df is not None:
            if 'planned_start' in project_df and 'planned_end' in project_df:
                fig.add_vline(
                    x=project_df['planned_start'], 
                    line_width=1.5, 
                    line_dash="solid", 
                    line_color="green",
                    annotation_text="Project Start",
                    annotation_position="top left"
                )
                
                fig.add_vline(
                    x=project_df['planned_end'], 
                    line_width=1.5, 
                    line_dash="solid", 
                    line_color="red",
                    annotation_text="Project End",
                    annotation_position="top right"
                )
    else:
        # Create empty figure with error message
        fig = go.Figure()
        fig.add_annotation(
            text="Task data missing required columns (task_name, planned_start, planned_end)",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
    
    return fig

def create_geospatial_map(projects_df, color_by='status', size_by='budget'):
    """
    Creates a geospatial map of projects
    
    Parameters:
    -----------
    projects_df : pandas DataFrame
        DataFrame containing project data with location information
    color_by : str
        Column to use for coloring points
    size_by : str
        Column to use for point sizes
        
    Returns:
    --------
    plotly.graph_objects.Figure
        Interactive map visualization
    """
    if projects_df.empty:
        # Create empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No project location data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        return fig
    
    # Check if latitude and longitude columns exist
    has_coords = 'latitude' in projects_df.columns and 'longitude' in projects_df.columns
    
    if not has_coords:
        # Try alternative column names
        alt_names = [('lat', 'lon'), ('lat', 'lng'), ('latitude', 'longitude')]
        found = False
        
        for lat_col, lon_col in alt_names:
            if lat_col in projects_df.columns and lon_col in projects_df.columns:
                projects_df = projects_df.rename(columns={lat_col: 'latitude', lon_col: 'longitude'})
                has_coords = True
                found = True
                break
                
        if not found:
            # Create empty figure with message
            fig = go.Figure()
            fig.add_annotation(
                text="Project data missing location coordinates",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False
            )
            return fig
    
    # Create map visualization
    if color_by in projects_df.columns and size_by in projects_df.columns:
        fig = px.scatter_mapbox(
            projects_df,
            lat="latitude",
            lon="longitude",
            color=color_by,
            size=size_by,
            hover_name="project_name" if "project_name" in projects_df.columns else None,
            zoom=3,
            height=600,
            title="Project Locations"
        )
        
        fig.update_layout(
            mapbox_style="open-street-map",
            margin=dict(l=0, r=0, t=50, b=0)
        )
    else:
        # Fallback to simple map without color/size differentiation
        fig = px.scatter_mapbox(
            projects_df,
            lat="latitude",
            lon="longitude",
            hover_name="project_name" if "project_name" in projects_df.columns else None,
            zoom=3,
            height=600,
            title="Project Locations"
        )
        
        fig.update_layout(
            mapbox_style="open-street-map",
            margin=dict(l=0, r=0, t=50, b=0)
        )
    
    return fig

def create_resource_heatmap(resources_df, time_period='month'):
    """
    Creates a resource allocation heatmap
    
    Parameters:
    -----------
    resources_df : pandas DataFrame
        DataFrame containing resource allocation data
    time_period : str
        Time period for aggregation ('day', 'week', 'month', 'quarter')
        
    Returns:
    --------
    plotly.graph_objects.Figure
        Heatmap visualization
    """
    if resources_df.empty:
        # Create empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No resource allocation data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        return fig
    
    # Aggregate data by resource and time period
    if 'resource_name' in resources_df.columns and 'date' in resources_df.columns and 'allocation' in resources_df.columns:
        # Convert date column to datetime if it's not already
        if not pd.api.types.is_datetime64_any_dtype(resources_df['date']):
            resources_df['date'] = pd.to_datetime(resources_df['date'])
        
        # Create time period column based on specified aggregation
        if time_period == 'day':
            resources_df['period'] = resources_df['date']
        elif time_period == 'week':
            resources_df['period'] = resources_df['date'].dt.to_period('W').dt.start_time
        elif time_period == 'month':
            resources_df['period'] = resources_df['date'].dt.to_period('M').dt.start_time
        elif time_period == 'quarter':
            resources_df['period'] = resources_df['date'].dt.to_period('Q').dt.start_time
        else:
            # Default to month
            resources_df['period'] = resources_df['date'].dt.to_period('M').dt.start_time
        
        # Aggregate allocations
        pivot_data = resources_df.pivot_table(
            index='resource_name',
            columns='period',
            values='allocation',
            aggfunc='sum',
            fill_value=0
        )
        
        # Create heatmap
        fig = px.imshow(
            pivot_data,
            color_continuous_scale=[
                [0, 'green'],
                [0.5, 'yellow'],
                [0.8, 'orange'],
                [1, 'red']
            ],
            labels=dict(x="Time Period", y="Resource", color="Allocation %"),
            title=f"Resource Allocation by {time_period.capitalize()}"
        )
        
        # Add annotations for overallocated resources (>100%)
        for i, resource in enumerate(pivot_data.index):
            for j, period in enumerate(pivot_data.columns):
                value = pivot_data.iloc[i, j]
                if value > 100:
                    fig.add_annotation(
                        x=j,
                        y=i,
                        text="!",
                        showarrow=False,
                        font=dict(color="white", size=12, family="Arial Black")
                    )
        
        # Customize layout
        fig.update_layout(
            height=max(400, len(pivot_data) * 30),  # Dynamic height based on number of resources
            yaxis={'categoryorder':'category ascending'},
            plot_bgcolor='white',
            margin=dict(l=10, r=10, t=50, b=10)
        )
    else:
        # Create empty figure with error message
        fig = go.Figure()
        fig.add_annotation(
            text="Resource data missing required columns (resource_name, date, allocation)",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
    
    return fig

def create_performance_radar(project_data, metrics=None, comparison_data=None):
    """
    Creates a radar chart of project performance metrics
    
    Parameters:
    -----------
    project_data : pandas Series
        Data for a specific project
    metrics : list, optional
        List of metrics to include in the radar chart
    comparison_data : pandas Series, optional
        Data for comparison (e.g., portfolio average)
        
    Returns:
    --------
    plotly.graph_objects.Figure
        Radar chart visualization
    """
    if project_data is None:
        # Create empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No project data available",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        return fig
    
    # Default metrics if not specified
    if metrics is None:
        metrics = [
            'spi', 'cpi', 'quality_index', 'risk_score', 
            'resource_efficiency', 'stakeholder_satisfaction'
        ]
    
    # Filter to metrics that exist in the data
    available_metrics = [m for m in metrics if m in project_data.index]
    
    if not available_metrics:
        # Create empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No performance metrics available for radar chart",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        return fig
    
    # Normalize values to 0-1 scale for radar chart
    # For metrics like SPI and CPI, 1.0 is the target
    normalized_values = []
    for metric in available_metrics:
        value = project_data[metric]
        
        # Apply specific normalization based on metric type
        if metric in ['spi', 'cpi']:
            # For SPI/CPI, scale around 1.0
            # <0.8 is bad (0), 1.0 is good (0.8), >1.2 is excellent (1.0)
            if value < 0.8:
                norm_value = value / 0.8 * 0.5  # 0-0.5 range for below threshold
            elif value < 1.0:
                norm_value = 0.5 + (value - 0.8) / 0.2 * 0.3  # 0.5-0.8 range
            elif value <= 1.2:
                norm_value = 0.8 + (value - 1.0) / 0.2 * 0.2  # 0.8-1.0 range for above target
            else:
                norm_value = 1.0  # Cap at 1.0
        elif metric == 'risk_score':
            # For risk_score, lower is better (inverse)
            # Assumes risk_score is 0-1 where 0 is no risk
            norm_value = 1.0 - value
        else:
            # For other metrics, normalize to 0-1 range
            # Assumes values are already in a meaningful scale
            norm_value = min(1.0, max(0.0, value))
        
        normalized_values.append(norm_value)
    
    # Create radar chart
    fig = go.Figure()
    
    # Add project data trace
    fig.add_trace(go.Scatterpolar(
        r=normalized_values,
        theta=available_metrics,
        fill='toself',
        name=project_data.name if hasattr(project_data, 'name') else 'Project',
        line=dict(color=ARCADIS_ORANGE)
    ))
    
    # Add comparison data if provided
    if comparison_data is not None:
        comparison_values = []
        for metric in available_metrics:
            if metric in comparison_data.index:
                value = comparison_data[metric]
                
                # Apply same normalization as above
                if metric in ['spi', 'cpi']:
                    if value < 0.8:
                        norm_value = value / 0.8 * 0.5
                    elif value < 1.0:
                        norm_value = 0.5 + (value - 0.8) / 0.2 * 0.3
                    elif value <= 1.2:
                        norm_value = 0.8 + (value - 1.0) / 0.2 * 0.2
                    else:
                        norm_value = 1.0
                elif metric == 'risk_score':
                    norm_value = 1.0 - value
                else:
                    norm_value = min(1.0, max(0.0, value))
                
                comparison_values.append(norm_value)
            else:
                comparison_values.append(0)
        
        fig.add_trace(go.Scatterpolar(
            r=comparison_values,
            theta=available_metrics,
            fill='toself',
            name=comparison_data.name if hasattr(comparison_data, 'name') else 'Comparison',
            line=dict(color=ARCADIS_TEAL)
        ))
    
    # Add a reference circle at "target" level
    fig.add_trace(go.Scatterpolar(
        r=[0.8] * len(available_metrics),
        theta=available_metrics,
        fill=None,
        mode='lines',
        line=dict(color='rgba(0,0,0,0.3)', dash='dash'),
        name='Target Level'
    ))
    
    # Customize layout
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 1]
            )
        ),
        showlegend=True,
        title="Performance Metrics Radar",
        height=500,
        width=500
    )
    
    return fig

def create_custom_dashboard(metrics_dict, layout_cols=2):
    """
    Creates a custom dashboard with multiple metrics
    
    Parameters:
    -----------
    metrics_dict : dict
        Dictionary of metric names and values
    layout_cols : int
        Number of columns in the layout
        
    Returns:
    --------
    plotly.graph_objects.Figure
        Dashboard visualization
    """
    if not metrics_dict:
        # Create empty figure with message
        fig = go.Figure()
        fig.add_annotation(
            text="No metrics provided for dashboard",
            xref="paper", yref="paper",
            x=0.5, y=0.5,
            showarrow=False
        )
        return fig
    
    # Determine layout
    num_metrics = len(metrics_dict)
    num_rows = (num_metrics + layout_cols - 1) // layout_cols  # Ceiling division
    
    # Create subplot figure
    fig = go.Figure()
    
    # Add individual metric visuals
    for i, (metric_name, metric_data) in enumerate(metrics_dict.items()):
        row = i // layout_cols + 1
        col = i % layout_cols + 1
        
        # Create an appropriate visualization based on metric type
        if isinstance(metric_data, (int, float)):
            # Simple numeric metric - use gauge
            fig.add_trace(go.Indicator(
                mode="gauge+number",
                value=metric_data['value'] if isinstance(metric_data, dict) else metric_data,
                title={"text": metric_name},
                domain={'row': row, 'column': col}
            ))
        elif isinstance(metric_data, dict) and 'trend' in metric_data:
            # Time series data - use sparkline
            x_data = list(range(len(metric_data['trend'])))
            y_data = metric_data['trend']
            
            fig.add_trace(go.Scatter(
                x=x_data,
                y=y_data,
                mode='lines',
                name=metric_name,
                line=dict(color=ARCADIS_TEAL)
            ))
        elif isinstance(metric_data, dict) and 'categories' in metric_data and 'values' in metric_data:
            # Categorical data - use bar chart
            fig.add_trace(go.Bar(
                x=metric_data['categories'],
                y=metric_data['values'],
                name=metric_name
            ))
    
    # Customize layout
    fig.update_layout(
        grid=dict(rows=num_rows, columns=layout_cols),
        title="Custom Dashboard",
        height=300 * num_rows,
        showlegend=False
    )
    
    return fig