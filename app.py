# -*- coding: utf-8 -*-
"""
ARCADIS PMO PULSE

A comprehensive Project Management Office (PMO) dashboard with advanced
visualization, analytics, and reporting capabilities for portfolio and
project performance tracking.

Author: Arcadis PMO Analytics Team
Version: 1.0
"""

# --- Import Libraries ---
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
import time
import io
import base64
import os
import json
import joblib
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor, IsolationForest
from sklearn.preprocessing import StandardScaler
import scipy.stats as stats
import xgboost as xgb

# App modules
from utils.data_generator import generate_mock_data
from utils.data_processor import (
    calculate_portfolio_metrics,
    calculate_project_metrics,
    get_projects_needing_attention,
    prepare_forecast_data
)
from utils.visualizations import (
    create_gauge_chart,
    create_trend_chart,
    create_portfolio_overview,
    create_risk_matrix,
    create_schedule_gantt
)
from utils.export import export_to_excel, export_to_powerpoint, generate_data_template

# Advanced features and enhanced analytics
from utils.predictive_analytics import (
    train_forecast_model,
    predict_project_completion,
    detect_anomalies,
    generate_natural_language_insights
)

# Enhanced Feature Modules
from utils.advanced_analytics import (
    detect_project_anomalies,
    optimize_resource_allocation,
    forecast_project_completion,
    generate_nlp_insights
)

from utils.enhanced_visualizations import (
    create_interactive_gantt,
    create_geospatial_map,
    create_resource_heatmap,
    create_performance_radar,
    create_custom_dashboard,
    get_theme_colors
)

from utils.data_integration import (
    DataIntegrationManager,
    JiraConnector,
    MSProjectConnector,
    ExcelConnector,
    RestApiConnector
)

from utils.collaboration import (
    CollaborationManager,
    CommentSystem,
    ActionTracker,
    AlertSystem,
    MeetingMode,
    ReportDistribution
)

# --- Page Configuration ---
st.set_page_config(
    page_title="ARCADIS PMO PULSE",
    page_icon="assets/images/favicon.svg",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Constants & Configuration ---
# Arcadis Brand Colors
ARCADIS_ORANGE = "#E67300"
ARCADIS_BLACK = "#000000"
ARCADIS_GREY = "#6c757d"
ARCADIS_DARK_GREY = "#646469"
ARCADIS_WHITE = "#FFFFFF"
ARCADIS_LIGHT_GREY = "#F5F5F5"
ARCADIS_TEAL = "#00A3A1"

COLOR_SUCCESS = "#2ECC71"
COLOR_WARNING = "#F1C40F"
COLOR_DANGER = "#E74C3C"
COLOR_INFO = "#3498DB"

# --- CSS Styling ---
with open('assets/app_styles.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# Add Font Awesome for icons
st.markdown(
    """
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    """, 
    unsafe_allow_html=True
)

# --- Initialize Session State ---
def initialize_session_state():
    """Initialize all session state variables"""
    
    # Core app state
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    if 'current_view' not in st.session_state:
        st.session_state.current_view = 'portfolio'
    if 'selected_project' not in st.session_state:
        st.session_state.selected_project = None
    if 'spi_threshold' not in st.session_state:
        st.session_state.spi_threshold = 0.90
    if 'cpi_threshold' not in st.session_state:
        st.session_state.cpi_threshold = 0.90
    if 'portfolio_kpis' not in st.session_state:
        st.session_state.portfolio_kpis = {}
    
    # Data frames
    if 'projects_df' not in st.session_state:
        st.session_state.projects_df = None
    if 'tasks_df' not in st.session_state:
        st.session_state.tasks_df = None
    if 'risks_df' not in st.session_state:
        st.session_state.risks_df = None
    if 'changes_df' not in st.session_state:
        st.session_state.changes_df = None
    if 'history_df' not in st.session_state:
        st.session_state.history_df = None
    if 'benefits_df' not in st.session_state:
        st.session_state.benefits_df = None
    if 'resources_df' not in st.session_state:
        st.session_state.resources_df = None
        
    # UI state
    if 'show_upload' not in st.session_state:
        st.session_state.show_upload = False
    if 'color_theme' not in st.session_state:
        st.session_state.color_theme = "Arcadis Default"
    if 'show_advanced_options' not in st.session_state:
        st.session_state.show_advanced_options = False
        
    # Feature visibility toggles
    if 'show_spi' not in st.session_state:
        st.session_state.show_spi = True
    if 'show_cpi' not in st.session_state:
        st.session_state.show_cpi = True
    if 'show_risks' not in st.session_state:
        st.session_state.show_risks = True
    if 'show_forecasts' not in st.session_state:
        st.session_state.show_forecasts = True
    if 'show_insights' not in st.session_state:
        st.session_state.show_insights = True
    if 'show_details' not in st.session_state:
        st.session_state.show_details = True
        
    # Analysis settings
    if 'forecast_horizon' not in st.session_state:
        st.session_state.forecast_horizon = 3  # months
    if 'anomaly_sensitivity' not in st.session_state:
        st.session_state.anomaly_sensitivity = 5  # scale 1-10
    if 'date_range' not in st.session_state:
        today = datetime.now().date()
        six_months_ago = today - relativedelta(months=6)
        st.session_state.date_range = (six_months_ago, today)
        
    # Cached data and analysis results
    if 'anomalies' not in st.session_state:
        st.session_state.anomalies = None
    if 'resource_optimization' not in st.session_state:
        st.session_state.resource_optimization = None
    if 'project_forecasts' not in st.session_state:
        st.session_state.project_forecasts = {}
    if 'portfolio_insights' not in st.session_state:
        st.session_state.portfolio_insights = []
    if 'project_insights' not in st.session_state:
        st.session_state.project_insights = {}
        
    # Managers for enhanced features
    if 'data_integration_manager' not in st.session_state:
        st.session_state.data_integration_manager = DataIntegrationManager()
    if 'collaboration_manager' not in st.session_state:
        st.session_state.collaboration_manager = CollaborationManager()
    if 'comment_system' not in st.session_state:
        st.session_state.comment_system = CommentSystem(st.session_state.collaboration_manager)
    if 'action_tracker' not in st.session_state:
        st.session_state.action_tracker = ActionTracker(st.session_state.collaboration_manager)
    if 'alert_system' not in st.session_state:
        st.session_state.alert_system = AlertSystem(st.session_state.collaboration_manager)
    if 'meeting_mode' not in st.session_state:
        st.session_state.meeting_mode = MeetingMode(st.session_state.collaboration_manager)
    if 'report_distribution' not in st.session_state:
        st.session_state.report_distribution = ReportDistribution(st.session_state.collaboration_manager)
        
    # User-specific state
    if 'user_role' not in st.session_state:
        st.session_state.user_role = "PMO Manager"
    if 'role_initialized' not in st.session_state:
        st.session_state.role_initialized = None
    
# Initialize session state
initialize_session_state()

# --- Helper Functions ---
def load_data():
    """Load mock data and calculate initial metrics"""
    with st.spinner("Loading project data..."):
        # Generate mock data
        projects_df, tasks_df, risks_df, changes_df, history_df, benefits_df = generate_mock_data()
        
        # Store dataframes in session state
        st.session_state.projects_df = projects_df
        st.session_state.tasks_df = tasks_df
        st.session_state.risks_df = risks_df
        st.session_state.changes_df = changes_df
        st.session_state.history_df = history_df
        st.session_state.benefits_df = benefits_df
        
        # Calculate portfolio metrics
        st.session_state.portfolio_kpis = calculate_portfolio_metrics(
            projects_df, tasks_df, history_df
        )
        
        # Set data_loaded flag
        st.session_state.data_loaded = True
        
        # Set selected project to the first project if it's not already set
        if st.session_state.selected_project is None and not projects_df.empty:
            st.session_state.selected_project = projects_df['project_name'].iloc[0]
        
        return True
    return False

def show_header():
    """Display the page header with logos and titles"""
    col1, col2 = st.columns([1, 4])
    with col1:
        st.image("assets/images/logo.svg", width=180)
    with col2:
        st.markdown("<h1 style='color: #E67300; margin-top: 10px;'>ARCADIS PMO PULSE</h1>", unsafe_allow_html=True)
        st.markdown("<p style='font-size: 16px; opacity: 0.8;'>Comprehensive project portfolio performance tracking and analytics</p>", unsafe_allow_html=True)
    
    # Add a modern divider with gradient
    st.markdown("""
        <div style="height: 4px; background: linear-gradient(to right, #E67300, #00A3A1); margin: 10px 0 30px 0; border-radius: 2px;"></div>
    """, unsafe_allow_html=True)

def display_welcome():
    """Display welcome message and dashboard overview"""
    st.markdown(
        """
        <div class="welcome-section">
            <h3><i class="fas fa-chart-line"></i> Welcome to PMO Pulse</h3>
            <p>
                PMO Pulse provides real-time insights into your project portfolio performance, 
                helping project managers and executives make informed decisions. 
                Track your KPIs, identify risks, and forecast future performance all in one place.
            </p>
        </div>
        """, 
        unsafe_allow_html=True
    )
    
    # Display dashboard capabilities
    st.markdown(
        """
        <div class="welcome-section">
            <h3><i class="fas fa-tools"></i> Dashboard Capabilities</h3>
            <ul class="styled-list">
                <li><b>Portfolio Overview:</b> Get a bird's-eye view of all your projects and their performance metrics</li>
                <li><b>Project Deep Dive:</b> Analyze individual project metrics, tasks, and timelines</li>
                <li><b>Risk Assessment:</b> Identify and track project risks before they impact delivery</li>
                <li><b>Performance Trends:</b> Visualize SPI and CPI trends over time</li>
                <li><b>Forecasting:</b> Predict project outcomes based on current performance</li>
                <li><b>Report Generation:</b> Export data and visualizations for presentations and reports</li>
            </ul>
        </div>
        """, 
        unsafe_allow_html=True
    )
    
    # Display dashboard images
    col1, col2 = st.columns(2)
    
    with col1:
        st.image("https://pixabay.com/get/g00a5c90128faa8d724cc9833de8842e4cf663dd3f6309e1c9a81d97a89c4cb68201165f235a6d0a7488e40f0b07e83d3ff3e2008474e5397e453663eb0b25561_1280.jpg", 
                 caption="Project Management Dashboard", use_container_width=True)
    with col2:
        st.image("https://pixabay.com/get/g545fb3673e3c1f0161c11de22f61cdeb13734ceb04b7f5d8270f983213fffdaad13cc9ddf38949e775cf7e98be026714a691ed13ebd2e19b8c8dab05c14ce27e_1280.jpg", 
                 caption="Data Visualization Charts", use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Display instructions
    st.markdown(
        """
        <div class="welcome-section">
            <h3><i class="fas fa-info-circle"></i> Getting Started</h3>
            <p>
                To begin exploring your project data, press the <b>"Load Mock Data"</b> button in the sidebar.
                Then, use the navigation tabs to explore different aspects of your portfolio.
            </p>
        </div>
        """, 
        unsafe_allow_html=True
    )

def display_sidebar():
    """Display sidebar with navigation, controls and customization options"""
    with st.sidebar:
        st.markdown("## PMO Pulse")
        
        # User Role Selection (new feature)
        st.markdown("### User Profile")
        role = st.selectbox(
            "Select View Type:",
            options=["Executive", "PMO Manager", "Project Manager", "Team Member", "Custom"],
            index=["Executive", "PMO Manager", "Project Manager", "Team Member", "Custom"].index(st.session_state.user_role),
            help="Set your user role to customize the dashboard view",
            key="user_role_select"
        )
        
        # Update user role in session state if changed
        if role != st.session_state.user_role:
            st.session_state.user_role = role
            st.session_state.role_initialized = None  # Reset so role-specific settings can be applied
        
        # Set up role-specific defaults if not already set
        if st.session_state.role_initialized != role:
            if role == "Executive":
                # Executive view focuses on high-level portfolio metrics and forecasts
                st.session_state.show_spi = True
                st.session_state.show_cpi = True
                st.session_state.show_risks = True
                st.session_state.show_forecasts = True
                st.session_state.show_insights = True
                st.session_state.show_details = False
                st.session_state.color_theme = "Arcadis Default"
            elif role == "PMO Manager":
                # PMO view includes all metrics and detailed tracking
                st.session_state.show_spi = True
                st.session_state.show_cpi = True
                st.session_state.show_risks = True
                st.session_state.show_forecasts = True
                st.session_state.show_insights = True
                st.session_state.show_details = True
                st.session_state.color_theme = "Arcadis Default"
            elif role == "Project Manager":
                # Project Manager focuses on their specific projects
                st.session_state.show_spi = True
                st.session_state.show_cpi = True
                st.session_state.show_risks = True
                st.session_state.show_forecasts = True
                st.session_state.show_insights = True
                st.session_state.show_details = True
                st.session_state.color_theme = "Colorful"
            elif role == "Team Member":
                # Team member has limited view focused on tasks
                st.session_state.show_spi = True
                st.session_state.show_cpi = False
                st.session_state.show_risks = False
                st.session_state.show_forecasts = False
                st.session_state.show_insights = True
                st.session_state.show_details = True
                st.session_state.color_theme = "Monochrome"
            
            # Mark that we've initialized for this role
            st.session_state.role_initialized = role
        
        st.markdown("---")
        
        # Data Loading Section
        st.markdown("### Data Options")
        if not st.session_state.data_loaded:
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Load Mock Data", key="load_data"):
                    success = load_data()
                    if success:
                        st.success("Data loaded successfully!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("Failed to load data. Please try again.")
            with col2:
                if st.button("Upload Data", key="upload_data"):
                    st.session_state.show_upload = True
                    st.rerun()
            
            # Add download template option
            st.markdown("#### Data Templates")
            template_data = generate_data_template()
            st.download_button(
                label="Download Data Template",
                data=template_data,
                file_name="pmo_dashboard_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download an Excel template with the required format for data upload"
            )
        else:
            st.success("✅ Data loaded")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Refresh Data", key="refresh_data"):
                    load_data()
                    st.success("Data refreshed!")
                    time.sleep(1)
                    st.rerun()
            with col2:
                if st.button("Upload Data", key="upload_data_loaded"):
                    st.session_state.show_upload = True
                    st.rerun()
            
            # Add download template option even when data is loaded
            with st.expander("Data Templates"):
                st.info("Download templates to prepare your data for import")
                template_data = generate_data_template()
                st.download_button(
                    label="Download Data Template",
                    data=template_data,
                    file_name="pmo_dashboard_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download an Excel template with the required format for data upload"
                )
                    
            # API Connection option (new feature)
            with st.expander("External Data Sources"):
                source_type = st.selectbox(
                    "Connect to:",
                    options=["None", "Jira", "Microsoft Project", "Excel Online", "Custom API"],
                    index=0,
                    key="external_source_type"
                )
                
                # Only show connection fields if an external source is selected
                if source_type != "None":
                    source_name = st.text_input("Source Name:", value=f"My {source_type}")
                    
                    # Connection-specific fields based on source type
                    if source_type == "Jira":
                        url = st.text_input("Jira URL:", placeholder="https://your-domain.atlassian.net")
                        auth_type = st.radio("Authentication:", ["API Token", "Basic Auth"], horizontal=True)
                        if auth_type == "API Token":
                            username = st.text_input("Email:")
                            api_token = st.text_input("API Token:", type="password")
                        else:
                            username = st.text_input("Username:")
                            password = st.text_input("Password:", type="password")
                            
                    elif source_type == "Microsoft Project":
                        url = st.text_input("Project Online URL:", placeholder="https://your-tenant.sharepoint.com/sites/pwa")
                        token = st.text_input("Access Token:", type="password")
                        
                    elif source_type == "Excel Online":
                        url = st.text_input("Excel File URL:", placeholder="https://your-tenant.sharepoint.com/sites/site/Documents/file.xlsx")
                        token = st.text_input("Access Token:", type="password")
                        
                    elif source_type == "Custom API":
                        url = st.text_input("API Base URL:", placeholder="https://api.example.com")
                        auth_type = st.radio("Authentication:", ["None", "API Key", "Bearer Token", "Basic Auth"], horizontal=True)
                        
                        if auth_type == "API Key":
                            key_location = st.radio("API Key Location:", ["Header", "Query Parameter"], horizontal=True)
                            key_name = st.text_input("Parameter Name:", value="api_key")
                            api_key = st.text_input("API Key:", type="password")
                        elif auth_type == "Bearer Token":
                            token = st.text_input("Bearer Token:", type="password")
                        elif auth_type == "Basic Auth":
                            username = st.text_input("Username:")
                            password = st.text_input("Password:", type="password")
                    
                    # Add connection button
                    if st.button("Connect Source"):
                        # This would add the source to the data integration manager
                        st.session_state.data_source_added = True
                        st.success(f"Connected to {source_type}")
                        time.sleep(1)
                        st.rerun()
        
        st.markdown("---")
        
        # Navigation Section
        if st.session_state.data_loaded:
            st.markdown("### Navigation")
            
            # View selection
            view_options = ["Portfolio View", "Project Details", "Risk Management", "Analytics & Forecasting"]
            
            # Add collaboration view if not in Team Member role 
            if role != "Team Member":
                view_options.append("Collaboration Hub")
            
            selected_view = st.radio("Select View:", view_options)
            
            # Map selection to internal state
            view_mapping = {
                "Portfolio View": "portfolio",
                "Project Details": "project",
                "Risk Management": "risk",
                "Analytics & Forecasting": "analytics",
                "Collaboration Hub": "collaboration"
            }
            st.session_state.current_view = view_mapping[selected_view]
            
            # Project filtering and selection (enhanced)
            if st.session_state.current_view in ["project", "risk"]:
                with st.expander("Project Filters", expanded=False):
                    # Add sector/department filter (new feature)
                    if 'sector' in st.session_state.projects_df.columns:
                        all_sectors = sorted(st.session_state.projects_df['sector'].unique().tolist())
                        selected_sector = st.selectbox(
                            "Filter by Sector:",
                            options=["All Sectors"] + all_sectors,
                            index=0,
                            help="Filter projects by sector/department"
                        )
                        
                        if selected_sector != "All Sectors":
                            filtered_projects = st.session_state.projects_df[st.session_state.projects_df['sector'] == selected_sector]
                        else:
                            filtered_projects = st.session_state.projects_df
                    else:
                        filtered_projects = st.session_state.projects_df
                    
                    # Project manager filter (new feature)
                    if 'project_manager' in st.session_state.projects_df.columns:
                        all_pms = sorted(filtered_projects['project_manager'].unique().tolist())
                        selected_pm = st.selectbox(
                            "Filter by Project Manager:",
                            options=["All Project Managers"] + all_pms,
                            index=0,
                            help="Filter projects by their assigned manager"
                        )
                        
                        if selected_pm != "All Project Managers":
                            filtered_projects = filtered_projects[filtered_projects['project_manager'] == selected_pm]
                    
                    # Status filter (new feature)
                    if 'status' in st.session_state.projects_df.columns:
                        all_statuses = sorted(filtered_projects['status'].unique().tolist())
                        selected_status = st.multiselect(
                            "Filter by Status:",
                            options=all_statuses,
                            default=all_statuses
                        )
                        
                        if selected_status and len(selected_status) < len(all_statuses):
                            filtered_projects = filtered_projects[filtered_projects['status'].isin(selected_status)]
                
                # Project selection from filtered list
                if not filtered_projects.empty:
                    project_options = filtered_projects['project_name'].tolist()
                    default_index = 0
                    
                    # Try to maintain previously selected project if it's in the filtered list
                    if st.session_state.selected_project in project_options:
                        default_index = project_options.index(st.session_state.selected_project)
                    
                    selected_project = st.selectbox(
                        "Select Project:",
                        options=project_options,
                        index=default_index
                    )
                    st.session_state.selected_project = selected_project
                else:
                    st.warning("No projects match the selected filters.")
            
            st.markdown("---")
            
            # Customization & Settings section (new feature)
            with st.expander("Dashboard Customization", expanded=False):
                st.markdown("#### Select Visible Metrics")
                
                # Metric visibility toggles
                st.session_state.show_spi = st.checkbox("Schedule Performance", value=st.session_state.show_spi)
                st.session_state.show_cpi = st.checkbox("Cost Performance", value=st.session_state.show_cpi)
                st.session_state.show_risks = st.checkbox("Risk Analytics", value=st.session_state.show_risks)
                st.session_state.show_forecasts = st.checkbox("Forecasts & Predictions", value=st.session_state.show_forecasts)
                st.session_state.show_insights = st.checkbox("AI-Generated Insights", value=st.session_state.show_insights)
                st.session_state.show_details = st.checkbox("Detailed Metrics", value=st.session_state.show_details)
                
                # Visual preferences
                st.markdown("#### Visual Preferences")
                
                # Chart type preferences
                chart_types = ["Line Chart", "Bar Chart", "Area Chart", "Scatter Plot"]
                default_chart = st.selectbox(
                    "Default Chart Type:", 
                    options=chart_types, 
                    index=0
                )
                
                # Color theme customization
                color_themes = ["Arcadis Default", "Monochrome", "Colorful", "High Contrast"]
                st.session_state.color_theme = st.selectbox(
                    "Color Theme:", 
                    options=color_themes, 
                    index=color_themes.index(st.session_state.color_theme)
                )
            
            # Filters & Thresholds section
            with st.expander("Filters & Thresholds"):
                # Date range filter
                today = datetime.now().date()
                six_months_ago = today - relativedelta(months=6)
                
                st.session_state.date_range = st.date_input(
                    "Date Range:",
                    value=st.session_state.date_range if isinstance(st.session_state.date_range, tuple) else (six_months_ago, today),
                    min_value=today - relativedelta(years=3),
                    max_value=today
                )
                
                # KPI Thresholds
                st.markdown("#### KPI Thresholds")
                
                st.session_state.spi_threshold = st.slider(
                    "SPI Threshold:",
                    min_value=0.7,
                    max_value=1.0,
                    value=st.session_state.spi_threshold,
                    step=0.01,
                    help="Projects with SPI below this threshold will be flagged"
                )
                
                st.session_state.cpi_threshold = st.slider(
                    "CPI Threshold:",
                    min_value=0.7,
                    max_value=1.0,
                    value=st.session_state.cpi_threshold,
                    step=0.01,
                    help="Projects with CPI below this threshold will be flagged"
                )
                
                # Advanced analysis options (new feature)
                st.markdown("#### Advanced Options")
                st.session_state.forecast_horizon = st.slider(
                    "Forecast Horizon (months)",
                    min_value=1,
                    max_value=12,
                    value=st.session_state.forecast_horizon,
                    help="Number of months to forecast into the future"
                )
                
                st.session_state.anomaly_sensitivity = st.slider(
                    "Anomaly Detection Sensitivity",
                    min_value=1,
                    max_value=10,
                    value=st.session_state.anomaly_sensitivity,
                    help="Sensitivity of anomaly detection (higher = more sensitive)"
                )
            
            # Export Options (new feature)
            with st.expander("Export & Share", expanded=False):
                if st.button("Export to Excel"):
                    output = export_to_excel(
                        st.session_state.projects_df,
                        st.session_state.tasks_df,
                        st.session_state.risks_df,
                        st.session_state.history_df
                    )
                    st.download_button(
                        label="Download Excel",
                        data=output,
                        file_name="pmo_pulse_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                if st.button("Export to PowerPoint"):
                    output = export_to_powerpoint(
                        st.session_state.projects_df,
                        st.session_state.tasks_df,
                        st.session_state.risks_df,
                        st.session_state.history_df,
                        st.session_state.portfolio_kpis
                    )
                    st.download_button(
                        label="Download PowerPoint",
                        data=output,
                        file_name="pmo_pulse_report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            
            st.markdown("---")

def create_geospatial_map(projects_df, color_by='status', size_by='budget'):
    """
    Create an interactive geospatial map of projects
    
    Parameters:
    -----------
    projects_df : pandas.DataFrame
        Project data with latitude and longitude columns
    color_by : str
        Column to use for coloring points
    size_by : str
        Column to use for sizing points
        
    Returns:
    --------
    plotly.graph_objects.Figure
        Map figure
    """
    if 'latitude' not in projects_df.columns or 'longitude' not in projects_df.columns:
        # If geo data is missing, return a placeholder map
        fig = px.scatter_mapbox(
            lat=[39.8283], lon=[-98.5795], # Geographic center of USA
            zoom=3
        )
        fig.update_layout(
            mapbox_style="open-street-map",
            margin=dict(l=0, r=0, t=0, b=0),
            height=500
        )
        return fig
    
    # Color mapping based on status
    if color_by == 'status':
        color_map = {
            'On Track': '#00A3A1',      # Teal
            'Minor Issues': '#FFC107',  # Amber 
            'At Risk': '#FF9800',       # Orange
            'Delayed': '#E67300',       # Burnt Orange
            'Completed': '#28A745'      # Green
        }
        color_discrete_map = color_map
    else:
        color_discrete_map = None
    
    # Size reference based on budget
    if size_by == 'budget':
        size_max = 25
        sizes = projects_df[size_by] / projects_df[size_by].max() * size_max + 8
    else:
        sizes = 15  # Default size
    
    # Create the map
    hover_data = {
        'project_name': True,
        'sector': True,
        'budget': True,
        'status': True,
        'project_manager': True,
        'latitude': False,
        'longitude': False
    }
    
    # Format budget as currency
    if 'budget' in projects_df.columns:
        projects_df['budget_display'] = projects_df['budget'].apply(lambda x: f"${x:,.0f}")
        hover_data['budget_display'] = True
        hover_data['budget'] = False
    
    fig = px.scatter_mapbox(
        projects_df,
        lat="latitude",
        lon="longitude",
        color=color_by,
        size_max=25,
        zoom=3,
        hover_name="project_name",
        hover_data=hover_data,
        color_discrete_map=color_discrete_map
    )
    
    # Update marker size
    fig.update_traces(marker=dict(size=sizes))
    
    # Layout
    fig.update_layout(
        mapbox_style="open-street-map",
        margin=dict(l=0, r=0, t=0, b=0),
        height=500,
        legend_title_text=color_by.replace('_', ' ').title()
    )
    
    return fig

def display_portfolio_view():
    """Display portfolio overview with KPIs and charts"""
    st.markdown("## Portfolio Overview")
    
    # Check if data is loaded and KPIs are available
    if not hasattr(st.session_state, 'portfolio_kpis') or st.session_state.portfolio_kpis is None:
        st.warning("Please load data first to see portfolio metrics.")
        return
        
    # Create tabs for different views
    portfolio_tabs = st.tabs(["Dashboard", "Geographic View", "Financial Analysis"])
    
    with portfolio_tabs[0]:
        # Summary KPIs
        kpis = st.session_state.portfolio_kpis
        
    # Geographic View Tab
    with portfolio_tabs[1]:
        st.markdown("### Project Locations")
        
        # Get the filtered projects dataframe
        projects_df = st.session_state.projects_df
        
        if 'latitude' not in projects_df.columns or 'longitude' not in projects_df.columns:
            st.warning("Geographic data not available. Please ensure your project data includes latitude and longitude coordinates.")
        else:
            # Controls for the map
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col1:
                color_by = st.selectbox(
                    "Color projects by:", 
                    options=["status", "sector", "strategic_alignment"],
                    index=0
                )
            
            with col2:
                size_by = st.selectbox(
                    "Size projects by:", 
                    options=["budget", "uniform"],
                    index=0
                )
            
            # Create the map
            try:
                map_fig = create_geospatial_map(
                    projects_df, 
                    color_by=color_by, 
                    size_by=size_by
                )
                st.plotly_chart(map_fig, use_container_width=True, key="portfolio_geo_map")
                
                # Add map insights
                col1, col2 = st.columns(2)
                with col1:
                    # Add some statistics about projects by location
                    st.markdown("#### Projects by Location")
                    if 'location' in projects_df.columns:
                        location_counts = projects_df['location'].value_counts().head(5)
                        for location, count in location_counts.items():
                            st.markdown(f"**{location}**: {count} projects")
                    else:
                        st.markdown("Location data not available")
                
                with col2:
                    # Add interesting insights based on geographical distribution
                    st.markdown("#### Geographical Insights")
                    
                    # Calculate average SPI and CPI by sector if available
                    if all(col in projects_df.columns for col in ['sector', 'spi', 'cpi']):
                        perf_by_sector = projects_df.groupby('sector')[['spi', 'cpi']].mean()
                        
                        # Find best and worst performing sectors
                        if not perf_by_sector.empty:
                            best_sector = perf_by_sector['spi'].idxmax()
                            worst_sector = perf_by_sector['spi'].idxmin()
                            
                            st.markdown(f"**Best performing sector**: {best_sector} (SPI: {perf_by_sector.loc[best_sector, 'spi']:.2f})")
                            st.markdown(f"**Needs attention**: {worst_sector} (SPI: {perf_by_sector.loc[worst_sector, 'spi']:.2f})")
                    
                    # Regional distribution analysis
                    st.markdown("#### Regional Distribution")
                    if 'sector' in projects_df.columns:
                        # Pie chart of projects by sector
                        sector_counts = projects_df['sector'].value_counts()
                        fig = px.pie(
                            values=sector_counts.values, 
                            names=sector_counts.index,
                            title="Projects by Sector"
                        )
                        st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Error generating map: {e}")
                st.markdown("Please ensure your project data includes valid latitude and longitude coordinates.")
    
    # Add custom CSS for KPI metrics
    st.markdown("""
    <style>
    .kpi-container {
        display: flex;
        justify-content: space-between;
        flex-wrap: wrap;
        margin-bottom: 20px;
    }
    .kpi-box {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        width: 22%;
        min-width: 180px;
        margin-bottom: 15px;
        position: relative;
        overflow: hidden;
    }
    .kpi-box::before {
        content: "";
        position: absolute;
        left: 0;
        top: 0;
        height: 100%;
        width: 6px;
        background: linear-gradient(to bottom, #E67300, #00A3A1);
        border-radius: 3px;
    }
    .kpi-label {
        color: #646469;
        font-size: 14px;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 10px;
    }
    .kpi-value {
        color: #003366;
        font-size: 28px;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .kpi-delta {
        font-size: 14px;
        font-weight: 500;
        padding: 3px 8px;
        border-radius: 12px;
        display: inline-block;
        background-color: rgba(0,0,0,0.04);
    }
    .kpi-delta-positive {
        color: #2E7D32;
    }
    .kpi-delta-negative {
        color: #C62828;
    }
    .kpi-delta-neutral {
        color: #616161;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Display KPIs using columns instead for better reliability
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(
            f"""
            <div class="kpi-box">
                <div class="kpi-label">Portfolio SPI</div>
                <div class="kpi-value">{kpis['avg_spi']:.2f}</div>
                <div class="kpi-delta {'kpi-delta-positive' if kpis['spi_change'] >= 0 else 'kpi-delta-negative'}" style="white-space: nowrap;">{kpis['spi_change']:.2f}</div>
            </div>
            """, 
            unsafe_allow_html=True
        )
    
    with col2:
        st.markdown(
            f"""
            <div class="kpi-box">
                <div class="kpi-label">Portfolio CPI</div>
                <div class="kpi-value">{kpis['avg_cpi']:.2f}</div>
                <div class="kpi-delta {'kpi-delta-positive' if kpis['cpi_change'] >= 0 else 'kpi-delta-negative'}" style="white-space: nowrap;">{kpis['cpi_change']:.2f}</div>
            </div>
            """, 
            unsafe_allow_html=True
        )
    
    with col3:
        st.markdown(
            f"""
            <div class="kpi-box">
                <div class="kpi-label">Active Projects</div>
                <div class="kpi-value">{kpis['active_projects']}</div>
                <div class="kpi-delta kpi-delta-neutral" style="white-space: nowrap;">{kpis['active_projects_change']} from last month</div>
            </div>
            """, 
            unsafe_allow_html=True
        )
    
    with col4:
        st.markdown(
            f"""
            <div class="kpi-box">
                <div class="kpi-label">At Risk Projects</div>
                <div class="kpi-value">{kpis['at_risk_projects']}</div>
                <div class="kpi-delta {'kpi-delta-negative' if kpis['at_risk_projects_change'] > 0 else 'kpi-delta-positive'}" style="white-space: nowrap;">{kpis['at_risk_projects_change']} from last month</div>
            </div>
            """, 
            unsafe_allow_html=True
        )
    
    # Portfolio Performance Overview
    st.markdown("### Portfolio Performance")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Projects by Status
        fig = create_portfolio_overview(st.session_state.projects_df, "status")
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Projects by Sector
        fig = create_portfolio_overview(st.session_state.projects_df, "sector")
        st.plotly_chart(fig, use_container_width=True)
        
    # Geospatial Project Map
    st.markdown("### Project Geospatial Map")
    
    # Check for location data
    has_location_data = all(col in st.session_state.projects_df.columns for col in ['latitude', 'longitude'])
    
    if has_location_data:
        # Filter out projects without coordinates
        projects_with_location = st.session_state.projects_df.dropna(subset=['latitude', 'longitude'])
        
        if not projects_with_location.empty:
            # Create map visualization options
            col1, col2, col3 = st.columns([1, 1, 2])
            
            with col1:
                color_by = st.selectbox(
                    "Color by:", 
                    options=["status", "sector", "spi", "cpi", "priority"],
                    index=0,
                    key="map_color_by"
                )
                
            with col2:
                size_by = st.selectbox(
                    "Size by:", 
                    options=["budget", "actual_cost", "completion_pct", "risk_score"],
                    index=0,
                    key="map_size_by"
                )
            
            with col3:
                st.markdown("#### Map Legend")
                if color_by == "status":
                    statuses = projects_with_location['status'].unique()
                    for status in statuses:
                        st.markdown(f"🔹 **{status}**")
                elif color_by == "priority":
                    priorities = projects_with_location['priority'].unique()
                    for priority in priorities:
                        st.markdown(f"🔹 **{priority}**")
                        
            # Create the map using our enhanced visualization module
            try:
                map_fig = create_geospatial_map(
                    projects_with_location, 
                    color_by=color_by, 
                    size_by=size_by
                )
                st.plotly_chart(map_fig, use_container_width=True, key="project_geo_map")
                
                # Add map insights
                col1, col2 = st.columns(2)
                with col1:
                    # Count projects by country/region
                    if 'location' in projects_with_location.columns:
                        st.markdown("#### Projects by Location")
                        location_counts = projects_with_location['location'].value_counts().head(5)
                        for location, count in location_counts.items():
                            st.markdown(f"**{location}**: {count} projects")
                
                with col2:
                    # Add interesting insights based on geographical distribution
                    st.markdown("#### Geographical Insights")
                    
                    # Calculate average SPI and CPI by location if available
                    if all(col in projects_with_location.columns for col in ['location', 'spi', 'cpi']):
                        perf_by_location = projects_with_location.groupby('location')[['spi', 'cpi']].mean()
                        
                        # Find best and worst performing locations
                        if not perf_by_location.empty:
                            best_loc = perf_by_location['spi'].idxmax()
                            worst_loc = perf_by_location['spi'].idxmin()
                            
                            st.markdown(f"**Best performing location**: {best_loc} (SPI: {perf_by_location.loc[best_loc, 'spi']:.2f})")
                            st.markdown(f"**Needs attention**: {worst_loc} (SPI: {perf_by_location.loc[worst_loc, 'spi']:.2f})")
            
            except Exception as e:
                st.error(f"Error creating map visualization: {e}")
                st.markdown("Please ensure projects have valid latitude and longitude coordinates.")
        else:
            st.info("No projects with location coordinates found in the dataset.")
    else:
        # Show message if location data is missing
        st.info("Project location data (latitude/longitude) not available in the dataset. Add location coordinates to enable the geospatial map visualization.")
    
    # Projects Needing Attention
    st.markdown("### Projects Needing Attention")
    
    attention_projects = get_projects_needing_attention(
        st.session_state.projects_df,
        st.session_state.history_df,
        spi_threshold=st.session_state.spi_threshold,
        cpi_threshold=st.session_state.cpi_threshold
    )
    
    if len(attention_projects) > 0:
        # Display projects needing attention
        col1, col2 = st.columns([2, 1])
        
        with col1:
            # Projects table
            st.dataframe(
                attention_projects[['project_name', 'status', 'spi', 'cpi', 'budget', 'project_manager']],
                use_container_width=True
            )
        
        with col2:
            # Performance bubble chart with better layout and spacing
            fig = px.scatter(
                attention_projects,
                x="spi",
                y="cpi",
                size="budget",
                color="status",
                hover_name="project_name",
                size_max=40,  # Slightly smaller bubbles to prevent overlap
                opacity=0.8,  # Add transparency to help with overlapping points
                color_discrete_map={
                    'On Track': '#2ECC71',
                    'Minor Issues': '#F1C40F',
                    'At Risk': '#E74C3C',
                    'Delayed': '#8E44AD',
                    'Completed': '#3498DB'
                },
                title="SPI vs CPI by Project Size",
                labels={"spi": "Schedule Performance Index (SPI)", 
                        "cpi": "Cost Performance Index (CPI)",
                        "status": "Project Status"}
            )
            
            fig.add_shape(
                type="rect",
                x0=0, y0=0,
                x1=st.session_state.spi_threshold,
                y1=st.session_state.cpi_threshold,
                fillcolor="rgba(255, 0, 0, 0.1)",
                line=dict(width=0),
                layer="below"
            )
            
            fig.add_shape(
                type="line",
                x0=st.session_state.spi_threshold,
                y0=0,
                x1=st.session_state.spi_threshold,
                y1=2,
                line=dict(color="rgba(255, 0, 0, 0.5)", width=1, dash="dash")
            )
            
            fig.add_shape(
                type="line",
                x0=0,
                y0=st.session_state.cpi_threshold,
                x1=2,
                y1=st.session_state.cpi_threshold,
                line=dict(color="rgba(255, 0, 0, 0.5)", width=1, dash="dash")
            )
            
            fig.update_layout(
                xaxis=dict(range=[0.5, 1.5], title="SPI"),
                yaxis=dict(range=[0.5, 1.5], title="CPI"),
                plot_bgcolor='white',
                paper_bgcolor='white',
                margin=dict(l=30, r=50, t=50, b=30),
                height=450,  # Increase height for better proportions
                legend=dict(
                    yanchor="top",
                    y=0.99,
                    xanchor="right",
                    x=0.99,
                    bordercolor="LightGrey",
                    borderwidth=1
                ),
                font=dict(
                    family="Poppins, sans-serif",
                    size=12
                )
            )
            
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No projects currently need attention.")
    
    # Performance trends over time
    st.markdown("### Portfolio Performance Trends")
    
    # Get monthly portfolio performance
    monthly_performance = st.session_state.history_df.groupby('month').agg({
        'spi': 'mean',
        'cpi': 'mean'
    }).reset_index()
    
    # Convert month to datetime for proper sorting
    monthly_performance['month_dt'] = pd.to_datetime(monthly_performance['month'] + '-01')
    monthly_performance = monthly_performance.sort_values('month_dt')
    
    # Create trend chart
    fig = create_trend_chart(monthly_performance)
    st.plotly_chart(fig, use_container_width=True)

def display_project_view():
    """Display detailed project view with KPIs, timeline, and tasks"""
    # Check if data is loaded
    if not hasattr(st.session_state, 'projects_df') or st.session_state.projects_df is None:
        st.warning("Please load data first to see project details.")
        return
        
    # Check if a project is selected
    if not hasattr(st.session_state, 'selected_project') or st.session_state.selected_project is None:
        st.info("Please select a project from the sidebar to view details.")
        return
    
    project_name = st.session_state.selected_project
    st.markdown(f"## Project Details: {project_name}")
    
    # Get project data safely
    project_matches = st.session_state.projects_df[st.session_state.projects_df['project_name'] == project_name]
    if project_matches.empty:
        st.error(f"Project '{project_name}' not found in the loaded data.")
        return
        
    project_df = project_matches.iloc[0]
    project_id = project_df['project_id']
    
    # Get tasks for this project
    project_tasks = st.session_state.tasks_df[st.session_state.tasks_df['project_id'] == project_id]
    
    # Get historical data for this project
    project_history = st.session_state.history_df[st.session_state.history_df['project_id'] == project_id]
    project_history = project_history.sort_values('month')
    
    # Latest KPIs
    latest_metrics = calculate_project_metrics(project_df, project_tasks, project_history)
    
    # Project KPIs
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Schedule Performance Index (SPI)",
            f"{latest_metrics['spi']:.2f}",
            f"{latest_metrics['spi_change']:.2f}",
            delta_color="normal" if latest_metrics['spi_change'] >= 0 else "inverse"
        )
    
    with col2:
        st.metric(
            "Cost Performance Index (CPI)",
            f"{latest_metrics['cpi']:.2f}",
            f"{latest_metrics['cpi_change']:.2f}",
            delta_color="normal" if latest_metrics['cpi_change'] >= 0 else "inverse"
        )
    
    with col3:
        st.metric(
            "Budget Spent",
            f"${latest_metrics['actual_cost']:,.0f}",
            f"{latest_metrics['spend_percentage']:.1f}% of Budget",
            delta_color="off"
        )
    
    with col4:
        st.metric(
            "Schedule Progress",
            f"{latest_metrics['schedule_progress']:.1f}%",
            f"{latest_metrics['days_remaining']} days remaining",
            delta_color="off"
        )
    
    # Project Info & Gauge Charts
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.markdown("### Project Information")
        
        # Create a formatted display of project information
        info_html = f"""
        <div style='background-color: white; padding: 15px; border-radius: 5px; border: 1px solid #ddd;'>
            <p><strong>Project Manager:</strong> {project_df['project_manager']}</p>
            <p><strong>Sector:</strong> {project_df['sector']}</p>
            <p><strong>Budget:</strong> ${project_df['budget']:,.0f}</p>
            <p><strong>Start Date:</strong> {project_df['start_date'].strftime('%d %b %Y')}</p>
            <p><strong>Planned End Date:</strong> {project_df['planned_end_date'].strftime('%d %b %Y')}</p>
            <p><strong>Status:</strong> <span style='font-weight: bold; color: {
                '#2ECC71' if project_df['status'] == 'On Track' else 
                '#F1C40F' if project_df['status'] == 'Minor Issues' else
                '#E74C3C' if project_df['status'] == 'At Risk' else
                '#8E44AD' if project_df['status'] == 'Delayed' else
                '#3498DB'
            };'>{project_df['status']}</span></p>
        </div>
        """
        
        st.markdown(info_html, unsafe_allow_html=True)
    
    with col2:
        # SPI and CPI gauge charts
        gauge_col1, gauge_col2 = st.columns(2)
        
        with gauge_col1:
            fig = create_gauge_chart(latest_metrics['spi'], "SPI", ARCADIS_TEAL)
            st.plotly_chart(fig, use_container_width=True)
        
        with gauge_col2:
            fig = create_gauge_chart(latest_metrics['cpi'], "CPI", ARCADIS_ORANGE)
            st.plotly_chart(fig, use_container_width=True)
    
    # Performance Trends
    st.markdown("### Performance Trends")
    
    if not project_history.empty:
        fig = go.Figure()
        
        # Add SPI line
        fig.add_trace(go.Scatter(
            x=project_history['month'],
            y=project_history['spi'],
            mode='lines+markers',
            name='SPI',
            line=dict(color=ARCADIS_TEAL, width=3),
            marker=dict(size=8)
        ))
        
        # Add CPI line
        fig.add_trace(go.Scatter(
            x=project_history['month'],
            y=project_history['cpi'],
            mode='lines+markers',
            name='CPI',
            line=dict(color=ARCADIS_ORANGE, width=3),
            marker=dict(size=8)
        ))
        
        # Add threshold line
        fig.add_shape(
            type="line",
            x0=project_history['month'].iloc[0],
            y0=1.0,
            x1=project_history['month'].iloc[-1],
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
            yaxis=dict(range=[0.5, 1.5]),
            plot_bgcolor='rgba(255, 255, 255, 1)',
            paper_bgcolor='rgba(255, 255, 255, 1)',
            margin=dict(l=20, r=20, t=40, b=20)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No historical performance data available for this project.")
    
    # Project Schedule
    st.markdown("### Project Schedule")
    
    if not project_tasks.empty:
        fig = create_schedule_gantt(project_tasks, project_df)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No task data available for this project.")
    
    # Task Details
    st.markdown("### Task Details")
    
    if not project_tasks.empty:
        # Add progress and status calculations
        project_tasks = project_tasks.copy()
        project_tasks['progress'] = (project_tasks['earned_value'] / project_tasks['planned_cost']).clip(0, 1) * 100
        project_tasks['status'] = project_tasks.apply(
            lambda x: 'Completed' if x['progress'] >= 100 else
                     'On Track' if x['progress'] >= x['expected_progress'] else
                     'Behind Schedule',
            axis=1
        )
        
        # Format for display
        display_tasks = project_tasks[[
            'task_name', 'planned_start', 'planned_end', 'progress', 'status',
            'planned_cost', 'actual_cost'
        ]].copy()
        
        # Format dates and currency
        display_tasks['planned_start'] = display_tasks['planned_start'].dt.strftime('%d %b %Y')
        display_tasks['planned_end'] = display_tasks['planned_end'].dt.strftime('%d %b %Y')
        display_tasks['planned_cost'] = display_tasks['planned_cost'].map('${:,.0f}'.format)
        display_tasks['actual_cost'] = display_tasks['actual_cost'].map('${:,.0f}'.format)
        display_tasks['progress'] = display_tasks['progress'].map('{:.1f}%'.format)
        
        # Rename columns for display
        display_tasks.columns = [
            'Task Name', 'Start Date', 'End Date', 'Progress', 'Status',
            'Planned Cost', 'Actual Cost'
        ]
        
        st.dataframe(display_tasks, use_container_width=True)
    else:
        st.info("No task data available for this project.")

def display_risk_view():
    """Display risk management view with risk matrix and register"""
    # Check if data is loaded
    if not hasattr(st.session_state, 'projects_df') or st.session_state.projects_df is None:
        st.warning("Please load data first to see risk management information.")
        return
        
    # Check if a project is selected
    if not hasattr(st.session_state, 'selected_project') or st.session_state.selected_project is None:
        st.info("Please select a project from the sidebar to view risk details.")
        return
    
    project_name = st.session_state.selected_project
    st.markdown(f"## Risk Management: {project_name}")
    
    # Get project data safely
    project_matches = st.session_state.projects_df[st.session_state.projects_df['project_name'] == project_name]
    if project_matches.empty:
        st.error(f"Project '{project_name}' not found in the loaded data.")
        return
        
    project_df = project_matches.iloc[0]
    project_id = project_df['project_id']
    
    # Get risks for this project
    project_risks = st.session_state.risks_df[st.session_state.risks_df['project_id'] == project_id]
    
    # Risk Insights Panel
    st.markdown("""
    <div style="background-color: #f8f9fa; border-left: 4px solid #E67300; padding: 15px; border-radius: 4px; margin-bottom: 20px;">
        <h4 style="color: #E67300; margin-top: 0;">🧠 Risk Intelligence Insights</h4>
        <div id="risk-insights">
    """, unsafe_allow_html=True)
    
    # Generate dynamic insights based on risk data
    if not project_risks.empty:
        high_risks = project_risks[project_risks['risk_score'] >= 0.6]
        mitigated_risks = project_risks[project_risks['status'] == 'Mitigated']
        open_risks = project_risks[project_risks['status'] == 'Open']
        
        insights = []
        
        # Insight 1: High risk items
        if len(high_risks) > 0:
            insights.append(f"Project has <strong>{len(high_risks)}</strong> high-risk items requiring immediate attention.")
            
        # Insight 2: Financial exposure
        risk_exposure = project_risks['exposure'].sum()
        budget_pct = (risk_exposure / project_df['budget'] * 100) if project_df['budget'] > 0 else 0
        if budget_pct > 15:
            insights.append(f"Risk exposure is <strong>${risk_exposure:,.0f}</strong>, representing <strong>{budget_pct:.1f}%</strong> of the project budget - above recommended threshold of 15%.")
        else:
            insights.append(f"Financial exposure from identified risks is <strong>${risk_exposure:,.0f}</strong> ({budget_pct:.1f}% of budget).")
            
        # Insight 3: Most common risk category
        if not project_risks.empty:
            category_counts = project_risks['category'].value_counts()
            top_category = category_counts.idxmax()
            insights.append(f"<strong>{top_category}</strong> is the most common risk category ({category_counts[top_category]} risks), focus mitigation efforts in this area.")
            
        # Insight 4: Mitigation progress
        mitigated_pct = len(mitigated_risks) / len(project_risks) * 100 if len(project_risks) > 0 else 0
        if mitigated_pct < 30:
            insights.append(f"Only <strong>{mitigated_pct:.1f}%</strong> of risks have been mitigated, consider accelerating risk response strategies.")
        else:
            insights.append(f"<strong>{mitigated_pct:.1f}%</strong> of identified risks have been successfully mitigated.")
        
        # Insight 5: Open high risks that need addressing
        high_open_risks = high_risks[high_risks['status'] == 'Open']
        if len(high_open_risks) > 0:
            insights.append(f"<strong>{len(high_open_risks)}</strong> high-risk items remain open. Highest priority: <strong>{high_open_risks.iloc[0]['risk_name']}</strong> (Score: {high_open_risks.iloc[0]['risk_score']:.2f}).")
        
        # Display insights with bullet points
        for i, insight in enumerate(insights):
            st.markdown(f"<p>• {insight}</p>", unsafe_allow_html=True)
    else:
        st.markdown("<p>No risk data available to generate insights.</p>", unsafe_allow_html=True)
    
    st.markdown("</div></div>", unsafe_allow_html=True)
    
    # Risk Summary Metrics
    st.markdown("### Risk Summary")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_risks = len(project_risks) if not project_risks.empty else 0
        high_risks = len(project_risks[project_risks['risk_score'] >= 0.6]) if not project_risks.empty else 0
        
        st.metric(
            "Total Risks",
            f"{total_risks}",
            f"{high_risks} high risk" if high_risks > 0 else None
        )
    
    with col2:
        mitigated = len(project_risks[project_risks['status'] == 'Mitigated']) if not project_risks.empty else 0
        mitigated_pct = mitigated / total_risks * 100 if total_risks > 0 else 0
        
        st.metric(
            "Risks Mitigated",
            f"{mitigated}",
            f"{mitigated_pct:.1f}% of total"
        )
    
    with col3:
        risk_exposure = project_risks['exposure'].sum() if not project_risks.empty else 0
        
        st.metric(
            "Risk Exposure",
            f"${risk_exposure:,.0f}",
            f"{(risk_exposure / project_df['budget'] * 100):.1f}% of Budget" if project_df['budget'] > 0 else None,
            delta_color="off"
        )
    
    with col4:
        if not project_risks.empty:
            avg_score = project_risks['risk_score'].mean()
            st.metric(
                "Avg. Risk Score",
                f"{avg_score:.2f}",
                None
            )
        else:
            st.metric("Avg. Risk Score", "0.00")
    
    # Risk Visualizations
    st.markdown("### Risk Analysis")
    
    if not project_risks.empty:
        # Create tabs for different risk visualizations
        risk_tabs = st.tabs(["Risk Matrix", "Risk Distribution", "Risk Trends"])
        
        with risk_tabs[0]:
            # Risk Matrix visualization
            fig = create_risk_matrix(project_risks)
            
            # Enhance the matrix with more details
            fig.update_layout(
                title="Risk Heat Map by Probability and Impact",
                coloraxis_colorbar=dict(
                    title="Risk Score",
                    tickvals=[0, 0.3, 0.6, 0.9],
                    ticktext=["Low", "Medium", "High", "Critical"]
                ),
                plot_bgcolor='white',
                margin=dict(l=20, r=20, t=50, b=20)
            )
            
            # Add annotations to quadrants
            fig.add_annotation(
                x=0.25, y=0.25, text="LOW",
                showarrow=False, font=dict(color="darkgreen", size=14)
            )
            
            fig.add_annotation(
                x=0.25, y=0.75, text="MODERATE",
                showarrow=False, font=dict(color="darkgoldenrod", size=14)
            )
            
            fig.add_annotation(
                x=0.75, y=0.25, text="MODERATE",
                showarrow=False, font=dict(color="darkgoldenrod", size=14)
            )
            
            fig.add_annotation(
                x=0.75, y=0.75, text="HIGH",
                showarrow=False, font=dict(color="darkred", size=14)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
        with risk_tabs[1]:
            col1, col2 = st.columns(2)
            
            with col1:
                # Create bar chart for risk categories
                cat_counts = project_risks['category'].value_counts().reset_index()
                cat_counts.columns = ['Category', 'Count']
                
                fig = px.bar(
                    cat_counts,
                    x='Category',
                    y='Count',
                    color='Count',
                    color_continuous_scale=px.colors.sequential.Oranges,
                    labels={'Count': 'Number of Risks'},
                    title='Risk Distribution by Category'
                )
                
                fig.update_layout(
                    xaxis_title="Risk Category",
                    yaxis_title="Number of Risks",
                    plot_bgcolor='white'
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Create pie chart for risk status
                status_counts = project_risks['status'].value_counts().reset_index()
                status_counts.columns = ['Status', 'Count']
                
                fig = px.pie(
                    status_counts,
                    values='Count',
                    names='Status',
                    title='Risk Distribution by Status',
                    color_discrete_sequence=px.colors.sequential.Teal
                )
                
                fig.update_traces(
                    textposition='inside',
                    textinfo='percent+label',
                    hole=0.4,
                    marker=dict(line=dict(color='#FFFFFF', width=2))
                )
                
                fig.update_layout(
                    showlegend=True,
                    legend=dict(orientation="h", y=-0.1)
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            # Add Risk Level Distribution
            def risk_level_category(score):
                if score >= 0.6:
                    return 'High'
                elif score >= 0.3:
                    return 'Medium'
                else:
                    return 'Low'
                    
            risk_levels = [risk_level_category(score) for score in project_risks['risk_score']]
            
            level_counts = pd.Series(risk_levels).value_counts().reset_index()
            level_counts.columns = ['Risk Level', 'Count']
            
            # Set a specific order for the risk levels
            level_order = ['High', 'Medium', 'Low']
            level_counts['Risk Level'] = pd.Categorical(level_counts['Risk Level'], categories=level_order, ordered=True)
            level_counts = level_counts.sort_values('Risk Level')
            
            # Define colors
            level_colors = {'High': '#E74C3C', 'Medium': '#F1C40F', 'Low': '#3498DB'}
            colors = [level_colors[level] for level in level_counts['Risk Level']]
            
            fig = px.bar(
                level_counts,
                x='Risk Level',
                y='Count',
                title='Risk Distribution by Severity',
                color='Risk Level',
                color_discrete_map=level_colors
            )
            
            fig.update_layout(
                xaxis_title="Risk Level",
                yaxis_title="Number of Risks",
                plot_bgcolor='white'
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
        with risk_tabs[2]:
            st.markdown("#### Risk Score Trends Over Time")
            
            # Create a simple simulation of risk trends if historical data isn't available
            from datetime import datetime, timedelta
            import numpy as np
            
            # Get today's date and generate past data points
            current_date = datetime.now()
            months = pd.date_range(end=current_date, periods=6, freq='M')
            
            # Generate simulated risk score history based on current value
            if not project_risks.empty:
                current_score = project_risks['risk_score'].mean()
                np.random.seed(int(project_id))  # Use project_id for reproducible randomness
                
                # Generate historical score with slight random walk from current
                variations = np.random.normal(0, 0.05, size=5)
                historical_scores = [current_score]
                
                # Work backwards
                for var in variations:
                    prev_score = historical_scores[0] - var
                    # Ensure within bounds
                    prev_score = max(0.1, min(1.0, prev_score))
                    historical_scores.insert(0, prev_score)
                
                # Create the dataframe
                trend_data = pd.DataFrame({
                    'Date': months,
                    'Risk Score': historical_scores
                })
                
                # Plot the trend
                fig = px.line(
                    trend_data,
                    x='Date',
                    y='Risk Score',
                    title='Average Risk Score Trend',
                    markers=True
                )
                
                fig.update_traces(
                    line=dict(color=ARCADIS_ORANGE, width=3),
                    marker=dict(size=10, color=ARCADIS_ORANGE)
                )
                
                fig.update_layout(
                    xaxis_title="Date",
                    yaxis_title="Average Risk Score",
                    plot_bgcolor='white',
                    yaxis=dict(range=[0, 1])
                )
                
                # Add target threshold line
                fig.add_hline(
                    y=0.6,
                    line_dash="dash",
                    line_color="red",
                    annotation_text="High Risk Threshold",
                    annotation_position="top right"
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Add trend insight
                current = historical_scores[-1]
                previous = historical_scores[0]
                change = current - previous
                change_pct = (change / previous) * 100 if previous > 0 else 0
                
                if change > 0:
                    st.markdown(f"""
                    <div style="padding: 15px; background-color: #FFF3E0; border-left: 5px solid #FF9800; margin-top: 10px;">
                        <h4 style="margin-top: 0; color: #E65100;">⚠️ Risk Alert</h4>
                        <p>Risk score has increased by <strong>{change:.2f}</strong> ({change_pct:.1f}%) over the past 6 months. Review mitigation strategies.</p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="padding: 15px; background-color: #E8F5E9; border-left: 5px solid #4CAF50; margin-top: 10px;">
                        <h4 style="margin-top: 0; color: #2E7D32;">✅ Risk Improvement</h4>
                        <p>Risk score has decreased by <strong>{abs(change):.2f}</strong> ({abs(change_pct):.1f}%) over the past 6 months. Current mitigation strategies are working effectively.</p>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No risk data available for trend analysis.")
    else:
        st.info("No risk data available for this project.")
    
    # Top Risks Panel
    st.markdown("### Top Risk Items")
    
    if not project_risks.empty:
        # Sort risks by score (highest first)
        top_risks = project_risks.sort_values('risk_score', ascending=False).head(5)
        
        # Create two columns for top risks
        cols = st.columns(2)
        
        for i, (_, risk) in enumerate(top_risks.iterrows()):
            col_idx = i % 2  # Alternate between columns
            
            with cols[col_idx]:
                # Define risk level and icon
                def get_risk_level(score):
                    if score >= 0.6:
                        return 'High'
                    elif score >= 0.3:
                        return 'Medium'
                    else:
                        return 'Low'
                
                risk_level = get_risk_level(risk['risk_score'])
                icon = ('fa-triangle-exclamation' if risk_level == 'High' else 
                       'fa-exclamation-circle' if risk_level == 'Medium' else 'fa-info-circle')
                color = ('#E74C3C' if risk_level == 'High' else 
                        '#F1C40F' if risk_level == 'Medium' else '#3498DB')
                
                # Create styled risk box
                st.markdown(f"""
                <div style="background-color: white; border-left: 5px solid {color}; 
                            padding: 15px; margin-bottom: 15px; border-radius: 5px; 
                            box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <h4 style="margin: 0; color: {color};">
                            <i class="fas {icon}"></i> {risk['risk_name']}
                        </h4>
                        <span style="background-color: {color}; color: white; 
                                    padding: 3px 8px; border-radius: 10px; font-size: 0.8em;">
                            {risk_level}
                        </span>
                    </div>
                    <p style="margin-top: 10px; margin-bottom: 5px;">{risk['description']}</p>
                    <div style="display: flex; justify-content: space-between; margin-top: 10px; 
                                font-size: 0.9em; color: #555;">
                        <span>Probability: {risk['probability']:.0%}</span>
                        <span>Impact: ${risk['impact_cost']:,.0f}</span>
                        <span>Exposure: ${risk['exposure']:,.0f}</span>
                    </div>
                    <div style="margin-top: 10px; padding-top: 10px; border-top: 1px solid #eee;">
                        <span style="font-weight: 500;">Status: </span>
                        <span style="background-color: {('#4CAF50' if risk['status'] == 'Mitigated' else '#607D8B')}; 
                                    color: white; padding: 2px 6px; border-radius: 4px; font-size: 0.8em;">
                            {risk['status']}
                        </span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.info("No risk data available for this project.")
    
    # Full Risk Register
    st.markdown("### Risk Register")
    
    if not project_risks.empty:
        # Add risk level column
        display_risks = project_risks.copy()
        
        # Define risk level function
        def risk_level(score):
            if score >= 0.6:
                return '🔴 High'
            elif score >= 0.3:
                return '🟠 Medium'
            else:
                return '🟢 Low'
        
        display_risks['risk_level'] = display_risks['risk_score'].apply(risk_level)
        
        # Select columns for display
        display_risks = display_risks[[
            'risk_name', 'risk_level', 'category', 'probability', 'impact_cost', 
            'exposure', 'risk_score', 'status'
        ]].copy()
        
        # Format columns
        display_risks['probability'] = display_risks['probability'].map('{:.0%}'.format)
        display_risks['impact_cost'] = display_risks['impact_cost'].map('${:,.0f}'.format)
        display_risks['exposure'] = display_risks['exposure'].map('${:,.0f}'.format)
        display_risks['risk_score'] = display_risks['risk_score'].map('{:.2f}'.format)
        
        # Rename columns for display
        display_risks.columns = [
            'Risk Name', 'Risk Level', 'Category', 'Probability', 'Impact', 
            'Exposure', 'Risk Score', 'Status'
        ]
        
        # Add filters
        col1, col2, col3 = st.columns(3)
        
        with col1:
            level_filter = st.multiselect(
                "Filter by risk level:",
                options=['🔴 High', '🟠 Medium', '🟢 Low'],
                default=['🔴 High', '🟠 Medium', '🟢 Low']
            )
            
        with col2:
            category_filter = st.multiselect(
                "Filter by category:",
                options=sorted(project_risks['category'].unique()),
                default=sorted(project_risks['category'].unique())
            )
            
        with col3:
            status_filter = st.multiselect(
                "Filter by status:",
                options=sorted(project_risks['status'].unique()),
                default=sorted(project_risks['status'].unique())
            )
        
        # Apply filters
        filtered_risks = display_risks
        
        if level_filter:
            filtered_risks = filtered_risks[filtered_risks['Risk Level'].isin(level_filter)]
            
        if category_filter:
            filtered_risks = filtered_risks[filtered_risks['Category'].isin(category_filter)]
            
        if status_filter:
            filtered_risks = filtered_risks[filtered_risks['Status'].isin(status_filter)]
        
        # Display filtered dataframe
        st.dataframe(filtered_risks, use_container_width=True)
        
        # Add download button
        csv = filtered_risks.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Risk Register",
            data=csv,
            file_name=f"{project_name}_risk_register.csv",
            mime="text/csv",
        )
    else:
        st.info("No risk data available for this project.")

def forecast_project_completion(project, tasks_df, history_df):
    """
    Generate a forecast for project completion date and final cost
    
    Parameters:
    -----------
    project : pandas.Series
        Project data
    tasks_df : pandas.DataFrame
        Tasks for the project
    history_df : pandas.DataFrame
        Historical performance data
        
    Returns:
    --------
    dict
        Forecast results
    """
    # Initialize results dictionary
    forecast = {
        'forecast_completion_date': None,
        'completion_confidence': 0,
        'forecast_final_cost': None,
        'schedule_variance_days': 0,
        'cost_variance': 0,
        'risk_factors': []
    }
    
    try:
        # Calculate forecast using basic EVM methods
        if ('planned_end' in project and not pd.isnull(project['planned_end']) and 
            'planned_start' in project and not pd.isnull(project['planned_start']) and
            'spi' in project and not pd.isnull(project['spi']) and project['spi'] > 0):
            
            # Calculate planned duration
            if isinstance(project['planned_start'], str):
                planned_start = pd.to_datetime(project['planned_start'])
            else:
                planned_start = project['planned_start']
                
            if isinstance(project['planned_end'], str):
                planned_end = pd.to_datetime(project['planned_end'])
            else:
                planned_end = project['planned_end']
            
            planned_duration = (planned_end - planned_start).days
            
            # Calculate elapsed days
            today = datetime.now()
            if today < planned_start:
                elapsed_days = 0
            else:
                elapsed_days = (today - planned_start).days
            
            # Calculate remaining duration using SPI
            if project['spi'] > 0:
                remaining_duration = max(0, (planned_duration - elapsed_days) / project['spi'])
                forecast_completion_date = planned_start + timedelta(days=elapsed_days + remaining_duration)
                
                # Calculate schedule variance in days
                schedule_variance_days = (forecast_completion_date - planned_end).days
                
                # Set forecast results
                forecast['forecast_completion_date'] = forecast_completion_date
                forecast['schedule_variance_days'] = schedule_variance_days
                forecast['completion_confidence'] = int(70 * project['spi']) if project['spi'] <= 1 else 70
        
        # Calculate cost forecast if we have budget and CPI
        if ('budget' in project and not pd.isnull(project['budget']) and 
            'cpi' in project and not pd.isnull(project['cpi']) and project['cpi'] > 0):
            
            # For actual cost, use if available, otherwise estimate
            if 'actual_cost' in project and not pd.isnull(project['actual_cost']):
                actual_cost = project['actual_cost']
            elif 'completion_pct' in project and not pd.isnull(project['completion_pct']):
                # Estimate based on % complete and CPI
                pct_complete = min(100, max(0, project['completion_pct'])) / 100
                actual_cost = project['budget'] * pct_complete / project['cpi']
            else:
                # Default estimate
                actual_cost = project['budget'] * 0.5 / project['cpi']
            
            # Calculate Estimate at Completion (EAC)
            forecast_final_cost = actual_cost + (project['budget'] - actual_cost) / project['cpi']
            
            # Calculate cost variance
            cost_variance = forecast_final_cost - project['budget']
            
            # Set forecast results
            forecast['forecast_final_cost'] = forecast_final_cost
            forecast['cost_variance'] = cost_variance
        
        # Identify risk factors
        risk_factors = []
        
        # SPI-based risks
        if 'spi' in project and not pd.isnull(project['spi']):
            if project['spi'] < st.session_state.spi_threshold - 0.1:
                risk_factors.append("Schedule performance is significantly below threshold")
            elif project['spi'] < st.session_state.spi_threshold:
                risk_factors.append("Schedule performance is below threshold")
        
        # CPI-based risks
        if 'cpi' in project and not pd.isnull(project['cpi']):
            if project['cpi'] < st.session_state.cpi_threshold - 0.1:
                risk_factors.append("Cost performance is significantly below threshold")
            elif project['cpi'] < st.session_state.cpi_threshold:
                risk_factors.append("Cost performance is below threshold")
        
        # Schedule variance risks
        if 'schedule_variance_days' in forecast and forecast['schedule_variance_days'] > 0:
            if forecast['schedule_variance_days'] > 30:
                risk_factors.append(f"Significant delay of {forecast['schedule_variance_days']} days")
            elif forecast['schedule_variance_days'] > 10:
                risk_factors.append(f"Moderate delay of {forecast['schedule_variance_days']} days")
        
        # Cost variance risks
        if 'cost_variance' in forecast and forecast['cost_variance'] > 0:
            budget = project['budget'] if 'budget' in project and not pd.isnull(project['budget']) else 0
            if budget > 0:
                cost_overrun_pct = (forecast['cost_variance'] / budget) * 100
                if cost_overrun_pct > 20:
                    risk_factors.append(f"Significant budget overrun of {cost_overrun_pct:.1f}%")
                elif cost_overrun_pct > 10:
                    risk_factors.append(f"Moderate budget overrun of {cost_overrun_pct:.1f}%")
        
        # High risk count
        if 'risk_count' in project and not pd.isnull(project['risk_count']) and project['risk_count'] > 5:
            risk_factors.append(f"High number of risks: {project['risk_count']}")
        
        # Save risk factors
        forecast['risk_factors'] = risk_factors
        
    except Exception as e:
        print(f"Error in forecast_project_completion: {e}")
    
    return forecast

def generate_nlp_insights(project):
    """
    Generate natural language insights for a project based on its metrics
    
    Parameters:
    -----------
    project : pandas.Series
        Project data
        
    Returns:
    --------
    list
        List of natural language insights
    """
    insights = []
    
    try:
        # Performance insights
        if 'spi' in project and not pd.isnull(project['spi']):
            if project['spi'] > 1.1:
                insights.append(f"Schedule performance is excellent at {project['spi']:.2f}, significantly ahead of plan.")
            elif project['spi'] > 1.0:
                insights.append(f"Schedule performance is good at {project['spi']:.2f}, slightly ahead of plan.")
            elif project['spi'] > 0.9:
                insights.append(f"Schedule performance is acceptable at {project['spi']:.2f}, but monitor for potential delays.")
            elif project['spi'] > 0.8:
                insights.append(f"Schedule performance is concerning at {project['spi']:.2f}, consider corrective actions.")
            else:
                insights.append(f"Schedule performance is critical at {project['spi']:.2f}, immediate intervention required.")
        
        if 'cpi' in project and not pd.isnull(project['cpi']):
            if project['cpi'] > 1.1:
                insights.append(f"Cost performance is excellent at {project['cpi']:.2f}, significantly under budget.")
            elif project['cpi'] > 1.0:
                insights.append(f"Cost performance is good at {project['cpi']:.2f}, slightly under budget.")
            elif project['cpi'] > 0.9:
                insights.append(f"Cost performance is acceptable at {project['cpi']:.2f}, but monitor for potential overruns.")
            elif project['cpi'] > 0.8:
                insights.append(f"Cost performance is concerning at {project['cpi']:.2f}, consider cost control measures.")
            else:
                insights.append(f"Cost performance is critical at {project['cpi']:.2f}, immediate financial intervention required.")
        
        # Progress insights
        if 'completion_pct' in project and not pd.isnull(project['completion_pct']):
            if 'planned_end' in project and not pd.isnull(project['planned_end']):
                planned_end = project['planned_end']
                if isinstance(planned_end, str):
                    planned_end = pd.to_datetime(planned_end)
                
                days_to_end = (planned_end - datetime.now()).days
                
                if days_to_end > 0:
                    expected_progress = 100 - (days_to_end / (project['planned_duration'] if 'planned_duration' in project else 100)) * 100
                    progress_diff = project['completion_pct'] - expected_progress
                    
                    if progress_diff < -20:
                        insights.append(f"Project is significantly behind expected progress ({project['completion_pct']:.1f}% vs expected {expected_progress:.1f}%).")
                    elif progress_diff < -10:
                        insights.append(f"Project is behind expected progress ({project['completion_pct']:.1f}% vs expected {expected_progress:.1f}%).")
                    elif progress_diff > 20:
                        insights.append(f"Project is significantly ahead of expected progress ({project['completion_pct']:.1f}% vs expected {expected_progress:.1f}%).")
                    elif progress_diff > 10:
                        insights.append(f"Project is ahead of expected progress ({project['completion_pct']:.1f}% vs expected {expected_progress:.1f}%).")
        
        # Risk insights
        if 'risk_count' in project and not pd.isnull(project['risk_count']):
            if 'risk_score' in project and not pd.isnull(project['risk_score']):
                if project['risk_count'] > 10 and project['risk_score'] > 30:
                    insights.append(f"High risk profile with {project['risk_count']} risks and score of {project['risk_score']}.")
                elif project['risk_count'] > 5 and project['risk_score'] > 20:
                    insights.append(f"Elevated risk profile with {project['risk_count']} risks and score of {project['risk_score']}.")
                elif project['risk_count'] == 0:
                    insights.append("No risks identified. Ensure risk identification is complete.")
    
    except Exception as e:
        print(f"Error in generate_nlp_insights: {e}")
        
    return insights

def detect_project_anomalies(projects_df, metrics_df, contamination=0.1):
    """
    Detect anomalous projects using Isolation Forest algorithm
    
    Parameters:
    -----------
    projects_df : pandas.DataFrame
        Projects data with project_id as index
    metrics_df : pandas.DataFrame
        Metrics data for anomaly detection
    contamination : float, optional
        Expected proportion of outliers, by default 0.1
        
    Returns:
    --------
    pandas.DataFrame
        DataFrame with anomaly detection results
    """
    try:
        # Ensure we have data to work with
        if projects_df.empty or metrics_df.empty:
            return pd.DataFrame()
        
        # Make a copy of the metrics data to avoid modifying the original
        metrics = metrics_df.copy()
        
        # Fill NaN values with column mean
        metrics = metrics.fillna(metrics.mean())
        
        # Need at least 10 observations to make meaningful anomaly detection
        if len(metrics) < 10:
            # For small datasets, use simple statistical thresholds
            anomalies = pd.DataFrame(index=projects_df.index)
            anomalies['is_anomaly'] = False
            anomalies['anomaly_score'] = 0
            
            for col in metrics.columns:
                mean = metrics[col].mean()
                std = metrics[col].std()
                
                # Flag values > 3 standard deviations from mean
                outliers = (metrics[col] > mean + 3*std) | (metrics[col] < mean - 3*std)
                
                # Update anomaly scores
                anomalies.loc[outliers, 'is_anomaly'] = True
                anomalies.loc[outliers, 'anomaly_score'] += 1
            
            return anomalies
        
        # Standardize the data
        scaler = StandardScaler()
        scaled_data = scaler.fit_transform(metrics)
        
        # Apply Isolation Forest
        model = IsolationForest(
            n_estimators=100,
            contamination=float(contamination),
            random_state=42
        )
        
        # Fit the model
        model.fit(scaled_data)
        
        # Predict anomalies (-1 for outliers, 1 for inliers)
        predictions = model.predict(scaled_data)
        scores = model.decision_function(scaled_data)
        
        # Convert to DataFrame
        anomalies = pd.DataFrame(index=projects_df.index)
        anomalies['is_anomaly'] = predictions == -1
        
        # Convert scores to 0-1 range for easier interpretation
        # Lower score = more anomalous
        anomalies['anomaly_score'] = np.abs(scores)
        
        return anomalies
    
    except Exception as e:
        print(f"Error in detect_project_anomalies: {e}")
        return pd.DataFrame()

def generate_project_forecasts():
    """
    Generate predictive forecasts for all projects using historical data
    and machine learning models
    
    Returns:
    --------
    dict
        Dictionary containing forecast results by project
    """
    if not st.session_state.data_loaded:
        return {}
    
    forecasts = {}
    
    # Check if forecasts are already cached and valid
    if ('project_forecasts' in st.session_state and 
        st.session_state.project_forecasts and 
        'forecast_generated_at' in st.session_state and
        (datetime.now() - st.session_state.forecast_generated_at).total_seconds() < 3600):  # Valid for 1 hour
        return st.session_state.project_forecasts
    
    for _, project in st.session_state.projects_df.iterrows():
        project_id = project['project_id']
        
        # Get tasks for this project
        project_tasks = st.session_state.tasks_df[st.session_state.tasks_df['project_id'] == project_id]
        
        # Get historical data for this project
        if 'project_id' in st.session_state.history_df.columns:
            project_history = st.session_state.history_df[st.session_state.history_df['project_id'] == project_id]
        else:
            project_history = None
        
        # Generate forecast using advanced analytics
        forecast = forecast_project_completion(project, project_tasks, project_history)
        
        # Add natural language insights
        insights = generate_nlp_insights(project)
        
        # Extend forecast with additional insights
        forecast['insights'] = insights
        
        # Detect anomalies if available
        if 'anomalies' in st.session_state and st.session_state.anomalies is not None:
            anomaly_df = st.session_state.anomalies
            if project_id in anomaly_df.index:
                forecast['is_anomaly'] = anomaly_df.loc[project_id, 'is_anomaly']
                forecast['anomaly_score'] = anomaly_df.loc[project_id, 'anomaly_score']
            else:
                forecast['is_anomaly'] = False
                forecast['anomaly_score'] = 0
        else:
            forecast['is_anomaly'] = False
            forecast['anomaly_score'] = 0
                
        forecasts[project_id] = forecast
    
    # Cache the forecasts
    st.session_state.project_forecasts = forecasts
    st.session_state.forecast_generated_at = datetime.now()
    
    return forecasts

def generate_ensemble_forecast(project, tasks_df, history_df):
    """
    Generate an ensemble forecast combining multiple predictive models
    for more accurate project outcome predictions
    
    Parameters:
    -----------
    project : pandas.Series
        Project data
    tasks_df : pandas.DataFrame
        Tasks for the project
    history_df : pandas.DataFrame
        Historical performance data
        
    Returns:
    --------
    dict
        Forecast results with confidence intervals
    """
    # Initialize results dictionary
    results = {
        'completion_date': None,
        'completion_date_best': None,
        'completion_date_worst': None,
        'final_cost': None, 
        'final_cost_best': None,
        'final_cost_worst': None,
        'confidence': None,
        'risk_factors': [],
        'forecast_models': []
    }
    
    try:
        # For dates, ensure we have proper datetime objects
        if not pd.isnull(project['planned_start']):
            if isinstance(project['planned_start'], str):
                planned_start = pd.to_datetime(project['planned_start'])
            else:
                planned_start = project['planned_start']
        else:
            # Can't forecast without start date
            return results
            
        if not pd.isnull(project['planned_end']):
            if isinstance(project['planned_end'], str):
                planned_end = pd.to_datetime(project['planned_end'])
            else:
                planned_end = project['planned_end']
        else:
            # Can't forecast without end date
            return results
        
        # Calculate basic EVM forecast using current SPI/CPI
        if 'spi' in project and 'cpi' in project and project['spi'] > 0:
            planned_duration = (planned_end - planned_start).days
            elapsed_days = (datetime.now() - planned_start).days
            
            # SPI-based completion date forecast
            if project['spi'] > 0:
                remaining_duration = (planned_duration - elapsed_days) / project['spi']
                evm_completion_date = planned_start + timedelta(days=elapsed_days + remaining_duration)
                results['forecast_models'].append("EVM")
            else:
                evm_completion_date = planned_end + timedelta(days=30)  # Default if SPI = 0
            
            # Budget forecast using CPI
            if 'budget' in project and project['cpi'] > 0:
                if 'actual_cost' in project:
                    actual_cost = project['actual_cost']
                else:
                    # Estimate based on budget and progress
                    if 'completion_pct' in project and not pd.isnull(project['completion_pct']):
                        progress = project['completion_pct'] / 100
                    else:
                        progress = 0.5  # Default assumption
                    
                    actual_cost = project['budget'] * progress / project['cpi']
                
                # Calculate EAC (Estimate at Completion)
                evm_final_cost = actual_cost + ((project['budget'] - (project['budget'] * project['spi'])) / project['cpi'])
            else:
                evm_final_cost = project['budget'] * 1.1  # Default 10% overrun
        else:
            # Default forecast if metrics are missing
            evm_completion_date = planned_end + timedelta(days=15)  # Assume 15 days delay
            evm_final_cost = project['budget'] * 1.1 if 'budget' in project else None
            
        # Try to use ML-based forecast if we have sufficient history
        ml_completion_date = None
        ml_final_cost = None
        
        if history_df is not None and not history_df.empty and len(history_df) >= 3:
            # Use Linear Regression or XGBoost for time series forecast
            try:
                # Sort by month
                history_sorted = history_df.sort_values('month')
                
                # Prepare time features (months from start)
                history_sorted['month_num'] = range(len(history_sorted))
                
                # Train completion date model based on SPI trend
                if 'spi' in history_sorted.columns:
                    X = history_sorted[['month_num']].values.reshape(-1, 1)
                    y_spi = history_sorted['spi'].values
                    
                    model_spi = LinearRegression()
                    model_spi.fit(X, y_spi)
                    
                    # Forecast SPI for future periods
                    forecast_periods = st.session_state.forecast_horizon
                    future_months = np.array(range(len(history_sorted), len(history_sorted) + forecast_periods)).reshape(-1, 1)
                    forecasted_spi = model_spi.predict(future_months)
                    
                    # Calculate expected completion based on SPI trend
                    avg_future_spi = np.mean(forecasted_spi)
                    
                    # Only use ML forecast if it's reasonable (SPI > 0)
                    if avg_future_spi > 0:
                        remaining_duration = (planned_duration - elapsed_days) / avg_future_spi
                        ml_completion_date = planned_start + timedelta(days=elapsed_days + remaining_duration)
                        results['forecast_models'].append("Linear Regression")
                
                # Train cost model based on CPI trend
                if 'cpi' in history_sorted.columns and 'budget' in project and 'actual_cost' in project:
                    X = history_sorted[['month_num']].values.reshape(-1, 1)
                    y_cpi = history_sorted['cpi'].values
                    
                    model_cpi = LinearRegression()
                    model_cpi.fit(X, y_cpi)
                    
                    forecasted_cpi = model_cpi.predict(future_months)
                    avg_future_cpi = np.mean(forecasted_cpi)
                    
                    # Only use ML forecast if it's reasonable (CPI > 0)
                    if avg_future_cpi > 0:
                        remaining_budget = project['budget'] - project['actual_cost']
                        adjusted_remaining = remaining_budget / avg_future_cpi
                        ml_final_cost = project['actual_cost'] + adjusted_remaining
            except Exception as e:
                print(f"Error in ML forecast: {e}")
                ml_completion_date = None
                ml_final_cost = None
        
        # Combine forecasts (ensemble approach)
        if ml_completion_date is not None:
            # Weight the forecasts (75% ML, 25% EVM if we have sufficient history)
            completion_date = ml_completion_date * 0.75 + evm_completion_date * 0.25
            confidence = 0.8  # Higher confidence with ML
        else:
            completion_date = evm_completion_date
            confidence = 0.6  # Lower confidence with just EVM
            
        if ml_final_cost is not None:
            final_cost = ml_final_cost * 0.75 + evm_final_cost * 0.25
        else:
            final_cost = evm_final_cost
            
        # Create confidence intervals
        completion_buffer = (planned_end - planned_start).days * 0.15  # 15% buffer
        cost_buffer = project['budget'] * 0.15 if 'budget' in project else 0  # 15% buffer
        
        # Set results
        results['completion_date'] = completion_date
        results['completion_date_best'] = completion_date - timedelta(days=completion_buffer/2)
        results['completion_date_worst'] = completion_date + timedelta(days=completion_buffer)
        results['final_cost'] = final_cost
        results['final_cost_best'] = final_cost * 0.95
        results['final_cost_worst'] = final_cost * 1.2
        results['confidence'] = confidence
        
        # Identify risk factors
        risk_factors = []
        
        # Schedule risks
        if 'spi' in project and project['spi'] < st.session_state.spi_threshold:
            risk_factors.append("Low Schedule Performance Index")
        
        # Cost risks
        if 'cpi' in project and project['cpi'] < st.session_state.cpi_threshold:
            risk_factors.append("Low Cost Performance Index")
            
        # Trend risks
        if history_df is not None and not history_df.empty and 'spi' in history_df.columns:
            # Check last 3 periods for downward trend
            if len(history_df) >= 3:
                last_periods = history_df.sort_values('month').tail(3)
                if 'spi' in last_periods.columns and last_periods['spi'].is_monotonic_decreasing:
                    risk_factors.append("Declining SPI Trend")
                if 'cpi' in last_periods.columns and last_periods['cpi'].is_monotonic_decreasing:
                    risk_factors.append("Declining CPI Trend")
        
        # Significant forecast changes
        if planned_end is not None and completion_date is not None:
            delay_days = (completion_date - planned_end).days
            if delay_days > 30:
                risk_factors.append(f"Significant schedule delay ({delay_days} days)")
                
        if 'budget' in project and final_cost is not None:
            cost_overrun_pct = (final_cost - project['budget']) / project['budget'] * 100
            if cost_overrun_pct > 20:
                risk_factors.append(f"Significant cost overrun ({cost_overrun_pct:.1f}%)")
        
        results['risk_factors'] = risk_factors
        
    except Exception as e:
        print(f"Error in ensemble forecast: {e}")
        # Return default results on error
    
    return results

def display_analytics_view():
    """Display analytics and forecasting view with predictive capabilities"""
    st.markdown("## Analytics & Forecasting")
    
    # Check if data is loaded
    if not hasattr(st.session_state, 'history_df') or st.session_state.history_df is None:
        st.warning("Please load data first to see analytics and forecasting information.")
        return
        
    # Add tabs for different analytics views
    analytics_tabs = st.tabs(["Performance Distribution", "Performance Trends", "Predictive Forecasts", "Anomaly Detection"])
    
    # Tab 1: Portfolio Performance Distribution
    with analytics_tabs[0]:
        st.markdown("### Portfolio Performance Distribution")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # SPI Distribution
            latest_performance = st.session_state.history_df.groupby('project_id').apply(
                lambda x: x.sort_values('month').iloc[-1]
            ).reset_index(drop=True)
            
    # Tab 2: Performance Trends
    with analytics_tabs[1]:
        st.markdown("### Performance Trends Analysis")
        
        # Controls for trend analysis
        col1, col2 = st.columns([1, 2])
        
        with col1:
            # Filter controls
            st.markdown("#### Filter Options")
            
            # Time range selector
            months_to_show = st.slider(
                "Months to analyze:", 
                min_value=3, 
                max_value=24, 
                value=12
            )
            
            # Metric selector
            metrics = st.multiselect(
                "Select metrics to display:", 
                options=["SPI", "CPI", "Both"],
                default=["Both"]
            )
            
            # Project selector
            all_projects = st.session_state.projects_df['project_name'].tolist()
            selected_projects = st.multiselect(
                "Select specific projects:", 
                options=all_projects,
                default=[]
            )
            
            # Add a button to update the chart
            st.button("Update Trends", key="update_trends")
            
        with col2:
            st.markdown("#### Performance Over Time")
            
            # Get historical data
            history_df = st.session_state.history_df
            
            # If specific projects selected, filter to those
            if selected_projects:
                project_ids = st.session_state.projects_df[
                    st.session_state.projects_df['project_name'].isin(selected_projects)
                ]['project_id'].tolist()
                
                if project_ids:
                    history_df = history_df[history_df['project_id'].isin(project_ids)]
            
            # Limit to the selected time range
            all_months = sorted(history_df['month'].unique())
            if len(all_months) > months_to_show:
                history_df = history_df[history_df['month'].isin(all_months[-months_to_show:])]
            
            # Create trend visualization
            if "Both" in metrics or ("SPI" in metrics and "CPI" in metrics):
                # Create a line chart with both metrics
                df_avg = history_df.groupby('month')[['spi', 'cpi']].mean().reset_index()
                
                fig = go.Figure()
                
                # Add SPI line
                fig.add_trace(go.Scatter(
                    x=df_avg['month'], 
                    y=df_avg['spi'],
                    mode='lines+markers',
                    name='SPI',
                    line=dict(color='#E67300', width=3)
                ))
                
                # Add CPI line
                fig.add_trace(go.Scatter(
                    x=df_avg['month'], 
                    y=df_avg['cpi'],
                    mode='lines+markers',
                    name='CPI',
                    line=dict(color='#00A3A1', width=3)
                ))
                
                # Add reference line at 1.0
                fig.add_shape(
                    type="line",
                    x0=df_avg['month'].iloc[0],
                    y0=1.0,
                    x1=df_avg['month'].iloc[-1],
                    y1=1.0,
                    line=dict(color="gray", width=1, dash="dash")
                )
                
                # Update layout
                fig.update_layout(
                    title="SPI and CPI Trends",
                    xaxis_title="Month",
                    yaxis_title="Performance Index",
                    height=500,
                    template="plotly_white"
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            elif "SPI" in metrics:
                # Create SPI-only chart
                df_avg = history_df.groupby('month')[['spi']].mean().reset_index()
                fig = px.line(
                    df_avg, 
                    x='month', 
                    y='spi', 
                    markers=True,
                    title="SPI Trend",
                    labels={'spi': 'SPI', 'month': 'Month'}
                )
                fig.update_traces(line=dict(color='#E67300', width=3))
                
                # Add reference line at 1.0
                fig.add_shape(
                    type="line",
                    x0=df_avg['month'].iloc[0],
                    y0=1.0,
                    x1=df_avg['month'].iloc[-1],
                    y1=1.0,
                    line=dict(color="gray", width=1, dash="dash")
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            elif "CPI" in metrics:
                # Create CPI-only chart
                df_avg = history_df.groupby('month')[['cpi']].mean().reset_index()
                fig = px.line(
                    df_avg, 
                    x='month', 
                    y='cpi', 
                    markers=True,
                    title="CPI Trend",
                    labels={'cpi': 'CPI', 'month': 'Month'}
                )
                fig.update_traces(line=dict(color='#00A3A1', width=3))
                
                # Add reference line at 1.0
                fig.add_shape(
                    type="line",
                    x0=df_avg['month'].iloc[0],
                    y0=1.0,
                    x1=df_avg['month'].iloc[-1],
                    y1=1.0,
                    line=dict(color="gray", width=1, dash="dash")
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            # Add trendline insights
            st.markdown("#### Trend Insights")
            
            # Calculate trend direction and magnitude
            if not history_df.empty and len(df_avg) > 1:
                # Get first and last month values
                first_month = df_avg.iloc[0]
                last_month = df_avg.iloc[-1]
                
                # Calculate percent changes
                if "SPI" in metrics or "Both" in metrics:
                    spi_change = (last_month['spi'] - first_month['spi']) / first_month['spi'] * 100
                    st.markdown(f"**SPI Trend:** {'🔼' if spi_change > 0 else '🔽'} {abs(spi_change):.1f}% {'improvement' if spi_change > 0 else 'decline'} over {months_to_show} months")
                
                if "CPI" in metrics or "Both" in metrics:
                    cpi_change = (last_month['cpi'] - first_month['cpi']) / first_month['cpi'] * 100
                    st.markdown(f"**CPI Trend:** {'🔼' if cpi_change > 0 else '🔽'} {abs(cpi_change):.1f}% {'improvement' if cpi_change > 0 else 'decline'} over {months_to_show} months")
                
                # Add overall assessment
                if "Both" in metrics:
                    overall_trend = (spi_change + cpi_change) / 2
                    if overall_trend > 5:
                        st.markdown("**Overall Assessment:** Strong positive trend in project performance 🟢")
                    elif overall_trend > 0:
                        st.markdown("**Overall Assessment:** Slight positive trend in project performance 🟢")
                    elif overall_trend > -5:
                        st.markdown("**Overall Assessment:** Relatively stable project performance 🟠")
                    else:
                        st.markdown("**Overall Assessment:** Negative trend in project performance, requires attention 🔴")
    
    # Tab 3: Predictive Forecasts
    with analytics_tabs[2]:
        st.markdown("### Predictive Project Forecasts")
        
        # Project selector
        projects_df = st.session_state.projects_df
        active_projects = projects_df[projects_df['status'] != 'Completed']['project_name'].tolist()
        
        if not active_projects:
            st.warning("No active projects found for forecasting.")
        else:
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.markdown("#### Forecast Settings")
                
                # Select project to forecast
                selected_project = st.selectbox(
                    "Select project to forecast:",
                    options=active_projects
                )
                
                # Select forecast method
                forecast_method = st.radio(
                    "Forecast method:",
                    options=["Schedule Performance (SPI)", "Ensemble Forecast", "Monte Carlo Simulation"],
                    index=0
                )
                
                # Confidence level
                confidence_level = st.slider(
                    "Confidence level:",
                    min_value=70,
                    max_value=99,
                    value=85,
                    step=5
                )
                
                # Button to generate forecast
                generate_btn = st.button("Generate Forecast", key="generate_forecast")
                
            with col2:
                st.markdown("#### Project Forecast")
                
                # Get project data
                project_data = projects_df[projects_df['project_name'] == selected_project].iloc[0]
                project_id = project_data['project_id']
                
                # Get project tasks
                tasks_df = st.session_state.tasks_df
                project_tasks = tasks_df[tasks_df['project_id'] == project_id]
                
                # Get historical performance
                history_df = st.session_state.history_df
                project_history = history_df[history_df['project_id'] == project_id]
                
                if project_data.empty or project_tasks.empty or project_history.empty:
                    st.warning("Insufficient data for project forecasting.")
                else:
                    # Project summary info
                    st.markdown(f"**Project**: {selected_project}")
                    
                    # Current project status
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Status", project_data['status'])
                    with col2:
                        st.metric("Current SPI", f"{project_history['spi'].iloc[-1]:.2f}")
                    with col3:
                        st.metric("Current CPI", f"{project_history['cpi'].iloc[-1]:.2f}")
                    
                    # Generate forecast
                    if generate_btn:
                        st.markdown("### Forecast Results")
                        
                        # Format original plan data
                        planned_start = project_data['start_date']
                        planned_end = project_data['planned_end_date']
                        planned_duration = (planned_end - planned_start).days
                        
                        # Calculate forecast based on selected method
                        if forecast_method == "Schedule Performance (SPI)":
                            # Simple SPI-based forecast
                            current_spi = project_history['spi'].iloc[-1]
                            
                            # Current date
                            today = datetime.now()
                            
                            if current_spi > 0:
                                # Calculate elapsed and remaining time
                                elapsed_days = (today - planned_start).days
                                remaining_duration = max(0, (planned_duration - elapsed_days) / current_spi)
                                
                                # Calculate forecast date
                                forecast_date = today + timedelta(days=remaining_duration)
                                
                                # Calculate variance
                                date_variance = (forecast_date - planned_end).days
                                
                                # Display results
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    st.markdown("#### Schedule Forecast")
                                    st.markdown(f"**Planned End Date:** {planned_end.strftime('%Y-%m-%d')}")
                                    st.markdown(f"**Forecast End Date:** {forecast_date.strftime('%Y-%m-%d')}")
                                    
                                    variance_color = "green" if date_variance <= 0 else "red"
                                    st.markdown(f"**Schedule Variance:** <span style='color:{variance_color};'>{abs(date_variance)} days {'early' if date_variance <= 0 else 'delay'}</span>", unsafe_allow_html=True)
                                    
                                    # Calculate confidence interval
                                    std_dev_days = max(5, int(abs(planned_duration * 0.1)))
                                    z_value = 1.96  # ~95% confidence
                                    
                                    lower_bound = forecast_date - timedelta(days=int(std_dev_days * z_value))
                                    upper_bound = forecast_date + timedelta(days=int(std_dev_days * z_value))
                                    
                                    st.markdown(f"**Confidence Interval (95%):**")
                                    st.markdown(f"Earliest: {lower_bound.strftime('%Y-%m-%d')}")
                                    st.markdown(f"Latest: {upper_bound.strftime('%Y-%m-%d')}")
                                
                                with col2:
                                    # Create a Gantt chart showing plan vs forecast
                                    fig = go.Figure()
                                    
                                    # Planned duration
                                    fig.add_trace(go.Bar(
                                        y=['Project Timeline'],
                                        x=[planned_duration],
                                        base=[planned_start.timestamp() * 1000],
                                        orientation='h',
                                        name='Planned',
                                        marker=dict(color='rgba(0, 163, 161, 0.8)')
                                    ))
                                    
                                    # Forecasted remaining duration
                                    fig.add_trace(go.Bar(
                                        y=['Forecast'],
                                        x=[(forecast_date - today).days],
                                        base=[today.timestamp() * 1000],
                                        orientation='h',
                                        name='Forecast',
                                        marker=dict(color='rgba(230, 115, 0, 0.8)')
                                    ))
                                    
                                    # Add a marker for current date
                                    fig.add_vline(
                                        x=today.timestamp() * 1000,
                                        line_width=3,
                                        line_dash="dash",
                                        line_color="black",
                                        annotation_text="Today"
                                    )
                                    
                                    # Update layout
                                    fig.update_layout(
                                        title="Project Timeline Forecast",
                                        xaxis=dict(
                                            title="Date",
                                            type='date'
                                        ),
                                        yaxis=dict(
                                            title=""
                                        ),
                                        barmode='group',
                                        height=300
                                    )
                                    
                                    st.plotly_chart(fig, use_container_width=True)
                            else:
                                st.warning("Cannot generate forecast with current SPI value.")
                        
                        elif forecast_method == "Ensemble Forecast":
                            st.info("Generating ensemble forecast combining multiple predictive models...")
                            
                            # Here would be the code to call an advanced forecast function
                            # For now, we'll show a placeholder visualization
                            
                            st.markdown("#### Ensemble Forecast Results")
                            
                            # Sample data for visualization
                            today = datetime.now()
                            forecast_date = planned_end + timedelta(days=15)  # Example forecast
                            
                            # Create a visualization showing multiple forecast methods
                            fig = go.Figure()
                            
                            # Timeline axis
                            dates = [planned_start, today, planned_end, forecast_date]
                            dates.sort()
                            
                            methods = ['EVM', 'Linear Regression', 'Historical Pattern', 'Ensemble']
                            forecasts = [
                                planned_end + timedelta(days=10),
                                planned_end + timedelta(days=18),
                                planned_end + timedelta(days=12),
                                forecast_date
                            ]
                            
                            for method, forecast in zip(methods, forecasts):
                                fig.add_trace(go.Scatter(
                                    x=[forecast],
                                    y=[method],
                                    mode='markers',
                                    marker=dict(size=15, symbol='diamond'),
                                    name=method
                                ))
                            
                            # Add reference lines
                            fig.add_vline(
                                x=today.timestamp() * 1000,
                                line_width=2,
                                line_dash="dash",
                                line_color="black",
                                annotation_text="Today"
                            )
                            
                            fig.add_vline(
                                x=planned_end.timestamp() * 1000,
                                line_width=2,
                                line_dash="dash",
                                line_color="green",
                                annotation_text="Planned End"
                            )
                            
                            fig.update_layout(
                                title="Forecast Comparison by Method",
                                xaxis=dict(
                                    title="Forecast Date",
                                    type='date'
                                ),
                                yaxis=dict(
                                    title="Forecast Method"
                                ),
                                height=300
                            )
                            
                            st.plotly_chart(fig, use_container_width=True)
                            
                            # Ensemble results
                            st.markdown("#### Final Ensemble Forecast")
                            st.markdown(f"**Planned End Date:** {planned_end.strftime('%Y-%m-%d')}")
                            st.markdown(f"**Ensemble Forecast Date:** {forecast_date.strftime('%Y-%m-%d')}")
                            
                            variance_days = (forecast_date - planned_end).days
                            variance_color = "green" if variance_days <= 0 else "red"
                            st.markdown(f"**Schedule Variance:** <span style='color:{variance_color};'>{abs(variance_days)} days {'early' if variance_days <= 0 else 'delay'}</span>", unsafe_allow_html=True)
                        
                        elif forecast_method == "Monte Carlo Simulation":
                            st.info("Running Monte Carlo simulation...")
                            
                            # Sample data for visualization
                            simulation_results = {
                                'min_date': planned_end - timedelta(days=5),
                                'max_date': planned_end + timedelta(days=30),
                                'mean_date': planned_end + timedelta(days=12),
                                'median_date': planned_end + timedelta(days=10),
                                'p10_date': planned_end + timedelta(days=5),
                                'p90_date': planned_end + timedelta(days=20)
                            }
                            
                            # Generate normal distribution data for the histogram
                            mean_offset = (simulation_results['mean_date'] - planned_end).days
                            std_dev = 5  # Standard deviation in days
                            sim_data = np.random.normal(mean_offset, std_dev, 1000)
                            
                            # Create histogram of completion dates
                            fig = px.histogram(
                                sim_data, 
                                title="Monte Carlo Simulation: Project Completion Date",
                                labels={'value': 'Days from Planned Date', 'count': 'Frequency'},
                                opacity=0.7
                            )
                            
                            # Add markers for key dates
                            for label, date in [
                                ('P10', (simulation_results['p10_date'] - planned_end).days),
                                ('Mean', mean_offset),
                                ('P90', (simulation_results['p90_date'] - planned_end).days)
                            ]:
                                fig.add_vline(
                                    x=date,
                                    line_width=2,
                                    line_dash="dash",
                                    annotation_text=label
                                )
                            
                            st.plotly_chart(fig, use_container_width=True)
                            
                            # Summary of Monte Carlo results
                            st.markdown("#### Monte Carlo Simulation Results")
                            st.markdown(f"**Planned End Date:** {planned_end.strftime('%Y-%m-%d')}")
                            st.markdown(f"**Mean Forecast Date:** {simulation_results['mean_date'].strftime('%Y-%m-%d')}")
                            st.markdown(f"**P10 (Optimistic) Date:** {simulation_results['p10_date'].strftime('%Y-%m-%d')}")
                            st.markdown(f"**P90 (Conservative) Date:** {simulation_results['p90_date'].strftime('%Y-%m-%d')}")
                            
                            # Confidence analysis
                            selected_confidence = confidence_level
                            if selected_confidence <= 75:
                                confidence_date = simulation_results['p10_date'] + timedelta(days=int((selected_confidence - 10) / 65 * 10))
                            else:
                                confidence_date = simulation_results['p90_date'] - timedelta(days=int((90 - selected_confidence) / 15 * 10))
                            
                            st.markdown(f"**At {selected_confidence}% Confidence:** Complete by {confidence_date.strftime('%Y-%m-%d')}")
            
            fig = px.histogram(
                latest_performance, 
                x="spi",
                nbins=20,
                color_discrete_sequence=[ARCADIS_TEAL],
                title="SPI Distribution"
            )
            
            fig.add_shape(
                type="line",
                x0=1.0,
                y0=0,
                x1=1.0,
                y1=latest_performance.shape[0] // 3,
                line=dict(color="black", width=2, dash="dash")
            )
            
            fig.add_annotation(
                x=1.0,
                y=latest_performance.shape[0] // 3,
                text="SPI = 1.0",
                showarrow=True,
                arrowhead=1,
                ax=50,
                ay=-30
            )
        
            fig.update_layout(
                xaxis_title="SPI",
                yaxis_title="Count",
                plot_bgcolor='rgba(255, 255, 255, 1)',
                paper_bgcolor='rgba(255, 255, 255, 1)',
                margin=dict(l=20, r=20, t=40, b=20)
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
        with col2:
            # CPI Distribution
            fig = px.histogram(
                latest_performance, 
                x="cpi",
                nbins=20,
                color_discrete_sequence=[ARCADIS_ORANGE],
                title="CPI Distribution"
            )
            
            fig.add_shape(
                type="line",
                x0=1.0,
                y0=0,
                x1=1.0,
                y1=latest_performance.shape[0] // 3,
                line=dict(color="black", width=2, dash="dash")
            )
            
            fig.add_annotation(
                x=1.0,
                y=latest_performance.shape[0] // 3,
                text="CPI = 1.0",
                showarrow=True,
                arrowhead=1,
                ax=50,
                ay=-30
            )
            
            fig.update_layout(
                xaxis_title="CPI",
                yaxis_title="Count",
                plot_bgcolor='rgba(255, 255, 255, 1)',
                paper_bgcolor='rgba(255, 255, 255, 1)',
                margin=dict(l=20, r=20, t=40, b=20)
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
        # Forecasting
        st.markdown("### Project Forecasting")
    
        # Project selection for forecast
        projects = st.session_state.projects_df['project_name'].unique()
        forecast_project = st.selectbox(
            "Select Project for Forecast:",
            projects,
            index=0 if st.session_state.selected_project is None else list(projects).index(st.session_state.selected_project)
        )
        
        # Get project data
        project_df = st.session_state.projects_df[st.session_state.projects_df['project_name'] == forecast_project].iloc[0]
        project_id = project_df['project_id']
        
        # Get historical data for this project
        project_history = st.session_state.history_df[st.session_state.history_df['project_id'] == project_id]
        project_history = project_history.sort_values('month')
    
        # Generate forecast data
        forecast_data = prepare_forecast_data(project_df, project_history)
        
        col1, col2 = st.columns(2)
    
        with col1:
            st.markdown("#### Schedule Forecast")
        
            fig = go.Figure()
            
            # Historical SPI
            fig.add_trace(go.Scatter(
                x=forecast_data['historical_dates'],
                y=forecast_data['historical_spi'],
                mode='lines+markers',
                name='Historical SPI',
                line=dict(color=ARCADIS_TEAL, width=3),
                marker=dict(size=8)
            ))
            
            # Forecasted SPI
            fig.add_trace(go.Scatter(
                x=forecast_data['forecast_dates'],
                y=forecast_data['forecast_spi'],
                mode='lines+markers',
                name='Forecast SPI',
                line=dict(color=ARCADIS_TEAL, width=3, dash='dash'),
                marker=dict(size=8)
            ))
            
            # Add threshold line
            fig.add_shape(
                type="line",
                x0=forecast_data['historical_dates'].iloc[0],
                y0=1.0,
                x1=forecast_data['forecast_dates'].iloc[-1],
                y1=1.0,
                line=dict(color="rgba(0, 0, 0, 0.3)", width=1, dash="dash")
            )
        
            # Update layout
            fig.update_layout(
                title="SPI Forecast",
                xaxis_title="Date",
                yaxis_title="SPI",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                ),
                yaxis=dict(range=[0.5, 1.5]),
                plot_bgcolor='rgba(255, 255, 255, 1)',
                paper_bgcolor='rgba(255, 255, 255, 1)',
                margin=dict(l=20, r=20, t=40, b=20)
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
            # Schedule Forecast Metrics
            schedule_variance = forecast_data['schedule_variance_days']
            forecast_end_date = forecast_data['forecast_end_date']
        
            metric_color = "normal"
            if schedule_variance > 0:
                metric_color = "inverse"
            
            st.metric(
                "Forecast End Date",
                forecast_end_date.strftime('%d %b %Y'),
                f"{abs(schedule_variance)} days {'behind' if schedule_variance > 0 else 'ahead of'} schedule",
                delta_color=metric_color
            )
    
        with col2:
            st.markdown("#### Cost Forecast")
        
            fig = go.Figure()
            
            # Historical CPI
            fig.add_trace(go.Scatter(
                x=forecast_data['historical_dates'],
                y=forecast_data['historical_cpi'],
                mode='lines+markers',
                name='Historical CPI',
                line=dict(color=ARCADIS_ORANGE, width=3),
                marker=dict(size=8)
            ))
            
            # Forecasted CPI
            fig.add_trace(go.Scatter(
                x=forecast_data['forecast_dates'],
                y=forecast_data['forecast_cpi'],
                mode='lines+markers',
                name='Forecast CPI',
                line=dict(color=ARCADIS_ORANGE, width=3, dash='dash'),
                marker=dict(size=8)
            ))
        
            # Add threshold line
            fig.add_shape(
                type="line",
                x0=forecast_data['historical_dates'].iloc[0],
                y0=1.0,
                x1=forecast_data['forecast_dates'].iloc[-1],
                y1=1.0,
                line=dict(color="rgba(0, 0, 0, 0.3)", width=1, dash="dash")
            )
        
            # Update layout
            fig.update_layout(
                title="CPI Forecast",
                xaxis_title="Date",
                yaxis_title="CPI",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                ),
                yaxis=dict(range=[0.5, 1.5]),
                plot_bgcolor='rgba(255, 255, 255, 1)',
                paper_bgcolor='rgba(255, 255, 255, 1)',
                margin=dict(l=20, r=20, t=40, b=20)
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
            # Cost Forecast Metrics
            budget = project_df['budget']
            forecast_cost = forecast_data['forecast_cost']
            cost_variance = budget - forecast_cost
        
            metric_color = "normal"
            if cost_variance < 0:
                metric_color = "inverse"
            
            st.metric(
                "Forecast Cost",
                f"${forecast_cost:,.0f}",
                f"${abs(cost_variance):,.0f} {'over' if cost_variance < 0 else 'under'} budget",
                delta_color=metric_color
            )
    
    # Dataset insights
    st.markdown("### Data Insights")
    
    # Business Analytics Image
    col1, col2 = st.columns(2)
    
    with col1:
        st.image("https://pixabay.com/get/g481283eed83560e55d0b411a32e8896e941fcf1788e4d60bcc4b3388302d2f541de907d2fa5b1013d7ac01e475bff7a90242d3204a40b12c446a480ba6e86741_1280.jpg", 
                 caption="Analytics Dashboard", use_container_width=True)
    
    with col2:
        st.image("https://pixabay.com/get/g914c64869642fb2fc6033f30452184f0561bdf56cae5b8b8f3f6be60f1df1fa9244c55dfd9a446eb282b9bb640e6f0dd4ffdba17f82004af12773a1a8d13c639_1280.jpg", 
                 caption="Business Analytics", use_container_width=True)
    
    # Analytics findings
    st.markdown(
        """
        <div class="welcome-section">
            <h3><i class="fas fa-lightbulb"></i> Key Insights</h3>
            <ul class="styled-list">
                <li><b>Performance Trends:</b> Projects in the Energy Transition sector are showing the highest average SPI (1.05) and CPI (1.08) values, indicating effective management practices in this sector.</li>
                <li><b>Risk Patterns:</b> The most common risk categories are Resource Availability (25%) and Scope Changes (22%), suggesting areas for preventive action.</li>
                <li><b>Schedule Predictions:</b> Based on current performance trends, 15% of projects are forecasted to exceed their planned duration by more than 15%, requiring immediate attention.</li>
                <li><b>Cost Analysis:</b> Projects managed by Diana Prince show the highest average CPI (1.12), demonstrating effective cost management techniques that could be shared across the organization.</li>
            </ul>
        </div>
        """, 
        unsafe_allow_html=True
    )

# --- Main App Layout ---
def upload_data():
    """Handle file upload for custom data"""
    st.markdown("## Upload Project Data")
    st.info("Upload your project data files in CSV format. Each file should follow the required template format.")
    
    # Projects file upload
    projects_file = st.file_uploader("Upload Projects Data (CSV)", type="csv", key="projects_upload")
    
    # Tasks file upload
    tasks_file = st.file_uploader("Upload Tasks Data (CSV)", type="csv", key="tasks_upload")
    
    # Risks file upload
    risks_file = st.file_uploader("Upload Risks Data (CSV)", type="csv", key="risks_upload")
    
    # History file upload
    history_file = st.file_uploader("Upload Performance History Data (CSV)", type="csv", key="history_upload")
    
    if st.button("Process Uploaded Files"):
        if projects_file is None:
            st.error("Projects file is required. Please upload a projects data file.")
            return
        
        try:
            # Read uploaded files
            projects_df = pd.read_csv(projects_file)
            
            tasks_df = pd.read_csv(tasks_file) if tasks_file else pd.DataFrame()
            risks_df = pd.read_csv(risks_file) if risks_file else pd.DataFrame()
            history_df = pd.read_csv(history_file) if history_file else pd.DataFrame()
            
            # Store in session state
            st.session_state.projects_df = projects_df
            st.session_state.tasks_df = tasks_df
            st.session_state.risks_df = risks_df
            st.session_state.history_df = history_df
            
            # Calculate portfolio metrics
            st.session_state.portfolio_kpis = calculate_portfolio_metrics(
                projects_df, tasks_df, history_df
            )
            
            # Set data_loaded flag
            st.session_state.data_loaded = True
            
            # Hide upload form
            st.session_state.show_upload = False
            
            # Success message
            st.success("Data uploaded and processed successfully!")
            time.sleep(1)
            st.rerun()
            
        except Exception as e:
            st.error(f"Error processing uploaded files: {str(e)}")
    
    if st.button("Cancel", key="cancel_upload"):
        st.session_state.show_upload = False
        st.rerun()

def main():
    # Display sidebar
    display_sidebar()
    
    # Display header
    show_header()
    
    # Handle file upload if requested
    if st.session_state.show_upload:
        upload_data()
        return
    
    # Display appropriate view based on state
    if not st.session_state.data_loaded:
        display_welcome()
    else:
        # Create tabs for different views
        tab_options = ["Portfolio Overview", "Project Details", "Risk Management", "Analytics & Forecasting"]
        
        # Add Collaboration Hub tab if role permits
        if st.session_state.user_role != "Team Member":
            tab_options.append("Collaboration Hub")
            
        tabs = st.tabs(tab_options)
        
        with tabs[0]:
            # Always display portfolio view when data is loaded
            display_portfolio_view()
            # Set current view to portfolio
            st.session_state.current_view = "portfolio"
            
        with tabs[1]:
            # Always display project view when data is loaded
            display_project_view()
            
        with tabs[2]:
            # Always display risk view when data is loaded
            display_risk_view()
            
        with tabs[3]:
            # Always display analytics view when data is loaded
            display_analytics_view()
            
        # Display Collaboration Hub if available
        if st.session_state.user_role != "Team Member" and len(tabs) > 4:
            with tabs[4]:
                display_collaboration_hub()

def display_collaboration_hub():
    """Display collaboration features including comments, action items and meeting mode"""
    st.title("Collaboration Hub")
    st.markdown("Collaborate with team members, track action items, and share insights")
    
    # Initialize managers if needed
    if 'collaboration_manager' not in st.session_state:
        st.session_state.collaboration_manager = CollaborationManager()
        st.session_state.comment_system = CommentSystem(st.session_state.collaboration_manager)
        st.session_state.action_tracker = ActionTracker(st.session_state.collaboration_manager)
        st.session_state.alert_system = AlertSystem(st.session_state.collaboration_manager)
        st.session_state.meeting_mode = MeetingMode(st.session_state.collaboration_manager)
        st.session_state.report_distribution = ReportDistribution(st.session_state.collaboration_manager)
    
    # Create sub-tabs
    collab_tabs = st.tabs(["Action Items", "Comments & Annotations", "Alerts", "Meeting Mode", "Reports"])
    
    # Action Items Tab
    with collab_tabs[0]:
        st.header("Action Items")
        st.markdown("Track and manage project action items and assignments")
        
        # Action filters
        col1, col2, col3 = st.columns(3)
        with col1:
            status_filter = st.multiselect(
                "Status:", 
                ["Open", "In Progress", "Completed", "Deferred"],
                default=["Open", "In Progress"]
            )
        with col2:
            priority_filter = st.multiselect(
                "Priority:", 
                ["Critical", "High", "Medium", "Low"],
                default=["Critical", "High"]
            )
        with col3:
            project_filter = st.selectbox(
                "Project:",
                ["All Projects"] + list(st.session_state.projects_df['project_name'].unique())
            )
        
        # Get filtered actions
        project_id = None if project_filter == "All Projects" else project_filter
        actions = st.session_state.action_tracker.get_actions(
            status=status_filter,
            priority=priority_filter,
            project_id=project_id
        )
        
        # Display actions
        if actions:
            for action in actions:
                with st.expander(f"{action['title']} - {action['priority']} Priority - {action['status']}"):
                    st.write(f"**Description:** {action['description']}")
                    st.write(f"**Assigned To:** {action['assigned_to']}")
                    st.write(f"**Due Date:** {action['due_date']}")
                    
                    # Action status update
                    new_status = st.selectbox(
                        "Update Status:", 
                        ["Open", "In Progress", "Completed", "Deferred"],
                        index=["Open", "In Progress", "Completed", "Deferred"].index(action['status']),
                        key=f"status_{action['id']}"
                    )
                    
                    if new_status != action['status']:
                        if st.button("Update Status", key=f"update_{action['id']}"):
                            st.session_state.action_tracker.update_action(
                                action['id'], {'status': new_status}
                            )
                            st.success(f"Status updated to {new_status}")
                            st.rerun()
                    
                    # Comments section
                    st.write("#### Comments")
                    for comment in action.get('comments', []):
                        st.write(f"**{comment['user_name']}:** {comment['text']}")
                    
                    # Add comment
                    new_comment = st.text_area("Add Comment:", key=f"comment_{action['id']}")
                    if st.button("Add Comment", key=f"add_comment_{action['id']}"):
                        if new_comment:
                            st.session_state.action_tracker.add_action_comment(
                                action['id'], "current_user", st.session_state.user_role, new_comment
                            )
                            st.success("Comment added!")
                            st.rerun()
        else:
            st.info("No action items found with the selected filters.")
        
        # Add new action item
        st.markdown("### Add New Action Item")
        with st.form(key="new_action_form"):
            action_title = st.text_input("Title:")
            action_description = st.text_area("Description:")
            action_assigned = st.selectbox(
                "Assign To:",
                ["Executive", "PMO Manager", "Project Manager", "Team Member"]
            )
            col1, col2 = st.columns(2)
            with col1:
                action_priority = st.selectbox("Priority:", ["Critical", "High", "Medium", "Low"])
            with col2:
                action_due = st.date_input("Due Date:")
            
            action_project = st.selectbox(
                "Related Project:",
                ["None"] + list(st.session_state.projects_df['project_name'].unique())
            )
            
            submit_action = st.form_submit_button("Create Action Item")
            
            if submit_action:
                if action_title and action_description:
                    project_id = None if action_project == "None" else action_project
                    st.session_state.action_tracker.add_action(
                        title=action_title,
                        description=action_description,
                        assigned_to=action_assigned,
                        due_date=action_due.isoformat(),
                        priority=action_priority,
                        status="Open",
                        project_id=project_id
                    )
                    st.success("Action item created!")
                    st.rerun()
                else:
                    st.error("Title and description are required.")
    
    # Comments Tab
    with collab_tabs[1]:
        st.header("Comments & Annotations")
        st.markdown("Add comments and annotations to projects, tasks, and charts")
        
        # Filter options
        comment_target_type = st.selectbox(
            "Filter By:",
            ["All", "Project", "Chart", "Risk", "Metric"]
        )
        
        target_id = None
        if comment_target_type != "All":
            if comment_target_type == "Project":
                target_id = st.selectbox(
                    "Select Project:",
                    list(st.session_state.projects_df['project_name'].unique())
                )
        
        # Get filtered comments
        comments = st.session_state.comment_system.get_comments(
            target_type=None if comment_target_type == "All" else comment_target_type.lower(),
            target_id=target_id
        )
        
        # Display comments
        if comments:
            for comment in comments:
                with st.expander(f"{comment['user_name']} on {comment['target_type']} - {comment['created_at'][:10]}"):
                    st.write(f"**Comment:** {comment['text']}")
                    
                    # Display replies
                    if 'replies' in comment and comment['replies']:
                        st.write("#### Replies")
                        for reply in comment['replies']:
                            st.write(f"**{reply['user_name']}:** {reply['text']}")
                    
                    # Add reply
                    new_reply = st.text_area("Add Reply:", key=f"reply_{comment['id']}")
                    if st.button("Reply", key=f"add_reply_{comment['id']}"):
                        if new_reply:
                            st.session_state.comment_system.add_reply(
                                comment['id'], "current_user", st.session_state.user_role, new_reply
                            )
                            st.success("Reply added!")
                            st.rerun()
        else:
            st.info("No comments found with the selected filters.")
        
        # Add new comment
        st.markdown("### Add New Comment")
        with st.form(key="new_comment_form"):
            comment_target_type = st.selectbox(
                "Comment On:",
                ["Project", "Chart", "Risk", "Metric"],
                key="new_comment_target_type"
            )
            
            if comment_target_type == "Project":
                comment_target_id = st.selectbox(
                    "Select Project:",
                    list(st.session_state.projects_df['project_name'].unique()),
                    key="new_comment_target_id"
                )
            else:
                comment_target_id = st.text_input("Target ID:", key="new_comment_target_id")
            
            comment_text = st.text_area("Comment:")
            
            submit_comment = st.form_submit_button("Add Comment")
            
            if submit_comment:
                if comment_text and comment_target_id:
                    st.session_state.comment_system.add_comment(
                        user_id="current_user",
                        user_name=st.session_state.user_role,
                        target_type=comment_target_type.lower(),
                        target_id=comment_target_id,
                        text=comment_text
                    )
                    st.success("Comment added!")
                    st.rerun()
                else:
                    st.error("Comment and target are required.")
    
    # Alerts Tab
    with collab_tabs[2]:
        st.header("Alerts & Notifications")
        st.markdown("Configure automated alerts based on project metrics and KPIs")
        
        # Display existing alert rules
        alert_rules = st.session_state.alert_system.get_alert_rules()
        
        if alert_rules:
            for rule in alert_rules:
                with st.expander(f"{rule['name']} - {'Enabled' if rule['enabled'] else 'Disabled'}"):
                    st.write(f"**Description:** {rule['description']}")
                    st.write(f"**Metric:** {rule['metric_type']}")
                    st.write(f"**Condition:** {rule['comparison_operator']} {rule['threshold_value']}")
                    st.write(f"**Project:** {rule['project_id']}")
                    st.write(f"**Channels:** {', '.join(rule['alert_channels'])}")
                    st.write(f"**Recipients:** {', '.join(rule['recipients'])}")
                    st.write(f"**Frequency:** {rule['frequency']}")
                    
                    # Toggle enabled status
                    enabled = st.checkbox("Enabled", value=rule['enabled'], key=f"enabled_{rule['id']}")
                    
                    if enabled != rule['enabled']:
                        if st.button("Update Status", key=f"update_alert_{rule['id']}"):
                            st.session_state.alert_system.update_alert_rule(
                                rule['id'], {'enabled': enabled}
                            )
                            st.success(f"Alert rule {'enabled' if enabled else 'disabled'}")
                            st.rerun()
        else:
            st.info("No alert rules configured.")
        
        # Create new alert rule
        st.markdown("### Create New Alert Rule")
        with st.form(key="new_alert_form"):
            alert_name = st.text_input("Alert Name:")
            alert_description = st.text_area("Description:")
            
            col1, col2 = st.columns(2)
            with col1:
                alert_metric = st.selectbox(
                    "Metric:",
                    ["spi", "cpi", "budget_variance", "schedule_variance", "risk_score"]
                )
            with col2:
                alert_project = st.selectbox(
                    "Project:",
                    ["All Projects"] + list(st.session_state.projects_df['project_name'].unique())
                )
            
            col1, col2 = st.columns(2)
            with col1:
                alert_operator = st.selectbox(
                    "Condition:",
                    ["<", "<=", "=", ">=", ">"]
                )
            with col2:
                alert_threshold = st.number_input("Threshold Value:", value=0.9, step=0.05)
            
            alert_recipients = st.multiselect(
                "Recipients:",
                ["Executive", "PMO Manager", "Project Manager", "Team Member"]
            )
            
            alert_channels = st.multiselect(
                "Notification Channels:",
                ["dashboard", "email"]
            )
            
            alert_frequency = st.selectbox(
                "Check Frequency:",
                ["realtime", "hourly", "daily", "weekly"]
            )
            
            submit_alert = st.form_submit_button("Create Alert Rule")
            
            if submit_alert:
                if alert_name and alert_description and alert_recipients and alert_channels:
                    project_id = "all" if alert_project == "All Projects" else alert_project
                    st.session_state.alert_system.create_alert_rule(
                        name=alert_name,
                        description=alert_description,
                        metric_type=alert_metric,
                        project_id=project_id,
                        threshold_value=alert_threshold,
                        comparison_operator=alert_operator,
                        recipients=alert_recipients,
                        alert_channels=alert_channels,
                        frequency=alert_frequency
                    )
                    st.success("Alert rule created!")
                    st.rerun()
                else:
                    st.error("All fields are required.")
    
    # Meeting Mode Tab
    with collab_tabs[3]:
        st.header("Meeting Mode")
        st.markdown("Create and manage presentation views for meetings and reviews")
        
        # Display existing presentations
        presentations = st.session_state.meeting_mode.get_presentations()
        
        if presentations:
            selected_presentation = st.selectbox(
                "Select Presentation:",
                [p['title'] for p in presentations]
            )
            
            # Find the selected presentation
            presentation = next((p for p in presentations if p['title'] == selected_presentation), None)
            
            if presentation:
                st.write(f"**Title:** {presentation['title']}")
                if presentation.get('subtitle'):
                    st.write(f"**Subtitle:** {presentation['subtitle']}")
                st.write(f"**Last Used:** {presentation['last_used'][:10]}")
                st.write(f"**Slides:** {len(presentation['slides'])}")
                
                if st.button("Start Presentation", key="start_presentation"):
                    st.session_state.presentation_mode = True
                    st.session_state.current_presentation = presentation
                    st.success("Presentation mode activated!")
                    # In a real implementation, this would trigger a presentation view
        else:
            st.info("No saved presentations.")
        
        # Create new presentation
        st.markdown("### Create New Presentation")
        with st.form(key="new_presentation_form"):
            presentation_title = st.text_input("Title:")
            presentation_subtitle = st.text_input("Subtitle (optional):")
            
            # Select slides to include
            slide_options = {
                "portfolio_overview": "Portfolio Overview",
                "at_risk_projects": "At-Risk Projects",
                "performance_trends": "Performance Trends",
                "resource_allocation": "Resource Allocation",
                "upcoming_milestones": "Upcoming Milestones",
                "risk_summary": "Risk Summary"
            }
            
            selected_slides = st.multiselect(
                "Select Slides to Include:",
                options=list(slide_options.keys()),
                format_func=lambda x: slide_options[x],
                default=["portfolio_overview", "at_risk_projects"]
            )
            
            col1, col2 = st.columns(2)
            with col1:
                slide_duration = st.number_input("Slide Duration (seconds):", value=30, min_value=5, step=5)
            with col2:
                auto_advance = st.checkbox("Auto-advance Slides", value=False)
            
            show_annotations = st.checkbox("Show Comments & Annotations", value=True)
            
            submit_presentation = st.form_submit_button("Create Presentation")
            
            if submit_presentation:
                if presentation_title and selected_slides:
                    st.session_state.meeting_mode.create_presentation(
                        slides=selected_slides,
                        title=presentation_title,
                        subtitle=presentation_subtitle
                    )
                    st.success("Presentation created!")
                    st.rerun()
                else:
                    st.error("Title and at least one slide are required.")
    
    # Reports Tab
    with collab_tabs[4]:
        st.header("Report Distribution")
        st.markdown("Schedule and distribute automated reports")
        
        # Display existing report schedules
        report_schedules = st.session_state.report_distribution.get_report_schedules()
        
        if report_schedules:
            for schedule in report_schedules:
                with st.expander(f"{schedule['name']} - {'Enabled' if schedule['enabled'] else 'Disabled'}"):
                    st.write(f"**Description:** {schedule['description']}")
                    st.write(f"**Report Type:** {schedule['report_type']}")
                    st.write(f"**Format:** {schedule['format']}")
                    st.write(f"**Frequency:** {schedule['frequency']}")
                    st.write(f"**Recipients:** {', '.join(schedule['recipients'])}")
                    
                    if schedule.get('last_run'):
                        st.write(f"**Last Run:** {schedule['last_run'][:10]}")
                    
                    # Toggle enabled status
                    enabled = st.checkbox("Enabled", value=schedule['enabled'], key=f"enabled_report_{schedule['id']}")
                    
                    if enabled != schedule['enabled']:
                        if st.button("Update Status", key=f"update_report_{schedule['id']}"):
                            st.session_state.report_distribution.update_report_schedule(
                                schedule['id'], {'enabled': enabled}
                            )
                            st.success(f"Report schedule {'enabled' if enabled else 'disabled'}")
                            st.rerun()
        else:
            st.info("No report schedules configured.")
        
        # Create new report schedule
        st.markdown("### Create New Report Schedule")
        with st.form(key="new_report_form"):
            report_name = st.text_input("Schedule Name:")
            report_description = st.text_area("Description:")
            
            report_type = st.selectbox(
                "Report Type:",
                ["portfolio_overview", "project_detail", "risk_summary", "performance_trends"]
            )
            
            report_recipients = st.multiselect(
                "Recipients:",
                ["Executive", "PMO Manager", "Project Manager", "Team Member"]
            )
            
            col1, col2 = st.columns(2)
            with col1:
                report_frequency = st.selectbox(
                    "Frequency:",
                    ["daily", "weekly", "monthly"]
                )
            with col2:
                report_format = st.selectbox(
                    "Format:",
                    ["pdf", "excel", "powerpoint"]
                )
            
            report_parameters = {}
            
            if report_type == "project_detail":
                report_parameters['project'] = st.selectbox(
                    "Project:",
                    list(st.session_state.projects_df['project_name'].unique())
                )
            
            submit_report = st.form_submit_button("Create Report Schedule")
            
            if submit_report:
                if report_name and report_description and report_recipients:
                    st.session_state.report_distribution.create_report_schedule(
                        name=report_name,
                        description=report_description,
                        report_type=report_type,
                        recipients=report_recipients,
                        frequency=report_frequency,
                        parameters=report_parameters,
                        format=report_format
                    )
                    st.success("Report schedule created!")
                    st.rerun()
                else:
                    st.error("All fields are required.")

# Run the main application
if __name__ == "__main__":
    main()
