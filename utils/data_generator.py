"""
Data Generator Module for PMO Pulse Dashboard

Generates mock project data for demonstration and testing purposes.
"""

import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta

def generate_mock_data(num_projects=40, tasks_per_project=20, risks_per_project=7, history_months=12):
    """
    Generates mock project data including projects, tasks, risks, and historical performance.
    
    Parameters:
    -----------
    num_projects : int
        Number of projects to generate
    tasks_per_project : int
        Number of tasks per project to generate
    risks_per_project : int
        Number of risks per project to generate
    history_months : int
        Number of months of historical data to generate
        
    Returns:
    --------
    tuple
        (projects_df, tasks_df, risks_df, changes_df, history_df, benefits_df)
    """
    np.random.seed(42)  # For reproducibility
    today = datetime.date.today()
    first_possible_start = today - relativedelta(years=3)
    
    # Generate project data
    project_ids = np.arange(1, num_projects + 1)
    sectors = np.random.choice(
        ['Infrastructure', 'Buildings', 'Water', 'Environment', 'Energy Transition', 'Digital Transformation'], 
        num_projects
    )
    statuses = np.random.choice(
        ['On Track', 'Minor Issues', 'At Risk', 'Delayed', 'Completed'], 
        num_projects, 
        p=[0.35, 0.2, 0.15, 0.15, 0.15]
    )
    budgets = np.random.randint(500_000, 30_000_000, num_projects)
    start_dates = [
        first_possible_start + datetime.timedelta(days=np.random.randint(0, (today - first_possible_start).days - 180)) 
        for _ in range(num_projects)
    ]
    planned_durations = np.random.randint(180, 1200, num_projects)
    end_dates = [
        sd + datetime.timedelta(days=int(pd)) 
        for sd, pd in zip(start_dates, planned_durations)
    ]
    target_cpi = np.random.uniform(0.95, 1.02, num_projects)
    target_spi = np.random.uniform(0.98, 1.02, num_projects)
    
    # Generate geographical coordinates based on sector
    # Each sector will have a base location with random spread to create clusters
    sector_base_locations = {
        'Infrastructure': (40.7128, -74.0060),  # New York
        'Buildings': (34.0522, -118.2437),      # Los Angeles
        'Water': (29.7604, -95.3698),           # Houston
        'Environment': (47.6062, -122.3321),    # Seattle
        'Energy Transition': (41.8781, -87.6298),  # Chicago
        'Digital Transformation': (37.7749, -122.4194)  # San Francisco
    }
    
    # Generate latitudes and longitudes with some random scatter around the base
    latitudes = []
    longitudes = []
    locations = []
    
    for sector in sectors:
        base_lat, base_long = sector_base_locations.get(sector, (39.8283, -98.5795))  # Default to geographic center of US
        # Add some random scatter (approximately within 50-150 miles)
        lat = base_lat + np.random.uniform(-1.5, 1.5)
        long = base_long + np.random.uniform(-1.5, 1.5)
        latitudes.append(lat)
        longitudes.append(long)
        
        # Generate a location name based on coordinates
        if 'Infrastructure' in sector:
            locations.append(f"Infrastructure Project - {lat:.2f}, {long:.2f}")
        elif 'Building' in sector:
            locations.append(f"Building Site {np.random.randint(1,100)} - {lat:.2f}, {long:.2f}")
        elif 'Water' in sector:
            locations.append(f"Water Treatment Facility {np.random.randint(1,50)} - {lat:.2f}, {long:.2f}")
        elif 'Environment' in sector:
            locations.append(f"Environmental Site {np.random.randint(1,30)} - {lat:.2f}, {long:.2f}")
        elif 'Energy' in sector:
            locations.append(f"Energy Project {np.random.randint(1,40)} - {lat:.2f}, {long:.2f}")
        else:
            locations.append(f"Digital Transformation - {lat:.2f}, {long:.2f}")
    
    projects_df = pd.DataFrame({
        'project_id': project_ids,
        'project_name': [
            f'Project {chr(65+(i%26))}{i//26 if i//26 > 0 else ""} ({sectors[i][0:3]})' 
            for i in range(num_projects)
        ],
        'sector': sectors,
        'budget': budgets,
        'start_date': pd.to_datetime(start_dates),
        'planned_end_date': pd.to_datetime(end_dates),
        'status': statuses,
        'project_manager': np.random.choice(
            ['Alice Smith', 'Bob Johnson', 'Charlie Brown', 'Diana Prince', 'Ethan Hunt', 'Fiona Glenanne'], 
            num_projects
        ),
        'planned_duration_days': planned_durations,
        'target_cpi': target_cpi,
        'target_spi': target_spi,
        'strategic_alignment': np.random.choice(
            ['High', 'Medium', 'Low'], 
            num_projects, 
            p=[0.5, 0.4, 0.1]
        ),
        'latitude': latitudes,
        'longitude': longitudes,
        'location': locations
    })
    
    # Generate historical performance data
    historical_data = []
    for pid in project_ids:
        proj = projects_df[projects_df['project_id'] == pid].iloc[0]
        
        # Base performance metrics with some randomness
        base_cpi = np.random.normal(loc=1.0, scale=0.1)
        base_spi = np.random.normal(loc=1.0, scale=0.08)
        
        # Adjust base metrics based on project status
        if proj['status'] == 'Delayed':
            base_spi -= 0.15
            base_cpi -= 0.1
        elif proj['status'] == 'At Risk':
            base_spi -= 0.08
            base_cpi -= 0.05
        elif proj['status'] == 'Minor Issues':
            base_spi -= 0.03
            base_cpi -= 0.02
        
        # Generate historical data for each month
        for i in range(history_months):
            month_date = today - relativedelta(months=i)
            
            # Calculate metrics with some random fluctuation and gradual trend
            cpi = max(0.5, min(1.5, base_cpi + np.random.normal(scale=0.03) - (i * 0.002)))
            spi = max(0.5, min(1.5, base_spi + np.random.normal(scale=0.02) - (i * 0.003)))
            
            # Only add history if the month is after the project start date
            if month_date >= proj['start_date'].date():
                historical_data.append({
                    'project_id': pid,
                    'month': month_date.strftime('%Y-%m'),
                    'cpi': cpi,
                    'spi': spi
                })
    
    history_df = pd.DataFrame(historical_data)
    
    # Generate task data
    all_tasks_data = []
    task_id_counter = 1
    
    for pid in project_ids:
        proj = projects_df[projects_df['project_id'] == pid].iloc[0]
        proj_start = proj['start_date']
        proj_end = proj['planned_end_date']
        
        # Calculate project duration in days
        if pd.isna(proj_start) or pd.isna(proj_end) or proj_end <= proj_start:
            proj_duration = 0
        else:
            proj_duration = (proj_end - proj_start).days
        
        # Ensure base > 0 for random offset calculation
        proj_duration_offset_base = max(1, proj_duration - 30)
        
        task_categories = ['Design', 'Development', 'Testing', 'Implementation', 'Documentation']
        
        for i in range(tasks_per_project):
            # Random task start offset from project start
            task_planned_start_offset = np.random.randint(0, max(1, proj_duration_offset_base))
            
            # Random task duration based on project duration
            task_planned_duration = np.random.randint(
                10, 
                max(11, proj_duration // tasks_per_project if tasks_per_project > 0 else 90)
            )
            
            # Calculate planned start and end dates
            task_planned_start = proj_start + datetime.timedelta(days=task_planned_start_offset)
            task_planned_end = task_planned_start + datetime.timedelta(days=task_planned_duration)
            
            # Adjust metrics based on project status
            delay_factor = 0
            cost_factor = 1.0
            completion_factor = np.random.uniform(0.8, 1.0)
            
            if proj['status'] == 'Minor Issues':
                delay_factor = np.random.randint(0, 7)
                cost_factor = np.random.uniform(1.0, 1.07)
                completion_factor = np.random.uniform(0.7, 0.95)
            elif proj['status'] == 'At Risk':
                delay_factor = np.random.randint(2, 15)
                cost_factor = np.random.uniform(1.03, 1.18)
                completion_factor = np.random.uniform(0.5, 0.85)
            elif proj['status'] == 'Delayed':
                delay_factor = np.random.randint(10, 30)
                cost_factor = np.random.uniform(1.08, 1.30)
                completion_factor = np.random.uniform(0.3, 0.7)
            elif proj['status'] == 'Completed':
                delay_factor = np.random.randint(-5, 7)
                cost_factor = np.random.uniform(0.95, 1.1)
                completion_factor = 1.0
            
            # Calculate actual end date with delay
            task_actual_end = task_planned_end + datetime.timedelta(days=delay_factor)
            
            # Calculate costs
            task_planned_cost = np.random.randint(10000, 500000)
            task_actual_cost = task_planned_cost * cost_factor
            
            # Calculate earned value based on completion
            task_earned_value = task_planned_cost * completion_factor
            
            # Calculate expected progress based on current date
            days_since_start = (min(today, task_planned_end.date()) - task_planned_start.date()).days
            expected_progress = min(100, max(0, days_since_start / task_planned_duration * 100)) if task_planned_duration > 0 else 0
            
            # Create task category and name
            task_category = task_categories[i % len(task_categories)]
            task_name = f"{task_category} - Task {i+1}"
            
            all_tasks_data.append({
                'task_id': task_id_counter,
                'project_id': pid,
                'task_name': task_name,
                'category': task_category,
                'planned_start': task_planned_start,
                'planned_end': task_planned_end,
                'actual_end': task_actual_end,
                'planned_cost': task_planned_cost,
                'actual_cost': task_actual_cost,
                'earned_value': task_earned_value,
                'expected_progress': expected_progress
            })
            
            task_id_counter += 1
    
    tasks_df = pd.DataFrame(all_tasks_data)
    
    # Generate risk data
    all_risks_data = []
    risk_id_counter = 1
    risk_categories = [
        'Resource Availability', 
        'Scope Changes', 
        'Technical Issues', 
        'External Factors', 
        'Budget Constraints', 
        'Quality Issues', 
        'Communication Issues'
    ]
    
    for pid in project_ids:
        proj = projects_df[projects_df['project_id'] == pid].iloc[0]
        
        # Adjust number of risks based on project status
        status_risk_modifier = {
            'On Track': 0.7,
            'Minor Issues': 1.0,
            'At Risk': 1.5,
            'Delayed': 2.0,
            'Completed': 0.5
        }
        
        actual_risks = max(1, int(risks_per_project * status_risk_modifier.get(proj['status'], 1.0)))
        
        for i in range(actual_risks):
            risk_category = np.random.choice(risk_categories)
            
            # Probability based on project status
            if proj['status'] == 'On Track':
                probability = np.random.uniform(0.1, 0.3)
            elif proj['status'] == 'Minor Issues':
                probability = np.random.uniform(0.2, 0.4)
            elif proj['status'] == 'At Risk':
                probability = np.random.uniform(0.3, 0.6)
            elif proj['status'] == 'Delayed':
                probability = np.random.uniform(0.4, 0.7)
            else:  # Completed
                probability = np.random.uniform(0.1, 0.2)
            
            # Impact based on project budget
            impact_percentage = np.random.uniform(0.01, 0.1)
            impact_cost = proj['budget'] * impact_percentage
            
            # Calculate exposure
            exposure = impact_cost * probability
            
            # Calculate risk score (normalized 0-1)
            risk_score = min(1.0, probability * impact_percentage * 20)
            
            # Risk status
            if risk_score >= 0.6:
                status = np.random.choice(['Active', 'Mitigating'], p=[0.3, 0.7])
            elif risk_score >= 0.3:
                status = np.random.choice(['Active', 'Mitigating', 'Monitoring'], p=[0.2, 0.5, 0.3])
            else:
                status = np.random.choice(['Active', 'Monitoring', 'Closed'], p=[0.1, 0.6, 0.3])
            
            # Risk name and description
            risk_name = f"Risk {risk_id_counter}: {risk_category}"
            descriptions = {
                'Resource Availability': "Potential shortage of key personnel with required skills",
                'Scope Changes': "Client may request significant changes to project scope",
                'Technical Issues': "Technical challenges in system integration may cause delays",
                'External Factors': "Regulatory changes could impact project schedule and cost",
                'Budget Constraints': "Budget restrictions may limit available resources",
                'Quality Issues': "Deliverables may not meet expected quality standards",
                'Communication Issues': "Poor communication between stakeholders may lead to misalignment"
            }
            
            all_risks_data.append({
                'risk_id': risk_id_counter,
                'project_id': pid,
                'risk_name': risk_name,
                'description': descriptions.get(risk_category, "Risk description unavailable"),
                'category': risk_category,
                'probability': probability,
                'impact_cost': impact_cost,
                'exposure': exposure,
                'risk_score': risk_score,
                'status': status
            })
            
            risk_id_counter += 1
    
    risks_df = pd.DataFrame(all_risks_data)
    
    # Generate benefits data for ROI and benefits reporting
    today = datetime.date.today()
    history_months = max(12, history_months)  # Ensure we have at least 12 months for trends
    
    benefits_data = {
        'Month': [(today - relativedelta(months=i)).strftime('%Y-%m') for i in range(history_months)][::-1],
        'ReportingTimeSaved_hrs': (np.linspace(5, 40, history_months) + np.random.normal(0, 3, history_months)).clip(min=0),
        'CostOverrunsAvoided_k': (np.cumsum(np.random.uniform(10, 50, history_months)) + np.random.normal(0, 20, history_months)).clip(min=0),
        'ForecastAccuracy_perc': np.clip(np.linspace(60, 85, history_months) + np.random.normal(0, 4, history_months), 50, 95)
    }
    benefits_df = pd.DataFrame(benefits_data)
    
    # Generate change requests
    changes_data = []
    change_id = 1
    
    for pid in project_ids:
        proj = projects_df[projects_df['project_id'] == pid].iloc[0]
        num_changes = np.random.randint(0, 8)
        
        # Only generate changes if project has valid dates
        if pd.notna(proj['start_date']) and pd.notna(proj['planned_end_date']) and proj['planned_end_date'] > proj['start_date']:
            proj_duration = (proj['planned_end_date'] - proj['start_date']).days
            min_offset = 20
            max_offset = max(21, proj_duration - 20)
            
            if max_offset > min_offset:
                for i in range(num_changes):
                    # Impact cost based on project budget
                    min_impact = max(100, int(proj['budget'] * 0.001))
                    max_impact = max(200, int(proj['budget'] * 0.08))
                    impact_cost = np.random.randint(min_impact, max_impact)
                    
                    # Calculate submission date
                    submit_date = proj['start_date'] + datetime.timedelta(days=np.random.randint(min_offset, max_offset))
                    submit_date = min(submit_date.date(), today)
                    
                    changes_data.append({
                        'change_id': change_id,
                        'project_id': pid,
                        'description': f'Change Request {i+1}',
                        'impact_cost': impact_cost,
                        'impact_schedule_days': np.random.randint(-5, 45),
                        'status': np.random.choice(
                            ['Submitted', 'Approved', 'Rejected', 'Implemented', 'On Hold'], 
                            p=[0.3, 0.4, 0.1, 0.15, 0.05]
                        ),
                        'date_submitted': pd.to_datetime(submit_date)
                    })
                    change_id += 1
    
    changes_df = pd.DataFrame(changes_data) if changes_data else pd.DataFrame(
        columns=['change_id', 'project_id', 'description', 'impact_cost', 
                 'impact_schedule_days', 'status', 'date_submitted']
    )
    
    return projects_df, tasks_df, risks_df, changes_df, history_df, benefits_df
