#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Collaboration Module for PMO Pulse Dashboard

This module provides collaboration features including:
- Comments and annotations on metrics and charts
- Action item tracking
- Automated alerts system
- Meeting mode for presentations
- Report sharing and distribution

Author: Arcadis PMO Analytics Team
Version: 1.0
"""

import pandas as pd
import numpy as np
import os
import json
import uuid
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Define constants
COLLAB_DIR = 'collaboration'
if not os.path.exists(COLLAB_DIR):
    os.makedirs(COLLAB_DIR)
    
COMMENTS_FILE = os.path.join(COLLAB_DIR, 'comments.json')
ACTIONS_FILE = os.path.join(COLLAB_DIR, 'actions.json')
ALERTS_FILE = os.path.join(COLLAB_DIR, 'alerts.json')
SETTINGS_FILE = os.path.join(COLLAB_DIR, 'settings.json')

class CollaborationManager:
    """Manager class for collaboration features"""
    
    def __init__(self):
        """Initialize the collaboration manager"""
        self.comments = self._load_json(COMMENTS_FILE, [])
        self.actions = self._load_json(ACTIONS_FILE, [])
        self.alerts = self._load_json(ALERTS_FILE, [])
        self.settings = self._load_json(SETTINGS_FILE, {})
    
    def _load_json(self, file_path, default_value):
        """Helper to load JSON data with a default fallback"""
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading {file_path}: {e}")
        return default_value
    
    def _save_json(self, file_path, data):
        """Helper to save JSON data"""
        try:
            with open(file_path, 'w') as f:
                json.dump(data, f, indent=2, default=str)
            return True
        except Exception as e:
            print(f"Error saving to {file_path}: {e}")
            return False

class CommentSystem:
    """System for managing comments and annotations"""
    
    def __init__(self, collaboration_manager):
        """Initialize with a collaboration manager"""
        self.manager = collaboration_manager
        self.comments = self.manager.comments
    
    def add_comment(self, user_id, user_name, target_type, target_id, text, coordinates=None, attachments=None):
        """
        Add a new comment
        
        Parameters:
        -----------
        user_id : str
            ID of the user adding the comment
        user_name : str
            Display name of the user
        target_type : str
            Type of item being commented on ('project', 'metric', 'chart', etc.)
        target_id : str
            ID of the specific item
        text : str
            Comment text
        coordinates : dict, optional
            Coordinates for positioning (for annotations on charts)
        attachments : list, optional
            List of attachment file references
            
        Returns:
        --------
        str
            ID of the newly created comment
        """
        comment_id = str(uuid.uuid4())
        timestamp = datetime.now().isoformat()
        
        new_comment = {
            'id': comment_id,
            'user_id': user_id,
            'user_name': user_name,
            'target_type': target_type,
            'target_id': target_id,
            'text': text,
            'coordinates': coordinates,
            'attachments': attachments,
            'created_at': timestamp,
            'updated_at': timestamp,
            'replies': []
        }
        
        self.comments.append(new_comment)
        self.manager._save_json(COMMENTS_FILE, self.comments)
        
        return comment_id
    
    def add_reply(self, comment_id, user_id, user_name, text, attachments=None):
        """
        Add a reply to an existing comment
        
        Parameters:
        -----------
        comment_id : str
            ID of the comment to reply to
        user_id : str
            ID of the user adding the reply
        user_name : str
            Display name of the user
        text : str
            Reply text
        attachments : list, optional
            List of attachment file references
            
        Returns:
        --------
        str
            ID of the newly created reply
        """
        reply_id = str(uuid.uuid4())
        timestamp = datetime.now().isoformat()
        
        new_reply = {
            'id': reply_id,
            'user_id': user_id,
            'user_name': user_name,
            'text': text,
            'attachments': attachments,
            'created_at': timestamp,
            'updated_at': timestamp
        }
        
        # Find the comment to reply to
        for comment in self.comments:
            if comment['id'] == comment_id:
                if 'replies' not in comment:
                    comment['replies'] = []
                comment['replies'].append(new_reply)
                comment['updated_at'] = timestamp
                self.manager._save_json(COMMENTS_FILE, self.comments)
                return reply_id
        
        raise ValueError(f"Comment with ID {comment_id} not found")
    
    def edit_comment(self, comment_id, user_id, text, attachments=None):
        """
        Edit an existing comment
        
        Parameters:
        -----------
        comment_id : str
            ID of the comment to edit
        user_id : str
            ID of the user making the edit
        text : str
            Updated text
        attachments : list, optional
            Updated list of attachment file references
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        timestamp = datetime.now().isoformat()
        
        # Find the comment to edit
        for comment in self.comments:
            if comment['id'] == comment_id:
                # Check if user is allowed to edit this comment
                if comment['user_id'] != user_id:
                    raise PermissionError("User not authorized to edit this comment")
                
                comment['text'] = text
                if attachments is not None:
                    comment['attachments'] = attachments
                comment['updated_at'] = timestamp
                self.manager._save_json(COMMENTS_FILE, self.comments)
                return True
        
        # Check if it's a reply
        for comment in self.comments:
            if 'replies' in comment:
                for reply in comment['replies']:
                    if reply['id'] == comment_id:
                        # Check if user is allowed to edit this reply
                        if reply['user_id'] != user_id:
                            raise PermissionError("User not authorized to edit this reply")
                        
                        reply['text'] = text
                        if attachments is not None:
                            reply['attachments'] = attachments
                        reply['updated_at'] = timestamp
                        comment['updated_at'] = timestamp
                        self.manager._save_json(COMMENTS_FILE, self.comments)
                        return True
        
        raise ValueError(f"Comment or reply with ID {comment_id} not found")
    
    def delete_comment(self, comment_id, user_id, is_admin=False):
        """
        Delete a comment or reply
        
        Parameters:
        -----------
        comment_id : str
            ID of the comment or reply to delete
        user_id : str
            ID of the user requesting deletion
        is_admin : bool
            Whether the user has admin privileges
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        # Find and delete a top-level comment
        for i, comment in enumerate(self.comments):
            if comment['id'] == comment_id:
                # Check if user is allowed to delete this comment
                if comment['user_id'] != user_id and not is_admin:
                    raise PermissionError("User not authorized to delete this comment")
                
                del self.comments[i]
                self.manager._save_json(COMMENTS_FILE, self.comments)
                return True
        
        # Check if it's a reply
        for comment in self.comments:
            if 'replies' in comment:
                for i, reply in enumerate(comment['replies']):
                    if reply['id'] == comment_id:
                        # Check if user is allowed to delete this reply
                        if reply['user_id'] != user_id and not is_admin:
                            raise PermissionError("User not authorized to delete this reply")
                        
                        del comment['replies'][i]
                        comment['updated_at'] = datetime.now().isoformat()
                        self.manager._save_json(COMMENTS_FILE, self.comments)
                        return True
        
        raise ValueError(f"Comment or reply with ID {comment_id} not found")
    
    def get_comments(self, target_type=None, target_id=None):
        """
        Get comments, optionally filtered by target
        
        Parameters:
        -----------
        target_type : str, optional
            Type of item to filter by
        target_id : str, optional
            ID of the specific item to filter by
            
        Returns:
        --------
        list
            List of matching comments
        """
        if target_type is None and target_id is None:
            # Return all comments
            return self.comments
        
        filtered_comments = []
        
        for comment in self.comments:
            if target_type is not None and comment['target_type'] != target_type:
                continue
            if target_id is not None and comment['target_id'] != target_id:
                continue
            filtered_comments.append(comment)
        
        return filtered_comments

class ActionTracker:
    """System for tracking action items"""
    
    def __init__(self, collaboration_manager):
        """Initialize with a collaboration manager"""
        self.manager = collaboration_manager
        self.actions = self.manager.actions
    
    def add_action(self, title, description, assigned_to, due_date, priority, status="Open", project_id=None):
        """
        Add a new action item
        
        Parameters:
        -----------
        title : str
            Short title of the action
        description : str
            Detailed description
        assigned_to : str
            ID of the user assigned to the action
        due_date : str
            Due date in ISO format (YYYY-MM-DD)
        priority : str
            Priority level ('Low', 'Medium', 'High', 'Critical')
        status : str
            Status of the action ('Open', 'In Progress', 'Completed', 'Deferred')
        project_id : str, optional
            ID of the associated project
            
        Returns:
        --------
        str
            ID of the newly created action item
        """
        action_id = str(uuid.uuid4())
        timestamp = datetime.now().isoformat()
        
        new_action = {
            'id': action_id,
            'title': title,
            'description': description,
            'assigned_to': assigned_to,
            'due_date': due_date,
            'priority': priority,
            'status': status,
            'project_id': project_id,
            'created_at': timestamp,
            'updated_at': timestamp,
            'completed_at': None,
            'comments': []
        }
        
        self.actions.append(new_action)
        self.manager._save_json(ACTIONS_FILE, self.actions)
        
        return action_id
    
    def update_action(self, action_id, updated_fields):
        """
        Update an existing action item
        
        Parameters:
        -----------
        action_id : str
            ID of the action to update
        updated_fields : dict
            Dictionary of fields to update
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        timestamp = datetime.now().isoformat()
        
        for action in self.actions:
            if action['id'] == action_id:
                # Update the specified fields
                for key, value in updated_fields.items():
                    if key in action:
                        action[key] = value
                
                # Special handling for status changes
                if 'status' in updated_fields and updated_fields['status'] == 'Completed':
                    action['completed_at'] = timestamp
                
                action['updated_at'] = timestamp
                self.manager._save_json(ACTIONS_FILE, self.actions)
                return True
        
        raise ValueError(f"Action with ID {action_id} not found")
    
    def delete_action(self, action_id):
        """
        Delete an action item
        
        Parameters:
        -----------
        action_id : str
            ID of the action to delete
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        for i, action in enumerate(self.actions):
            if action['id'] == action_id:
                del self.actions[i]
                self.manager._save_json(ACTIONS_FILE, self.actions)
                return True
        
        raise ValueError(f"Action with ID {action_id} not found")
    
    def add_action_comment(self, action_id, user_id, user_name, text):
        """
        Add a comment to an action item
        
        Parameters:
        -----------
        action_id : str
            ID of the action
        user_id : str
            ID of the user adding the comment
        user_name : str
            Display name of the user
        text : str
            Comment text
            
        Returns:
        --------
        str
            ID of the newly created comment
        """
        comment_id = str(uuid.uuid4())
        timestamp = datetime.now().isoformat()
        
        new_comment = {
            'id': comment_id,
            'user_id': user_id,
            'user_name': user_name,
            'text': text,
            'created_at': timestamp
        }
        
        for action in self.actions:
            if action['id'] == action_id:
                if 'comments' not in action:
                    action['comments'] = []
                action['comments'].append(new_comment)
                action['updated_at'] = timestamp
                self.manager._save_json(ACTIONS_FILE, self.actions)
                return comment_id
        
        raise ValueError(f"Action with ID {action_id} not found")
    
    def get_actions(self, assigned_to=None, status=None, project_id=None, priority=None):
        """
        Get actions, optionally filtered
        
        Parameters:
        -----------
        assigned_to : str, optional
            Filter by assigned user ID
        status : str or list, optional
            Filter by status
        project_id : str, optional
            Filter by project ID
        priority : str or list, optional
            Filter by priority level
            
        Returns:
        --------
        list
            List of matching action items
        """
        filtered_actions = []
        
        for action in self.actions:
            if assigned_to is not None and action['assigned_to'] != assigned_to:
                continue
                
            if status is not None:
                if isinstance(status, list) and action['status'] not in status:
                    continue
                elif not isinstance(status, list) and action['status'] != status:
                    continue
                    
            if project_id is not None and action['project_id'] != project_id:
                continue
                
            if priority is not None:
                if isinstance(priority, list) and action['priority'] not in priority:
                    continue
                elif not isinstance(priority, list) and action['priority'] != priority:
                    continue
            
            filtered_actions.append(action)
        
        return filtered_actions
    
    def get_overdue_actions(self):
        """
        Get all overdue action items
        
        Returns:
        --------
        list
            List of overdue action items
        """
        today = datetime.now().date()
        overdue_actions = []
        
        for action in self.actions:
            if action['status'] in ['Open', 'In Progress']:
                due_date = datetime.fromisoformat(action['due_date']).date()
                if due_date < today:
                    overdue_actions.append(action)
        
        return overdue_actions

class AlertSystem:
    """System for automated alerts and notifications"""
    
    def __init__(self, collaboration_manager):
        """Initialize with a collaboration manager"""
        self.manager = collaboration_manager
        self.alerts = self.manager.alerts
        self.settings = self.manager.settings.get('alerts', {
            'email_enabled': False,
            'email_server': '',
            'email_port': 587,
            'email_username': '',
            'email_password': '',
            'email_from': '',
            'email_use_tls': True
        })
    
    def create_alert_rule(self, name, description, metric_type, project_id, threshold_value, 
                         comparison_operator, recipients, alert_channels, frequency="daily"):
        """
        Create a new alert rule
        
        Parameters:
        -----------
        name : str
            Name of the alert rule
        description : str
            Description of what the alert monitors
        metric_type : str
            Type of metric to monitor ('spi', 'cpi', 'budget_variance', etc.)
        project_id : str
            ID of the associated project, or 'all' for all projects
        threshold_value : float
            Threshold value that triggers the alert
        comparison_operator : str
            Operator for comparison ('<', '<=', '=', '>=', '>')
        recipients : list
            List of recipient user IDs
        alert_channels : list
            List of notification channels ('email', 'dashboard', etc.)
        frequency : str
            How often to check ('realtime', 'hourly', 'daily', 'weekly')
            
        Returns:
        --------
        str
            ID of the newly created alert rule
        """
        rule_id = str(uuid.uuid4())
        timestamp = datetime.now().isoformat()
        
        new_rule = {
            'id': rule_id,
            'name': name,
            'description': description,
            'metric_type': metric_type,
            'project_id': project_id,
            'threshold_value': threshold_value,
            'comparison_operator': comparison_operator,
            'recipients': recipients,
            'alert_channels': alert_channels,
            'frequency': frequency,
            'created_at': timestamp,
            'updated_at': timestamp,
            'enabled': True,
            'last_triggered': None,
            'triggered_count': 0
        }
        
        self.alerts.append(new_rule)
        self.manager._save_json(ALERTS_FILE, self.alerts)
        
        return rule_id
    
    def update_alert_rule(self, rule_id, updated_fields):
        """
        Update an existing alert rule
        
        Parameters:
        -----------
        rule_id : str
            ID of the rule to update
        updated_fields : dict
            Dictionary of fields to update
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        timestamp = datetime.now().isoformat()
        
        for rule in self.alerts:
            if rule['id'] == rule_id:
                # Update the specified fields
                for key, value in updated_fields.items():
                    if key in rule:
                        rule[key] = value
                
                rule['updated_at'] = timestamp
                self.manager._save_json(ALERTS_FILE, self.alerts)
                return True
        
        raise ValueError(f"Alert rule with ID {rule_id} not found")
    
    def delete_alert_rule(self, rule_id):
        """
        Delete an alert rule
        
        Parameters:
        -----------
        rule_id : str
            ID of the rule to delete
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        for i, rule in enumerate(self.alerts):
            if rule['id'] == rule_id:
                del self.alerts[i]
                self.manager._save_json(ALERTS_FILE, self.alerts)
                return True
        
        raise ValueError(f"Alert rule with ID {rule_id} not found")
    
    def check_alerts(self, project_data, user_data=None):
        """
        Check all alert rules against current data
        
        Parameters:
        -----------
        project_data : pandas DataFrame
            Current project performance data
        user_data : dict, optional
            User information for notification
            
        Returns:
        --------
        list
            List of triggered alerts
        """
        if project_data.empty:
            return []
        
        timestamp = datetime.now().isoformat()
        triggered_alerts = []
        
        for rule in self.alerts:
            if not rule['enabled']:
                continue
            
            metric_type = rule['metric_type']
            threshold = rule['threshold_value']
            operator = rule['comparison_operator']
            
            # Filter by project if specified
            if rule['project_id'] != 'all':
                relevant_data = project_data[project_data['project_id'] == rule['project_id']]
            else:
                relevant_data = project_data
            
            if relevant_data.empty:
                continue
            
            # Check the condition
            triggered_projects = []
            
            for _, project in relevant_data.iterrows():
                if metric_type in project:
                    value = project[metric_type]
                    project_id = project['project_id']
                    project_name = project.get('project_name', project_id)
                    
                    condition_met = False
                    if operator == '<' and value < threshold:
                        condition_met = True
                    elif operator == '<=' and value <= threshold:
                        condition_met = True
                    elif operator == '=' and value == threshold:
                        condition_met = True
                    elif operator == '>=' and value >= threshold:
                        condition_met = True
                    elif operator == '>' and value > threshold:
                        condition_met = True
                    
                    if condition_met:
                        triggered_projects.append({
                            'project_id': project_id,
                            'project_name': project_name,
                            'value': value
                        })
            
            if triggered_projects:
                # Update rule stats
                rule['last_triggered'] = timestamp
                rule['triggered_count'] += 1
                
                # Create alert notification
                alert = {
                    'id': str(uuid.uuid4()),
                    'rule_id': rule['id'],
                    'rule_name': rule['name'],
                    'metric_type': metric_type,
                    'threshold_value': threshold,
                    'comparison_operator': operator,
                    'triggered_at': timestamp,
                    'triggered_projects': triggered_projects,
                    'recipients': rule['recipients'],
                    'sent_notifications': []
                }
                
                triggered_alerts.append(alert)
                
                # Process notifications
                self._send_notifications(alert, user_data)
        
        # Save updated alert rules
        if triggered_alerts:
            self.manager._save_json(ALERTS_FILE, self.alerts)
        
        return triggered_alerts
    
    def _send_notifications(self, alert, user_data):
        """
        Send notifications for a triggered alert
        
        Parameters:
        -----------
        alert : dict
            Alert information
        user_data : dict
            User information for notifications
        """
        # Get the rule details
        rule_id = alert['rule_id']
        rule = None
        
        for r in self.alerts:
            if r['id'] == rule_id:
                rule = r
                break
        
        if not rule:
            return
        
        # Process each notification channel
        for channel in rule['alert_channels']:
            if channel == 'email' and self.settings['email_enabled']:
                self._send_email_notification(alert, rule, user_data)
            
            # Record the notification
            alert['sent_notifications'].append({
                'channel': channel,
                'timestamp': datetime.now().isoformat(),
                'status': 'sent'
            })
    
    def _send_email_notification(self, alert, rule, user_data):
        """
        Send email notification
        
        Parameters:
        -----------
        alert : dict
            Alert information
        rule : dict
            Alert rule details
        user_data : dict
            User information with email addresses
        """
        if not user_data:
            return
        
        # Prepare email content
        subject = f"PMO Pulse Alert: {rule['name']}"
        
        # Build the email body
        body = f"""<html>
        <body>
            <h2>PMO Pulse Dashboard Alert</h2>
            <p><strong>{rule['name']}</strong></p>
            <p>{rule['description']}</p>
            <p>Metric: {alert['metric_type']}</p>
            <p>Condition: {alert['comparison_operator']} {alert['threshold_value']}</p>
            <h3>Affected Projects:</h3>
            <ul>
        """
        
        for project in alert['triggered_projects']:
            body += f"<li>{project['project_name']}: {project['value']}</li>"
        
        body += """</ul>
            <p>Please log in to the PMO Pulse Dashboard for more details.</p>
        </body>
        </html>"""
        
        # Get recipient email addresses
        recipients = []
        for user_id in rule['recipients']:
            if user_id in user_data and 'email' in user_data[user_id]:
                recipients.append(user_data[user_id]['email'])
        
        if not recipients:
            return
        
        # Send the email
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.settings['email_from']
            msg['To'] = ", ".join(recipients)
            msg['Subject'] = subject
            
            # Add HTML body
            msg.attach(MIMEText(body, 'html'))
            
            # Connect to server and send
            server = smtplib.SMTP(self.settings['email_server'], self.settings['email_port'])
            if self.settings['email_use_tls']:
                server.starttls()
            
            if self.settings['email_username'] and self.settings['email_password']:
                server.login(self.settings['email_username'], self.settings['email_password'])
            
            server.send_message(msg)
            server.quit()
            
        except Exception as e:
            print(f"Error sending email notification: {e}")
    
    def update_email_settings(self, settings):
        """
        Update email notification settings
        
        Parameters:
        -----------
        settings : dict
            Dictionary of email settings
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        updated_settings = self.settings.copy()
        updated_settings.update(settings)
        
        self.settings = updated_settings
        self.manager.settings['alerts'] = updated_settings
        
        return self.manager._save_json(SETTINGS_FILE, self.manager.settings)
    
    def get_alert_rules(self, project_id=None, enabled_only=False):
        """
        Get alert rules, optionally filtered
        
        Parameters:
        -----------
        project_id : str, optional
            Filter by project ID
        enabled_only : bool
            Whether to return only enabled rules
            
        Returns:
        --------
        list
            List of matching alert rules
        """
        filtered_rules = []
        
        for rule in self.alerts:
            if enabled_only and not rule['enabled']:
                continue
                
            if project_id is not None and rule['project_id'] != project_id and rule['project_id'] != 'all':
                continue
            
            filtered_rules.append(rule)
        
        return filtered_rules

class MeetingMode:
    """System for presentation mode"""
    
    def __init__(self, collaboration_manager):
        """Initialize with a collaboration manager"""
        self.manager = collaboration_manager
        self.settings = self.manager.settings.get('meeting_mode', {
            'default_slides': [
                'portfolio_overview',
                'at_risk_projects',
                'performance_trends',
                'upcoming_milestones'
            ],
            'slide_duration': 30,  # seconds
            'auto_advance': False,
            'show_annotations': True,
            'presenter_notes_visible': False
        })
    
    def update_settings(self, settings):
        """
        Update meeting mode settings
        
        Parameters:
        -----------
        settings : dict
            Dictionary of meeting mode settings
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        updated_settings = self.settings.copy()
        updated_settings.update(settings)
        
        self.settings = updated_settings
        self.manager.settings['meeting_mode'] = updated_settings
        
        return self.manager._save_json(SETTINGS_FILE, self.manager.settings)
    
    def create_presentation(self, slides=None, title=None, subtitle=None):
        """
        Create a new presentation configuration
        
        Parameters:
        -----------
        slides : list, optional
            List of slide IDs to include (defaults to default_slides)
        title : str, optional
            Presentation title
        subtitle : str, optional
            Presentation subtitle
            
        Returns:
        --------
        dict
            Presentation configuration
        """
        if slides is None:
            slides = self.settings['default_slides']
        
        if title is None:
            title = "PMO Pulse Dashboard"
            
        timestamp = datetime.now().isoformat()
        
        presentation = {
            'id': str(uuid.uuid4()),
            'title': title,
            'subtitle': subtitle,
            'created_at': timestamp,
            'last_used': timestamp,
            'slides': slides,
            'duration': self.settings['slide_duration'],
            'auto_advance': self.settings['auto_advance'],
            'show_annotations': self.settings['show_annotations']
        }
        
        # Store in settings
        if 'presentations' not in self.manager.settings:
            self.manager.settings['presentations'] = []
        
        self.manager.settings['presentations'].append(presentation)
        self.manager._save_json(SETTINGS_FILE, self.manager.settings)
        
        return presentation
    
    def get_presentations(self):
        """
        Get all saved presentations
        
        Returns:
        --------
        list
            List of presentation configurations
        """
        return self.manager.settings.get('presentations', [])
    
    def get_presentation(self, presentation_id):
        """
        Get a specific presentation
        
        Parameters:
        -----------
        presentation_id : str
            ID of the presentation
            
        Returns:
        --------
        dict
            Presentation configuration
        """
        presentations = self.get_presentations()
        
        for presentation in presentations:
            if presentation['id'] == presentation_id:
                return presentation
        
        return None
    
    def update_presentation(self, presentation_id, updated_fields):
        """
        Update a presentation
        
        Parameters:
        -----------
        presentation_id : str
            ID of the presentation to update
        updated_fields : dict
            Dictionary of fields to update
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        presentations = self.get_presentations()
        
        for presentation in presentations:
            if presentation['id'] == presentation_id:
                # Update fields
                for key, value in updated_fields.items():
                    presentation[key] = value
                
                # Update last_used timestamp
                presentation['last_used'] = datetime.now().isoformat()
                
                # Save changes
                self.manager.settings['presentations'] = presentations
                return self.manager._save_json(SETTINGS_FILE, self.manager.settings)
        
        return False
    
    def delete_presentation(self, presentation_id):
        """
        Delete a presentation
        
        Parameters:
        -----------
        presentation_id : str
            ID of the presentation to delete
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        presentations = self.get_presentations()
        
        for i, presentation in enumerate(presentations):
            if presentation['id'] == presentation_id:
                del presentations[i]
                self.manager.settings['presentations'] = presentations
                return self.manager._save_json(SETTINGS_FILE, self.manager.settings)
        
        return False

class ReportDistribution:
    """System for sharing and distributing reports"""
    
    def __init__(self, collaboration_manager):
        """Initialize with a collaboration manager"""
        self.manager = collaboration_manager
        self.settings = self.manager.settings.get('report_distribution', {
            'default_format': 'pdf',
            'email_enabled': False,
            'email_template': "Please find attached the latest PMO Pulse report.",
            'auto_distribution_enabled': False,
            'default_recipients': []
        })
    
    def update_settings(self, settings):
        """
        Update report distribution settings
        
        Parameters:
        -----------
        settings : dict
            Dictionary of report distribution settings
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        updated_settings = self.settings.copy()
        updated_settings.update(settings)
        
        self.settings = updated_settings
        self.manager.settings['report_distribution'] = updated_settings
        
        return self.manager._save_json(SETTINGS_FILE, self.manager.settings)
    
    def create_report_schedule(self, name, description, report_type, recipients, 
                             frequency, parameters=None, format='pdf'):
        """
        Create a scheduled report
        
        Parameters:
        -----------
        name : str
            Name of the scheduled report
        description : str
            Description of the report
        report_type : str
            Type of report ('portfolio_overview', 'project_detail', etc.)
        recipients : list
            List of recipient user IDs
        frequency : str
            Schedule frequency ('daily', 'weekly', 'monthly')
        parameters : dict, optional
            Additional parameters for the report
        format : str
            Report format ('pdf', 'excel', 'powerpoint')
            
        Returns:
        --------
        str
            ID of the newly created report schedule
        """
        schedule_id = str(uuid.uuid4())
        timestamp = datetime.now().isoformat()
        
        new_schedule = {
            'id': schedule_id,
            'name': name,
            'description': description,
            'report_type': report_type,
            'recipients': recipients,
            'frequency': frequency,
            'parameters': parameters or {},
            'format': format,
            'created_at': timestamp,
            'updated_at': timestamp,
            'last_run': None,
            'enabled': True
        }
        
        # Store in settings
        if 'report_schedules' not in self.manager.settings:
            self.manager.settings['report_schedules'] = []
        
        self.manager.settings['report_schedules'].append(new_schedule)
        self.manager._save_json(SETTINGS_FILE, self.manager.settings)
        
        return schedule_id
    
    def update_report_schedule(self, schedule_id, updated_fields):
        """
        Update a report schedule
        
        Parameters:
        -----------
        schedule_id : str
            ID of the schedule to update
        updated_fields : dict
            Dictionary of fields to update
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        schedules = self.get_report_schedules()
        
        for schedule in schedules:
            if schedule['id'] == schedule_id:
                # Update fields
                for key, value in updated_fields.items():
                    schedule[key] = value
                
                schedule['updated_at'] = datetime.now().isoformat()
                
                # Save changes
                self.manager.settings['report_schedules'] = schedules
                return self.manager._save_json(SETTINGS_FILE, self.manager.settings)
        
        return False
    
    def delete_report_schedule(self, schedule_id):
        """
        Delete a report schedule
        
        Parameters:
        -----------
        schedule_id : str
            ID of the schedule to delete
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        schedules = self.get_report_schedules()
        
        for i, schedule in enumerate(schedules):
            if schedule['id'] == schedule_id:
                del schedules[i]
                self.manager.settings['report_schedules'] = schedules
                return self.manager._save_json(SETTINGS_FILE, self.manager.settings)
        
        return False
    
    def get_report_schedules(self, enabled_only=False):
        """
        Get all report schedules
        
        Parameters:
        -----------
        enabled_only : bool
            Whether to return only enabled schedules
            
        Returns:
        --------
        list
            List of report schedules
        """
        schedules = self.manager.settings.get('report_schedules', [])
        
        if enabled_only:
            return [s for s in schedules if s.get('enabled', True)]
        
        return schedules
    
    def send_report(self, report_data, recipients, format='pdf', subject=None, message=None):
        """
        Send a report to recipients
        
        Parameters:
        -----------
        report_data : bytes
            Report file data
        recipients : list
            List of email addresses
        format : str
            Report format ('pdf', 'excel', 'powerpoint')
        subject : str, optional
            Email subject
        message : str, optional
            Email message
            
        Returns:
        --------
        bool
            True if successful, False otherwise
        """
        if not self.settings['email_enabled']:
            return False
        
        # Get email settings from alerts
        alert_settings = self.manager.settings.get('alerts', {})
        
        if not alert_settings.get('email_enabled'):
            return False
        
        # Prepare email
        if subject is None:
            subject = "PMO Pulse Dashboard Report"
        
        if message is None:
            message = self.settings['email_template']
        
        # Create email message
        msg = MIMEMultipart()
        msg['From'] = alert_settings['email_from']
        msg['To'] = ", ".join(recipients)
        msg['Subject'] = subject
        
        # Add message body
        msg.attach(MIMEText(message, 'plain'))
        
        # Add attachment
        filename = f"pmo_pulse_report.{format}"
        attachment = MIMEApplication(report_data)
        attachment['Content-Disposition'] = f'attachment; filename="{filename}"'
        msg.attach(attachment)
        
        # Send email
        try:
            server = smtplib.SMTP(alert_settings['email_server'], alert_settings['email_port'])
            if alert_settings['email_use_tls']:
                server.starttls()
            
            if alert_settings['email_username'] and alert_settings['email_password']:
                server.login(alert_settings['email_username'], alert_settings['email_password'])
            
            server.send_message(msg)
            server.quit()
            return True
            
        except Exception as e:
            print(f"Error sending report email: {e}")
            return False
    
    def process_scheduled_reports(self, data_manager, export_manager, user_manager=None):
        """
        Process all scheduled reports that need to be run
        
        Parameters:
        -----------
        data_manager : object
            Data manager with access to project data
        export_manager : object
            Export manager for generating reports
        user_manager : object, optional
            User manager for recipient information
            
        Returns:
        --------
        list
            List of processed report schedules
        """
        schedules = self.get_report_schedules(enabled_only=True)
        now = datetime.now()
        processed = []
        
        for schedule in schedules:
            # Check if it's time to run the schedule
            if schedule['last_run'] is not None:
                last_run = datetime.fromisoformat(schedule['last_run'])
                
                # Determine if the schedule should run based on frequency
                if schedule['frequency'] == 'daily':
                    if (now - last_run).days < 1:
                        continue
                elif schedule['frequency'] == 'weekly':
                    if (now - last_run).days < 7:
                        continue
                elif schedule['frequency'] == 'monthly':
                    if (now - last_run).days < 30:
                        continue
            
            # Generate the report
            try:
                # This would call the export manager to generate a report
                # For example: report_data = export_manager.generate_report(schedule['report_type'], schedule['parameters'], schedule['format'])
                report_data = b''  # Placeholder
                
                # Get recipient email addresses
                recipient_emails = []
                if user_manager is not None:
                    for user_id in schedule['recipients']:
                        user = user_manager.get_user(user_id)
                        if user and 'email' in user:
                            recipient_emails.append(user['email'])
                
                # Send the report
                if recipient_emails:
                    success = self.send_report(
                        report_data, 
                        recipient_emails, 
                        schedule['format'],
                        f"PMO Pulse {schedule['name']} Report"
                    )
                    
                    if success:
                        schedule['last_run'] = now.isoformat()
                        processed.append(schedule)
            except Exception as e:
                print(f"Error processing scheduled report {schedule['id']}: {e}")
        
        # Save updated schedules
        if processed:
            self.manager.settings['report_schedules'] = schedules
            self.manager._save_json(SETTINGS_FILE, self.manager.settings)
        
        return processed

# Main function to test the module
def test_collaboration():
    """Test function for the collaboration module"""
    manager = CollaborationManager()
    
    # Test comment system
    comments = CommentSystem(manager)
    comment_id = comments.add_comment(
        user_id="user1",
        user_name="Test User",
        target_type="project",
        target_id="proj1",
        text="This is a test comment"
    )
    print(f"Added comment with ID: {comment_id}")
    
    # Test action tracker
    actions = ActionTracker(manager)
    action_id = actions.add_action(
        title="Test Action",
        description="This is a test action",
        assigned_to="user1",
        due_date=datetime.now().date().isoformat(),
        priority="Medium",
        status="Open",
        project_id="proj1"
    )
    print(f"Added action with ID: {action_id}")
    
    # Test alert system
    alerts = AlertSystem(manager)
    rule_id = alerts.create_alert_rule(
        name="SPI Alert",
        description="Alert when SPI drops below threshold",
        metric_type="spi",
        project_id="all",
        threshold_value=0.8,
        comparison_operator="<",
        recipients=["user1"],
        alert_channels=["dashboard"],
        frequency="daily"
    )
    print(f"Added alert rule with ID: {rule_id}")
    
    # Clean up - comment out to see data persisted to files
    # import os
    # for file in [COMMENTS_FILE, ACTIONS_FILE, ALERTS_FILE, SETTINGS_FILE]:
    #     if os.path.exists(file):
    #         os.remove(file)
    # print("Test files removed")

if __name__ == "__main__":
    test_collaboration()