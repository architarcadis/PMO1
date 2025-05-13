#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Data Integration Module for PMO Pulse Dashboard

This module provides connectors to external data sources including:
- Jira for issue and project tracking
- Microsoft Project for project schedules
- Excel Online/SharePoint for custom data
- REST API connector for other systems
- Database connectors for SQL-based systems

Author: Arcadis PMO Analytics Team
Version: 1.0
"""

import pandas as pd
import numpy as np
import requests
import json
import os
import io
import base64
from datetime import datetime, timedelta
import urllib.parse

# Define constants
CONFIG_DIR = 'config'
if not os.path.exists(CONFIG_DIR):
    os.makedirs(CONFIG_DIR)
    
CONFIG_FILE = os.path.join(CONFIG_DIR, 'data_sources.json')

class DataIntegrationManager:
    """Manager class for data integration sources"""
    
    def __init__(self):
        """Initialize the data integration manager"""
        self.sources = {}
        self.load_config()
    
    def load_config(self):
        """Load configuration from saved file"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    self.sources = json.load(f)
            except Exception as e:
                print(f"Error loading configuration: {e}")
    
    def save_config(self):
        """Save current configuration to file"""
        try:
            with open(CONFIG_FILE, 'w') as f:
                # Don't save sensitive information like API keys
                safe_config = {}
                for source_id, source in self.sources.items():
                    safe_config[source_id] = {
                        k: v for k, v in source.items() 
                        if k not in ['api_key', 'password', 'token', 'secret']
                    }
                json.dump(safe_config, f, indent=2)
        except Exception as e:
            print(f"Error saving configuration: {e}")
    
    def add_source(self, source_id, source_type, name, connection_params, credentials=None):
        """
        Add a new data source
        
        Parameters:
        -----------
        source_id : str
            Unique identifier for the source
        source_type : str
            Type of source ('jira', 'ms_project', 'excel', 'api', 'database')
        name : str
            Display name for the source
        connection_params : dict
            Connection parameters (urls, endpoints, etc.)
        credentials : dict, optional
            Authentication credentials (will not be saved to disk)
        """
        self.sources[source_id] = {
            'type': source_type,
            'name': name,
            'connection': connection_params,
            'last_sync': None
        }
        
        # Store credentials in memory but not in saved config
        if credentials:
            self.sources[source_id].update(credentials)
            
        self.save_config()
    
    def remove_source(self, source_id):
        """Remove a data source"""
        if source_id in self.sources:
            del self.sources[source_id]
            self.save_config()
    
    def get_source(self, source_id):
        """Get a specific data source"""
        return self.sources.get(source_id, None)
    
    def list_sources(self):
        """List all configured data sources"""
        return [
            {
                'id': source_id,
                'name': source['name'],
                'type': source['type'],
                'last_sync': source.get('last_sync', None)
            }
            for source_id, source in self.sources.items()
        ]
    
    def fetch_data(self, source_id, entity_type, params=None):
        """
        Fetch data from a configured source
        
        Parameters:
        -----------
        source_id : str
            ID of the source to fetch from
        entity_type : str
            Type of entity to fetch ('projects', 'tasks', 'issues', etc.)
        params : dict, optional
            Additional parameters for the query
            
        Returns:
        --------
        pandas.DataFrame
            Data fetched from the source
        """
        source = self.get_source(source_id)
        if not source:
            raise ValueError(f"Data source '{source_id}' not found")
        
        # Create appropriate connector based on source type
        if source['type'] == 'jira':
            connector = JiraConnector(source)
        elif source['type'] == 'ms_project':
            connector = MSProjectConnector(source)
        elif source['type'] == 'excel':
            connector = ExcelConnector(source)
        elif source['type'] == 'api':
            connector = RestApiConnector(source)
        elif source['type'] == 'database':
            connector = DatabaseConnector(source)
        else:
            raise ValueError(f"Unsupported source type: {source['type']}")
        
        # Fetch and return the data
        result = connector.fetch(entity_type, params)
        
        # Update last sync time
        source['last_sync'] = datetime.now().isoformat()
        self.save_config()
        
        return result

class BaseConnector:
    """Base class for data connectors"""
    
    def __init__(self, source_config):
        """Initialize with source configuration"""
        self.config = source_config
    
    def fetch(self, entity_type, params=None):
        """
        Fetch data for the specified entity type
        
        Parameters:
        -----------
        entity_type : str
            Type of entity to fetch
        params : dict, optional
            Additional parameters for the query
            
        Returns:
        --------
        pandas.DataFrame
            Fetched data
        """
        raise NotImplementedError("Subclasses must implement fetch method")
    
    def test_connection(self):
        """
        Test the connection to the data source
        
        Returns:
        --------
        bool
            True if connection is successful, False otherwise
        """
        raise NotImplementedError("Subclasses must implement test_connection method")

class JiraConnector(BaseConnector):
    """Connector for Jira"""
    
    def fetch(self, entity_type, params=None):
        """Fetch data from Jira"""
        if 'api_key' not in self.config and 'token' not in self.config:
            raise ValueError("Authentication credentials missing for Jira connector")
        
        # Get basic connection parameters
        base_url = self.config['connection'].get('url')
        if not base_url:
            raise ValueError("Jira URL not specified in connector configuration")
        
        # Construct the API endpoint based on entity type
        if entity_type == 'projects':
            endpoint = f"{base_url}/rest/api/2/project"
        elif entity_type == 'issues':
            endpoint = f"{base_url}/rest/api/2/search"
        elif entity_type == 'sprints':
            endpoint = f"{base_url}/rest/agile/1.0/board/{params.get('board_id', '')}/sprint"
        else:
            raise ValueError(f"Unsupported entity type for Jira: {entity_type}")
        
        # Prepare authentication
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        
        if 'api_key' in self.config:
            headers['Authorization'] = f"Bearer {self.config['api_key']}"
        elif 'token' in self.config:
            auth_str = f"{self.config['username']}:{self.config['token']}"
            auth_bytes = auth_str.encode('ascii')
            base64_bytes = base64.b64encode(auth_bytes)
            headers['Authorization'] = f"Basic {base64_bytes.decode('ascii')}"
        
        # Prepare request parameters
        request_params = {}
        if params:
            if entity_type == 'issues':
                request_params['jql'] = params.get('jql', 'project IS NOT EMPTY')
                request_params['maxResults'] = params.get('max_results', 100)
                request_params['startAt'] = params.get('start_at', 0)
                request_params['fields'] = params.get('fields', '*all')
        
        # Make the request (simplified, real implementation would handle pagination)
        try:
            response = requests.get(endpoint, headers=headers, params=request_params)
            response.raise_for_status()
            data = response.json()
            
            # Convert to DataFrame (structure depends on entity type)
            if entity_type == 'projects':
                df = pd.json_normalize(data)
            elif entity_type == 'issues':
                df = pd.json_normalize(data['issues'])
            elif entity_type == 'sprints':
                df = pd.json_normalize(data['values'])
            
            return df
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data from Jira: {e}")
            return pd.DataFrame()
    
    def test_connection(self):
        """Test Jira connection"""
        try:
            # Try to fetch a simple endpoint that requires authentication
            base_url = self.config['connection'].get('url')
            endpoint = f"{base_url}/rest/api/2/myself"
            
            headers = {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            }
            
            if 'api_key' in self.config:
                headers['Authorization'] = f"Bearer {self.config['api_key']}"
            elif 'token' in self.config:
                auth_str = f"{self.config['username']}:{self.config['token']}"
                auth_bytes = auth_str.encode('ascii')
                base64_bytes = base64.b64encode(auth_bytes)
                headers['Authorization'] = f"Basic {base64_bytes.decode('ascii')}"
            
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()
            return True
            
        except Exception as e:
            print(f"Jira connection test failed: {e}")
            return False

class MSProjectConnector(BaseConnector):
    """Connector for Microsoft Project"""
    
    def fetch(self, entity_type, params=None):
        """Fetch data from Microsoft Project Online"""
        # Project Online uses SharePoint's REST API
        if 'token' not in self.config:
            raise ValueError("Authentication token missing for MS Project connector")
        
        # Get basic connection parameters
        base_url = self.config['connection'].get('url')
        if not base_url:
            raise ValueError("MS Project URL not specified in connector configuration")
        
        # Construct the API endpoint based on entity type
        if entity_type == 'projects':
            endpoint = f"{base_url}/_api/ProjectData/Projects"
        elif entity_type == 'tasks':
            endpoint = f"{base_url}/_api/ProjectData/Tasks"
        elif entity_type == 'resources':
            endpoint = f"{base_url}/_api/ProjectData/Resources"
        elif entity_type == 'assignments':
            endpoint = f"{base_url}/_api/ProjectData/Assignments"
        else:
            raise ValueError(f"Unsupported entity type for MS Project: {entity_type}")
        
        # Prepare authentication headers
        headers = {
            'Accept': 'application/json',
            'Authorization': f"Bearer {self.config['token']}"
        }
        
        # Prepare request parameters (filtering, etc.)
        request_params = {
            '$format': 'json'
        }
        
        if params:
            if 'filter' in params:
                request_params['$filter'] = params['filter']
            if 'select' in params:
                request_params['$select'] = ','.join(params['select'])
            if 'top' in params:
                request_params['$top'] = params['top']
            if 'skip' in params:
                request_params['$skip'] = params['skip']
        
        # Make the request
        try:
            response = requests.get(endpoint, headers=headers, params=request_params)
            response.raise_for_status()
            data = response.json()
            
            # Convert to DataFrame
            if 'value' in data:
                df = pd.DataFrame(data['value'])
            else:
                df = pd.DataFrame([data])
            
            return df
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data from MS Project: {e}")
            return pd.DataFrame()
    
    def test_connection(self):
        """Test MS Project connection"""
        try:
            # Try to fetch a simple endpoint
            base_url = self.config['connection'].get('url')
            endpoint = f"{base_url}/_api/ProjectData/$metadata"
            
            headers = {
                'Accept': 'application/json',
                'Authorization': f"Bearer {self.config['token']}"
            }
            
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()
            return True
            
        except Exception as e:
            print(f"MS Project connection test failed: {e}")
            return False

class ExcelConnector(BaseConnector):
    """Connector for Excel Online / SharePoint"""
    
    def fetch(self, entity_type, params=None):
        """Fetch data from Excel Online / SharePoint"""
        if 'token' not in self.config:
            raise ValueError("Authentication token missing for Excel connector")
        
        # Get basic connection parameters
        base_url = self.config['connection'].get('url')
        if not base_url:
            raise ValueError("Excel file URL not specified in connector configuration")
        
        # For Excel, entity_type represents the worksheet name
        worksheet = entity_type
        
        # Microsoft Graph API endpoints
        # Assuming the URL is in format: https://tenant.sharepoint.com/sites/site/Shared%20Documents/file.xlsx
        site_parts = base_url.split('/')
        tenant_url = '/'.join(site_parts[:3])  # e.g., https://tenant.sharepoint.com
        
        # Extract site path and file path
        site_path = '/'.join(site_parts[3:5])  # e.g., sites/site
        file_path = '/'.join(site_parts[5:])   # e.g., Shared%20Documents/file.xlsx
        
        # Construct the Graph API endpoint
        endpoint = f"https://graph.microsoft.com/v1.0/sites/{tenant_url}/{site_path}/drive/root:/{file_path}:/workbook/worksheets/{worksheet}/range(address='A1:XFD1048576')"
        
        # Prepare authentication headers
        headers = {
            'Authorization': f"Bearer {self.config['token']}",
            'Accept': 'application/json'
        }
        
        # Make the request
        try:
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()
            data = response.json()
            
            # Convert the response to a DataFrame
            if 'values' in data:
                values = data['values']
                if len(values) > 1:  # At least a header row and one data row
                    headers = values[0]
                    data_rows = values[1:]
                    df = pd.DataFrame(data_rows, columns=headers)
                    return df
                else:
                    return pd.DataFrame()
            else:
                return pd.DataFrame()
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data from Excel Online: {e}")
            return pd.DataFrame()
    
    def test_connection(self):
        """Test Excel Online connection"""
        try:
            # Try to fetch the workbook properties
            base_url = self.config['connection'].get('url')
            site_parts = base_url.split('/')
            tenant_url = '/'.join(site_parts[:3])
            site_path = '/'.join(site_parts[3:5])
            file_path = '/'.join(site_parts[5:])
            
            endpoint = f"https://graph.microsoft.com/v1.0/sites/{tenant_url}/{site_path}/drive/root:/{file_path}"
            
            headers = {
                'Authorization': f"Bearer {self.config['token']}",
                'Accept': 'application/json'
            }
            
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()
            return True
            
        except Exception as e:
            print(f"Excel Online connection test failed: {e}")
            return False

class RestApiConnector(BaseConnector):
    """Connector for generic REST APIs"""
    
    def fetch(self, entity_type, params=None):
        """Fetch data from a REST API"""
        # Get basic connection parameters
        base_url = self.config['connection'].get('url')
        if not base_url:
            raise ValueError("API URL not specified in connector configuration")
        
        # Construct the API endpoint
        endpoint = f"{base_url}/{entity_type}"
        
        # Prepare authentication headers
        headers = {
            'Accept': 'application/json'
        }
        
        # Add authentication if provided
        if 'api_key' in self.config:
            # Check if API key should be in header or query param
            if self.config['connection'].get('auth_type') == 'header':
                headers[self.config['connection'].get('auth_header', 'Authorization')] = self.config['api_key']
            else:
                # Will be added to query params
                if not params:
                    params = {}
                params[self.config['connection'].get('auth_param', 'api_key')] = self.config['api_key']
        elif 'token' in self.config:
            headers['Authorization'] = f"Bearer {self.config['token']}"
        elif 'username' in self.config and 'password' in self.config:
            auth = (self.config['username'], self.config['password'])
        else:
            auth = None
        
        # Make the request
        try:
            if 'username' in self.config and 'password' in self.config:
                response = requests.get(endpoint, headers=headers, params=params, auth=auth)
            else:
                response = requests.get(endpoint, headers=headers, params=params)
                
            response.raise_for_status()
            data = response.json()
            
            # Convert to DataFrame - structure depends on the API response format
            if isinstance(data, list):
                df = pd.DataFrame(data)
            elif isinstance(data, dict) and 'results' in data:
                df = pd.DataFrame(data['results'])
            elif isinstance(data, dict) and 'data' in data:
                df = pd.DataFrame(data['data'])
            elif isinstance(data, dict) and 'items' in data:
                df = pd.DataFrame(data['items'])
            else:
                # Fallback - just flatten the JSON
                df = pd.json_normalize(data)
            
            return df
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data from REST API: {e}")
            return pd.DataFrame()
    
    def test_connection(self):
        """Test REST API connection"""
        try:
            # Try a simple GET request to the base URL
            base_url = self.config['connection'].get('url')
            
            headers = {
                'Accept': 'application/json'
            }
            
            # Add authentication if provided
            if 'api_key' in self.config:
                if self.config['connection'].get('auth_type') == 'header':
                    headers[self.config['connection'].get('auth_header', 'Authorization')] = self.config['api_key']
                else:
                    params = {self.config['connection'].get('auth_param', 'api_key'): self.config['api_key']}
                    response = requests.get(base_url, headers=headers, params=params)
            elif 'token' in self.config:
                headers['Authorization'] = f"Bearer {self.config['token']}"
                response = requests.get(base_url, headers=headers)
            elif 'username' in self.config and 'password' in self.config:
                auth = (self.config['username'], self.config['password'])
                response = requests.get(base_url, headers=headers, auth=auth)
            else:
                response = requests.get(base_url, headers=headers)
            
            response.raise_for_status()
            return True
            
        except Exception as e:
            print(f"REST API connection test failed: {e}")
            return False

class DatabaseConnector(BaseConnector):
    """Connector for SQL databases"""
    
    def fetch(self, entity_type, params=None):
        """Fetch data from a SQL database"""
        try:
            import sqlalchemy
            from sqlalchemy import create_engine, text
        except ImportError:
            raise ImportError("SQLAlchemy is required for database connections. Install with pip install sqlalchemy")
        
        # Get connection parameters
        db_type = self.config['connection'].get('type', 'postgresql')
        host = self.config['connection'].get('host')
        port = self.config['connection'].get('port')
        database = self.config['connection'].get('database')
        
        if not all([host, database]):
            raise ValueError("Database connection parameters incomplete")
        
        # Construct connection string
        if db_type == 'postgresql':
            conn_str = f"postgresql://{self.config.get('username')}:{self.config.get('password')}@{host}:{port}/{database}"
        elif db_type == 'mysql':
            conn_str = f"mysql+pymysql://{self.config.get('username')}:{self.config.get('password')}@{host}:{port}/{database}"
        elif db_type == 'mssql':
            conn_str = f"mssql+pyodbc://{self.config.get('username')}:{self.config.get('password')}@{host}:{port}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        else:
            raise ValueError(f"Unsupported database type: {db_type}")
        
        # Create engine and connection
        try:
            engine = create_engine(conn_str)
            
            # entity_type can be a table name or a SQL query
            if entity_type.lower().startswith('select '):
                # It's a SQL query
                query = entity_type
            else:
                # It's a table name
                table_name = entity_type
                
                # Apply filters if provided
                where_clause = ""
                if params and 'filter' in params:
                    where_clause = f" WHERE {params['filter']}"
                
                limit_clause = ""
                if params and 'limit' in params:
                    limit_clause = f" LIMIT {params['limit']}"
                
                offset_clause = ""
                if params and 'offset' in params:
                    offset_clause = f" OFFSET {params['offset']}"
                
                query = f"SELECT * FROM {table_name}{where_clause}{limit_clause}{offset_clause}"
            
            # Execute query and return results
            with engine.connect() as connection:
                df = pd.read_sql(text(query), connection)
            
            return df
            
        except Exception as e:
            print(f"Error fetching data from database: {e}")
            return pd.DataFrame()
    
    def test_connection(self):
        """Test database connection"""
        try:
            import sqlalchemy
            from sqlalchemy import create_engine
        except ImportError:
            print("SQLAlchemy is required for database connections")
            return False
        
        try:
            # Get connection parameters
            db_type = self.config['connection'].get('type', 'postgresql')
            host = self.config['connection'].get('host')
            port = self.config['connection'].get('port')
            database = self.config['connection'].get('database')
            
            # Construct connection string
            if db_type == 'postgresql':
                conn_str = f"postgresql://{self.config.get('username')}:{self.config.get('password')}@{host}:{port}/{database}"
            elif db_type == 'mysql':
                conn_str = f"mysql+pymysql://{self.config.get('username')}:{self.config.get('password')}@{host}:{port}/{database}"
            elif db_type == 'mssql':
                conn_str = f"mssql+pyodbc://{self.config.get('username')}:{self.config.get('password')}@{host}:{port}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
            else:
                raise ValueError(f"Unsupported database type: {db_type}")
            
            # Try to connect
            engine = create_engine(conn_str)
            with engine.connect() as connection:
                # Just test the connection with a simple query
                connection.execute(sqlalchemy.text("SELECT 1"))
            
            return True
            
        except Exception as e:
            print(f"Database connection test failed: {e}")
            return False

# Main function to test the module
def test_integration():
    """Test function for the data integration module"""
    manager = DataIntegrationManager()
    print("Data Integration Manager initialized")
    
    # Example: Add a REST API source
    manager.add_source(
        source_id="demo_api",
        source_type="api",
        name="Demo API",
        connection_params={
            "url": "https://jsonplaceholder.typicode.com",
            "auth_type": "none"
        }
    )
    
    # List configured sources
    sources = manager.list_sources()
    print(f"Configured sources: {sources}")
    
    # Test fetching data
    try:
        posts_df = manager.fetch_data("demo_api", "posts")
        print(f"Fetched {len(posts_df)} posts from demo API")
        print(posts_df.head())
    except Exception as e:
        print(f"Error fetching data: {e}")
    
    # Clean up
    manager.remove_source("demo_api")
    print("Test source removed")

if __name__ == "__main__":
    test_integration()