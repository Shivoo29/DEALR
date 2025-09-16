"""
Validation utilities for ZERF Automation System
"""

import re
import os
from datetime import datetime
from pathlib import Path
from typing import Optional, Union, List
from urllib.parse import urlparse

from .exceptions import ValidationError

class Validators:
    """Collection of validation utilities"""
    
    @staticmethod
    def validate_date_format(date_string: str, format_string: str = "%m/%d/%Y") -> bool:
        """Validate date format"""
        try:
            datetime.strptime(date_string, format_string)
            return True
        except ValueError:
            return False
    
    @staticmethod
    def validate_date_range(start_date: str, end_date: str, format_string: str = "%m/%d/%Y") -> bool:
        """Validate that start_date is before end_date"""
        try:
            start = datetime.strptime(start_date, format_string)
            end = datetime.strptime(end_date, format_string)
            return start <= end
        except ValueError:
            return False
    
    @staticmethod
    def validate_time_format(time_string: str) -> bool:
        """Validate time format (HH:MM)"""
        pattern = r'^([01]?[0-9]|2[0-3]):[0-5][0-9]$'
        return bool(re.match(pattern, time_string))
    
    @staticmethod
    def validate_url(url: str) -> bool:
        """Validate URL format"""
        try:
            result = urlparse(url)
            return all([result.scheme, result.netloc])
        except Exception:
            return False
    
    @staticmethod
    def validate_sharepoint_url(url: str) -> bool:
        """Validate SharePoint URL format"""
        if not Validators.validate_url(url):
            return False
        
        # Check for SharePoint-specific patterns
        sharepoint_patterns = [
            r'\.sharepoint\.com',
            r'/sites/',
            r'sharepoint'
        ]
        
        return any(re.search(pattern, url, re.IGNORECASE) for pattern in sharepoint_patterns)
    
    @staticmethod
    def validate_file_path(file_path: str, must_exist: bool = False) -> bool:
        """Validate file path"""
        try:
            path = Path(file_path)
            if must_exist:
                return path.exists() and path.is_file()
            else:
                # Check if parent directory exists or can be created
                return path.parent.exists() or path.parent.parent.exists()
        except Exception:
            return False
    
    @staticmethod
    def validate_directory_path(dir_path: str, must_exist: bool = False, create_if_missing: bool = False) -> bool:
        """Validate directory path"""
        try:
            path = Path(dir_path)
            
            if must_exist:
                return path.exists() and path.is_dir()
            
            if create_if_missing:
                path.mkdir(parents=True, exist_ok=True)
                return True
            
            return True  # Path is valid even if it doesn't exist
        except Exception:
            return False
    
    @staticmethod
    def validate_email(email: str) -> bool:
        """Validate email format"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))
    
    @staticmethod
    def validate_excel_file(file_path: str) -> bool:
        """Validate Excel file"""
        if not Validators.validate_file_path(file_path, must_exist=True):
            return False
        
        valid_extensions = ['.xlsx', '.xls']
        return Path(file_path).suffix.lower() in valid_extensions
    
    @staticmethod
    def validate_config_completeness(config_dict: dict, required_fields: List[str]) -> tuple[bool, List[str]]:
        """Validate that all required configuration fields are present and not empty"""
        missing_fields = []
        
        for field in required_fields:
            if '.' in field:
                # Handle nested fields like 'SharePoint.site_url'
                keys = field.split('.')
                value = config_dict
                
                try:
                    for key in keys:
                        value = value[key]
                except (KeyError, TypeError):
                    missing_fields.append(field)
                    continue
                
                if not value or str(value).strip() in ['', 'None', '<Your Name>', '<Your Password>', 'n/A']:
                    missing_fields.append(field)
            else:
                # Handle simple fields
                if field not in config_dict or not config_dict[field] or str(config_dict[field]).strip() in ['', 'None']:
                    missing_fields.append(field)
        
        return len(missing_fields) == 0, missing_fields

class ConfigValidator:
    """Specialized validator for ZERF configuration"""
    
    REQUIRED_FIELDS = [
        'DateRange.start_date',
        'DateRange.end_date',
        'Paths.download_folder',
        'Schedule.run_time'
    ]
    
    SHAREPOINT_FIELDS = [
        'SharePoint.site_url',
        'SharePoint.username',
        'SharePoint.password'
    ]
    
    @classmethod
    def validate_config(cls, config_dict: dict, require_sharepoint: bool = False) -> tuple[bool, List[str]]:
        """Validate complete configuration"""
        errors = []
        
        # Check required fields
        is_complete, missing = Validators.validate_config_completeness(config_dict, cls.REQUIRED_FIELDS)
        if not is_complete:
            errors.extend([f"Missing required field: {field}" for field in missing])
        
        # Check SharePoint fields if required
        if require_sharepoint:
            sp_complete, sp_missing = Validators.validate_config_completeness(config_dict, cls.SHAREPOINT_FIELDS)
            if not sp_complete:
                errors.extend([f"Missing SharePoint field: {field}" for field in sp_missing])
        
        # Validate date formats
        try:
            start_date = config_dict.get('DateRange', {}).get('start_date', '')
            end_date = config_dict.get('DateRange', {}).get('end_date', '')
            
            if start_date and not Validators.validate_date_format(start_date):
                errors.append("Invalid start_date format (expected MM/DD/YYYY)")
            
            if end_date and not Validators.validate_date_format(end_date):
                errors.append("Invalid end_date format (expected MM/DD/YYYY)")
            
            if start_date and end_date and not Validators.validate_date_range(start_date, end_date):
                errors.append("start_date must be before or equal to end_date")
        except Exception as e:
            errors.append(f"Date validation error: {e}")
        
        # Validate time format
        try:
            run_time = config_dict.get('Schedule', {}).get('run_time', '')
            if run_time and not Validators.validate_time_format(run_time):
                errors.append("Invalid run_time format (expected HH:MM)")
        except Exception as e:
            errors.append(f"Time validation error: {e}")
        
        # Validate SharePoint URL if provided
        try:
            site_url = config_dict.get('SharePoint', {}).get('site_url', '')
            if site_url and site_url not in ['', 'n/A'] and not Validators.validate_sharepoint_url(site_url):
                errors.append("Invalid SharePoint site URL")
        except Exception as e:
            errors.append(f"SharePoint URL validation error: {e}")
        
        # Validate paths
        try:
            download_folder = config_dict.get('Paths', {}).get('download_folder', '')
            if download_folder and not Validators.validate_directory_path(download_folder, create_if_missing=True):
                errors.append("Invalid download folder path")
        except Exception as e:
            errors.append(f"Path validation error: {e}")
        
        return len(errors) == 0, errors
    
    @classmethod
    def validate_and_raise(cls, config_dict: dict, require_sharepoint: bool = False):
        """Validate configuration and raise ValidationError if invalid"""
        is_valid, errors = cls.validate_config(config_dict, require_sharepoint)
        if not is_valid:
            raise ValidationError(f"Configuration validation failed: {'; '.join(errors)}")
        
        return True