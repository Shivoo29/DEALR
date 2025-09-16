"""
Configuration management for ZERF Automation System
"""

import os
import configparser
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any
import keyring

from .logger import get_logger
from .exceptions import ConfigurationError
from .validators import ConfigValidator

logger = get_logger(__name__)

class ConfigManager:
    """Enhanced configuration manager with secure credential storage"""
    
    DEFAULT_CONFIG = {
        'Paths': {
            'download_folder': 'downloads',
            'vbs_script': 'scripts/zerf_automation.vbs',
            'backup_folder': 'backup'
        },
        'SharePoint': {
            'site_url': '',
            'username': '',
            'folder_path': 'ERF Reporting_Data Analytics & Power BI'
        },
        'DateRange': {
            'start_date': '08/03/2025',
            'end_date': datetime.now().strftime("%m/%d/%Y")
        },
        'Schedule': {
            'run_time': '08:00',
            'check_interval': '30'
        },
        'Settings': {
            'auto_start': 'true',
            'cleanup_old_files': 'true',
            'max_retries': '3',
            'timeout_minutes': '10',
            'log_level': 'INFO'
        }
    }
    
    def __init__(self, config_file: str = None):
        """Initialize configuration manager"""
        if config_file is None:
            config_file = Path("config") / "zerf_config.ini"
        
        self.config_file = Path(config_file)
        self.config = configparser.RawConfigParser()
        self._ensure_config_dir()
        self.load_config()
    
    def _ensure_config_dir(self):
        """Ensure configuration directory exists"""
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
    
    def load_config(self):
        """Load configuration from file or create default"""
        if self.config_file.exists():
            try:
                self.config.read(self.config_file)
                logger.info(f"Configuration loaded from {self.config_file}")
                self._migrate_config_if_needed()
            except Exception as e:
                logger.error(f"Failed to load config: {e}")
                self._create_default_config()
        else:
            logger.info("No configuration file found, creating default")
            self._create_default_config()
    
    def _create_default_config(self):
        """Create default configuration file"""
        self.config = configparser.RawConfigParser()
        
        for section, settings in self.DEFAULT_CONFIG.items():
            self.config.add_section(section)
            for key, value in settings.items():
                self.config.set(section, key, str(value))
        
        self.save_config()
        logger.info(f"Default configuration created: {self.config_file}")
    
    def _migrate_config_if_needed(self):
        """Migrate old configuration format to new format"""
        migrated = False
        
        # Add missing sections
        for section in self.DEFAULT_CONFIG:
            if not self.config.has_section(section):
                self.config.add_section(section)
                migrated = True
        
        # Add missing keys with defaults
        for section, settings in self.DEFAULT_CONFIG.items():
            for key, default_value in settings.items():
                if not self.config.has_option(section, key):
                    self.config.set(section, key, str(default_value))
                    migrated = True
        
        if migrated:
            self.save_config()
            logger.info("Configuration migrated to latest format")
    
    def save_config(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                self.config.write(f)
            logger.debug(f"Configuration saved to {self.config_file}")
        except Exception as e:
            logger.error(f"Failed to save config: {e}")
            raise ConfigurationError(f"Failed to save configuration: {e}")
    
    def get(self, section: str, key: str, fallback: Any = None) -> str:
        """Get configuration value"""
        try:
            return self.config.get(section, key, fallback=fallback)
        except (configparser.NoSectionError, configparser.NoOptionError):
            if fallback is not None:
                return str(fallback)
            raise ConfigurationError(f"Configuration key not found: {section}.{key}")
    
    def set(self, section: str, key: str, value: Any):
        """Set configuration value"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))
    
    def get_int(self, section: str, key: str, fallback: int = None) -> int:
        """Get integer configuration value"""
        try:
            return self.config.getint(section, key, fallback=fallback)
        except (configparser.NoSectionError, configparser.NoOptionError):
            if fallback is not None:
                return fallback
            raise ConfigurationError(f"Configuration key not found: {section}.{key}")
    
    def get_bool(self, section: str, key: str, fallback: bool = None) -> bool:
        """Get boolean configuration value"""
        try:
            return self.config.getboolean(section, key, fallback=fallback)
        except (configparser.NoSectionError, configparser.NoOptionError):
            if fallback is not None:
                return fallback
            raise ConfigurationError(f"Configuration key not found: {section}.{key}")
    
    # Specific getters for common configuration values
    def get_download_folder(self) -> Path:
        """Get download folder path"""
        folder = self.get('Paths', 'download_folder', 'downloads')
        return Path(folder)
    
    def get_backup_folder(self) -> Path:
        """Get backup folder path"""
        folder = self.get('Paths', 'backup_folder', 'backup')
        return Path(folder)
    
    def get_vbs_script_path(self) -> Path:
        """Get VBS script path"""
        script = self.get('Paths', 'vbs_script', 'scripts/zerf_automation.vbs')
        return Path(script)
    
    def get_start_date(self) -> str:
        """Get start date"""
        return self.get('DateRange', 'start_date', '08/03/2025')
    
    def get_end_date(self) -> str:
        """Get end date"""
        return self.get('DateRange', 'end_date', datetime.now().strftime("%m/%d/%Y"))
    
    def set_start_date(self, date: str):
        """Set start date"""
        self.set('DateRange', 'start_date', date)
    
    def set_end_date(self, date: str):
        """Set end date"""
        self.set('DateRange', 'end_date', date)
    
    def get_run_time(self) -> str:
        """Get scheduled run time"""
        return self.get('Schedule', 'run_time', '08:00')
    
    def get_check_interval(self) -> int:
        """Get check interval in seconds"""
        return self.get_int('Schedule', 'check_interval', 30)
    
    def get_max_retries(self) -> int:
        """Get maximum retry attempts"""
        return self.get_int('Settings', 'max_retries', 3)
    
    def get_timeout_minutes(self) -> int:
        """Get timeout in minutes"""
        return self.get_int('Settings', 'timeout_minutes', 10)
    
    def get_log_level(self) -> str:
        """Get logging level"""
        return self.get('Settings', 'log_level', 'INFO')
    
    # SharePoint configuration with secure credential storage
    def get_sharepoint_url(self) -> str:
        """Get SharePoint site URL"""
        return self.get('SharePoint', 'site_url', '')
    
    def get_sharepoint_username(self) -> str:
        """Get SharePoint username"""
        return self.get('SharePoint', 'username', '')
    
    def get_sharepoint_password(self) -> str:
        """Get SharePoint password from secure storage"""
        username = self.get_sharepoint_username()
        if not username:
            return ''
        
        try:
            # Try to get from keyring first
            password = keyring.get_password("zerf_automation", username)
            if password:
                return password
            
            # Fallback to config file (less secure)
            return self.get('SharePoint', 'password', '')
        except Exception as e:
            logger.warning(f"Failed to retrieve password from keyring: {e}")
            return self.get('SharePoint', 'password', '')
    
    def set_sharepoint_credentials(self, site_url: str, username: str, password: str, use_keyring: bool = True):
        """Set SharePoint credentials"""
        self.set('SharePoint', 'site_url', site_url)
        self.set('SharePoint', 'username', username)
        
        if use_keyring and username:
            try:
                keyring.set_password("zerf_automation", username, password)
                # Clear password from config file for security
                self.set('SharePoint', 'password', '')
                logger.info("Password stored securely in system keyring")
            except Exception as e:
                logger.warning(f"Failed to store password in keyring: {e}")
                self.set('SharePoint', 'password', password)
                logger.warning("Password stored in config file (less secure)")
        else:
            self.set('SharePoint', 'password', password)
    
    def get_sharepoint_folder(self) -> str:
        """Get SharePoint folder path"""
        return self.get('SharePoint', 'folder_path', 'ERF Reporting_Data Analytics & Power BI')
    
    def validate_configuration(self, require_sharepoint: bool = False) -> tuple[bool, list]:
        """Validate current configuration"""
        config_dict = self.to_dict()
        return ConfigValidator.validate_config(config_dict, require_sharepoint)
    
    def to_dict(self) -> Dict[str, Dict[str, str]]:
        """Convert configuration to dictionary"""
        result = {}
        for section_name in self.config.sections():
            result[section_name] = dict(self.config.items(section_name))
        return result
    
    def update_from_dict(self, config_dict: Dict[str, Dict[str, str]]):
        """Update configuration from dictionary"""
        for section_name, section_data in config_dict.items():
            if not self.config.has_section(section_name):
                self.config.add_section(section_name)
            
            for key, value in section_data.items():
                self.config.set(section_name, key, str(value))
    
    def get_environment_override(self, key: str, default: str = None) -> str:
        """Get value from environment variable if available"""
        env_key = f"ZERF_{key.upper().replace('.', '_')}"
        return os.getenv(env_key, default)
    
    def load_environment_overrides(self):
        """Load configuration overrides from environment variables"""
        env_mappings = {
            'ZERF_SHAREPOINT_URL': ('SharePoint', 'site_url'),
            'ZERF_SHAREPOINT_USERNAME': ('SharePoint', 'username'),
            'ZERF_SHAREPOINT_PASSWORD': ('SharePoint', 'password'),
            'ZERF_DOWNLOAD_FOLDER': ('Paths', 'download_folder'),
            'ZERF_RUN_TIME': ('Schedule', 'run_time'),
            'ZERF_START_DATE': ('DateRange', 'start_date'),
            'ZERF_END_DATE': ('DateRange', 'end_date'),
            'ZERF_LOG_LEVEL': ('Settings', 'log_level'),
        }
        
        for env_var, (section, key) in env_mappings.items():
            value = os.getenv(env_var)
            if value:
                self.set(section, key, value)
                logger.info(f"Configuration override from environment: {section}.{key}")
    
    def export_config(self, export_path: Path, include_passwords: bool = False):
        """Export configuration to file"""
        export_config = configparser.RawConfigParser()
        
        for section_name in self.config.sections():
            export_config.add_section(section_name)
            for key, value in self.config.items(section_name):
                # Skip passwords unless explicitly requested
                if not include_passwords and key == 'password':
                    export_config.set(section_name, key, '***HIDDEN***')
                else:
                    export_config.set(section_name, key, value)
        
        with open(export_path, 'w') as f:
            export_config.write(f)
        
        logger.info(f"Configuration exported to {export_path}")

# Convenience function
def get_config_manager(config_file: str = None) -> ConfigManager:
    """Get a configuration manager instance"""
    return ConfigManager(config_file)