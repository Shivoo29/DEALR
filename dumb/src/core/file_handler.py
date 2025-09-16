"""
File handling and management for ZERF Automation System
"""

import os
import time
import shutil
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, List, Dict
import hashlib

from ..utils.logger import get_logger
from ..utils.exceptions import FileNotFoundError, TemporaryFileError
from ..utils.validators import Validators

logger = get_logger(__name__)

class FileHandler:
    """Enhanced file handler with robust file detection and management"""
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.download_folder = config_manager.get_download_folder()
        self.backup_folder = config_manager.get_backup_folder()
        
        # Ensure directories exist
        self.download_folder.mkdir(parents=True, exist_ok=True)
        self.backup_folder.mkdir(parents=True, exist_ok=True)
        
        # File detection settings
        self.max_wait_minutes = 5
        self.check_interval_seconds = 10
        
    def find_latest_download(self, max_wait_minutes: int = None, expected_pattern: str = None) -> Optional[Path]:
        """Find the most recently downloaded Excel file with improved detection"""
        try:
            max_wait = max_wait_minutes or self.max_wait_minutes
            logger.info(f"Searching for downloaded files (max wait: {max_wait} minutes)")
            
            # Search locations in order of preference
            search_locations = [
                self.download_folder,
                Path.home() / "Downloads",
                Path.home() / "Desktop"
            ]
            
            start_time = time.time()
            max_wait_seconds = max_wait * 60
            
            while time.time() - start_time < max_wait_seconds:
                candidates = []
                
                for location in search_locations:
                    if not location.exists():
                        continue
                    
                    logger.debug(f"Searching in: {location}")
                    location_candidates = self._find_candidates_in_location(
                        location, expected_pattern
                    )
                    candidates.extend(location_candidates)
                
                if candidates:
                    # Sort by creation time and select the most recent
                    best_candidate = self._select_best_candidate(candidates)
                    
                    if best_candidate and self._validate_file_accessibility(best_candidate):
                        logger.info(f"✅ Found suitable file: {best_candidate}")
                        return best_candidate
                
                # Wait before next check
                elapsed_minutes = (time.time() - start_time) / 60
                logger.info(f"Waiting for files... ({elapsed_minutes:.1f} min elapsed)")
                time.sleep(self.check_interval_seconds)
            
            logger.warning(f"⚠️ No suitable files found after {max_wait} minutes")
            return None
            
        except Exception as e:
            logger.error(f"File detection error: {e}")
            return None
    
    def _find_candidates_in_location(self, location: Path, pattern: str = None) -> List[Dict]:
        """Find candidate files in a specific location"""
        candidates = []
        
        try:
            # Look for Excel files
            excel_patterns = ["*.xlsx", "*.xls"]
            current_time = time.time()
            
            for pattern_str in excel_patterns:
                for file_path in location.glob(pattern_str):
                    try:
                        # Skip temporary files
                        if self._is_temporary_file(file_path):
                            continue
                        
                        # Check if file is recent (within last hour)
                        file_age_seconds = current_time - os.path.getctime(file_path)
                        if file_age_seconds > 3600:  # 1 hour
                            continue
                        
                        # Check if pattern matches (if specified)
                        if pattern and pattern.lower() not in file_path.name.lower():
                            continue
                        
                        candidates.append({
                            'path': file_path,
                            'age_seconds': file_age_seconds,
                            'age_minutes': file_age_seconds / 60,
                            'size': file_path.stat().st_size,
                            'location': location
                        })
                        
                        logger.debug(f"Candidate: {file_path.name} (age: {file_age_seconds/60:.1f} min)")
                        
                    except (OSError, PermissionError) as e:
                        logger.debug(f"Skipping inaccessible file {file_path}: {e}")
                        continue
            
        except Exception as e:
            logger.warning(f"Error searching in {location}: {e}")
        
        return candidates
    
    def _is_temporary_file(self, file_path: Path) -> bool:
        """Check if file is a temporary file that should be ignored"""
        temp_indicators = [
            '~$',  # Excel temporary files
            '.tmp',  # General temporary files
            '.temp',  # Temporary files
            '.crdownload',  # Chrome download in progress
            '.partial'  # Partial download
        ]
        
        filename = file_path.name.lower()
        return any(indicator in filename for indicator in temp_indicators)
    
    def _select_best_candidate(self, candidates: List[Dict]) -> Optional[Path]:
        """Select the best candidate file from the list"""
        if not candidates:
            return None
        
        # Sort by creation time (most recent first)
        candidates.sort(key=lambda x: x['age_seconds'])
        
        # Prefer files in the primary download folder
        primary_folder_candidates = [
            c for c in candidates 
            if c['location'] == self.download_folder
        ]
        
        if primary_folder_candidates:
            return primary_folder_candidates[0]['path']
        
        # Fallback to any recent file
        return candidates[0]['path']
    
    def _validate_file_accessibility(self, file_path: Path) -> bool:
        """Validate that file is accessible and not locked"""
        try:
            # Check if file exists
            if not file_path.exists():
                return False
            
            # Check if file is not empty
            if file_path.stat().st_size == 0:
                logger.debug(f"File is empty: {file_path}")
                return False
            
            # Try to open file to ensure it's not locked
            with open(file_path, 'rb') as f:
                f.read(1)  # Try to read at least one byte
            
            logger.debug(f"File accessibility validated: {file_path}")
            return True
            
        except (PermissionError, IOError) as e:
            logger.debug(f"File not accessible: {file_path} - {e}")
            return False
        except Exception as e:
            logger.warning(f"File validation error: {file_path} - {e}")
            return False
    
    def backup_file(self, file_path: Path, backup_type: str = "auto") -> Optional[Path]:
        """Create backup of a file with timestamp and type"""
        try:
            if not file_path.exists():
                logger.error(f"Cannot backup non-existent file: {file_path}")
                return None
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_stem = file_path.stem
            file_suffix = file_path.suffix
            
            # Create backup filename
            backup_filename = f"{file_stem}_{backup_type}_{timestamp}{file_suffix}"
            backup_path = self.backup_folder / backup_filename
            
            # Copy file to backup location
            shutil.copy2(file_path, backup_path)
            
            # Verify backup was created
            if backup_path.exists():
                logger.info(f"✅ Backup created: {backup_path.name}")
                return backup_path
            else:
                logger.error(f"❌ Backup creation failed: {backup_path}")
                return None
                
        except Exception as e:
            logger.error(f"Backup operation failed: {e}")
            return None
    
    def cleanup_old_files(self, directory: Path, days_to_keep: int = 30) -> int:
        """Clean up old files in directory"""
        try:
            if not directory.exists():
                logger.warning(f"Directory does not exist: {directory}")
                return 0
            
            cutoff_time = datetime.now() - timedelta(days=days_to_keep)
            cutoff_timestamp = cutoff_time.timestamp()
            
            cleaned_count = 0
            
            for file_path in directory.iterdir():
                if file_path.is_file():
                    try:
                        file_mtime = file_path.stat().st_mtime
                        
                        if file_mtime < cutoff_timestamp:
                            file_age_days = (datetime.now().timestamp() - file_mtime) / (24 * 3600)
                            file_path.unlink()
                            cleaned_count += 1
                            logger.debug(f"Deleted old file: {file_path.name} (age: {file_age_days:.1f} days)")
                            
                    except (OSError, PermissionError) as e:
                        logger.warning(f"Failed to delete {file_path}: {e}")
                        continue
            
            if cleaned_count > 0:
                logger.info(f"✅ Cleaned up {cleaned_count} old files from {directory}")
            else:
                logger.debug(f"No old files to clean up in {directory}")
            
            return cleaned_count
            
        except Exception as e:
            logger.error(f"Cleanup operation failed: {e}")
            return 0
    
    def get_file_info(self, file_path: Path) -> Dict:
        """Get comprehensive file information"""
        try:
            if not file_path.exists():
                return {'exists': False}
            
            stat_info = file_path.stat()
            
            return {
                'exists': True,
                'path': str(file_path),
                'name': file_path.name,
                'size_bytes': stat_info.st_size,
                'size_mb': stat_info.st_size / (1024 * 1024),
                'created': datetime.fromtimestamp(stat_info.st_ctime),
                'modified': datetime.fromtimestamp(stat_info.st_mtime),
                'is_excel': Validators.validate_excel_file(str(file_path)),
                'is_accessible': self._validate_file_accessibility(file_path),
                'md5_hash': self._calculate_file_hash(file_path)
            }
            
        except Exception as e:
            logger.error(f"Failed to get file info: {e}")
            return {'exists': False, 'error': str(e)}
    
    def _calculate_file_hash(self, file_path: Path) -> Optional[str]:
        """Calculate MD5 hash of file for integrity checking"""
        try:
            hash_md5 = hashlib.md5()
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
        except Exception as e:
            logger.debug(f"Failed to calculate hash for {file_path}: {e}")
            return None
    
    def move_file(self, source: Path, destination: Path) -> bool:
        """Move file with error handling and validation"""
        try:
            if not source.exists():
                logger.error(f"Source file does not exist: {source}")
                return False
            
            # Ensure destination directory exists
            destination.parent.mkdir(parents=True, exist_ok=True)
            
            # Move file
            shutil.move(str(source), str(destination))
            
            # Verify move was successful
            if destination.exists() and not source.exists():
                logger.info(f"✅ File moved: {source.name} → {destination}")
                return True
            else:
                logger.error(f"❌ File move verification failed")
                return False
                
        except Exception as e:
            logger.error(f"File move failed: {e}")
            return False
    
    def copy_file(self, source: Path, destination: Path) -> bool:
        """Copy file with error handling and validation"""
        try:
            if not source.exists():
                logger.error(f"Source file does not exist: {source}")
                return False
            
            # Ensure destination directory exists
            destination.parent.mkdir(parents=True, exist_ok=True)
            
            # Copy file
            shutil.copy2(str(source), str(destination))
            
            # Verify copy was successful
            if destination.exists():
                # Verify file sizes match
                if source.stat().st_size == destination.stat().st_size:
                    logger.info(f"✅ File copied: {source.name} → {destination}")
                    return True
                else:
                    logger.error(f"❌ File copy size mismatch")
                    return False
            else:
                logger.error(f"❌ File copy failed - destination not created")
                return False
                
        except Exception as e:
            logger.error(f"File copy failed: {e}")
            return False
    
    def wait_for_file_stability(self, file_path: Path, max_wait_seconds: int = 30) -> bool:
        """Wait for file to become stable (not being written to)"""
        try:
            if not file_path.exists():
                return False
            
            last_size = -1
            last_mtime = -1
            stable_checks = 0
            required_stable_checks = 3  # File must be stable for 3 consecutive checks
            
            start_time = time.time()
            
            while time.time() - start_time < max_wait_seconds:
                try:
                    stat_info = file_path.stat()
                    current_size = stat_info.st_size
                    current_mtime = stat_info.st_mtime
                    
                    if current_size == last_size and current_mtime == last_mtime:
                        stable_checks += 1
                        if stable_checks >= required_stable_checks:
                            logger.debug(f"File stable: {file_path.name}")
                            return True
                    else:
                        stable_checks = 0
                    
                    last_size = current_size
                    last_mtime = current_mtime
                    
                except (OSError, PermissionError):
                    # File might be locked, continue waiting
                    stable_checks = 0
                
                time.sleep(2)  # Check every 2 seconds
            
            logger.warning(f"File stability timeout: {file_path.name}")
            return False
            
        except Exception as e:
            logger.error(f"File stability check failed: {e}")
            return False
    
    def get_directory_summary(self, directory: Path) -> Dict:
        """Get summary information about a directory"""
        try:
            if not directory.exists():
                return {'exists': False}
            
            total_files = 0
            total_size = 0
            file_types = {}
            oldest_file = None
            newest_file = None
            
            for file_path in directory.iterdir():
                if file_path.is_file():
                    total_files += 1
                    
                    try:
                        stat_info = file_path.stat()
                        file_size = stat_info.st_size
                        file_mtime = stat_info.st_mtime
                        
                        total_size += file_size
                        
                        # Track file types
                        file_ext = file_path.suffix.lower()
                        file_types[file_ext] = file_types.get(file_ext, 0) + 1
                        
                        # Track oldest and newest files
                        if oldest_file is None or file_mtime < oldest_file['mtime']:
                            oldest_file = {'path': file_path, 'mtime': file_mtime}
                        
                        if newest_file is None or file_mtime > newest_file['mtime']:
                            newest_file = {'path': file_path, 'mtime': file_mtime}
                            
                    except (OSError, PermissionError):
                        continue
            
            return {
                'exists': True,
                'path': str(directory),
                'total_files': total_files,
                'total_size_bytes': total_size,
                'total_size_mb': total_size / (1024 * 1024),
                'file_types': file_types,
                'oldest_file': oldest_file['path'].name if oldest_file else None,
                'newest_file': newest_file['path'].name if newest_file else None,
                'oldest_file_date': datetime.fromtimestamp(oldest_file['mtime']) if oldest_file else None,
                'newest_file_date': datetime.fromtimestamp(newest_file['mtime']) if newest_file else None
            }
            
        except Exception as e:
            logger.error(f"Failed to get directory summary: {e}")
            return {'exists': False, 'error': str(e)}