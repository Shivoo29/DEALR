"""
Data processing engine for ZERF Excel files
"""

import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List, Tuple

from ..utils.logger import get_logger, create_progress_logger
from ..utils.exceptions import DataProcessingError, ValidationError
from ..utils.validators import Validators

logger = get_logger(__name__)

class DataProcessor:
    """Advanced data processor for ZERF Excel files"""
    
    # Data cleaning rules configuration
    CLEANING_RULES = {
        'duplicate_removal': {
            'enabled': True,
            'columns': ['Unique_ID'],
            'keep': 'first'
        },
        'status_filtering': {
            'enabled': True,
            'column': 'Engineering Request Form Status',
            'exclude_values': ['Draft', 'Presubmit', 'Submit']
        },
        'blank_status_removal': {
            'enabled': True,
            'column': 'ERF Sched Line Status'
        },
        'commodity_filtering': {
            'enabled': True,
            'column': 'Commodity Type',
            'exclude_pattern': 'Indirect'
        },
        'plant_filtering': {
            'enabled': True,
            'column': 'Ship-To-Plant',
            'allowed_values': [6100, 6200, 6300, '6100', '6200', '6300']
        },
        'pgr_filtering': {
            'enabled': True,
            'column': 'PGr',
            'exclude_values': ['W91', 'Z05']
        }
    }
    
    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.processing_stats = {}
    
    def process_file(self, input_file: Path, custom_rules: Dict = None) -> Optional[Path]:
        """Process Excel file with data cleaning rules"""
        try:
            logger.info(f"Starting data processing: {input_file}")
            
            # Validate input file
            if not self._validate_input_file(input_file):
                raise DataProcessingError(f"Invalid input file: {input_file}")
            
            # Generate output filename
            output_file = self._generate_output_filename(input_file)
            
            # Load and process data
            progress = create_progress_logger(logger, 8, "Data Processing")
            
            progress.step("Loading Excel file...")
            excel_data = self._load_excel_file(input_file)
            
            progress.step("Analyzing data structure...")
            self._analyze_data_structure(excel_data)
            
            progress.step("Creating Unique_ID column...")
            processed_data = self._create_unique_id(excel_data)
            
            progress.step("Removing duplicates...")
            processed_data = self._remove_duplicates(processed_data)
            
            progress.step("Filtering by status...")
            processed_data = self._filter_by_status(processed_data)
            
            progress.step("Applying business rules...")
            processed_data = self._apply_business_rules(processed_data, custom_rules)
            
            progress.step("Validating processed data...")
            self._validate_processed_data(processed_data)
            
            progress.step("Saving cleaned file...")
            self._save_processed_data(processed_data, output_file)
            
            # Log processing statistics
            self._log_processing_stats()
            
            logger.info(f"✅ Data processing completed: {output_file}")
            return output_file
            
        except Exception as e:
            logger.error(f"Data processing failed: {e}", exc_info=True)
            raise DataProcessingError(f"Data processing failed: {e}")
    
    def _validate_input_file(self, file_path: Path) -> bool:
        """Validate input Excel file"""
        try:
            if not file_path.exists():
                logger.error(f"File does not exist: {file_path}")
                return False
            
            if not Validators.validate_excel_file(str(file_path)):
                logger.error(f"Invalid Excel file format: {file_path}")
                return False
            
            # Check file size (warn if very large)
            file_size = file_path.stat().st_size
            if file_size > 50 * 1024 * 1024:  # 50MB
                logger.warning(f"Large file detected ({file_size / 1024 / 1024:.1f}MB): {file_path}")
            
            return True
            
        except Exception as e:
            logger.error(f"File validation error: {e}")
            return False
    
    def _generate_output_filename(self, input_file: Path) -> Path:
        """Generate output filename for processed file"""
        download_folder = self.config_manager.get_download_folder()
        today = datetime.now().strftime("%m-%d-%Y")
        output_filename = f"zerf_{today}_cleaned.xlsx"
        return download_folder / output_filename
    
    def _load_excel_file(self, file_path: Path) -> Dict[str, pd.DataFrame]:
        """Load Excel file and return dictionary of DataFrames"""
        try:
            excel_file = pd.ExcelFile(file_path)
            data = {}
            
            for sheet_name in excel_file.sheet_names:
                logger.debug(f"Loading sheet: {sheet_name}")
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                data[sheet_name] = df
                logger.debug(f"Sheet {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
            
            return data
            
        except Exception as e:
            logger.error(f"Failed to load Excel file: {e}")
            raise DataProcessingError(f"Failed to load Excel file: {e}")
    
    def _analyze_data_structure(self, excel_data: Dict[str, pd.DataFrame]):
        """Analyze and log data structure"""
        try:
            total_rows = sum(len(df) for df in excel_data.values())
            total_sheets = len(excel_data)
            
            logger.info(f"Data structure: {total_sheets} sheets, {total_rows} total rows")
            
            # Check for required columns
            required_columns = ['ERF Number', 'Item']
            for sheet_name, df in excel_data.items():
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    logger.warning(f"Sheet {sheet_name} missing required columns: {missing_columns}")
                
                # Log column names for debugging
                logger.debug(f"Sheet {sheet_name} columns: {list(df.columns)}")
            
            self.processing_stats['input'] = {
                'sheets': total_sheets,
                'total_rows': total_rows,
                'structure': {name: {'rows': len(df), 'columns': len(df.columns)} 
                            for name, df in excel_data.items()}
            }
            
        except Exception as e:
            logger.error(f"Data structure analysis failed: {e}")
    
    def _create_unique_id(self, excel_data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
        """Create Unique_ID column by combining ERF Number and Item"""
        try:
            processed_data = {}
            
            for sheet_name, df in excel_data.items():
                df_copy = df.copy()
                
                if 'ERF Number' in df_copy.columns and 'Item' in df_copy.columns:
                    # Find ERF Number column position
                    erf_position = df_copy.columns.get_loc('ERF Number')
                    
                    # Create Unique_ID values
                    unique_id_values = (
                        df_copy['ERF Number'].astype(str) + '-' + 
                        df_copy['Item'].astype(str)
                    )
                    
                    # Insert Unique_ID column after ERF Number
                    df_copy.insert(erf_position + 1, 'Unique_ID', unique_id_values)
                    
                    logger.info(f"Sheet {sheet_name}: Created Unique_ID column")
                else:
                    logger.warning(f"Sheet {sheet_name}: Cannot create Unique_ID - missing required columns")
                
                processed_data[sheet_name] = df_copy
            
            return processed_data
            
        except Exception as e:
            logger.error(f"Unique_ID creation failed: {e}")
            raise DataProcessingError(f"Unique_ID creation failed: {e}")
    
    def _remove_duplicates(self, excel_data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
        """Remove duplicate records based on Unique_ID"""
        try:
            processed_data = {}
            
            for sheet_name, df in excel_data.items():
                initial_rows = len(df)
                
                if 'Unique_ID' in df.columns:
                    df_deduplicated = df.drop_duplicates(subset=['Unique_ID'], keep='first')
                    duplicates_removed = initial_rows - len(df_deduplicated)
                    
                    if duplicates_removed > 0:
                        logger.info(f"Sheet {sheet_name}: Removed {duplicates_removed} duplicate records")
                    
                    processed_data[sheet_name] = df_deduplicated
                else:
                    logger.warning(f"Sheet {sheet_name}: No Unique_ID column for deduplication")
                    processed_data[sheet_name] = df
            
            return processed_data
            
        except Exception as e:
            logger.error(f"Duplicate removal failed: {e}")
            raise DataProcessingError(f"Duplicate removal failed: {e}")
    
    def _filter_by_status(self, excel_data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
        """Filter out specific status values"""
        try:
            processed_data = {}
            
            for sheet_name, df in excel_data.items():
                df_filtered = df.copy()
                initial_rows = len(df_filtered)
                
                # Remove specific ERF statuses
                if 'Engineering Request Form Status' in df_filtered.columns:
                    status_to_remove = ['Draft', 'Presubmit', 'Submit']
                    df_filtered = df_filtered[
                        ~df_filtered['Engineering Request Form Status'].isin(status_to_remove)
                    ]
                    status_removed = initial_rows - len(df_filtered)
                    if status_removed > 0:
                        logger.info(f"Sheet {sheet_name}: Removed {status_removed} rows with excluded statuses")
                
                # Remove blank ERF Sched Line Status
                if 'ERF Sched Line Status' in df_filtered.columns:
                    current_rows = len(df_filtered)
                    df_filtered = df_filtered.dropna(subset=['ERF Sched Line Status'])
                    df_filtered = df_filtered[df_filtered['ERF Sched Line Status'] != '']
                    blank_removed = current_rows - len(df_filtered)
                    if blank_removed > 0:
                        logger.info(f"Sheet {sheet_name}: Removed {blank_removed} rows with blank ERF Sched Line Status")
                
                processed_data[sheet_name] = df_filtered
            
            return processed_data
            
        except Exception as e:
            logger.error(f"Status filtering failed: {e}")
            raise DataProcessingError(f"Status filtering failed: {e}")
    
    def _apply_business_rules(self, excel_data: Dict[str, pd.DataFrame], custom_rules: Dict = None) -> Dict[str, pd.DataFrame]:
        """Apply business-specific filtering rules"""
        try:
            processed_data = {}
            
            for sheet_name, df in excel_data.items():
                df_filtered = df.copy()
                initial_rows = len(df_filtered)
                
                # Remove Indirect commodity types
                if 'Commodity Type' in df_filtered.columns:
                    current_rows = len(df_filtered)
                    df_filtered = df_filtered[
                        ~df_filtered['Commodity Type'].astype(str).str.contains('Indirect', case=False, na=False)
                    ]
                    indirect_removed = current_rows - len(df_filtered)
                    if indirect_removed > 0:
                        logger.info(f"Sheet {sheet_name}: Removed {indirect_removed} rows with Indirect commodity type")
                
                # Keep only specific Ship-To-Plant values
                if 'Ship-To-Plant' in df_filtered.columns:
                    current_rows = len(df_filtered)
                    allowed_plants = [6100, 6200, 6300, '6100', '6200', '6300']
                    df_filtered = df_filtered[df_filtered['Ship-To-Plant'].isin(allowed_plants)]
                    plant_removed = current_rows - len(df_filtered)
                    if plant_removed > 0:
                        logger.info(f"Sheet {sheet_name}: Removed {plant_removed} rows with non-allowed plants")
                
                # Remove specific PGr values
                if 'PGr' in df_filtered.columns:
                    current_rows = len(df_filtered)
                    pgr_to_remove = ['W91', 'Z05']
                    df_filtered = df_filtered[~df_filtered['PGr'].isin(pgr_to_remove)]
                    pgr_removed = current_rows - len(df_filtered)
                    if pgr_removed > 0:
                        logger.info(f"Sheet {sheet_name}: Removed {pgr_removed} rows with PGr values W91 or Z05")
                
                # Apply custom rules if provided
                if custom_rules:
                    df_filtered = self._apply_custom_rules(df_filtered, custom_rules, sheet_name)
                
                final_rows = len(df_filtered)
                total_removed = initial_rows - final_rows
                
                logger.info(f"Sheet {sheet_name}: {initial_rows} → {final_rows} rows (removed {total_removed})")
                processed_data[sheet_name] = df_filtered
            
            return processed_data
            
        except Exception as e:
            logger.error(f"Business rules application failed: {e}")
            raise DataProcessingError(f"Business rules application failed: {e}")
    
    def _apply_custom_rules(self, df: pd.DataFrame, custom_rules: Dict, sheet_name: str) -> pd.DataFrame:
        """Apply custom filtering rules"""
        try:
            df_filtered = df.copy()
            
            for rule_name, rule_config in custom_rules.items():
                if not rule_config.get('enabled', True):
                    continue
                
                initial_rows = len(df_filtered)
                
                if rule_name == 'column_filter':
                    column = rule_config.get('column')
                    values = rule_config.get('exclude_values', [])
                    if column in df_filtered.columns and values:
                        df_filtered = df_filtered[~df_filtered[column].isin(values)]
                
                elif rule_name == 'regex_filter':
                    column = rule_config.get('column')
                    pattern = rule_config.get('pattern')
                    if column in df_filtered.columns and pattern:
                        df_filtered = df_filtered[
                            ~df_filtered[column].astype(str).str.contains(pattern, case=False, na=False)
                        ]
                
                elif rule_name == 'range_filter':
                    column = rule_config.get('column')
                    min_val = rule_config.get('min_value')
                    max_val = rule_config.get('max_value')
                    if column in df_filtered.columns:
                        if min_val is not None:
                            df_filtered = df_filtered[df_filtered[column] >= min_val]
                        if max_val is not None:
                            df_filtered = df_filtered[df_filtered[column] <= max_val]
                
                final_rows = len(df_filtered)
                if final_rows != initial_rows:
                    logger.info(f"Sheet {sheet_name}: Custom rule '{rule_name}' removed {initial_rows - final_rows} rows")
            
            return df_filtered
            
        except Exception as e:
            logger.error(f"Custom rules application failed: {e}")
            return df
    
    def _validate_processed_data(self, excel_data: Dict[str, pd.DataFrame]):
        """Validate processed data quality"""
        try:
            validation_results = []
            
            for sheet_name, df in excel_data.items():
                # Check for empty dataframes
                if len(df) == 0:
                    validation_results.append(f"Sheet {sheet_name} is empty after processing")
                    continue
                
                # Check for required columns
                required_columns = ['ERF Number', 'Item']
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    validation_results.append(f"Sheet {sheet_name} missing columns: {missing_columns}")
                
                # Check for data quality issues
                if 'Unique_ID' in df.columns:
                    duplicate_count = df['Unique_ID'].duplicated().sum()
                    if duplicate_count > 0:
                        validation_results.append(f"Sheet {sheet_name} has {duplicate_count} duplicate Unique_IDs")
                
                # Check for excessive null values
                null_percentages = (df.isnull().sum() / len(df)) * 100
                high_null_columns = null_percentages[null_percentages > 50].index.tolist()
                if high_null_columns:
                    validation_results.append(f"Sheet {sheet_name} has high null rates in: {high_null_columns}")
            
            if validation_results:
                logger.warning("Data validation warnings:")
                for result in validation_results:
                    logger.warning(f"  - {result}")
            else:
                logger.info("✅ Data validation passed")
            
            # Store validation results
            self.processing_stats['validation'] = validation_results
            
        except Exception as e:
            logger.error(f"Data validation failed: {e}")
    
    def _save_processed_data(self, excel_data: Dict[str, pd.DataFrame], output_file: Path):
        """Save processed data to Excel file"""
        try:
            # Ensure output directory exists
            output_file.parent.mkdir(parents=True, exist_ok=True)
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in excel_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    logger.debug(f"Saved sheet {sheet_name}: {len(df)} rows")
            
            # Verify file was created
            if not output_file.exists():
                raise DataProcessingError("Output file was not created")
            
            file_size = output_file.stat().st_size
            logger.info(f"Processed file saved: {output_file} ({file_size / 1024:.1f} KB)")
            
        except Exception as e:
            logger.error(f"Failed to save processed data: {e}")
            raise DataProcessingError(f"Failed to save processed data: {e}")
    
    def _log_processing_stats(self):
        """Log comprehensive processing statistics"""
        try:
            if not self.processing_stats:
                return
            
            logger.info("=" * 50)
            logger.info("DATA PROCESSING STATISTICS")
            logger.info("=" * 50)
            
            # Input statistics
            if 'input' in self.processing_stats:
                input_stats = self.processing_stats['input']
                logger.info(f"Input: {input_stats['sheets']} sheets, {input_stats['total_rows']} total rows")
                
                for sheet_name, stats in input_stats['structure'].items():
                    logger.info(f"  - {sheet_name}: {stats['rows']} rows, {stats['columns']} columns")
            
            # Processing results would be logged here
            logger.info("=" * 50)
            
        except Exception as e:
            logger.error(f"Failed to log processing stats: {e}")
    
    def get_processing_summary(self) -> Dict:
        """Get summary of last processing operation"""
        return {
            'stats': self.processing_stats,
            'rules_applied': list(self.CLEANING_RULES.keys()),
            'timestamp': datetime.now().isoformat()
        }
    
    def process_file_with_preview(self, input_file: Path, max_preview_rows: int = 100) -> Dict:
        """Process file and return preview of changes"""
        try:
            logger.info(f"Processing file with preview: {input_file}")
            
            # Load original data
            original_data = self._load_excel_file(input_file)
            
            # Process data
            processed_data = {}
            changes_summary = {}
            
            for sheet_name, df in original_data.items():
                original_rows = len(df)
                
                # Apply processing steps
                df_processed = df.copy()
                
                # Create Unique_ID
                if 'ERF Number' in df_processed.columns and 'Item' in df_processed.columns:
                    erf_position = df_processed.columns.get_loc('ERF Number')
                    unique_id_values = (
                        df_processed['ERF Number'].astype(str) + '-' + 
                        df_processed['Item'].astype(str)
                    )
                    df_processed.insert(erf_position + 1, 'Unique_ID', unique_id_values)
                
                # Apply all cleaning rules
                df_processed = self._remove_duplicates({sheet_name: df_processed})[sheet_name]
                df_processed = self._filter_by_status({sheet_name: df_processed})[sheet_name]
                df_processed = self._apply_business_rules({sheet_name: df_processed})[sheet_name]
                
                processed_rows = len(df_processed)
                rows_removed = original_rows - processed_rows
                
                changes_summary[sheet_name] = {
                    'original_rows': original_rows,
                    'processed_rows': processed_rows,
                    'rows_removed': rows_removed,
                    'removal_percentage': (rows_removed / original_rows * 100) if original_rows > 0 else 0
                }
                
                # Create preview
                preview_data = df_processed.head(max_preview_rows) if len(df_processed) > 0 else df_processed
                processed_data[sheet_name] = preview_data
            
            return {
                'preview_data': processed_data,
                'changes_summary': changes_summary,
                'total_original_rows': sum(s['original_rows'] for s in changes_summary.values()),
                'total_processed_rows': sum(s['processed_rows'] for s in changes_summary.values()),
                'preview_limited': any(len(df) > max_preview_rows for df in original_data.values())
            }
            
        except Exception as e:
            logger.error(f"Preview processing failed: {e}")
            raise DataProcessingError(f"Preview processing failed: {e}")
    
    def validate_file_format(self, file_path: Path) -> Dict[str, Any]:
        """Validate file format and structure"""
        try:
            validation_result = {
                'valid': True,
                'errors': [],
                'warnings': [],
                'file_info': {},
                'column_analysis': {}
            }
            
            # Basic file validation
            if not file_path.exists():
                validation_result['valid'] = False
                validation_result['errors'].append("File does not exist")
                return validation_result
            
            if not Validators.validate_excel_file(str(file_path)):
                validation_result['valid'] = False
                validation_result['errors'].append("Invalid Excel file format")
                return validation_result
            
            # File info
            file_stat = file_path.stat()
            validation_result['file_info'] = {
                'size_bytes': file_stat.st_size,
                'size_mb': file_stat.st_size / (1024 * 1024),
                'modified': datetime.fromtimestamp(file_stat.st_mtime).isoformat()
            }
            
            # Load and analyze structure
            try:
                excel_data = self._load_excel_file(file_path)
                
                for sheet_name, df in excel_data.items():
                    # Check required columns
                    required_columns = ['ERF Number', 'Item']
                    missing_required = [col for col in required_columns if col not in df.columns]
                    
                    if missing_required:
                        validation_result['errors'].append(
                            f"Sheet '{sheet_name}' missing required columns: {missing_required}"
                        )
                        validation_result['valid'] = False
                    
                    # Check optional but expected columns
                    expected_columns = [
                        'Engineering Request Form Status',
                        'ERF Sched Line Status',
                        'Commodity Type',
                        'Ship-To-Plant',
                        'PGr'
                    ]
                    
                    missing_expected = [col for col in expected_columns if col not in df.columns]
                    if missing_expected:
                        validation_result['warnings'].append(
                            f"Sheet '{sheet_name}' missing expected columns: {missing_expected}"
                        )
                    
                    # Column analysis
                    validation_result['column_analysis'][sheet_name] = {
                        'total_columns': len(df.columns),
                        'total_rows': len(df),
                        'columns': list(df.columns),
                        'null_counts': df.isnull().sum().to_dict(),
                        'data_types': df.dtypes.astype(str).to_dict()
                    }
            
            except Exception as e:
                validation_result['valid'] = False
                validation_result['errors'].append(f"Failed to analyze file structure: {e}")
            
            return validation_result
            
        except Exception as e:
            logger.error(f"File format validation failed: {e}")
            return {
                'valid': False,
                'errors': [f"Validation failed: {e}"],
                'warnings': [],
                'file_info': {},
                'column_analysis': {}
            }
    
    def export_processing_rules(self) -> Dict:
        """Export current processing rules configuration"""
        return {
            'cleaning_rules': self.CLEANING_RULES,
            'version': '2.0',
            'exported_at': datetime.now().isoformat()
        }
    
    def import_processing_rules(self, rules_config: Dict):
        """Import processing rules configuration"""
        try:
            if 'cleaning_rules' in rules_config:
                self.CLEANING_RULES.update(rules_config['cleaning_rules'])
                logger.info("Processing rules imported successfully")
            else:
                raise ValidationError("Invalid rules configuration format")
        except Exception as e:
            logger.error(f"Failed to import processing rules: {e}")
            raise DataProcessingError(f"Failed to import processing rules: {e}")