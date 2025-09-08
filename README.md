# ZERF Data Automation System



A comprehensive Python automation tool that streamlines the ZERF (SAP) data extraction, cleaning, and SharePoint upload process with a user-friendly GUI interface.
## Data Extraction Automation For Lam Research(DEALR)

## Overview

The ZERF Data Automation System automates the complete workflow of:
1. Extracting data from SAP using VBS scripts
2. Cleaning and filtering Excel data according to business rules
3. Uploading processed files to SharePoint
4. Creating backups and maintaining audit logs

## Features

### Core Functionality
- **Automated SAP Data Extraction**: VBS script automation for SAP ZERF transaction
- **Advanced Data Cleaning**: 7-step data cleaning process with business rule filters
- **SharePoint Integration**: Automatic upload to configured SharePoint locations
- **Date Range Configuration**: User-configurable start and end dates for data extraction
- **Backup Management**: Automatic backup of original and processed files
- **Comprehensive Logging**: Detailed logs with timestamps and error tracking

### GUI Features
- **User-Friendly Interface**: Tabbed interface with Configuration, Control, and Logs
- **Real-time Monitoring**: Live log display with color-coded messages
- **Date Picker Widgets**: Calendar-based date selection (with fallback to text entry)
- **Manual Processing**: Option to process files manually outside of automation
- **Scheduler Control**: Start/stop scheduled automation with status monitoring

### Data Cleaning Rules
1. **Unique ID Creation**: Combines ERF Number + Item for duplicate detection
2. **Duplicate Removal**: Removes duplicate records based on Unique_ID
3. **Status Filtering**: Excludes Draft, Presubmit, and Submit statuses
4. **Blank Status Removal**: Filters out blank ERF Sched Line Status entries
5. **Commodity Type Filter**: Removes records with "Indirect" commodity types
6. **Plant Filtering**: Keeps only specific Ship-To-Plant values (6100, 6200, 6300)
7. **PGr Filtering**: Removes records with PGr values "W91" and "Z05"

## Installation

### Prerequisites
- Python 3.7 or higher
- Windows OS (for SAP GUI automation)
- SAP GUI installed and configured
- Excel/Office applications

### Required Python Packages
```bash
pip install pandas openpyxl schedule configparser pathlib
```

### Optional Packages
```bash
# For enhanced date picker widgets
pip install tkcalendar

# For SharePoint integration (choose one)
pip install Office365-REST-Python-Client
# OR
pip install sharepy
```

## Setup Instructions

### 1. Initial Setup
1. Download the `zerf_automation.py` file to your desired directory
2. Ensure SAP GUI is installed and you can access the ZERF transaction
3. Install required Python packages using pip commands above

### 2. Configuration
Run the application in GUI mode:
```bash
python zerf_automation.py --gui
```

#### Configure the following in the Configuration tab:

**Date Range Configuration:**
- **Start Date**: Default is 08/03/2025 (MM/DD/YYYY format)
- **End Date**: Default is current date (MM/DD/YYYY format)
- Use "Set to Today" button to quickly update end date

**SharePoint Configuration:**
- **Site URL**: Your SharePoint site URL
- **Username**: Your SharePoint username
- **Password**: Your SharePoint password

**Paths Configuration:**
- **Download Folder**: Where Excel files will be saved (default: ./downloads)

**Schedule Configuration:**
- **Daily Run Time**: Time for automatic execution (HH:MM format, default: 08:00)

### 3. Save Configuration
Click "Save Configuration" to store your settings in `zerf_config.ini`

## Usage

### GUI Mode (Recommended)
```bash
python zerf_automation.py --gui
```

#### Control Tab Options:
- **Run Now**: Execute the workflow immediately with current settings
- **Start Scheduler**: Begin daily scheduled automation
- **Stop**: Stop the automation system
- **Test File Detection**: Test the file detection functionality
- **Process File Manually**: Select and process an Excel file manually

### Command Line Mode
```bash
# Run immediately with default settings
python zerf_automation.py --run-now

# Run with custom date range
python zerf_automation.py --run-now --start-date "08/03/2025" --end-date "12/31/2025"

# Run in background (scheduled mode)
python zerf_automation.py --background
```

## Workflow Process

### Automated Workflow Steps:
1. **VBS Script Execution**: Automates SAP GUI to extract ZERF data
2. **File Detection**: Automatically locates the downloaded Excel file
3. **Data Cleaning**: Applies 7-step cleaning process
4. **SharePoint Upload**: Uploads cleaned file to configured location
5. **Backup Creation**: Creates timestamped backups of original and cleaned files
6. **Logging**: Records all activities with detailed timestamps

### Manual Processing:
- Use "Process File Manually" to clean any Excel file
- Bypasses SAP automation but applies same cleaning rules
- Useful for processing historical files or testing

## File Structure

```
project_directory/
├── zerf_automation.py          # Main application file
├── zerf_config.ini            # Configuration file (auto-generated)
├── zerf_automation.vbs        # SAP automation script (auto-generated)
├── downloads/                 # Downloaded files directory
├── backup/                    # Backup files directory
└── logs/                      # Log files directory
    └── zerf_automation_YYYYMMDD.log
```

## Configuration File (zerf_config.ini)

The system automatically creates and manages a configuration file with the following sections:

```ini
[DateRange]
start_date = 08/03/2025
end_date = 09/07/2025

[SharePoint]
site_url = https://your-site.sharepoint.com/sites/your-site
username = your-username
password = your-password
folder_path = ERF Reporting_Data Analytics & Power BI

[Paths]
download_folder = downloads
vbs_script = zerf_automation.vbs
backup_folder = backup

[Schedule]
run_time = 08:00
check_interval = 30

[Settings]
auto_start = true
cleanup_old_files = true
max_retries = 3
```

## Troubleshooting

### Common Issues:

#### SAP Connection Issues:
- Ensure SAP GUI is running and logged in
- Verify access to ZERF transaction
- Check that SAP GUI scripting is enabled

#### File Detection Problems:
- Check that Excel files are being saved to the expected location
- Ensure files are not locked by other applications
- Use "Test File Detection" button to diagnose issues

#### SharePoint Upload Failures:
- Verify SharePoint credentials are correct
- Check network connectivity to SharePoint
- Ensure you have write permissions to the target folder

#### Date Format Issues:
- Use MM/DD/YYYY format for all dates
- Ensure start date is before end date
- Check that dates are valid calendar dates

### Log Files:
- Check daily log files in the `logs/` directory
- Logs include detailed error messages and timestamps
- Use "Refresh Logs" button in GUI to view latest entries

## Security Considerations

- **Password Storage**: SharePoint passwords are stored in plain text in the config file
- **Recommendation**: Use application-specific passwords or consider implementing encryption
- **File Permissions**: Ensure config files have appropriate access restrictions
- **Network Security**: Be aware that credentials are transmitted over the network

## Support and Maintenance

### Regular Maintenance:
- Monitor log files for errors or warnings
- Clean up old backup files periodically
- Update SharePoint credentials as needed
- Test automation periodically to ensure SAP compatibility

### System Requirements:
- Windows environment for SAP GUI integration
- Adequate disk space for downloads and backups
- Stable network connection for SharePoint uploads
- SAP GUI user access and permissions

## Customization

### Adding New Data Cleaning Rules:
Modify the `clean_excel_data()` method in the main class to add additional filtering steps.

### Changing File Naming Conventions:
Update the `get_today_filename()` method to modify output file naming patterns.

### Extending SharePoint Integration:
Additional SharePoint libraries can be integrated by modifying the upload methods.

## Version History

- **v2.0**: Added date range configuration and PGr filtering
- **v1.0**: Initial release with basic automation functionality

## Contact

For technical support or questions about this automation system, contact your IT team or the system administrator.

---

**Note**: This system requires proper SAP GUI access and SharePoint permissions. Ensure you have the necessary credentials and access rights before deployment.