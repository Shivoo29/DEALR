#!/usr/bin/env python3
"""
ZERF Data Automation System - Main Entry Point
==============================================
Data Extraction Automation For Lam Research
==============================================
"""

import sys
import argparse
from pathlib import Path

# Add src to Python path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from utils.logger import get_logger
from core.automation_engine import ZERFAutomationEngine
from gui.main_window import ZERFAutomationGUI

logger = get_logger(__name__)

def main():
    """Main entry point for the ZERF Automation System"""
    parser = argparse.ArgumentParser(
        description='ZERF Data Automation System - SAP Data Processing Automation',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py --gui                    Launch GUI interface
  python main.py --run-now               Run workflow immediately
  python main.py --background            Run in background scheduler mode
  python main.py --run-now --start-date "08/03/2025" --end-date "09/15/2025"
        """
    )
    
    # Operation modes
    parser.add_argument('--gui', action='store_true', 
                       help='Launch GUI interface (default if no other mode specified)')
    parser.add_argument('--background', action='store_true', 
                       help='Run in background scheduler mode')
    parser.add_argument('--run-now', action='store_true', 
                       help='Run workflow immediately')
    
    # Configuration overrides
    parser.add_argument('--start-date', metavar='MM/DD/YYYY',
                       help='Override start date (MM/DD/YYYY format)')
    parser.add_argument('--end-date', metavar='MM/DD/YYYY',
                       help='Override end date (MM/DD/YYYY format)')
    parser.add_argument('--config', metavar='PATH',
                       help='Path to configuration file (default: config/zerf_config.ini)')
    
    # Utility options
    parser.add_argument('--validate-config', action='store_true',
                       help='Validate configuration and exit')
    parser.add_argument('--test-sharepoint', action='store_true',
                       help='Test SharePoint connection and exit')
    parser.add_argument('--version', action='version', version='ZERF Automation System 2.0')
    
    args = parser.parse_args()
    
    try:
        # Default to GUI if no mode specified
        if not any([args.gui, args.background, args.run_now, args.validate_config, args.test_sharepoint]):
            args.gui = True
        
        if args.gui:
            logger.info("Starting ZERF Automation System GUI")
            app = ZERFAutomationGUI(config_file=args.config)
            app.run()
            
        elif args.validate_config:
            logger.info("Validating configuration...")
            engine = ZERFAutomationEngine(config_file=args.config)
            if engine.validate_configuration():
                print("✅ Configuration is valid")
                return 0
            else:
                print("❌ Configuration validation failed")
                return 1
                
        elif args.test_sharepoint:
            logger.info("Testing SharePoint connection...")
            engine = ZERFAutomationEngine(config_file=args.config)
            if engine.test_sharepoint_connection():
                print("✅ SharePoint connection successful")
                return 0
            else:
                print("❌ SharePoint connection failed")
                return 1
                
        elif args.run_now:
            logger.info("Running ZERF automation workflow immediately")
            engine = ZERFAutomationEngine(config_file=args.config)
            
            # Override dates if provided
            if args.start_date:
                engine.config_manager.set_start_date(args.start_date)
            if args.end_date:
                engine.config_manager.set_end_date(args.end_date)
            
            success = engine.run_full_workflow()
            return 0 if success else 1
            
        elif args.background:
            logger.info("Starting ZERF automation in background scheduler mode")
            engine = ZERFAutomationEngine(config_file=args.config)
            try:
                engine.start_scheduler()
            except KeyboardInterrupt:
                logger.info("Received shutdown signal")
                engine.stop()
                return 0
                
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user")
        return 0
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        return 1

if __name__ == "__main__":
    sys.exit(main())