"""
Logging utilities for ZERF Automation System
"""

import logging
import logging.handlers
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional

try:
    import colorlog
    COLORLOG_AVAILABLE = True
except ImportError:
    COLORLOG_AVAILABLE = False

class ZERFLogger:
    """Enhanced logger for ZERF Automation System"""
    
    _loggers = {}
    _log_level = logging.INFO
    _log_dir = Path("logs")
    
    @classmethod
    def setup_logging(cls, log_level: str = "INFO", log_dir: str = "logs"):
        """Setup global logging configuration"""
        cls._log_level = getattr(logging, log_level.upper())
        cls._log_dir = Path(log_dir)
        cls._log_dir.mkdir(exist_ok=True)
    
    @classmethod
    def get_logger(cls, name: str, enable_console: bool = True, enable_file: bool = True) -> logging.Logger:
        """Get or create a logger with the specified configuration"""
        
        if name in cls._loggers:
            return cls._loggers[name]
        
        logger = logging.getLogger(name)
        logger.setLevel(cls._log_level)
        
        # Clear existing handlers to avoid duplicates
        logger.handlers.clear()
        
        # File handler with rotation
        if enable_file:
            log_file = cls._log_dir / f"zerf_automation_{datetime.now().strftime('%Y%m%d')}.log"
            file_handler = logging.handlers.RotatingFileHandler(
                log_file,
                maxBytes=10*1024*1024,  # 10MB
                backupCount=5,
                encoding='utf-8'
            )
            file_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            file_handler.setFormatter(file_formatter)
            logger.addHandler(file_handler)
        
        # Console handler with colors
        if enable_console:
            console_handler = logging.StreamHandler(sys.stdout)
            
            if COLORLOG_AVAILABLE:
                console_formatter = colorlog.ColoredFormatter(
                    '%(log_color)s%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S',
                    log_colors={
                        'DEBUG': 'cyan',
                        'INFO': 'green',
                        'WARNING': 'yellow',
                        'ERROR': 'red',
                        'CRITICAL': 'red,bg_white',
                    }
                )
            else:
                console_formatter = logging.Formatter(
                    '%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S'
                )
            
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
        
        # Prevent propagation to avoid duplicate logs
        logger.propagate = False
        
        cls._loggers[name] = logger
        return logger

class GUILogHandler(logging.Handler):
    """Custom log handler for GUI applications"""
    
    def __init__(self, gui_callback):
        super().__init__()
        self.gui_callback = gui_callback
        self.setFormatter(logging.Formatter('%(message)s'))
    
    def emit(self, record):
        if self.gui_callback:
            try:
                msg = self.format(record)
                level = self._get_level_name(record.levelno)
                self.gui_callback(msg, level)
            except Exception:
                # Avoid recursive logging errors
                pass
    
    def _get_level_name(self, levelno: int) -> str:
        """Convert log level number to string for GUI"""
        if levelno >= logging.CRITICAL:
            return "CRITICAL"
        elif levelno >= logging.ERROR:
            return "ERROR"
        elif levelno >= logging.WARNING:
            return "WARNING"
        elif levelno >= logging.INFO:
            return "INFO"
        else:
            return "DEBUG"

class ProgressLogger:
    """Logger with progress tracking capabilities"""
    
    def __init__(self, logger: logging.Logger, total_steps: int, description: str = ""):
        self.logger = logger
        self.total_steps = total_steps
        self.current_step = 0
        self.description = description
        self.start_time = datetime.now()
    
    def step(self, message: str = "", increment: int = 1):
        """Log a step in the progress"""
        self.current_step += increment
        progress = (self.current_step / self.total_steps) * 100
        
        elapsed = datetime.now() - self.start_time
        
        if message:
            full_message = f"[{progress:.1f}%] {self.description} - {message}"
        else:
            full_message = f"[{progress:.1f}%] {self.description} - Step {self.current_step}/{self.total_steps}"
        
        self.logger.info(full_message)
        
        if self.current_step >= self.total_steps:
            self.logger.info(f"✅ {self.description} completed in {elapsed.total_seconds():.1f}s")
    
    def error(self, message: str):
        """Log an error during progress"""
        elapsed = datetime.now() - self.start_time
        self.logger.error(f"❌ {self.description} failed at step {self.current_step}: {message} (after {elapsed.total_seconds():.1f}s)")

# Convenience functions
def get_logger(name: str, level: str = "INFO") -> logging.Logger:
    """Get a logger instance with default configuration"""
    ZERFLogger.setup_logging(level)
    return ZERFLogger.get_logger(name)

def setup_logging(level: str = "INFO", log_dir: str = "logs"):
    """Setup global logging configuration"""
    ZERFLogger.setup_logging(level, log_dir)

def create_progress_logger(logger: logging.Logger, total_steps: int, description: str = "") -> ProgressLogger:
    """Create a progress logger"""
    return ProgressLogger(logger, total_steps, description)