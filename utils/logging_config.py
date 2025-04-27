"""
Logging configuration for the Billing PDF Automation project.
"""
import os
import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler

def configure_logging(logs_dir='logs', console_level=logging.INFO, file_level=logging.DEBUG):
    """
    Configure application logging with timestamped files and console output.
    
    Args:
        logs_dir (str): Directory to store log files
        console_level: Logging level for console output
        file_level: Logging level for file output
        
    Returns:
        logger: Configured root logger
    """
    # Create logs directory if it doesn't exist
    os.makedirs(logs_dir, exist_ok=True)
    
    # Create a performance logs subdirectory
    perf_logs_dir = os.path.join(logs_dir, 'performance')
    os.makedirs(perf_logs_dir, exist_ok=True)
    
    # Get current timestamp for log filenames
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Main application log (daily rotation)
    main_log_file = os.path.join(logs_dir, f'app_{timestamp}.log')
    
    # Performance log
    perf_log_file = os.path.join(perf_logs_dir, f'performance_{timestamp}.log')
    
    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)  # Capture all logs, handlers filter levels
    
    # Clear any existing handlers
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    # Create formatters
    verbose_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s'
    )
    console_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    )
    perf_formatter = logging.Formatter(
        '%(asctime)s - %(message)s'
    )
    
    # Create file handler for main log
    file_handler = logging.FileHandler(main_log_file)
    file_handler.setLevel(file_level)
    file_handler.setFormatter(verbose_formatter)
    root_logger.addHandler(file_handler)
    
    # Create console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(console_level)
    console_handler.setFormatter(console_formatter)
    root_logger.addHandler(console_handler)
    
    # Create a separate logger for performance logs
    perf_logger = logging.getLogger('performance')
    perf_logger.setLevel(logging.INFO)
    perf_logger.propagate = False  # Don't send to root logger
    
    # Add a file handler for performance logs
    perf_file_handler = logging.FileHandler(perf_log_file)
    perf_file_handler.setFormatter(perf_formatter)
    perf_logger.addHandler(perf_file_handler)
    
    # Log startup message
    logging.info(f"Logging initialized: console={console_level}, file={file_level}")
    logging.info(f"Log files: {main_log_file}, {perf_log_file}")
    
    return root_logger