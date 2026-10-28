# logger_config.py

import logging
from pathlib import Path
from typing import Optional, Any
from datetime import datetime

class TestLogger:
    """Centralized logging configuration for test automation."""
    
    _instance: Optional['TestLogger'] = None
    
    def __new__(cls) -> 'TestLogger':
        """Implement singleton pattern for logger."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize_logger()
        return cls._instance

    def _initialize_logger(self) -> None:
        """Initialize logger with file and console handlers."""
        self.logger = logging.getLogger('TestAutomation')
        self.logger.setLevel(logging.INFO)

        if not self.logger.handlers:
            # Create formatters
            console_formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            
            file_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )

            # Console Handler
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(console_formatter)
            console_handler.setLevel(logging.INFO)
            
            # File Handler
            log_dir = Path(__file__).parent / 'logs'
            log_dir.mkdir(exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = log_dir / f'test_automation_{timestamp}.log'
            
            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(file_formatter)
            file_handler.setLevel(logging.DEBUG)

            self.logger.addHandler(console_handler)
            self.logger.addHandler(file_handler)

    def get_logger(self) -> logging.Logger:
        """Get the configured logger instance."""
        return self.logger

    @staticmethod
    def log_method_call(method_name: str, **kwargs) -> None:
        """
        Utility method to log method calls with parameters.
        
        Args:
            method_name: Name of the method being called
            kwargs: Method parameters to log
        """
        logger = TestLogger().get_logger()
        params = ', '.join(f'{k}={v}' for k, v in kwargs.items())
        logger.debug(f"Calling method '{method_name}' with parameters: {params}")

    @staticmethod
    def log_method_completion(method_name: str, message: str = None) -> None:
        """
        Utility method to log method completion.
        
        Args:
            method_name: Name of the completed method
            message: Optional completion message
        """
        logger = TestLogger().get_logger()
        if message:
            logger.debug(f"Method '{method_name}' completed: {message}")
        else:
            logger.debug(f"Method '{method_name}' completed successfully")

    @staticmethod
    def log_method(func):
        """Decorator for logging method calls and completions."""
        def wrapper(self, *args, **kwargs):
            method_name = func.__name__
            TestLogger.log_method_call(method_name)
            try:
                result = func(self, *args, **kwargs)
                TestLogger.log_method_completion(method_name)
                return result
            except Exception as e:
                self.logger.error(f"Failed to execute {method_name}: {str(e)}")
                raise RuntimeError(f"{method_name} execution failed: {str(e)}")
        return wrapper