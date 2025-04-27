"""
Decorator functions for the Billing PDF Automation project.
"""
import time
import functools
import logging
from datetime import datetime
import os

def performance_logger(output_dir=None, log_to_console=True):
    """
    Decorator that logs the execution time of functions.
    
    Args:
        output_dir (str, optional): Directory to save performance logs.
            If None, only logs to console/main log. Defaults to None.
        log_to_console (bool, optional): Whether to print timing to console.
            Defaults to True.
            
    Returns:
        decorator: The performance measurement decorator
    
    Example:
        @performance_logger()
        def my_function():
            # Function code here
            
        @performance_logger(output_dir='logs', log_to_console=False)
        def another_function():
            # Function code here
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Start time measurement
            start_time = time.perf_counter()
            
            # Call the function
            result = func(*args, **kwargs)
            
            # End time measurement
            end_time = time.perf_counter()
            execution_time = end_time - start_time
            
            # Prepare log message
            function_name = func.__name__
            module_name = func.__module__
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Prepare arguments string (limited to prevent huge logs)
            args_str = str(args)[:50] + ('...' if len(str(args)) > 50 else '')
            kwargs_str = str(kwargs)[:50] + ('...' if len(str(kwargs)) > 50 else '')
            
            log_message = (f"PERF: {timestamp} | {module_name}.{function_name} | "
                          f"Time: {execution_time:.4f}s | "
                          f"Args: {args_str} | Kwargs: {kwargs_str}")
            
            # Log to main application log
            logging.info(log_message)
            
            # Log to console if requested
            if log_to_console:
                print(log_message)
            
            # Save to performance log file if directory specified
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
                log_file = os.path.join(output_dir, 'performance.log')
                
                with open(log_file, 'a') as f:
                    f.write(f"{log_message}\n")
            
            return result
        return wrapper
    return decorator


def cache_result(expiry_seconds=3600):
    """
    Cache function results with time-based expiration.
    
    Args:
        expiry_seconds (int, optional): Cache lifetime in seconds. 
            Defaults to 3600 (1 hour).
            
    Returns:
        decorator: The caching decorator
        
    Example:
        @cache_result(expiry_seconds=1800)  # Cache for 30 minutes
        def get_invoice_data(invoice_id):
            # Expensive operation to fetch invoice data
            return data
    """
    # Storage for cached results
    cache = {}
    
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Create a cache key from function name and arguments
            # Using repr for simple serialization - for more complex needs, consider hash or json
            key = (func.__name__, repr(args), repr(kwargs))
            
            # Check if we have a cached result and it's still valid
            now = time.time()
            if key in cache:
                result, timestamp = cache[key]
                # If the cache hasn't expired, return the cached result
                if now - timestamp < expiry_seconds:
                    logging.debug(f"Cache hit for {func.__name__}")
                    return result
                logging.debug(f"Cache expired for {func.__name__}")
            else:
                logging.debug(f"Cache miss for {func.__name__}")
            
            # Call the function and cache the result
            result = func(*args, **kwargs)
            cache[key] = (result, now)
            return result
        return wrapper
    return decorator


def retry(max_attempts=3, delay=1, backoff=2, exceptions=(Exception,)):
    """
    Retry decorator with exponential backoff.
    
    Args:
        max_attempts (int, optional): Maximum number of retry attempts. 
            Defaults to 3.
        delay (float, optional): Initial delay between retries in seconds. 
            Defaults to 1.
        backoff (float, optional): Backoff multiplier. 
            Defaults to 2.
        exceptions (tuple, optional): Exceptions that trigger a retry. 
            Defaults to (Exception,).
            
    Returns:
        decorator: The retry decorator
        
    Example:
        @retry(max_attempts=5, delay=2, exceptions=(ConnectionError, TimeoutError))
        def api_call():
            # Code that might fail temporarily
            return response
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Initialize variables for retry logic
            attempts = 0
            current_delay = delay
            
            while attempts < max_attempts:
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    attempts += 1
                    if attempts == max_attempts:
                        logging.error(f"Function {func.__name__} failed after {attempts} attempts. Error: {str(e)}")
                        raise
                    
                    # Log the retry attempt
                    logging.warning(
                        f"Retry {attempts}/{max_attempts} for {func.__name__} after error: {str(e)}. "
                        f"Waiting {current_delay}s before next attempt."
                    )
                    
                    # Wait before the next attempt
                    time.sleep(current_delay)
                    
                    # Increase the delay for the next attempt (exponential backoff)
                    current_delay *= backoff
        return wrapper
    return decorator