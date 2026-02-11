# core/logger.py
import logging
import os
from logging.handlers import RotatingFileHandler

LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs")
LOG_FILE = os.path.join(LOG_DIR, "system.log")

# Global flag to disable file logging in worker processes
_FILE_LOGGING_DISABLED = False

def setup_logger(name: str, disable_file_logging: bool = False) -> logging.Logger:
    """
    Setup logger with console and optional file handlers.
    
    Args:
        name: Logger name
        disable_file_logging: If True, only log to console (for multiprocessing compatibility)
    """
    global _FILE_LOGGING_DISABLED
    
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        formatter = logging.Formatter(
            "[%(asctime)s] [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        logger.addHandler(ch)

        # Rotating file handler (disabled in multiprocessing to avoid file locks)
        # Check both parameter and global flag
        if not disable_file_logging and not _FILE_LOGGING_DISABLED:
            fh = RotatingFileHandler(LOG_FILE, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8")
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(formatter)
            logger.addHandler(fh)

    return logger


def disable_file_logging():
    """
    Disable file logging for all existing loggers and prevent future file handlers.
    Call this in worker processes to avoid file lock conflicts.
    """
    global _FILE_LOGGING_DISABLED
    _FILE_LOGGING_DISABLED = True
    
    # Remove file handlers from all existing loggers
    for logger_name in list(logging.Logger.manager.loggerDict.keys()):
        logger_obj = logging.getLogger(logger_name)
        handlers_to_remove = []
        for handler in logger_obj.handlers:
            if isinstance(handler, RotatingFileHandler):
                handlers_to_remove.append(handler)
        
        for handler in handlers_to_remove:
            logger_obj.removeHandler(handler)
            handler.close()


logger = setup_logger("default")