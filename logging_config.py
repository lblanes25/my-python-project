# logging_config.py
import logging

def setup_logging():
    """Configure logging for the application"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("qa_analytics.log"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("qa_analytics")