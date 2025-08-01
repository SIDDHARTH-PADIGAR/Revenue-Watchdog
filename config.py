"""
Configuration constants for Revenue Watchdog application
"""

# API Configuration
DEFAULT_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "mistralai/mistral-7b-instruct"
API_TIMEOUT = 30

# File Processing
SUPPORTED_FORMATS = ['.csv', '.txt', '.pdf']
MAX_PDF_AMOUNTS = 5
MIN_DEAL_AMOUNT = 1000

# Data Processing
COLUMN_MAPPING = {
    'customer': 'customer_name',
    'client': 'customer_name',
    'account': 'customer_name',
    'deal_value': 'deal_size',
    'contract_value': 'deal_size',
    'amount': 'deal_size',
    'discount': 'discount_percent',
    'disc_%': 'discount_percent',
    'price': 'unit_price',
    'renewal_date': 'renewal',
    'close': 'close_date',
    'status': 'deal_status'
}

STANDARD_FIELDS = {
    'deal_id': lambda i: f"DEAL_{i:04d}",
    'customer_name': 'Unknown Customer',
    'deal_size': 0,
    'discount_percent': 0,
    'close_date': '',
    'renewal': '',
    'deal_status': 'Open'
}

# Analysis Rules
HIGH_DISCOUNT_THRESHOLD = 20  # Percentage
OPPORTUNITY_COST_FACTOR = 0.1

# UI Configuration
WINDOW_SIZE = "1000x700"
APP_TITLE = "Revenue Leakage & Margin Watchdog"

# Export Settings
EXPORT_COMMENT_PREFIX = "#"
DATETIME_FORMAT = '%Y-%m-%d %H:%M:%S'