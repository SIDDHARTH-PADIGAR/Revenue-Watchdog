"""
Utility functions for Revenue Watchdog application
"""

import pandas as pd
from datetime import datetime
from typing import List


def is_date_past(date_str: str) -> bool:
    """Check if date string represents a past date"""
    try:
        date_obj = pd.to_datetime(date_str)
        return date_obj < datetime.now()
    except:
        return False


def check_dependencies() -> List[str]:
    """Check for missing required dependencies"""
    missing_deps = []
    
    try:
        import pandas
    except ImportError:
        missing_deps.append("pandas")
    
    try:
        import requests
    except ImportError:
        missing_deps.append("requests")
    
    return missing_deps


def format_currency(amount: float) -> str:
    """Format amount as currency string"""
    return f"${amount:,.2f}"


def center_window(root, width: int, height: int):
    """Center a tkinter window on screen"""
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")


def normalize_column_name(col_name: str) -> str:
    """Normalize column name to standard format"""
    return col_name.lower().replace(' ', '_')


def validate_deal_data(deal: dict) -> bool:
    """Validate that a deal record has required fields"""
    required_fields = ['deal_id', 'customer_name', 'deal_size']
    return all(field in deal for field in required_fields)