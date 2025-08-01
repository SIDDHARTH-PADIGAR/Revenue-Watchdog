#!/usr/bin/env python3
"""
Revenue Leakage & Margin Watchdog
A desktop finance tool for analyzing deal data and identifying revenue risks.

Features:
- File ingestion (CSV, PDF, TXT)
- Automated data analysis with AI-powered insights
- Risk and margin surfacing
- Export capabilities
- Action suggestions

Author: AI Assistant
Requirements: Python 3.7+, tkinter, pandas, requests
Optional: PyPDF2 or pdfplumber for PDF processing
"""

import tkinter as tk
import sys

from gui.main_window import RevenueWatchdogApp
from utils.helpers import check_dependencies
from config import WINDOW_SIZE


def main():
    """Main application entry point"""
    
    # Check required dependencies
    missing_deps = check_dependencies()
    
    if missing_deps:
        print(f"Missing required dependencies: {', '.join(missing_deps)}")
        print("Install with: pip install " + " ".join(missing_deps))
        return
    
    # Initialize and run application
    root = tk.Tk()
    app = RevenueWatchdogApp(root)
    
    print("Starting Revenue Leakage & Margin Watchdog...")
    print("=" * 50)
    print("SETUP INSTRUCTIONS:")
    print("1. Get an API key from OpenRouter.ai")
    print("2. Enter your API key in the application")
    print("3. Upload CSV/TXT/PDF files with deal data")
    print("4. Click 'Analyze Data' for AI-powered insights")
    print("5. Export results for further action")
    print("=" * 50)
    
    root.mainloop()


if __name__ == "__main__":
    main()