"""
Data parsing module for Revenue Watchdog
Handles file ingestion and data parsing for CSV, TXT, and PDF files
"""

import pandas as pd
import os
import sys
import re
import tempfile
from pathlib import Path
from typing import Dict, List, Any

from config import SUPPORTED_FORMATS, COLUMN_MAPPING, STANDARD_FIELDS, MIN_DEAL_AMOUNT, MAX_PDF_AMOUNTS

# Optional PDF processing imports
try:
    import PyPDF2
    PDF_AVAILABLE = True
except ImportError:
    try:
        import pdfplumber
        PDF_AVAILABLE = True
    except ImportError:
        PDF_AVAILABLE = False


class DataParser:
    """Handles file ingestion and data parsing"""
    
    def __init__(self):
        self.supported_formats = SUPPORTED_FORMATS
        
    def parse_file(self, file_path: str) -> List[Dict[str, Any]]:
        """Parse uploaded file and extract deal data"""
        file_ext = Path(file_path).suffix.lower()
        
        try:
            if file_ext == '.csv':
                return self._parse_csv(file_path)
            elif file_ext == '.txt':
                return self._parse_txt(file_path)
            elif file_ext == '.pdf':
                return self._parse_pdf(file_path)
            else:
                raise ValueError(f"Unsupported file format: {file_ext}")
        except Exception as e:
            raise Exception(f"Error parsing {file_path}: {str(e)}")
    
    def _parse_csv(self, file_path: str) -> List[Dict[str, Any]]:
        """Parse CSV file containing deal data"""
        df = pd.read_csv(file_path)
        
        # Normalize column names
        df.columns = [col.lower().replace(' ', '_') for col in df.columns]
        
        # Map common column variations to standard names
        for old_col, new_col in COLUMN_MAPPING.items():
            if old_col in df.columns and new_col not in df.columns:
                df.rename(columns={old_col: new_col}, inplace=True)
        
        # Add missing standard fields with defaults
        for field, default in STANDARD_FIELDS.items():
            if field not in df.columns:
                if callable(default):
                    df[field] = [default(i) for i in range(len(df))]
                else:
                    df[field] = default
        
        return df.to_dict('records')
    
    def _parse_txt(self, file_path: str) -> List[Dict[str, Any]]:
        """Parse text file - assumes structured format or CSV-like"""
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Try to detect if it's CSV-like
        lines = content.strip().split('\n')
        if len(lines) > 1 and (',' in lines[0] or '\t' in lines[0]):
            # Treat as CSV
            delimiter = ',' if ',' in lines[0] else '\t'
            
            # Write to temp CSV and parse
            with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
                tmp.write(content.replace('\t', ',') if delimiter == '\t' else content)
                tmp_path = tmp.name
            
            try:
                result = self._parse_csv(tmp_path)
                os.unlink(tmp_path)
                return result
            except:
                os.unlink(tmp_path)
                raise
        
        # TODO: Implement more sophisticated text parsing
        # For now, create a single dummy record
        return [{
            'deal_id': 'TXT_001',
            'customer_name': 'From Text File',
            'deal_size': 10000,
            'discount_percent': 0,
            'close_date': '',
            'renewal': '',
            'deal_status': 'Imported from Text'
        }]
    
    def _parse_pdf(self, file_path: str) -> List[Dict[str, Any]]:
        """Parse PDF file - extract text and look for deal data"""
        if not PDF_AVAILABLE:
            raise Exception("PDF parsing requires PyPDF2 or pdfplumber. Install with: pip install PyPDF2")
        
        text_content = ""
        
        try:
            # Try PyPDF2 first
            if 'PyPDF2' in sys.modules:
                with open(file_path, 'rb') as f:
                    reader = PyPDF2.PdfReader(f)
                    for page in reader.pages:
                        text_content += page.extract_text() + "\n"
            
            # TODO: Add pdfplumber support for better table extraction
            # elif 'pdfplumber' in sys.modules:
            #     with pdfplumber.open(file_path) as pdf:
            #         for page in pdf.pages:
            #             text_content += page.extract_text() + "\n"
            
        except Exception as e:
            raise Exception(f"Error reading PDF: {str(e)}")
        
        # Basic text parsing - look for numbers that might be deal values
        amounts = re.findall(r'\$?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)', text_content)
        
        # Create dummy deal records from PDF
        deals = []
        for i, amount in enumerate(amounts[:MAX_PDF_AMOUNTS]):  # Limit to first 5 amounts found
            clean_amount = float(amount.replace(',', ''))
            if clean_amount > MIN_DEAL_AMOUNT:  # Only consider significant amounts
                deals.append({
                    'deal_id': f'PDF_{i+1:03d}',
                    'customer_name': f'PDF Customer {i+1}',
                    'deal_size': clean_amount,
                    'discount_percent': 0,
                    'close_date': '',
                    'renewal': '',
                    'deal_status': 'Extracted from PDF'
                })
        
        if not deals:
            # Create at least one record
            deals.append({
                'deal_id': 'PDF_001',
                'customer_name': 'PDF Import',
                'deal_size': 50000,
                'discount_percent': 0,
                'close_date': '',
                'renewal': '',
                'deal_status': 'PDF Content'
            })
        
        return deals