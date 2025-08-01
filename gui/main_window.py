"""
Main GUI window for Revenue Watchdog application
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime
from pathlib import Path

from core.data_parser import DataParser
from core.llm_interface import LLMInterface
from config import WINDOW_SIZE, APP_TITLE, DATETIME_FORMAT, EXPORT_COMMENT_PREFIX
from utils.helpers import center_window, format_currency


class RevenueWatchdogApp:
    """Main application class"""
    
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        
        # Parse window size from config
        width, height = map(int, WINDOW_SIZE.split('x'))
        center_window(root, width, height)
        
        # Initialize components
        self.data_parser = DataParser()
        self.llm_interface = LLMInterface()
        
        # Data storage
        self.parsed_data = []
        self.analysis_results = {}
        
        # Create UI
        self.create_ui()
        
    def create_ui(self):
        """Create the main user interface"""
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text=APP_TITLE, 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # API Key section
        self._create_api_section(main_frame)
        
        # File operations section
        self._create_file_section(main_frame)
        
        # Results section
        self._create_results_section(main_frame)
        
    def _create_api_section(self, parent):
        """Create API configuration section"""
        api_frame = ttk.LabelFrame(parent, text="LLM Configuration", padding="5")
        api_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        api_frame.columnconfigure(1, weight=1)
        
        ttk.Label(api_frame, text="OpenRouter API Key:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        self.api_key_var = tk.StringVar()
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*", width=50)
        api_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(api_frame, text="Set Key", command=self.set_api_key).grid(row=0, column=2)
        
    def _create_file_section(self, parent):
        """Create file operations section"""
        file_frame = ttk.LabelFrame(parent, text="File Operations", padding="5")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="Upload Files", command=self.upload_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(file_frame, text="Analyze Data", command=self.analyze_data).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(file_frame, text="Export Results", command=self.export_results).pack(side=tk.LEFT, padx=(0, 5))
        
        self.status_var = tk.StringVar(value="Ready - Upload files to begin")
        status_label = ttk.Label(file_frame, textvariable=self.status_var)
        status_label.pack(side=tk.RIGHT)
        
    def _create_results_section(self, parent):
        """Create results display section"""
        results_frame = ttk.LabelFrame(parent, text="Analysis Results", padding="5")
        results_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Create notebook for tabbed results
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self._create_summary_tab()
        self._create_deals_tab()
        self._create_data_tab()
        
    def _create_summary_tab(self):
        """Create executive summary tab"""
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="Executive Summary")
        
        self.summary_text = scrolledtext.ScrolledText(self.summary_frame, wrap=tk.WORD, 
                                                     height=15, width=80)
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def _create_deals_tab(self):
        """Create flagged deals tab"""
        self.deals_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.deals_frame, text="Flagged Deals")
        
        # Create treeview for deals
        columns = ('Deal ID', 'Customer', 'Risk Type', 'Impact ($)', 'Suggestion')
        self.deals_tree = ttk.Treeview(self.deals_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.deals_tree.heading(col, text=col)
            self.deals_tree.column(col, width=150)
        
        # Add scrollbars
        deals_scrollbar_y = ttk.Scrollbar(self.deals_frame, orient=tk.VERTICAL, command=self.deals_tree.yview)
        deals_scrollbar_x = ttk.Scrollbar(self.deals_frame, orient=tk.HORIZONTAL, command=self.deals_tree.xview)
        self.deals_tree.configure(yscrollcommand=deals_scrollbar_y.set, xscrollcommand=deals_scrollbar_x.set)
        
        self.deals_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        deals_scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        deals_scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        self.deals_frame.columnconfigure(0, weight=1)
        self.deals_frame.rowconfigure(0, weight=1)
        
    def _create_data_tab(self):
        """Create raw data tab"""
        self.data_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_frame, text="Raw Data")
        
        self.data_text = scrolledtext.ScrolledText(self.data_frame, wrap=tk.WORD, height=15, width=80)
        self.data_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def set_api_key(self):
        """Set the API key for LLM interface"""
        api_key = self.api_key_var.get().strip()
        if api_key:
            self.llm_interface.api_key = api_key
            self.status_var.set("API key configured")
            messagebox.showinfo("Success", "API key has been set")
        else:
            messagebox.showwarning("Warning", "Please enter a valid API key")
    
    def upload_files(self):
        """Handle file upload and parsing"""
        file_types = [
            ("All Supported", "*.csv;*.txt;*.pdf"),
            ("CSV files", "*.csv"),
            ("Text files", "*.txt"),
            ("PDF files", "*.pdf")
        ]
        
        # Remove PDF option if not available
        if not hasattr(self.data_parser, '_parse_pdf'):
            file_types.pop()
        
        file_paths = filedialog.askopenfilenames(
            title="Select deal data files",
            filetypes=file_types
        )
        
        if not file_paths:
            return
        
        self.status_var.set("Parsing files...")
        self.root.update()
        
        all_data = []
        successful_files = 0
        
        for file_path in file_paths:
            try:
                file_data = self.data_parser.parse_file(file_path)
                all_data.extend(file_data)
                successful_files += 1
            except Exception as e:
                messagebox.showerror("Error", f"Failed to parse {Path(file_path).name}:\n{str(e)}")
        
        if successful_files > 0:
            self.parsed_data = all_data
            self.display_raw_data()
            self.status_var.set(f"Loaded {len(all_data)} deals from {successful_files} files")
            messagebox.showinfo("Success", f"Successfully loaded {len(all_data)} deals from {successful_files} file(s)")
        else:
            self.status_var.set("No files loaded")
    
    def analyze_data(self):
        """Perform AI-powered analysis on parsed data"""
        if not self.parsed_data:
            messagebox.showwarning("Warning", "Please upload data files first")
            return
        
        self.status_var.set("Analyzing data with AI...")
        self.root.update()
        
        try:
            # Perform analysis
            self.analysis_results = self.llm_interface.analyze_deals(self.parsed_data)
            
            # Display results
            self.display_analysis_results()
            self.status_var.set("Analysis complete")
            
        except Exception as e:
            messagebox.showerror("Error", f"Analysis failed:\n{str(e)}")
            self.status_var.set("Analysis failed")
    
    def display_raw_data(self):
        """Display raw parsed data"""
        self.data_text.delete(1.0, tk.END)
        
        if self.parsed_data:
            # Create a summary view
            summary = f"Loaded {len(self.parsed_data)} deals\n\n"
            summary += "Sample data:\n" + "="*50 + "\n"
            
            # Show first few records
            for i, deal in enumerate(self.parsed_data[:5]):
                summary += f"\nDeal {i+1}:\n"
                for key, value in deal.items():
                    summary += f"  {key}: {value}\n"
            
            if len(self.parsed_data) > 5:
                summary += f"\n... and {len(self.parsed_data) - 5} more deals"
            
            self.data_text.insert(tk.END, summary)
    
    def display_analysis_results(self):
        """Display analysis results in the UI"""
        if not self.analysis_results:
            return
        
        # Update summary tab
        self.summary_text.delete(1.0, tk.END)
        
        summary = self.analysis_results.get('summary', {})
        flagged_deals = self.analysis_results.get('flagged_deals', [])
        recommendations = self.analysis_results.get('recommendations', [])
        
        summary_text = f"""REVENUE LEAKAGE & MARGIN ANALYSIS REPORT
Generated: {datetime.now().strftime(DATETIME_FORMAT)}

EXECUTIVE SUMMARY
=================
Total Estimated Revenue Leakage: {format_currency(summary.get('total_leakage', 0))}
High Risk Deals Identified: {summary.get('high_risk_deals', 0)}
Total Issues Found: {summary.get('issues_found', 0)}

RISK BREAKDOWN
==============
"""
        
        # Group flagged deals by risk type
        risk_types = {}
        for deal in flagged_deals:
            risk_type = deal.get('risk_type', 'unknown')
            if risk_type not in risk_types:
                risk_types[risk_type] = {'count': 0, 'impact': 0}
            risk_types[risk_type]['count'] += 1
            risk_types[risk_type]['impact'] += deal.get('impact', 0)
        
        for risk_type, data in risk_types.items():
            summary_text += f"{risk_type.replace('_', ' ').title()}: {data['count']} deals, {format_currency(data['impact'])} impact\n"
        
        summary_text += f"\nRECOMMENDATIONS\n===============\n"
        for i, rec in enumerate(recommendations, 1):
            summary_text += f"{i}. {rec}\n"
        
        self.summary_text.insert(tk.END, summary_text)
        
        # Update flagged deals tab
        self._update_deals_tree(flagged_deals)
    
    def _update_deals_tree(self, flagged_deals):
        """Update the deals treeview with flagged deals"""
        # Clear existing items
        for item in self.deals_tree.get_children():
            self.deals_tree.delete(item)
        
        # Add flagged deals
        for deal in flagged_deals:
            self.deals_tree.insert('', tk.END, values=(
                deal.get('deal_id', 'N/A'),
                deal.get('customer_name', 'N/A'),
                deal.get('risk_type', '').replace('_', ' ').title(),
                format_currency(deal.get('impact', 0)),
                deal.get('suggestion', 'N/A')
            ))
    
    def export_results(self):
        """Export analysis results to CSV"""
        if not self.analysis_results:
            messagebox.showwarning("Warning", "No analysis results to export")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Save analysis results",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        
        if not file_path:
            return
        
        try:
            flagged_deals = self.analysis_results.get('flagged_deals', [])
            
            # Create DataFrame for export
            df = pd.DataFrame(flagged_deals)
            
            # Add summary information as header comments
            summary = self.analysis_results.get('summary', {})
            
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                f.write(f"{EXPORT_COMMENT_PREFIX} Revenue Leakage Analysis Report\n")
                f.write(f"{EXPORT_COMMENT_PREFIX} Generated: {datetime.now().strftime(DATETIME_FORMAT)}\n")
                f.write(f"{EXPORT_COMMENT_PREFIX} Total Leakage: {format_currency(summary.get('total_leakage', 0))}\n")
                f.write(f"{EXPORT_COMMENT_PREFIX} High Risk Deals: {summary.get('high_risk_deals', 0)}\n")
                f.write(f"{EXPORT_COMMENT_PREFIX}\n")
                
                # Write CSV data
                df.to_csv(f, index=False)
            
            self.status_var.set(f"Results exported to {Path(file_path).name}")
            messagebox.showinfo("Success", f"Results exported successfully to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{str(e)}")