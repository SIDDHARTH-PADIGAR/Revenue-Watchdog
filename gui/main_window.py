"""
Main GUI window for Revenue Watchdog application - Polished Version
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
    """Main application class with polished UI/UX"""
    
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        
        # Configure style
        self.setup_styles()
        
        # Parse window size from config
        width, height = map(int, WINDOW_SIZE.split('x'))
        center_window(root, width, height)
        
        # Initialize components
        self.data_parser = DataParser()
        self.llm_interface = LLMInterface()
        
        # Data storage
        self.parsed_data = []
        self.analysis_results = {}
        
        # UI state variables
        self.api_configured = False
        self.analysis_in_progress = False
        
        # Create UI
        self.create_ui()
        self.update_ui_state()
        
    def setup_styles(self):
        """Configure custom styles for better appearance"""
        style = ttk.Style()
        
        # Configure colors and fonts
        style.configure("Title.TLabel", font=('Segoe UI', 18, 'bold'), foreground='#2c3e50')
        style.configure("Subtitle.TLabel", font=('Segoe UI', 9), foreground='#7f8c8d')
        style.configure("Success.TLabel", foreground='#27ae60')
        style.configure("Warning.TLabel", foreground='#e67e22')
        style.configure("Error.TLabel", foreground='#e74c3c')
        
        # Button styles
        style.map("Accent.TButton",
                 background=[('active', '#3498db'), ('pressed', '#2980b9')])
        
    def create_ui(self):
        """Create the main user interface with improved layout"""
        
        # Main container with padding
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Header section
        self._create_header(main_container)
        
        # Configuration section
        self._create_config_section(main_container)
        
        # Action buttons section
        self._create_action_section(main_container)
        
        # Status bar
        self._create_status_bar(main_container)
        
        # Results section (expandable)
        self._create_results_section(main_container)
        
    def _create_header(self, parent):
        """Create application header with title and description"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # App title
        title_label = ttk.Label(header_frame, text=APP_TITLE, style="Title.TLabel")
        title_label.pack(anchor=tk.W)
        
        # Subtitle
        subtitle_label = ttk.Label(header_frame, 
                                  text="AI-powered revenue leakage detection and deal analysis", 
                                  style="Subtitle.TLabel")
        subtitle_label.pack(anchor=tk.W, pady=(2, 0))
        
        # Add separator
        ttk.Separator(header_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(10, 0))
        
    def _create_config_section(self, parent):
        """Create configuration section with improved layout"""
        config_frame = ttk.LabelFrame(parent, text=" 🔑 API Configuration", padding=15)
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create grid layout
        config_frame.columnconfigure(1, weight=1)
        
        # API Key label and entry
        ttk.Label(config_frame, text="OpenRouter API Key:", 
                 font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.api_key_var = tk.StringVar()
        self.api_entry = ttk.Entry(config_frame, textvariable=self.api_key_var, 
                                  show="*", font=('Consolas', 9))
        self.api_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.api_button = ttk.Button(config_frame, text="Configure", 
                                    command=self.set_api_key, style="Accent.TButton")
        self.api_button.grid(row=0, column=2)
        
        # API status indicator
        self.api_status_frame = ttk.Frame(config_frame)
        self.api_status_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        self.api_status_label = ttk.Label(self.api_status_frame, text="⚠️ API key required")
        self.api_status_label.pack(side=tk.LEFT)
        
    def _create_action_section(self, parent):
        """Create action buttons section with better organization"""
        action_frame = ttk.LabelFrame(parent, text=" 📁 File Operations", padding=15)
        action_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Create button container
        button_container = ttk.Frame(action_frame)
        button_container.pack(fill=tk.X)
        
        # Upload button with icon
        self.upload_button = ttk.Button(button_container, text="📂 Upload Files", 
                                       command=self.upload_files)
        self.upload_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Analyze button with icon
        self.analyze_button = ttk.Button(button_container, text="🤖 Analyze Data", 
                                        command=self.analyze_data, style="Accent.TButton")
        self.analyze_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Export button with icon
        self.export_button = ttk.Button(button_container, text="💾 Export Results", 
                                       command=self.export_results)
        self.export_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Progress bar (initially hidden)
        self.progress_frame = ttk.Frame(action_frame)
        self.progress_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.progress_label = ttk.Label(self.progress_frame, text="")
        
        # File info display
        self.file_info_frame = ttk.Frame(action_frame)
        self.file_info_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.file_info_label = ttk.Label(self.file_info_frame, text="")
        self.file_info_label.pack(anchor=tk.W)
        
    def _create_status_bar(self, parent):
        """Create status bar with improved styling"""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready - Configure API key and upload files to begin")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                     relief=tk.SUNKEN, padding=(10, 5))
        self.status_label.pack(fill=tk.X)
        
    def _create_results_section(self, parent):
        """Create results section with improved tabbed interface"""
        results_frame = ttk.LabelFrame(parent, text=" 📊 Analysis Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook with better styling
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        self._create_summary_tab()
        self._create_deals_tab()
        self._create_insights_tab()
        self._create_data_tab()
        
    def _create_summary_tab(self):
        """Create executive summary tab with better formatting"""
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="📋 Executive Summary")
        
        # Add toolbar for summary tab
        summary_toolbar = ttk.Frame(self.summary_frame)
        summary_toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))
        
        ttk.Button(summary_toolbar, text="📋 Copy Summary", 
                  command=self.copy_summary).pack(side=tk.RIGHT)
        
        self.summary_text = scrolledtext.ScrolledText(
            self.summary_frame, wrap=tk.WORD, font=('Consolas', 10),
            bg='#f8f9fa', relief=tk.FLAT, borderwidth=1
        )
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def _create_deals_tab(self):
        """Create flagged deals tab with enhanced table"""
        self.deals_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.deals_frame, text="⚠️ Flagged Deals")
        
        # Add toolbar for deals tab
        deals_toolbar = ttk.Frame(self.deals_frame)
        deals_toolbar.pack(fill=tk.X, padx=5, pady=(5, 0))
        
        ttk.Label(deals_toolbar, text="Filter by risk:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.risk_filter_var = tk.StringVar(value="All")
        self.risk_filter_combo = ttk.Combobox(deals_toolbar, textvariable=self.risk_filter_var, 
                                             values=["All"], state="readonly", width=15)
        self.risk_filter_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.risk_filter_combo.bind('<<ComboboxSelected>>', self.filter_deals)
        
        # Create enhanced treeview
        tree_frame = ttk.Frame(self.deals_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        columns = ('Deal ID', 'Customer', 'Risk Type', 'Impact ($)', 'Confidence', 'Suggestion')
        self.deals_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=12)
        
        # Configure column widths and headings
        column_widths = {'Deal ID': 100, 'Customer': 150, 'Risk Type': 120, 
                        'Impact ($)': 100, 'Confidence': 80, 'Suggestion': 200}
        
        for col in columns:
            self.deals_tree.heading(col, text=col, command=lambda c=col: self.sort_deals(c))
            self.deals_tree.column(col, width=column_widths.get(col, 150))
        
        # Add scrollbars
        deals_scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.deals_tree.yview)
        deals_scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.deals_tree.xview)
        self.deals_tree.configure(yscrollcommand=deals_scrollbar_y.set, xscrollcommand=deals_scrollbar_x.set)
        
        self.deals_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        deals_scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        deals_scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Add context menu
        self.deals_tree.bind("<Button-3>", self.show_deal_context_menu)
        
    def _create_insights_tab(self):
        """Create insights tab for recommendations and trends"""
        self.insights_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.insights_frame, text="💡 Insights")
        
        # Create sections for different insights
        recommendations_frame = ttk.LabelFrame(self.insights_frame, text="Recommendations", padding=10)
        recommendations_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.recommendations_text = scrolledtext.ScrolledText(
            recommendations_frame, wrap=tk.WORD, height=8, font=('Segoe UI', 10)
        )
        self.recommendations_text.pack(fill=tk.BOTH, expand=True)
        
    def _create_data_tab(self):
        """Create raw data tab with better formatting"""
        self.data_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_frame, text="📄 Raw Data")
        
        # Add data statistics
        stats_frame = ttk.Frame(self.data_frame)
        stats_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.data_stats_label = ttk.Label(stats_frame, text="No data loaded")
        self.data_stats_label.pack(anchor=tk.W)
        
        self.data_text = scrolledtext.ScrolledText(
            self.data_frame, wrap=tk.WORD, font=('Consolas', 9),
            bg='#f8f9fa', relief=tk.FLAT, borderwidth=1
        )
        self.data_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def update_ui_state(self):
        """Update UI state based on current application state"""
        # Update button states
        self.analyze_button['state'] = 'normal' if (self.api_configured and self.parsed_data and not self.analysis_in_progress) else 'disabled'
        self.export_button['state'] = 'normal' if self.analysis_results else 'disabled'
        
        # Update API status
        if self.api_configured:
            self.api_status_label.config(text="✅ API configured", style="Success.TLabel")
        else:
            self.api_status_label.config(text="⚠️ API key required", style="Warning.TLabel")
            
    def show_progress(self, message):
        """Show progress bar with message"""
        self.progress_label.config(text=message)
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))
        self.progress_bar.pack(fill=tk.X)
        self.progress_bar.start(10)
        
    def hide_progress(self):
        """Hide progress bar"""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.progress_label.pack_forget()
        
    def set_api_key(self):
        """Set the API key for LLM interface with improved feedback"""
        api_key = self.api_key_var.get().strip()
        if api_key:
            self.llm_interface.api_key = api_key
            self.api_configured = True
            self.status_var.set("✅ API key configured successfully")
            self.update_ui_state()
            messagebox.showinfo("Success", "API key has been configured successfully!")
        else:
            messagebox.showwarning("Warning", "Please enter a valid API key")
    
    def upload_files(self):
        """Handle file upload with better progress indication"""
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
        
        self.show_progress("Parsing files...")
        self.root.update()
        
        all_data = []
        successful_files = 0
        failed_files = []
        
        for file_path in file_paths:
            try:
                file_data = self.data_parser.parse_file(file_path)
                all_data.extend(file_data)
                successful_files += 1
            except Exception as e:
                failed_files.append(f"{Path(file_path).name}: {str(e)}")
        
        self.hide_progress()
        
        if successful_files > 0:
            self.parsed_data = all_data
            self.display_raw_data()
            self.file_info_label.config(
                text=f"📁 {len(all_data)} deals loaded from {successful_files} file(s)"
            )
            self.status_var.set(f"✅ Successfully loaded {len(all_data)} deals")
            self.update_ui_state()
            
            # Show success message with details
            success_msg = f"Successfully loaded {len(all_data)} deals from {successful_files} file(s)"
            if failed_files:
                success_msg += f"\n\nFailed files:\n" + "\n".join(failed_files[:3])
                if len(failed_files) > 3:
                    success_msg += f"\n... and {len(failed_files) - 3} more"
            
            messagebox.showinfo("Upload Complete", success_msg)
        else:
            self.status_var.set("❌ No files could be loaded")
            if failed_files:
                messagebox.showerror("Upload Failed", 
                                   "All files failed to load:\n" + "\n".join(failed_files[:5]))
    
    def analyze_data(self):
        """Perform AI analysis with better progress indication"""
        if not self.parsed_data:
            messagebox.showwarning("Warning", "Please upload data files first")
            return
        
        if not self.api_configured:
            messagebox.showwarning("Warning", "Please configure your API key first")
            return
        
        self.analysis_in_progress = True
        self.update_ui_state()
        self.show_progress("Analyzing data with AI... This may take a few minutes")
        
        try:
            # Perform analysis
            self.analysis_results = self.llm_interface.analyze_deals(self.parsed_data)
            
            # Display results
            self.display_analysis_results()
            self.status_var.set("✅ Analysis completed successfully")
            
            # Switch to summary tab
            self.notebook.select(0)
            
        except Exception as e:
            messagebox.showerror("Analysis Error", f"Analysis failed:\n{str(e)}")
            self.status_var.set("❌ Analysis failed")
        finally:
            self.analysis_in_progress = False
            self.hide_progress()
            self.update_ui_state()
    
    def display_raw_data(self):
        """Display raw parsed data with better formatting"""
        self.data_text.delete(1.0, tk.END)
        
        if self.parsed_data:
            # Update statistics
            self.data_stats_label.config(text=f"📊 Dataset: {len(self.parsed_data)} deals loaded")
            
            # Create formatted view
            summary = f"DATASET OVERVIEW\n{'='*50}\n"
            summary += f"Total Deals: {len(self.parsed_data)}\n"
            summary += f"Loaded: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
            
            # Show field analysis
            if self.parsed_data:
                all_fields = set()
                for deal in self.parsed_data:
                    all_fields.update(deal.keys())
                
                summary += f"FIELDS DETECTED ({len(all_fields)}):\n{'-'*30}\n"
                for field in sorted(all_fields):
                    summary += f"• {field}\n"
                
                summary += f"\nSAMPLE RECORDS:\n{'-'*30}\n"
                
                # Show first few records with better formatting
                for i, deal in enumerate(self.parsed_data[:3]):
                    summary += f"\n[Record {i+1}]\n"
                    for key, value in deal.items():
                        summary += f"  {key}: {value}\n"
                
                if len(self.parsed_data) > 3:
                    summary += f"\n... and {len(self.parsed_data) - 3} more records\n"
            
            self.data_text.insert(tk.END, summary)
    
    def display_analysis_results(self):
        """Display analysis results with enhanced formatting"""
        if not self.analysis_results:
            return
        
        summary = self.analysis_results.get('summary', {})
        flagged_deals = self.analysis_results.get('flagged_deals', [])
        recommendations = self.analysis_results.get('recommendations', [])
        
        # Update summary tab with rich formatting
        self.summary_text.delete(1.0, tk.END)
        
        summary_text = f"""REVENUE LEAKAGE & MARGIN ANALYSIS REPORT
{'='*60}
Generated: {datetime.now().strftime(DATETIME_FORMAT)}

🎯 EXECUTIVE SUMMARY
{'-'*40}
💰 Total Estimated Revenue Leakage: {format_currency(summary.get('total_leakage', 0))}
⚠️  High Risk Deals Identified: {summary.get('high_risk_deals', 0)}
🔍 Total Issues Found: {summary.get('issues_found', 0)}
📊 Deals Analyzed: {len(self.parsed_data)}

🚨 RISK BREAKDOWN
{'-'*40}"""
        
        # Group flagged deals by risk type
        risk_types = {}
        for deal in flagged_deals:
            risk_type = deal.get('risk_type', 'unknown')
            if risk_type not in risk_types:
                risk_types[risk_type] = {'count': 0, 'impact': 0}
            risk_types[risk_type]['count'] += 1
            risk_types[risk_type]['impact'] += deal.get('impact', 0)
        
        for risk_type, data in risk_types.items():
            icon = self._get_risk_icon(risk_type)
            summary_text += f"{icon} {risk_type.replace('_', ' ').title()}: {data['count']} deals, {format_currency(data['impact'])} impact\n"
        
        summary_text += f"\n💡 KEY RECOMMENDATIONS\n{'-'*40}\n"
        for i, rec in enumerate(recommendations, 1):
            summary_text += f"{i}. {rec}\n"
        
        self.summary_text.insert(tk.END, summary_text)
        
        # Update flagged deals tab
        self._update_deals_tree(flagged_deals)
        
        # Update insights tab
        self._update_insights_tab(recommendations)
        
        # Update risk filter options
        risk_options = ["All"] + list(risk_types.keys())
        self.risk_filter_combo['values'] = risk_options
    
    def _get_risk_icon(self, risk_type):
        """Get appropriate icon for risk type"""
        icons = {
            'pricing_anomaly': '💰',
            'margin_compression': '📉',
            'unusual_discount': '🔥',
            'payment_terms': '⏰',
            'contract_risk': '📄',
            'default': '⚠️'
        }
        return icons.get(risk_type, icons['default'])
    
    def _update_deals_tree(self, flagged_deals):
        """Update deals tree with enhanced data"""
        # Clear existing items
        for item in self.deals_tree.get_children():
            self.deals_tree.delete(item)
        
        # Add flagged deals with confidence scores
        for deal in flagged_deals:
            confidence = deal.get('confidence', 0.5)
            confidence_text = f"{confidence*100:.0f}%"
            
            # Add color coding based on impact
            impact = deal.get('impact', 0)
            tags = []
            if impact > 50000:
                tags.append('high_impact')
            elif impact > 10000:
                tags.append('medium_impact')
            else:
                tags.append('low_impact')
            
            self.deals_tree.insert('', tk.END, values=(
                deal.get('deal_id', 'N/A'),
                deal.get('customer_name', 'N/A'),
                deal.get('risk_type', '').replace('_', ' ').title(),
                format_currency(impact),
                confidence_text,
                deal.get('suggestion', 'N/A')[:50] + "..." if len(deal.get('suggestion', '')) > 50 else deal.get('suggestion', 'N/A')
            ), tags=tags)
        
        # Configure tag colors
        self.deals_tree.tag_configure('high_impact', background='#ffebee')
        self.deals_tree.tag_configure('medium_impact', background='#fff3e0')
        self.deals_tree.tag_configure('low_impact', background='#f1f8e9')
    
    def _update_insights_tab(self, recommendations):
        """Update insights tab with recommendations"""
        self.recommendations_text.delete(1.0, tk.END)
        
        insights_text = "ACTIONABLE RECOMMENDATIONS\n" + "="*50 + "\n\n"
        
        for i, rec in enumerate(recommendations, 1):
            insights_text += f"💡 Recommendation #{i}\n"
            insights_text += f"{'-'*30}\n"
            insights_text += f"{rec}\n\n"
        
        self.recommendations_text.insert(tk.END, insights_text)
    
    def copy_summary(self):
        """Copy summary to clipboard"""
        summary_content = self.summary_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(summary_content)
        self.status_var.set("📋 Summary copied to clipboard")
    
    def filter_deals(self, event=None):
        """Filter deals by risk type"""
        selected_risk = self.risk_filter_var.get()
        
        # Clear current items
        for item in self.deals_tree.get_children():
            self.deals_tree.delete(item)
        
        # Get flagged deals from analysis results
        flagged_deals = self.analysis_results.get('flagged_deals', [])
        
        # Filter and display deals
        for deal in flagged_deals:
            if selected_risk == "All" or deal.get('risk_type', '') == selected_risk:
                confidence = deal.get('confidence', 0.5)
                confidence_text = f"{confidence*100:.0f}%"
                
                impact = deal.get('impact', 0)
                tags = []
                if impact > 50000:
                    tags.append('high_impact')
                elif impact > 10000:
                    tags.append('medium_impact')
                else:
                    tags.append('low_impact')
                
                suggestion = deal.get('suggestion', 'N/A')
                if len(suggestion) > 50:
                    suggestion = suggestion[:50] + "..."
                
                self.deals_tree.insert('', tk.END, values=(
                    deal.get('deal_id', 'N/A'),
                    deal.get('customer_name', 'N/A'),
                    deal.get('risk_type', '').replace('_', ' ').title(),
                    format_currency(impact),
                    confidence_text,
                    suggestion
                ), tags=tags)
    
    def sort_deals(self, column):
        """Sort deals by column"""
        # Implementation for sorting deals
        pass
    
    def show_deal_context_menu(self, event):
        """Show context menu for deal details"""
        # Implementation for context menu
        pass
    
    def export_results(self):
        """Export results with enhanced options"""
        if not self.analysis_results:
            messagebox.showwarning("Warning", "No analysis results to export")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Save analysis results",
            defaultextension=".csv",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
                ("JSON files", "*.json")
            ]
        )
        
        if not file_path:
            return
        
        try:
            self.show_progress("Exporting results...")
            
            flagged_deals = self.analysis_results.get('flagged_deals', [])
            summary = self.analysis_results.get('summary', {})
            
            if file_path.endswith('.csv'):
                self._export_csv(file_path, flagged_deals, summary)
            elif file_path.endswith('.xlsx'):
                self._export_excel(file_path, flagged_deals, summary)
            elif file_path.endswith('.json'):
                self._export_json(file_path, self.analysis_results)
            
            self.hide_progress()
            self.status_var.set(f"✅ Results exported to {Path(file_path).name}")
            messagebox.showinfo("Export Complete", f"Results exported successfully to:\n{file_path}")
            
        except Exception as e:
            self.hide_progress()
            messagebox.showerror("Export Error", f"Export failed:\n{str(e)}")
    
    def _export_csv(self, file_path, flagged_deals, summary):
        """Export to CSV format"""
        df = pd.DataFrame(flagged_deals)
        
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            f.write(f"{EXPORT_COMMENT_PREFIX} Revenue Leakage Analysis Report\n")
            f.write(f"{EXPORT_COMMENT_PREFIX} Generated: {datetime.now().strftime(DATETIME_FORMAT)}\n")
            f.write(f"{EXPORT_COMMENT_PREFIX} Total Leakage: {format_currency(summary.get('total_leakage', 0))}\n")
            f.write(f"{EXPORT_COMMENT_PREFIX} High Risk Deals: {summary.get('high_risk_deals', 0)}\n")
            f.write(f"{EXPORT_COMMENT_PREFIX}\n")
            
            df.to_csv(f, index=False)
    
    def _export_excel(self, file_path, flagged_deals, summary):
        """Export to Excel format with multiple sheets"""
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Flagged deals sheet
            df_deals = pd.DataFrame(flagged_deals)
            df_deals.to_excel(writer, sheet_name='Flagged Deals', index=False)
            
            # Summary sheet
            summary_data = {
                'Metric': ['Total Leakage', 'High Risk Deals', 'Issues Found', 'Total Deals Analyzed'],
                'Value': [summary.get('total_leakage', 0), 
                         summary.get('high_risk_deals', 0),
                         summary.get('issues_found', 0),
                         len(self.parsed_data)]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
    
    def _export_json(self, file_path, analysis_results):
        """Export to JSON format"""
        import json
        
        # Add metadata
        export_data = {
            'metadata': {
                'generated': datetime.now().strftime(DATETIME_FORMAT),
                'total_deals_analyzed': len(self.parsed_data),
                'export_version': '1.0'
            },
            'analysis_results': analysis_results
        }
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, indent=2, default=str)