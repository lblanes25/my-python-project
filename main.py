import os
import sys
import logging
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List
from config_manager import ConfigManager
from data_processor import DataProcessor
from report_generator import ReportGenerator
from logging_config import setup_logging

logger = setup_logging()

class QAAnalyticsApp:
    """Main application with GUI interface"""

    def __init__(self, root):
        """Initialize the application"""
        self.root = root
        self.root.title("QA Analytics Automation")
        self.root.geometry("800x600")

        # Load configuration
        self.config_manager = ConfigManager()
        self.available_analytics = self.config_manager.get_available_analytics()

        # Set up UI components
        self._setup_ui()

    def _setup_ui(self):
        """Set up the user interface"""
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Analytics selection
        ttk.Label(main_frame, text="Select QA-ID:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        self.analytic_var = tk.StringVar()
        self.analytic_combo = ttk.Combobox(main_frame, textvariable=self.analytic_var, state="readonly", width=50)
        self.analytic_combo["values"] = [f"{id} - {name}" for id, name in self.available_analytics]
        if self.available_analytics:
            self.analytic_combo.current(0)
        self.analytic_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

        # Source file selection
        ttk.Label(main_frame, text="Source Data File:").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))

        source_frame = ttk.Frame(main_frame)
        source_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

        self.source_var = tk.StringVar()
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_var, width=50)
        self.source_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        source_btn = ttk.Button(source_frame, text="Browse...", command=self._browse_source)
        source_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Output directory
        ttk.Label(main_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))

        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(0, 5))

        self.output_var = tk.StringVar(value="output")
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=50)
        self.output_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        output_btn = ttk.Button(output_frame, text="Browse...", command=self._browse_output)
        output_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Execution frame
        exec_frame = ttk.Frame(main_frame)
        exec_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        self.progress = ttk.Progressbar(exec_frame, orient="horizontal", length=200, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        exec_btn = ttk.Button(exec_frame, text="Run Analysis", command=self._run_analysis)
        exec_btn.pack(side=tk.RIGHT)

        # Status log
        ttk.Label(main_frame, text="Status Log:").grid(row=4, column=0, sticky=tk.W, pady=(10, 5))

        self.log_text = tk.Text(main_frame, height=15, width=80, wrap=tk.WORD)
        self.log_text.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)

        # Add scrollbar to log
        log_scroll = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=5, column=2, sticky=(tk.N, tk.S))
        self.log_text.config(yscrollcommand=log_scroll.set)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Configure resizing
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

        # Set up log handler
        self._setup_log_handler()

    def _setup_log_handler(self):
        """Set up log handler to redirect to text widget"""

        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                logging.Handler.__init__(self)
                self.text_widget = text_widget

            def emit(self, record):
                msg = self.format(record)

                def append():
                    self.text_widget.config(state=tk.NORMAL)
                    self.text_widget.insert(tk.END, msg + "\n")
                    self.text_widget.see(tk.END)
                    self.text_widget.config(state=tk.DISABLED)

                # Schedule to be executed in the main thread
                self.text_widget.after(0, append)

        # Create a handler and add it to the logger
        text_handler = TextHandler(self.log_text)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S')
        text_handler.setFormatter(formatter)
        logger.addHandler(text_handler)

    def _browse_source(self):
        """Browse for source data file"""
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")],
            title="Select Source Data File"
        )
        if filename:
            self.source_var.set(filename)

    def _browse_output(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_var.set(directory)

    def _run_analysis(self):
        """Run the analysis process"""
        # Validate inputs
        if not self.analytic_var.get():
            messagebox.showerror("Error", "Please select a QA-ID")
            return

        if not self.source_var.get():
            messagebox.showerror("Error", "Please select a source data file")
            return

        if not os.path.exists(self.source_var.get()):
            messagebox.showerror("Error", "Source data file does not exist")
            return

        # Get the analytic ID from selection
        analytic_id = self.analytic_var.get().split(" - ")[0]

        # Start progress bar
        self.progress.start()
        self.status_var.set("Processing...")

        # Run in a separate thread to avoid freezing the UI
        threading.Thread(target=self._process_data, args=(analytic_id,), daemon=True).start()

    def _process_data(self, analytic_id):
        """Process data in a separate thread"""
        try:
            # Get configuration
            config = self.config_manager.get_config(analytic_id)

            # Create output directory if needed
            output_dir = self.output_var.get()
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # Initialize processor
            processor = DataProcessor(config)

            # Process data
            logger.info(f"Starting processing for QA-ID {analytic_id}")
            success, message = processor.process_data(self.source_var.get())

            if not success:
                self.root.after(0, lambda: messagebox.showerror("Error", message))
                return

            # Generate reports
            logger.info("Generating reports...")
            report_generator = ReportGenerator(config, processor.results)

            # Generate main report
            main_report_path = os.path.join(
                output_dir,
                f"QA_{analytic_id}_Main_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            main_report = report_generator.generate_main_report(main_report_path)

            # Generate individual reports
            individual_reports = report_generator.generate_individual_reports()

            # Show completion message
            report_count = 1 + len(individual_reports)
            completion_msg = f"Processing complete. Generated {report_count} reports."
            logger.info(completion_msg)

            self.root.after(0, lambda: messagebox.showinfo("Success", completion_msg))
            self.root.after(0, lambda: self.status_var.set("Ready"))

        except Exception as e:
            logger.error(f"Error in processing: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {e}"))

        finally:
            # Stop progress bar
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self.status_var.set("Ready"))


# Application entry point
if __name__ == "__main__":
    root = tk.Tk()
    app = QAAnalyticsApp(root)
    root.mainloop()