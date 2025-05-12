import os
import sys
import logging
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Tuple
import datetime
from config_manager import ConfigManager
from data_processor import DataProcessor
from consolidated_report_generator import ConsolidatedReportGenerator
from logging_config import setup_logging

logger = setup_logging()


class QAAnalyticsApp:
    """Main application with GUI interface for consolidated analytics"""

    def __init__(self, root):
        """Initialize the application"""
        self.root = root
        self.root.title("QA Analytics Automation")
        self.root.geometry("800x700")

        # Load configuration
        self.config_manager = ConfigManager()
        self.available_analytics = self.config_manager.get_available_analytics()

        # Dictionary to store source files for selected analytics
        self.source_files = {}

        # Set up UI components
        self._setup_ui()

    def _setup_ui(self):
        """Set up the user interface"""
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Analytics selection (listbox with checkboxes)
        ttk.Label(main_frame, text="Select QA Analytics:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        analytics_frame = ttk.Frame(main_frame)
        analytics_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N), pady=(0, 10))

        # Create a frame for the listbox and scrollbar
        list_frame = ttk.Frame(analytics_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Listbox with checkboxes (using Treeview as a workaround)
        self.analytics_tree = ttk.Treeview(list_frame, columns=("ID", "Name", "Source"), show="headings", height=5)
        self.analytics_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure scrollbar
        scrollbar.config(command=self.analytics_tree.yview)
        self.analytics_tree.config(yscrollcommand=scrollbar.set)

        # Configure columns
        self.analytics_tree.column("ID", width=50, anchor=tk.CENTER)
        self.analytics_tree.column("Name", width=250)
        self.analytics_tree.column("Source", width=300)

        # Configure headings
        self.analytics_tree.heading("ID", text="QA-ID")
        self.analytics_tree.heading("Name", text="Analytic Name")
        self.analytics_tree.heading("Source", text="Source File")

        # Populate the listbox with available analytics
        for analytic_id, name in self.available_analytics:
            self.analytics_tree.insert("", tk.END, values=(analytic_id, name, "Click to select file"))

        # Bind double-click to select source file
        self.analytics_tree.bind("<Double-1>", self._select_source_file)

        # Output directory
        ttk.Label(main_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, pady=(10, 5))

        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(10, 5))

        self.output_var = tk.StringVar(value="output")
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=50)
        self.output_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        output_btn = ttk.Button(output_frame, text="Browse...", command=self._browse_output)
        output_btn.pack(side=tk.RIGHT, padx=(5, 0))

        # Consolidated reports checkbox
        self.consolidated_var = tk.BooleanVar(value=True)
        consolidated_check = ttk.Checkbutton(
            main_frame,
            text="Generate consolidated reports by Audit Leader",
            variable=self.consolidated_var
        )
        consolidated_check.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5, 10))

        # Execution frame
        exec_frame = ttk.Frame(main_frame)
        exec_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 10))

        self.progress = ttk.Progressbar(exec_frame, orient="horizontal", length=200, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        exec_btn = ttk.Button(exec_frame, text="Run Selected Analytics", command=self._run_analytics)
        exec_btn.pack(side=tk.RIGHT)

        # Status log
        ttk.Label(main_frame, text="Status Log:").grid(row=5, column=0, sticky=tk.W, pady=(5, 5))

        self.log_text = tk.Text(main_frame, height=15, width=80, wrap=tk.WORD)
        self.log_text.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)

        # Add scrollbar to log
        log_scroll = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=6, column=2, sticky=(tk.N, tk.S))
        self.log_text.config(yscrollcommand=log_scroll.set)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Configure resizing
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)

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

    def _select_source_file(self, event):
        """Handle double-click on treeview to select a source file"""
        # Get selected item
        selected_item = self.analytics_tree.focus()
        if not selected_item:
            return

        # Get the analytic ID from the selected item
        values = self.analytics_tree.item(selected_item, "values")
        if not values or len(values) < 2:
            return

        analytic_id = values[0]

        # Open file dialog to select source file
        filename = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")],
            title=f"Select Source Data File for QA-ID {analytic_id}"
        )

        if filename:
            # Store the selected file and update treeview
            self.source_files[analytic_id] = filename
            self.analytics_tree.item(selected_item, values=(analytic_id, values[1], filename))

    def _browse_output(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_var.set(directory)

    def _run_analytics(self):
        """Run the selected analytics"""
        # Get selected analytics
        selected_analytics = []
        selected_files = {}

        for item in self.analytics_tree.get_children():
            values = self.analytics_tree.item(item, "values")
            analytic_id = values[0]
            source_file = values[2]

            # Check if this analytic has a source file selected
            if source_file and source_file != "Click to select file":
                selected_analytics.append(analytic_id)
                selected_files[analytic_id] = source_file

        # Validate selections
        if not selected_analytics:
            messagebox.showerror("Error", "Please select at least one analytic and provide a source file")
            return

        for analytic_id, source_file in selected_files.items():
            if not os.path.exists(source_file):
                messagebox.showerror("Error", f"Source file for QA-ID {analytic_id} does not exist: {source_file}")
                return

        # Prepare output directory
        output_dir = self.output_var.get()
        if not output_dir:
            messagebox.showerror("Error", "Please specify an output directory")
            return

        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create output directory: {e}")
                return

        # Start progress bar
        self.progress.start()
        self.status_var.set("Processing...")

        # Get consolidated reports preference
        generate_consolidated = self.consolidated_var.get()

        # Run in a separate thread to avoid freezing the UI
        threading.Thread(
            target=self._process_analytics,
            args=(selected_analytics, selected_files, output_dir, generate_consolidated),
            daemon=True
        ).start()

    def _process_analytics(self, analytics_ids, source_files, output_dir, generate_consolidated):
        """Process selected analytics in a separate thread"""
        try:
            logger.info(f"Starting processing for {len(analytics_ids)} analytics: {', '.join(analytics_ids)}")

            # Track individually generated reports
            individual_reports = []
            main_report = None

            if not generate_consolidated:
                # Process each analytic individually
                for analytic_id in analytics_ids:
                    try:
                        # Get configuration
                        config = self.config_manager.get_config(analytic_id)

                        # Initialize processor
                        processor = DataProcessor(config)

                        # Process data
                        logger.info(f"Processing QA-ID {analytic_id}: {config['analytic_name']}")
                        success, message = processor.process_data(source_files[analytic_id])

                        if not success:
                            logger.error(f"Failed to process QA-ID {analytic_id}: {message}")
                            continue

                        # Generate report
                        from report_generator import ReportGenerator
                        report_generator = ReportGenerator(config, processor.results)

                        # Main report
                        main_report_path = os.path.join(
                            output_dir,
                            f"QA_{analytic_id}_Main_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        )

                        report_path = report_generator.generate_main_report(main_report_path)
                        if report_path:
                            individual_reports.append(report_path)

                        # Individual reports
                        leader_reports = report_generator.generate_individual_reports()
                        individual_reports.extend(leader_reports)

                        logger.info(f"Generated reports for QA-ID {analytic_id}")

                    except Exception as e:
                        logger.error(f"Error processing QA-ID {analytic_id}: {e}")
            else:
                # Use consolidated report generator
                consolidated_generator = ConsolidatedReportGenerator(output_dir=output_dir)

                # Run all selected analytics
                results_by_analytic = consolidated_generator.run_analytics(analytics_ids, source_files)

                if not results_by_analytic:
                    self.root.after(0, lambda: messagebox.showerror(
                        "Error", "Failed to generate results for any selected analytics"))
                    return

                # Generate consolidated reports
                reports_by_leader = consolidated_generator.generate_consolidated_reports(results_by_analytic)

                # Check for main report
                if "__MAIN_REPORT__" in reports_by_leader:
                    main_report = reports_by_leader.pop("__MAIN_REPORT__")

                # Add other reports to individual reports list
                for leader, report_path in reports_by_leader.items():
                    individual_reports.append(report_path)

            # Show completion message
            report_count = len(individual_reports)
            completion_msg = f"Processing complete. Generated {report_count} individual reports."

            if main_report:
                completion_msg += f"\n\nDepartment-level consolidated report generated:\n{main_report}"

            logger.info(completion_msg)

            # List generated reports in log
            if individual_reports:
                logger.info("\nGenerated individual reports:")
                for path in individual_reports:
                    logger.info(f"  {path}")

            if main_report:
                logger.info("\nGenerated department-level consolidated report:")
                logger.info(f"  {main_report}")
                logger.info("\nThis report contains a comprehensive department-wide view for QA evaluation.")

            self.root.after(0, lambda: messagebox.showinfo("Success", completion_msg))

        except Exception as e:
            logger.error(f"Error in processing: {e}")
            import traceback
            logger.error(traceback.format_exc())
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