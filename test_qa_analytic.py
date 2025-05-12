import os
import sys
import pandas as pd
from config_manager import ConfigManager
from data_processor import DataProcessor
from report_generator import ReportGenerator
from logging_config import setup_logging

# Set up logging
logger = setup_logging()


class TestQAAnalytic:
    """Test class to run a QA analytic without the GUI"""

    def __init__(self, analytic_id, source_file, output_dir="output"):
        """Initialize test with analytic ID and source file"""
        self.analytic_id = analytic_id
        self.source_file = source_file
        self.output_dir = output_dir

        # Ensure output directory exists
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        # Load configuration
        self.config_manager = ConfigManager()
        self.available_analytics = self.config_manager.get_available_analytics()

        logger.info(f"Available analytics: {self.available_analytics}")

    def run_test(self):
        """Run the QA analytic test"""
        try:
            # Get configuration
            config = self.config_manager.get_config(self.analytic_id)
            logger.info(f"Loaded config for QA-ID {self.analytic_id}: {config['analytic_name']}")

            # Initialize processor
            processor = DataProcessor(config)

            # Process data
            logger.info(f"Processing data from {self.source_file}")
            success, message = processor.process_data(self.source_file)

            if not success:
                logger.error(f"Processing failed: {message}")
                return False

            logger.info(f"Processing completed: {message}")

            # Generate reports
            logger.info("Generating reports...")
            report_generator = ReportGenerator(config, processor.results)

            # Generate main report
            import datetime
            main_report_path = os.path.join(
                self.output_dir,
                f"QA_{self.analytic_id}_Main_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            main_report = report_generator.generate_main_report(main_report_path)

            # Generate individual reports
            individual_reports = report_generator.generate_individual_reports()

            # Show completion message
            report_count = 1 + len(individual_reports)
            completion_msg = f"Processing complete. Generated {report_count} reports."
            logger.info(completion_msg)

            # Analyze results
            self._analyze_results(processor.results)

            return True

        except Exception as e:
            logger.error(f"Error in test: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def _analyze_results(self, results):
        """Analyze and print the test results"""
        if 'summary' not in results or 'detail' not in results:
            logger.error("Results missing summary or detail data")
            return

        summary = results['summary']
        detail = results['detail']

        # Overall statistics
        total_records = len(detail)
        gc_count = sum(detail['Compliance'] == 'GC')
        dnc_count = sum(detail['Compliance'] == 'DNC')
        pc_count = sum(detail['Compliance'] == 'PC')

        logger.info("\n" + "=" * 80)
        logger.info(f"TEST RESULTS SUMMARY FOR QA-ID {self.analytic_id}")
        logger.info("=" * 80)
        logger.info(f"Total records processed: {total_records}")
        logger.info(f"Generally Conforms (GC): {gc_count} ({gc_count / total_records * 100:.1f}%)")
        logger.info(f"Does Not Conform (DNC): {dnc_count} ({dnc_count / total_records * 100:.1f}%)")
        logger.info(f"Partially Conforms (PC): {pc_count} ({pc_count / total_records * 100:.1f}%)")
        logger.info("=" * 80)

        # Check validation columns
        validation_columns = [col for col in detail.columns if col.startswith('Valid_')]

        if validation_columns:
            logger.info("\nValidation Rule Results:")
            for col in validation_columns:
                rule_name = col.replace('Valid_', '')
                pass_count = sum(detail[col])
                fail_count = sum(~detail[col])
                logger.info(f"  {rule_name}: {pass_count} passed, {fail_count} failed")

        # Group by summary
        if not summary.empty:
            logger.info("\nSummary by Group:")
            pd.set_option('display.width', 120)
            pd.set_option('display.max_columns', None)
            logger.info("\n" + str(summary))

        # List the output reports
        logger.info("\nGenerated Reports:")
        for file in os.listdir(self.output_dir):
            if file.startswith(f"QA_{self.analytic_id}") and file.endswith(".xlsx"):
                logger.info(f"  {os.path.join(self.output_dir, file)}")

        logger.info("=" * 80)


if __name__ == "__main__":
    # First check that we generated test data
    if not os.path.exists('test_data/qa_77_test_data.xlsx'):
        print("Test data not found. Run the test_data_generator.py script first.")
        sys.exit(1)

    # Run the test
    analytic_id = "77"
    source_file = "test_data/qa_77_test_data.xlsx"

    print(f"Running test for QA-ID {analytic_id} using {source_file}")
    test = TestQAAnalytic(analytic_id, source_file)
    success = test.run_test()

    if success:
        print("Test completed successfully. See log for details.")
    else:
        print("Test failed. See log for details.")