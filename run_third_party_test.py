import os
import sys
from test_qa_analytic import TestQAAnalytic
from test_data_generator2 import create_third_party_test_data

# Ensure directories exist
if not os.path.exists('test_data'):
    os.makedirs('test_data')
if not os.path.exists('configs'):
    os.makedirs('configs')


def run_third_party_risk_test():
    """Run the third party risk assessment validation test"""

    # First generate test data if it doesn't exist
    if not os.path.exists('test_data/qa_78_test_data.xlsx'):
        print("Generating test data for third party risk assessment...")
        create_third_party_test_data()

    # Define analytics ID and source file
    analytic_id = "78"  # Third Party Risk Assessment Validation
    source_file = "test_data/qa_78_test_data.xlsx"

    print(f"Running test for QA-ID {analytic_id} using {source_file}")

    # Run the test using the TestQAAnalytic class
    test = TestQAAnalytic(analytic_id, source_file)
    success = test.run_test()

    if success:
        print("Test completed successfully. See log for details.")
    else:
        print("Test failed. See log for details.")

    return success


if __name__ == "__main__":
    # Ensure the config file exists - you'll need to create the qa_78.yaml file
    if not os.path.exists('configs/qa_78.yaml'):
        print("Configuration file 'configs/qa_78.yaml' not found.")
        print("Please create the configuration file first.")
        sys.exit(1)

    # Run the test
    run_third_party_risk_test()