import pandas as pd
import numpy as np
import datetime
import os

# Ensure test data directory exists
if not os.path.exists('test_data'):
    os.makedirs('test_data')


def create_third_party_test_data():
    """Create test data for QA-ID 78 - Third Party Risk Assessment Validation"""

    # Define audit leaders
    audit_leaders = [
        "Maria Johnson",
        "James Smith",
        "Sarah Brown",
        "David Wilson",
        "Emily Davis"
    ]

    # Define risk levels
    risk_levels = ["Critical", "High", "Medium", "Low", "N/A"]

    # Create records
    records = []

    # Generate 50 test records
    for i in range(1, 51):
        # Determine which test case this record represents
        test_case = i % 10

        record = {
            'Audit Entity ID': f'AE-2025-{i:03d}',
            'Audit Name': f'Audit Entity {i}',
            'Audit Leader': np.random.choice(audit_leaders)
        }

        # CASE 1-3: Fully compliant records with third parties and risk assigned
        if test_case in [1, 2, 3]:
            # Generate 1-5 third parties
            num_third_parties = np.random.randint(1, 6)
            third_parties = ', '.join([f'TP-{np.random.randint(1, 100)}' for _ in range(num_third_parties)])

            record['Third Parties'] = third_parties
            record['L1 Third Party Risk'] = np.random.choice(risk_levels[:-1])  # Exclude N/A

        # CASE 4-5: Audit with no third parties and N/A risk (compliant)
        elif test_case in [4, 5]:
            record['Third Parties'] = ""  # No third parties
            record['L1 Third Party Risk'] = "N/A"  # Appropriate when no third parties

        # CASE 6: Non-compliant - third parties exist but risk is N/A
        elif test_case == 6:
            # Generate 1-5 third parties
            num_third_parties = np.random.randint(1, 6)
            third_parties = ', '.join([f'TP-{np.random.randint(1, 100)}' for _ in range(num_third_parties)])

            record['Third Parties'] = third_parties
            record['L1 Third Party Risk'] = "N/A"  # This is the violation

        # CASE 7: Non-compliant - third parties exist but risk field is empty
        elif test_case == 7:
            # Generate 1-5 third parties
            num_third_parties = np.random.randint(1, 6)
            third_parties = ', '.join([f'TP-{np.random.randint(1, 100)}' for _ in range(num_third_parties)])

            record['Third Parties'] = third_parties
            record['L1 Third Party Risk'] = ""  # This is the violation - empty risk

        # CASE 8: Non-compliant - many third parties (10-30) but risk is N/A
        elif test_case == 8:
            # Generate 10-30 third parties
            num_third_parties = np.random.randint(10, 31)
            third_parties = ', '.join([f'TP-{np.random.randint(1, 100)}' for _ in range(num_third_parties)])

            record['Third Parties'] = third_parties
            record['L1 Third Party Risk'] = "N/A"  # This is the violation

        # CASE 9-0: Other compliant records - various risk levels
        else:
            # Generate 1-5 third parties
            num_third_parties = np.random.randint(1, 6)
            third_parties = ', '.join([f'TP-{np.random.randint(1, 100)}' for _ in range(num_third_parties)])

            record['Third Parties'] = third_parties
            record['L1 Third Party Risk'] = np.random.choice(risk_levels[:-1])  # Exclude N/A

        records.append(record)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(records)
    output_path = 'test_data/qa_78_test_data.xlsx'
    df.to_excel(output_path, index=False)
    print(f"Created third party test data with {len(records)} records at {output_path}")

    # Display sample of the data
    print("\nSample of test data:")
    print(df.head())

    return df


if __name__ == "__main__":
    print("Generating test data for QA-ID 78 - Third Party Risk Assessment Validation...")
    create_third_party_test_data()
    print("\nTest data generation complete.")
    print("Test data file: test_data/qa_78_test_data.xlsx")