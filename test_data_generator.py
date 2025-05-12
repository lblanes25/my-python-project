import pandas as pd
import numpy as np
import datetime
import os
import yaml

# Ensure output directory exists
if not os.path.exists('test_data'):
    os.makedirs('test_data')


# Create reference data for HR titles
def create_hr_reference_data():
    """Create HR reference data with employee titles"""

    employees = [
        {'Employee_Name': 'John Doe', 'Title': 'Audit Team Lead'},
        {'Employee_Name': 'Jane Smith', 'Title': 'Audit Leader'},
        {'Employee_Name': 'Alice Johnson', 'Title': 'Auditor'},
        {'Employee_Name': 'Bob Brown', 'Title': 'Audit Leader'},
        {'Employee_Name': 'Charlie Davis', 'Title': 'Audit Team Lead'},
        {'Employee_Name': 'Diana Evans', 'Title': 'Executive Auditor'},
        {'Employee_Name': 'Edward Franklin', 'Title': 'Audit Manager'},
        {'Employee_Name': 'Fiona Garcia', 'Title': 'Auditor'},
        {'Employee_Name': 'George Harris', 'Title': 'Analyst'},
        {'Employee_Name': 'Hannah Ingram', 'Title': 'Audit Team Lead'},
        {'Employee_Name': 'Ian Jackson', 'Title': 'Audit Leader'},
    ]

    hr_df = pd.DataFrame(employees)
    os.makedirs('ref_data', exist_ok=True)
    hr_df.to_excel('ref_data/hr_titles.xlsx', index=False)
    print(f"Created HR reference data with {len(employees)} employees")

    return hr_df


# Create test data for QA-ID 77
def create_test_data(hr_data):
    """Create test data with a mix of compliant and non-compliant records"""

    # Extract employees by role for use in our test data
    auditors = hr_data[hr_data['Title'] == 'Auditor']['Employee_Name'].tolist()
    team_leads = hr_data[hr_data['Title'] == 'Audit Team Lead']['Employee_Name'].tolist()
    audit_leaders = hr_data[hr_data['Title'] == 'Audit Leader']['Employee_Name'].tolist()

    # Base date for our test data
    base_date = datetime.datetime(2025, 5, 1)

    # Create records
    records = []

    # Create 50 records with a mix of compliant and non-compliant data
    for i in range(1, 51):
        # Basic workpaper information
        record = {
            'Audit TW ID': f'WP-2025-{i:03d}',
        }

        # Determine which test case this record represents
        test_case = i % 10

        # CASE 1-4: Fully compliant records
        if test_case in [1, 2, 3, 4]:
            record['TW submitter'] = np.random.choice(auditors)
            record['TL approver'] = np.random.choice(team_leads)
            record['AL approver'] = np.random.choice(audit_leaders)

            # Sequential dates
            submit_date = base_date + datetime.timedelta(days=i)
            tl_date = submit_date + datetime.timedelta(days=np.random.randint(1, 5))
            al_date = tl_date + datetime.timedelta(days=np.random.randint(1, 5))

            record['Submit Date'] = submit_date
            record['TL Approval Date'] = tl_date
            record['AL Approval Date'] = al_date

        # CASE 5: Segregation of duties violation - submitter is also TL approver
        elif test_case == 5:
            person = np.random.choice(team_leads)
            record['TW submitter'] = person
            record['TL approver'] = person  # Violation: Same person
            record['AL approver'] = np.random.choice(audit_leaders)

            # Sequential dates
            submit_date = base_date + datetime.timedelta(days=i)
            tl_date = submit_date + datetime.timedelta(days=np.random.randint(1, 5))
            al_date = tl_date + datetime.timedelta(days=np.random.randint(1, 5))

            record['Submit Date'] = submit_date
            record['TL Approval Date'] = tl_date
            record['AL Approval Date'] = al_date

        # CASE 6: Segregation of duties violation - submitter is also AL approver
        elif test_case == 6:
            person = np.random.choice(audit_leaders)
            record['TW submitter'] = person
            record['TL approver'] = np.random.choice(team_leads)
            record['AL approver'] = person  # Violation: Same person

            # Sequential dates
            submit_date = base_date + datetime.timedelta(days=i)
            tl_date = submit_date + datetime.timedelta(days=np.random.randint(1, 5))
            al_date = tl_date + datetime.timedelta(days=np.random.randint(1, 5))

            record['Submit Date'] = submit_date
            record['TL Approval Date'] = tl_date
            record['AL Approval Date'] = al_date

        # CASE 7: Approval sequence violation - TL approved before submission
        elif test_case == 7:
            record['TW submitter'] = np.random.choice(auditors)
            record['TL approver'] = np.random.choice(team_leads)
            record['AL approver'] = np.random.choice(audit_leaders)

            # Non-sequential dates
            submit_date = base_date + datetime.timedelta(days=i)
            tl_date = submit_date - datetime.timedelta(days=np.random.randint(1, 5))  # TL before submit
            al_date = submit_date + datetime.timedelta(days=np.random.randint(1, 5))

            record['Submit Date'] = submit_date
            record['TL Approval Date'] = tl_date
            record['AL Approval Date'] = al_date

        # CASE 8: Approval sequence violation - AL approved before TL
        elif test_case == 8:
            record['TW submitter'] = np.random.choice(auditors)
            record['TL approver'] = np.random.choice(team_leads)
            record['AL approver'] = np.random.choice(audit_leaders)

            # Non-sequential dates
            submit_date = base_date + datetime.timedelta(days=i)
            al_date = submit_date + datetime.timedelta(days=np.random.randint(1, 3))
            tl_date = al_date + datetime.timedelta(days=np.random.randint(1, 3))  # TL after AL

            record['Submit Date'] = submit_date
            record['TL Approval Date'] = tl_date
            record['AL Approval Date'] = al_date

        # CASE 9: Multiple violations - segregation and sequence
        elif test_case == 9:
            person = np.random.choice(team_leads)
            record['TW submitter'] = person
            record['TL approver'] = person  # Violation: Same person
            record['AL approver'] = np.random.choice(audit_leaders)

            # Non-sequential dates
            submit_date = base_date + datetime.timedelta(days=i)
            al_date = submit_date + datetime.timedelta(days=np.random.randint(1, 3))
            tl_date = al_date + datetime.timedelta(days=np.random.randint(1, 3))  # TL after AL

            record['Submit Date'] = submit_date
            record['TL Approval Date'] = tl_date
            record['AL Approval Date'] = al_date

        # CASE 0: Missing data
        else:
            record['TW submitter'] = np.random.choice(auditors)
            record['TL approver'] = np.random.choice(team_leads)
            record['AL approver'] = None  # Missing approver

            # Missing dates
            submit_date = base_date + datetime.timedelta(days=i)
            record['Submit Date'] = submit_date
            record['TL Approval Date'] = submit_date + datetime.timedelta(days=np.random.randint(1, 5))
            record['AL Approval Date'] = None  # Missing date

        records.append(record)

    # Convert to DataFrame and write to Excel
    df = pd.DataFrame(records)
    output_path = 'test_data/qa_77_test_data.xlsx'
    df.to_excel(output_path, index=False)
    print(f"Created test data with {len(records)} records at {output_path}")

    # Display sample of the data
    print("\nSample of test data:")
    print(df.head())

    return df


# Create config directory if it doesn't exist
os.makedirs('configs', exist_ok=True)


# Create the enhanced config file
def create_enhanced_config_file():
    """Create the enhanced config file for QA-ID 77"""

    config = {
        'analytic_id': 77,
        'analytic_name': 'Audit Test Workpaper Approvals',
        'analytic_description': 'This analytic evaluates workpaper approvals to ensure proper segregation of duties, correct approval sequences, and appropriate approval authority based on job titles. It helps identify process violations that could impact audit quality.',

        'source': {
            'file_type': 'xlsx',
            'required_columns': [
                {'name': 'Audit TW ID', 'alias': ['TW_ID', 'Workpaper ID']},
                {'name': 'TW submitter', 'alias': ['Submitter', 'Prepared By']},
                {'name': 'TL approver', 'alias': ['Team Lead', 'TL']},
                {'name': 'AL approver', 'alias': ['Audit Leader', 'AL']},
                {'name': 'Submit Date', 'alias': ['Submission Date', 'Date Submitted']},
                {'name': 'TL Approval Date', 'alias': ['TL Date']},
                {'name': 'AL Approval Date', 'alias': ['AL Date']}
            ]
        },

        'reference_files': [
            {
                'name': 'HR_Titles',
                'path': 'ref_data/hr_titles.xlsx',
                'key_column': 'Employee_Name',
                'value_column': 'Title'
            }
        ],

        'validations': [
            {
                'rule': 'segregation_of_duties',
                'description': 'Submitter cannot be TL or AL',
                'rationale': 'Ensures independent review by preventing the submitter from also being an approver, which maintains the integrity of the review process.',
                'parameters': {
                    'submitter_field': 'TW submitter',
                    'approver_fields': ['TL approver', 'AL approver']
                }
            },
            {
                'rule': 'approval_sequence',
                'description': 'Approvals must be in order - Submit -> TL -> AL',
                'rationale': 'Maintains proper workflow sequence to ensure the Team Lead reviews before the Audit Leader and that no approvals happen before submission.',
                'parameters': {
                    'date_fields_in_order': ['Submit Date', 'TL Approval Date', 'AL Approval Date']
                }
            },
            {
                'rule': 'title_based_approval',
                'description': 'AL must have appropriate title',
                'rationale': 'Ensures approval authority is limited to those with appropriate job titles, maintaining organizational hierarchy and approval standards.',
                'parameters': {
                    'approver_field': 'AL approver',
                    'allowed_titles': ['Audit Leader', 'Executive Auditor', 'Audit Manager'],
                    'title_reference': 'HR_Titles'
                }
            }
        ],

        'thresholds': {
            'error_percentage': 5.0,
            'rationale': 'Industry standard for audit workpapers allows for up to 5% error rate. Higher rates require remediation and process review.'
        },

        'reporting': {
            'group_by': 'AL approver',
            'summary_fields': ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage'],
            'detail_required': True
        },

        'report_metadata': {
            'owner': 'Quality Assurance Team',
            'review_frequency': 'Monthly',
            'last_revised': '2025-05-01',
            'version': '1.0',
            'contact_email': 'qa_analytics@example.com',
            'regulatory_requirement': 'SOX 404',
            'business_impact': 'High',
            'notes': 'This analytic is part of the core audit quality monitoring program and results should be reviewed monthly by the QA Director.'
        }
    }

    config_path = 'configs/qa_77.yaml'
    with open(config_path, 'w', encoding='utf-8') as f:
        yaml.dump(config, f, default_flow_style=False, sort_keys=False)

    print(f"Created enhanced config file at {config_path}")


# Run the functions to create all test files
if __name__ == "__main__":
    print("Generating enhanced test data for QA-ID 77...")
    hr_data = create_hr_reference_data()
    test_data = create_test_data(hr_data)
    create_enhanced_config_file()
    print("\nEnhanced test data generation complete. You can now run the QA Analytics app and test with QA-ID 77.")
    print("Test data file: test_data/qa_77_test_data.xlsx")
    print("Reference data: ref_data/hr_titles.xlsx")
    print("Enhanced config file: configs/qa_77.yaml")