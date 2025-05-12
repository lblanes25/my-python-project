import os
import yaml
from typing import Dict, List, Tuple
from logging_config import setup_logging

logger = setup_logging()


class ConfigManager:
    """Manages loading and validation of configuration files"""

    def __init__(self, config_dir: str = "configs"):
        """Initialize config manager with directory of config files"""
        self.config_dir = config_dir
        self.configs = {}
        self.load_all_configs()

    def load_all_configs(self) -> None:
        """Load all configuration files from the config directory"""
        try:
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir)
                self._create_sample_config()

            for filename in os.listdir(self.config_dir):
                if filename.endswith(('.yaml', '.yml')):
                    config_path = os.path.join(self.config_dir, filename)
                    try:
                        with open(config_path, 'r', encoding='utf-8') as file:
                            config = yaml.safe_load(file)
                            if self._validate_config(config):
                                analytic_id = str(config.get('analytic_id'))
                                self.configs[analytic_id] = config
                                logger.info(f"Loaded config for QA-ID {analytic_id}")
                    except Exception as e:
                        logger.error(f"Error loading config {filename}: {e}")
        except Exception as e:
            logger.error(f"Error accessing config directory: {e}")

    def _validate_config(self, config: Dict) -> bool:
        """Validate that a configuration has all required elements"""
        required_keys = ['analytic_id', 'analytic_name', 'source', 'validations', 'thresholds', 'reporting']

        # Check required top-level keys
        for key in required_keys:
            if key not in config:
                logger.error(f"Missing required config key: {key}")
                return False

        # Check source configuration
        if 'required_columns' not in config['source']:
            logger.error("Source config missing required_columns")
            return False

        return True

    def _create_sample_config(self) -> None:
        """Create a sample configuration file with enhanced fields"""
        sample_config = {
            'analytic_id': 77,
            'analytic_name': 'Audit Test Workpaper Approvals',
            'analytic_description': 'This analytic evaluates workpaper approvals to ensure proper segregation of duties, correct approval sequences, and appropriate approval authority based on job titles.',
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
                    'rationale': 'Ensures independent review by preventing the submitter from also being an approver.',
                    'parameters': {
                        'submitter_field': 'TW submitter',
                        'approver_fields': ['TL approver', 'AL approver']
                    }
                },
                {
                    'rule': 'approval_sequence',
                    'description': 'Approvals must be in order: Submit -> TL -> AL',
                    'rationale': 'Maintains proper workflow sequence to ensure the Team Lead reviews before the Audit Leader.',
                    'parameters': {
                        'date_fields_in_order': ['Submit Date', 'TL Approval Date', 'AL Approval Date']
                    }
                },
                {
                    'rule': 'title_based_approval',
                    'description': 'AL must have appropriate title',
                    'rationale': 'Ensures approval authority is limited to those with appropriate job titles.',
                    'parameters': {
                        'approver_field': 'AL approver',
                        'allowed_titles': ['Audit Leader', 'Executive Auditor', 'Audit Manager'],
                        'title_reference': 'HR_Titles'
                    }
                }
            ],
            'thresholds': {
                'error_percentage': 5.0,
                'rationale': 'Industry standard for audit workpapers allows for up to 5% error rate.'
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
                'contact_email': 'qa_analytics@example.com'
            }
        }

        sample_path = os.path.join(self.config_dir, 'sample_qa_77.yaml')
        with open(sample_path, 'w', encoding='utf-8') as file:
            yaml.dump(sample_config, file, default_flow_style=False)

        logger.info(f"Created enhanced sample config at {sample_path}")

    def get_config(self, analytic_id: str) -> Dict:
        """Get configuration for a specific analytic ID"""
        if analytic_id in self.configs:
            return self.configs[analytic_id]
        else:
            logger.error(f"No configuration found for QA-ID {analytic_id}")
            raise ValueError(f"No configuration found for QA-ID {analytic_id}")

    def save_config(self, config: Dict) -> bool:
        """Save configuration to file"""
        if 'analytic_id' not in config:
            logger.error("Cannot save config: missing analytic_id")
            return False

        try:
            analytic_id = str(config['analytic_id'])
            filename = f"qa_{analytic_id}.yaml"
            file_path = os.path.join(self.config_dir, filename)

            with open(file_path, 'w', encoding='utf-8') as file:
                yaml.dump(config, file, default_flow_style=False)

            # Update in-memory config
            self.configs[analytic_id] = config
            logger.info(f"Saved config for QA-ID {analytic_id} to {file_path}")
            return True

        except Exception as e:
            logger.error(f"Error saving config: {e}")
            return False

    def get_available_analytics(self) -> List[Tuple[str, str]]:
        """Get list of available analytics as (id, name) tuples"""
        return [(analytic_id, config.get('analytic_name', 'Unnamed'))
                for analytic_id, config in self.configs.items()]