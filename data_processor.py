import os
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional
from validation_rules import ValidationRules
from logging_config import setup_logging

logger = setup_logging()

class DataProcessor:
    """Processes data files according to configuration rules"""

    def __init__(self, config: Dict):
        """Initialize with configuration dictionary"""
        self.config = config
        self.validation_rules = ValidationRules()
        self.reference_data = {}
        self.source_data = None
        self.results = None

    def load_source_data(self, file_path: str) -> bool:
        """
        Load source data from file

        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Load data from Excel
            self.source_data = pd.read_excel(file_path)

            # Map column aliases to standard names
            self._map_column_aliases()

            # Validate required columns
            missing_columns = self._check_required_columns()
            if missing_columns:
                logger.error(f"Missing required columns: {', '.join(missing_columns)}")
                return False

            # Clean and prepare the data
            self._clean_data()

            logger.info(f"Successfully loaded source data with {len(self.source_data)} rows")
            return True

        except Exception as e:
            logger.error(f"Error loading source data: {e}")
            return False

    def _map_column_aliases(self) -> None:
        """Map column aliases to standard names based on configuration"""
        if not self.source_data is not None:
            return

        column_mapping = {}
        for column_info in self.config['source']['required_columns']:
            std_name = column_info['name']
            aliases = column_info.get('alias', [])

            # Check if standard name exists in DataFrame
            if std_name in self.source_data.columns:
                continue

            # Check if any alias exists in DataFrame
            for alias in aliases:
                if alias in self.source_data.columns:
                    column_mapping[alias] = std_name
                    break

        # Rename columns if needed
        if column_mapping:
            self.source_data = self.source_data.rename(columns=column_mapping)

    def _check_required_columns(self) -> List[str]:
        """Check that all required columns are present"""
        if self.source_data is None:
            return []

        required_columns = [col['name'] for col in self.config['source']['required_columns']]
        missing = [col for col in required_columns if col not in self.source_data.columns]
        return missing

    def _clean_data(self) -> None:
        """Clean and prepare data for analysis"""
        if self.source_data is None:
            return

        # Convert date columns to datetime
        for col_info in self.config['source']['required_columns']:
            col_name = col_info['name']
            if 'date' in col_name.lower() and col_name in self.source_data.columns:
                try:
                    self.source_data[col_name] = pd.to_datetime(
                        self.source_data[col_name],
                        errors='coerce'
                    )
                except Exception as e:
                    logger.warning(f"Error converting {col_name} to datetime: {e}")

        # Strip whitespace from string columns
        for col in self.source_data.columns:
            if self.source_data[col].dtype == 'object':
                self.source_data[col] = self.source_data[col].str.strip()

    def load_reference_data(self) -> bool:
        """
        Load reference data files specified in configuration

        Returns:
            bool: True if successful or no reference data needed
        """
        if 'reference_files' not in self.config:
            return True

        ref_files = self.config.get('reference_files', [])
        if not ref_files:
            return True

        success = True
        for ref_file_info in ref_files:
            try:
                name = ref_file_info['name']
                path = ref_file_info['path']
                key_col = ref_file_info['key_column']
                value_col = ref_file_info['value_column']

                if not os.path.exists(path):
                    logger.error(f"Reference file not found: {path}")
                    success = False
                    continue

                # Load the reference file
                ref_df = pd.read_excel(path)

                # Create a dictionary from key->value columns
                ref_dict = dict(zip(ref_df[key_col], ref_df[value_col]))

                # Store in reference data dictionary
                self.reference_data[name] = ref_dict

                logger.info(f"Loaded reference data '{name}' with {len(ref_dict)} entries")

            except Exception as e:
                logger.error(f"Error loading reference data {ref_file_info.get('name')}: {e}")
                success = False

        return success

    def run_validations(self) -> None:
        """Run all validation rules and compile results"""
        if self.source_data is None:
            logger.error("Cannot run validations - no source data loaded")
            return

        # Create result columns for each validation
        validation_results = {}

        for validation in self.config['validations']:
            rule_name = validation['rule']
            params = validation.get('parameters', {})

            # Get the validation method by name
            if hasattr(self.validation_rules, rule_name) and callable(getattr(self.validation_rules, rule_name)):
                validation_method = getattr(self.validation_rules, rule_name)

                # Run the validation
                try:
                    if rule_name == 'title_based_approval':
                        result = validation_method(self.source_data, params, self.reference_data)
                    else:
                        result = validation_method(self.source_data, params)

                    validation_results[rule_name] = result
                    logger.info(f"Validation '{rule_name}' completed - {result.sum()} of {len(result)} records conform")

                except Exception as e:
                    logger.error(f"Error running validation '{rule_name}': {e}")
                    validation_results[rule_name] = pd.Series(False, index=self.source_data.index)
            else:
                logger.error(f"Validation rule '{rule_name}' not found")
                validation_results[rule_name] = pd.Series(False, index=self.source_data.index)

        # Calculate overall result - "GC", "PC", or "DNC"
        if validation_results:
            # Create a DataFrame with all validation results
            result_df = pd.DataFrame(validation_results)

            # Add to source data
            for col in result_df.columns:
                self.source_data[f"Valid_{col}"] = result_df[col]

            # Determine if ALL validations pass - Generally Conforms (GC)
            all_valid = result_df.all(axis=1)

            # Calculate overall compliance result
            self.source_data['Compliance'] = np.where(
                all_valid,
                'GC',  # All validations pass - Generally Conforms
                'DNC'  # Some validations fail - Does Not Conform
            )

            # Add a column to help when manually validating DNCs
            self.source_data['DNC_Validated'] = np.where(
                self.source_data['Compliance'] == 'DNC',
                'TBD',  # To be validated manually
                'N/A'  # Not applicable for GC items
            )

            logger.info(f"Validation complete: {all_valid.sum()} GC, {len(all_valid) - all_valid.sum()} DNC")
        else:
            logger.warning("No validation results calculated")
            self.source_data['Compliance'] = 'N/A'

    def generate_summary(self) -> pd.DataFrame:
        """Generate summary statistics by group"""
        if self.source_data is None or 'Compliance' not in self.source_data:
            logger.error("Cannot generate summary - validation not complete")
            return None

        group_by_field = self.config['reporting']['group_by']

        if group_by_field not in self.source_data.columns:
            logger.error(f"Group by field '{group_by_field}' not found in data")
            return None

        # Count records by group and compliance status
        summary = self.source_data.groupby([group_by_field, 'Compliance']).size().unstack(fill_value=0)

        # Ensure all compliance categories exist
        for category in ['GC', 'PC', 'DNC']:
            if category not in summary.columns:
                summary[category] = 0

        # Calculate totals and percentages
        summary['Total'] = summary.sum(axis=1)
        summary['DNC_Percentage'] = (summary['DNC'] / summary['Total'] * 100).round(2)

        # Compare against threshold
        threshold = self.config['thresholds']['error_percentage']
        summary['Exceeds_Threshold'] = summary['DNC_Percentage'] > threshold

        # Order columns
        ordered_cols = ['GC', 'PC', 'DNC', 'Total', 'DNC_Percentage', 'Exceeds_Threshold']
        summary = summary[ordered_cols]

        # Reset index to make the group field a column
        summary = summary.reset_index()

        return summary

    def process_data(self, source_file: str) -> Tuple[bool, str]:
        """
        Process data file according to configuration

        Args:
            source_file: Path to source data file

        Returns:
            Tuple of (success, message)
        """
        # Step 1: Load source data
        if not self.load_source_data(source_file):
            return False, "Failed to load source data"

        # Step 2: Load reference data if needed
        if not self.load_reference_data():
            return False, "Failed to load reference data"

        # Step 3: Run validations
        self.run_validations()

        # Step 4: Generate summary
        summary = self.generate_summary()
        if summary is None:
            return False, "Failed to generate summary"

        self.results = {
            'detail': self.source_data,
            'summary': summary
        }

        return True, "Processing complete"
