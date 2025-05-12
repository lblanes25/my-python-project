import pandas as pd
import numpy as np
from typing import Dict, List
from logging_config import setup_logging

logger = setup_logging()


class ValidationRules:
    """Library of validation rules that can be applied to data"""

    @staticmethod
    def segregation_of_duties(df: pd.DataFrame, params: Dict) -> pd.Series:
        """
        Validates segregation of duties - submitter cannot be an approver

        Args:
            df: DataFrame containing the data
            params: Dict with 'submitter_field' and 'approver_fields' keys

        Returns:
            Series with True for rows that conform, False for non-conforming
        """
        submitter_field = params.get('submitter_field')
        approver_fields = params.get('approver_fields', [])

        if not submitter_field or not approver_fields:
            logger.error("Missing required parameters for segregation_of_duties")
            return pd.Series(False, index=df.index)

        # Standardize names to lowercase for comparison and handle None values
        df_clean = df.copy()
        df_clean[submitter_field] = df_clean[submitter_field].str.lower() if df_clean[
                                                                                 submitter_field].dtype == 'object' else \
            df_clean[submitter_field]

        # Initialize result as all True
        result = pd.Series(True, index=df.index)

        # Check each approver field
        for approver_field in approver_fields:
            if approver_field in df.columns:
                df_clean[approver_field] = df_clean[approver_field].str.lower() if df_clean[
                                                                                       approver_field].dtype == 'object' else \
                    df_clean[approver_field]
                # Mark false where submitter = approver (ignoring nulls)
                submitter_is_approver = (df_clean[submitter_field].notna() &
                                         df_clean[approver_field].notna() &
                                         (df_clean[submitter_field] == df_clean[approver_field]))
                result = result & ~submitter_is_approver

        return result

    @staticmethod
    def approval_sequence(df: pd.DataFrame, params: Dict) -> pd.Series:
        """
        Validates that approvals happened in the correct sequence

        Args:
            df: DataFrame containing the data
            params: Dict with 'date_fields_in_order' key

        Returns:
            Series with True for rows that conform, False for non-conforming
        """
        date_fields = params.get('date_fields_in_order', [])

        if not date_fields or len(date_fields) < 2:
            logger.error("Not enough date fields for approval_sequence")
            return pd.Series(False, index=df.index)

        # Convert date columns to datetime if they aren't already
        df_dates = df.copy()
        for field in date_fields:
            if field in df.columns:
                try:
                    df_dates[field] = pd.to_datetime(df_dates[field], errors='coerce')
                except Exception as e:
                    logger.error(f"Error converting {field} to datetime: {e}")

        # Initialize result as all True
        result = pd.Series(True, index=df.index)

        # Check sequential dates
        for i in range(len(date_fields) - 1):
            field1 = date_fields[i]
            field2 = date_fields[i + 1]

            if field1 in df.columns and field2 in df.columns:
                # Both dates present - field1 should be before field2
                both_present = df_dates[field1].notna() & df_dates[field2].notna()
                correct_order = df_dates[field1] <= df_dates[field2]

                # Update result - only check ordering if both dates are present
                result = result & (~both_present | correct_order)

        return result

    @staticmethod
    def title_based_approval(df: pd.DataFrame, params: Dict, ref_data: Dict) -> pd.Series:
        """
        Validates that approvers have appropriate titles

        Args:
            df: DataFrame containing the data
            params: Dict with fields and allowed titles
            ref_data: Reference data containing title information

        Returns:
            Series with True for rows that conform, False for non-conforming
        """
        approver_field = params.get('approver_field')
        allowed_titles = params.get('allowed_titles', [])
        title_ref_name = params.get('title_reference')

        if not approver_field or not title_ref_name or title_ref_name not in ref_data:
            logger.error("Missing parameters for title_based_approval")
            return pd.Series(False, index=df.index)

        # Get title reference data
        title_dict = ref_data[title_ref_name]

        # Check each approver's title
        result = pd.Series(False, index=df.index)

        for idx, row in df.iterrows():
            approver = row[approver_field]
            if pd.isna(approver):
                result[idx] = True  # No approver, so can't check
                continue

            # Look up title in reference data
            approver_title = title_dict.get(approver)
            if approver_title and approver_title in allowed_titles:
                result[idx] = True

        return result

    @staticmethod
    def third_party_risk_validation(df: pd.DataFrame, params: Dict) -> pd.Series:
        """
        Validates that third party risk assessment is properly completed when
        third parties are present

        Args:
            df: DataFrame containing the data
            params: Dict with 'third_party_field' and 'risk_level_field' keys

        Returns:
            Series with True for rows that conform, False for non-conforming
        """
        third_party_field = params.get('third_party_field')
        risk_level_field = params.get('risk_level_field')

        if not third_party_field or not risk_level_field:
            logger.error("Missing required parameters for third_party_risk_validation")
            return pd.Series(False, index=df.index)

        # Initialize result as all False
        result = pd.Series(False, index=df.index)

        for idx, row in df.iterrows():
            # Get third party and risk values for this row
            third_parties = row[third_party_field]
            risk_level = row[risk_level_field]

            # Case 1: No third parties and risk level is N/A - this is correct
            if (pd.isna(third_parties) or third_parties == ""):
                if risk_level == "N/A":
                    result[idx] = True
            # Case 2: Third parties exist and risk level is NOT N/A - this is correct
            elif not pd.isna(third_parties) and third_parties != "":
                if not pd.isna(risk_level) and risk_level != "" and risk_level != "N/A":
                    result[idx] = True

        return result