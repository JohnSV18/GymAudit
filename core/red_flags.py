"""
Red Flag Detection Module
Defines and checks for membership data anomalies
Supports multiple membership types and locations
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional


class RedFlag:
    """Represents a single red flag violation"""
    def __init__(self, flag_type: str, description: str, value: Any = None):
        self.flag_type = flag_type
        self.description = description
        self.value = value

    def __str__(self):
        return self.description


class RedFlagChecker:
    """Checks membership records for red flags based on configurable rules"""

    # Old format column indices (17-column format)
    COL_LAST_NAME = 0
    COL_FIRST_NAME = 1
    COL_MEMBER = 2
    COL_JOIN_DATE = 3
    COL_EXP_DATE = 4
    COL_TYPE = 5
    COL_GROUP = 6
    COL_CODE = 7
    COL_PAY_TYPE = 8
    COL_DUES_AMT = 9
    COL_CYCLE = 10
    COL_BALANCE = 11
    COL_START_DRAFT = 12
    COL_END_DRAFT = 13
    COL_FULFILLMENT = 14
    COL_MEMBERSHIP_LENGTH = 15
    COL_SALES_REP = 16

    # New format column indices (20-column format)
    # last_name, first_name, member_number, transaction_date, transaction_reference,
    # receipt, amount, join_date, expiration_date, member_type, member_group, code,
    # payment_method, dues_amount, balance, start_draft, end_draft, contract_date,
    # site_number, postedby
    NEW_FORMAT_COLUMNS = {
        'last_name': 0,
        'first_name': 1,
        'member_number': 2,
        'transaction_date': 3,
        'transaction_reference': 4,
        'receipt': 5,
        'amount': 6,
        'join_date': 7,
        'expiration_date': 8,
        'member_type': 9,
        'member_group': 10,
        'code': 11,
        'payment_method': 12,
        'dues_amount': 13,
        'balance': 14,
        'start_draft': 15,
        'end_draft': 16,
        'contract_date': 17,
        'site_number': 18,
        'postedby': 19
    }

    def __init__(self, membership_type: str, location: str, config_path: str = None, format_type: str = 'old'):
        """
        Initialize with membership type and location

        Args:
            membership_type: Key from config (e.g., '1_year_paid_in_full')
            location: Key from config (e.g., 'bqe', 'greenpoint', 'lic')
            config_path: Path to red_flag_rules.json (optional)
            format_type: 'old' for 17-column format, 'new' for 20-column format
        """
        self.membership_type = membership_type
        self.location = location
        self.format_type = format_type

        # Load config
        if config_path is None:
            config_path = Path(__file__).parent.parent / 'config' / 'red_flag_rules.json'

        with open(config_path, 'r') as f:
            self.config = json.load(f)

        # Get membership type config
        self.type_config = self.config['membership_types'].get(membership_type, {})
        self.rules = self.type_config.get('rules', {})
        self.pricing = self.type_config.get('pricing', {})

        # Get expected dues for this location
        self.expected_dues = self.pricing.get(location, 0) or 0

    def get_column_index(self, column_name: str) -> int:
        """
        Get the column index for a given column name based on format type.

        Args:
            column_name: Name of the column (e.g., 'join_date', 'code')

        Returns:
            Column index for the current format
        """
        if self.format_type == 'new':
            return self.NEW_FORMAT_COLUMNS.get(column_name, -1)

        # Old format mapping
        old_format_mapping = {
            'last_name': self.COL_LAST_NAME,
            'first_name': self.COL_FIRST_NAME,
            'member_number': self.COL_MEMBER,
            'join_date': self.COL_JOIN_DATE,
            'expiration_date': self.COL_EXP_DATE,
            'member_type': self.COL_TYPE,
            'member_group': self.COL_GROUP,
            'code': self.COL_CODE,
            'payment_method': self.COL_PAY_TYPE,
            'dues_amt': self.COL_DUES_AMT,
            'dues_amount': self.COL_DUES_AMT,  # Alias for consistency with new format
            'cycle': self.COL_CYCLE,
            'balance': self.COL_BALANCE,
            'start_draft': self.COL_START_DRAFT,
            'end_draft': self.COL_END_DRAFT,
        }
        return old_format_mapping.get(column_name, -1)

    def get_bp_detection_columns(self) -> Dict[str, int]:
        """
        Get column indices needed for BP (Billing Problem) detection.

        Returns:
            Dictionary with 'code', 'member_type', and 'member_group' column indices
        """
        return {
            'code': self.get_column_index('code'),
            'member_type': self.get_column_index('member_type'),
            'member_group': self.get_column_index('member_group')
        }

    @staticmethod
    def parse_date(date_str: str) -> Optional[datetime]:
        """Parse date in various formats (M/D/YY or M/D/YYYY)"""
        if not date_str or not date_str.strip():
            return None
        date_str = date_str.strip()
        if date_str.lower() in ('nan', 'nat', 'none', ''):
            return None

        formats = [
            '%m/%d/%Y',  # 1/15/2025 (4-digit year)
            '%m/%d/%y',  # 1/15/25 (2-digit year)
            '%Y-%m-%d',  # 2025-01-15 (ISO format)
            '%Y/%m/%d',  # 2025/12/30 (year first with slashes)
        ]
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return None

    @staticmethod
    def parse_currency(currency_str: str) -> Optional[float]:
        """Parse currency value, handling commas and quotes"""
        if not currency_str:
            return None
        try:
            cleaned = currency_str.replace(',', '').replace('"', '').replace('$', '').strip()
            return float(cleaned)
        except:
            return None

    def check_date_difference(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if join date and expiration date meet the membership type requirements
        """
        # Use format-aware column lookup
        join_idx = self.get_column_index('join_date')
        exp_idx = self.get_column_index('expiration_date')

        join_date = self.parse_date(row[join_idx])
        exp_date = self.parse_date(row[exp_idx])

        if not join_date or not exp_date:
            return True, RedFlag("date_invalid", "Invalid date format")

        diff_days = (exp_date - join_date).days
        date_rule_type = self.rules.get('date_rule_type', 'exact_range')

        if date_rule_type == 'exact_range':
            # For 1 Year: must be exactly 365-366 days
            min_days = self.rules.get('date_diff_min_days', 365)
            max_days = self.rules.get('date_diff_max_days', 366)

            if not (min_days <= diff_days <= max_days):
                return True, RedFlag(
                    "date_mismatch",
                    f"Exp date not within expected range ({diff_days} days, expected {min_days}-{max_days})",
                    diff_days
                )

        elif date_rule_type == 'max_only':
            # For 3 Month / 1 Month: flag if exceeds max
            max_days = self.rules.get('date_diff_max_days', 31)

            if diff_days > max_days:
                return True, RedFlag(
                    "date_mismatch",
                    f"Exp date exceeds maximum ({diff_days} days, max {max_days})",
                    diff_days
                )

        return False, None

    def check_expiration_year(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if expiration date year matches expected (for Month-to-Month)
        """
        expected_year = self.rules.get('expected_exp_year')
        if expected_year is None:
            return False, None

        # Use format-aware column lookup
        exp_idx = self.get_column_index('expiration_date')
        exp_date = self.parse_date(row[exp_idx])
        if not exp_date:
            return True, RedFlag("date_invalid", "Invalid expiration date")

        if exp_date.year != expected_year:
            return True, RedFlag(
                "exp_year_wrong",
                f"Exp year should be {expected_year} (found {exp_date.year})",
                exp_date.year
            )

        return False, None

    def check_dues_amount(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if dues amount meets the threshold percentage
        """
        if self.expected_dues == 0:
            return False, None  # No dues check if pricing not set

        # Use format-aware column lookup
        dues_idx = self.get_column_index('dues_amount')
        dues_amt = self.parse_currency(row[dues_idx])

        if dues_amt is None:
            return True, RedFlag("dues_invalid", "Invalid dues amount")

        threshold_percent = self.rules.get('payment_threshold_percent', 90)
        min_dues = self.expected_dues * (threshold_percent / 100)

        if dues_amt < min_dues:
            return True, RedFlag(
                "dues_low",
                f"Dues ${dues_amt:.2f} < {threshold_percent}% of ${self.expected_dues:.2f} (min ${min_dues:.2f})",
                dues_amt
            )

        return False, None

    def check_cycle(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if cycle value meets requirements
        """
        # Check if cycle checking is disabled in config
        if not self.rules.get('check_cycle', True):
            return False, None

        # Use format-aware column lookup
        cycle_idx = self.get_column_index('cycle')
        if cycle_idx < 0 or cycle_idx >= len(row):
            return False, None  # Cycle column not available in this format

        try:
            cycle = int(row[cycle_idx])
        except:
            return True, RedFlag("cycle_invalid", "Invalid cycle value")

        cycle_rule_type = self.rules.get('cycle_rule_type', 'exact')

        if cycle_rule_type == 'exact':
            expected = self.rules.get('expected_cycle')
            if expected is not None and cycle != expected:
                return True, RedFlag(
                    "cycle_wrong",
                    f"Cycle should be {expected} (found {cycle})",
                    cycle
                )

        elif cycle_rule_type == 'max':
            max_cycle = self.rules.get('cycle_max')
            if max_cycle is not None and cycle > max_cycle:
                return True, RedFlag(
                    "cycle_exceeds_max",
                    f"Cycle {cycle} exceeds maximum of {max_cycle}",
                    cycle
                )

        return False, None

    def check_balance(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if balance is as expected (usually 0)
        """
        if not self.rules.get('check_balance', True):
            return False, None

        # Use format-aware column lookup
        balance_idx = self.get_column_index('balance')
        balance = self.parse_currency(row[balance_idx])

        if balance is None:
            return True, RedFlag("balance_invalid", "Invalid balance")

        expected_balance = self.rules.get('expected_balance', 0)

        if balance != expected_balance:
            balance_type = "credit" if balance < 0 else "debit"
            return True, RedFlag(
                f"balance_{balance_type}",
                f"Balance: ${balance:.2f} ({balance_type}), expected ${expected_balance:.2f}",
                balance
            )

        return False, None

    def check_end_draft_date(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if end draft date meets requirements
        """
        # Use format-aware column lookup
        end_draft_idx = self.get_column_index('end_draft')

        if not self.rules.get('check_end_draft', False):
            # Check for year-based rule (Month-to-Month)
            expected_year = self.rules.get('expected_end_draft_year')
            if expected_year is None:
                return False, None

            end_draft_str = row[end_draft_idx].strip()
            end_draft = self.parse_date(end_draft_str)

            if not end_draft:
                return True, RedFlag("end_draft_invalid", "Invalid end draft date")

            if end_draft.year != expected_year:
                return True, RedFlag(
                    "end_draft_year_wrong",
                    f"End draft year should be {expected_year} (found {end_draft.year})",
                    end_draft.year
                )

            return False, None

        # Original exact match check
        end_draft = row[end_draft_idx].strip()
        expected = self.rules.get('expected_end_draft', '12/31/99')

        if end_draft != expected:
            return True, RedFlag(
                "end_draft_wrong",
                f"End Draft: {end_draft} (expected {expected})",
                end_draft
            )

        return False, None

    def check_draft_date(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if draft date is within acceptable range from join date (for Month-to-Month)
        """
        max_months = self.rules.get('draft_date_max_months_from_join')
        if max_months is None:
            return False, None

        # Use format-aware column lookup
        join_idx = self.get_column_index('join_date')
        start_draft_idx = self.get_column_index('start_draft')

        join_date = self.parse_date(row[join_idx])
        draft_date = self.parse_date(row[start_draft_idx])

        if not join_date or not draft_date:
            return False, None  # Can't check without valid dates

        diff_days = (draft_date - join_date).days
        max_days = max_months * 31  # Approximate

        if diff_days > max_days:
            return True, RedFlag(
                "draft_date_too_far",
                f"Draft date is {diff_days} days from join date (max ~{max_days} days / {max_months} months)",
                diff_days
            )

        return False, None

    def check_transaction_amount(self, row: List[str]) -> Tuple[bool, Optional[RedFlag]]:
        """
        Check if transaction amount is less than 90% of expected price.
        Applies to both charges (positive) and payments (negative).
        Only runs for new format files with 'amount' column.
        """
        if self.format_type != 'new':
            return False, None

        if self.expected_dues == 0:
            return False, None  # No price configured for this location

        # Get amount column (index 6 in new format)
        amount_idx = self.get_column_index('amount')
        if amount_idx < 0 or amount_idx >= len(row):
            return False, None

        amount = self.parse_currency(row[amount_idx])
        if amount is None:
            return False, None

        # Calculate 90% threshold
        threshold_percent = self.rules.get('payment_threshold_percent', 90)
        min_expected = self.expected_dues * (threshold_percent / 100)

        abs_amount = abs(amount)

        # Check if amount is less than threshold (including $0)
        if abs_amount < min_expected:
            if amount == 0:
                txn_type = "Transaction"
            else:
                txn_type = "Charge" if amount > 0 else "Payment"
            return True, RedFlag(
                "low_amount",
                f"{txn_type} ${abs_amount:.2f} is less than {threshold_percent}% of expected ${self.expected_dues:.2f} (min ${min_expected:.2f})",
                amount
            )

        return False, None

    def check_charge_needs_verification(
        self,
        row: List[str],
        prev_row: Optional[List[str]],
        next_row: Optional[List[str]]
    ) -> Optional[RedFlag]:
        """Check if charge passes threshold but has no matching payment in adjacent rows."""
        if self.format_type != 'new':
            return None

        amount_idx = self.get_column_index('amount')
        member_idx = self.get_column_index('member_number')

        if amount_idx < 0 or amount_idx >= len(row):
            return None

        amount = self.parse_currency(row[amount_idx])
        if amount is None or amount <= 0:
            return None

        # Only check charges that PASS the threshold
        threshold_percent = self.rules.get('payment_threshold_percent', 90)
        min_expected = self.expected_dues * (threshold_percent / 100)
        if amount < min_expected:
            return None

        # Get current row's member_number for comparison
        current_member = row[member_idx].strip() if member_idx >= 0 and member_idx < len(row) else None

        # Check adjacent rows for matching payment FROM SAME MEMBER
        def get_adjacent_amount(adj_row):
            if adj_row is None or amount_idx >= len(adj_row):
                return None
            # Check if same member
            if current_member and member_idx >= 0 and member_idx < len(adj_row):
                adj_member = adj_row[member_idx].strip()
                if adj_member != current_member:
                    return None  # Different member, don't count as matching
            return self.parse_currency(adj_row[amount_idx])

        prev_amount = get_adjacent_amount(prev_row)
        next_amount = get_adjacent_amount(next_row)
        tolerance = 1.0

        has_matching_payment = False
        if prev_amount is not None and prev_amount < 0:
            if abs(amount + prev_amount) <= tolerance:
                has_matching_payment = True
        if next_amount is not None and next_amount < 0:
            if abs(amount + next_amount) <= tolerance:
                has_matching_payment = True

        if not has_matching_payment:
            return RedFlag(
                "needs_verification",
                f"Charge ${amount:.2f} - no matching payment in adjacent rows, verify in gym software",
                amount
            )
        return None

    def get_min_monthly_fee(self) -> float:
        """
        Get the minimum monthly fee for the current location.

        Returns:
            Minimum monthly fee amount
        """
        min_fee = self.rules.get('min_monthly_fee', 0)
        if isinstance(min_fee, dict):
            return min_fee.get(self.location, 0) or 0
        return min_fee or 0

    def get_grace_period_months(self) -> int:
        """
        Get the grace period in months for payment verification.

        Returns:
            Number of grace period months
        """
        return self.rules.get('grace_period_months', 3)

    def check_all(self, row: List[str]) -> List[RedFlag]:
        """
        Run all applicable red flag checks on a row

        Args:
            row: List representing a CSV row

        Returns:
            List of RedFlag objects (empty if no flags)
        """
        red_flags = []

        # List of all checks to run
        checks = [
            self.check_date_difference,
            self.check_expiration_year,
            self.check_dues_amount,
            self.check_cycle,
            self.check_balance,
            self.check_end_draft_date,
            self.check_draft_date,
            self.check_transaction_amount,  # Transaction amount check for new format files
        ]

        for check_func in checks:
            is_flagged, flag = check_func(row)
            if is_flagged and flag:
                red_flags.append(flag)

        return red_flags

    def calculate_membership_age(self, row: List[str]) -> Optional[int]:
        """Calculate days since join date"""
        # Use format-aware column lookup
        join_idx = self.get_column_index('join_date')
        join_date = self.parse_date(row[join_idx])
        if join_date:
            return (datetime.now() - join_date).days
        return None

    def is_membership_expired(self, row: List[str]) -> Optional[bool]:
        """Check if membership is expired"""
        # Use format-aware column lookup
        exp_idx = self.get_column_index('expiration_date')
        exp_date = self.parse_date(row[exp_idx])
        if exp_date:
            return datetime.now() > exp_date
        return None

    def get_financial_impact(self, row: List[str], red_flags: List[RedFlag]) -> float:
        """
        Calculate financial impact of red flags

        Returns total dollar amount at risk/missing
        """
        impact = 0.0
        # Use format-aware column lookup
        dues_idx = self.get_column_index('dues_amount')

        for flag in red_flags:
            if flag.flag_type in ['dues_low', 'dues_invalid']:
                # Missing dues (expected - actual)
                actual = self.parse_currency(row[dues_idx]) or 0
                if self.expected_dues > 0:
                    threshold = self.expected_dues * (self.rules.get('payment_threshold_percent', 90) / 100)
                    if actual < threshold:
                        impact += (threshold - actual)

            elif flag.flag_type.startswith('balance_'):
                # Outstanding balance
                balance = abs(flag.value) if flag.value else 0
                impact += balance

        return impact

    def get_financial_impact_breakdown(self, row: List[str], red_flags: List[RedFlag]) -> Dict[str, float]:
        """
        Calculate financial impact broken down by category

        Returns:
            Dictionary with 'dues_impact', 'balance_impact', and 'total'
        """
        dues_impact = 0.0
        balance_impact = 0.0
        # Use format-aware column lookup
        dues_idx = self.get_column_index('dues_amount')

        for flag in red_flags:
            if flag.flag_type in ['dues_low', 'dues_invalid']:
                actual = self.parse_currency(row[dues_idx]) or 0
                if self.expected_dues > 0:
                    threshold = self.expected_dues * (self.rules.get('payment_threshold_percent', 90) / 100)
                    if actual < threshold:
                        dues_impact += (threshold - actual)

            elif flag.flag_type.startswith('balance_'):
                balance = abs(flag.value) if flag.value else 0
                balance_impact += balance

        return {
            'dues_impact': dues_impact,
            'balance_impact': balance_impact,
            'total': dues_impact + balance_impact
        }


def load_config(config_path: str = None) -> Dict[str, Any]:
    """
    Load the red flag rules configuration

    Args:
        config_path: Path to config file (optional)

    Returns:
        Configuration dictionary
    """
    if config_path is None:
        config_path = Path(__file__).parent.parent / 'config' / 'red_flag_rules.json'

    with open(config_path, 'r') as f:
        return json.load(f)


def get_locations(config: Dict[str, Any] = None) -> Dict[str, str]:
    """
    Get available locations

    Returns:
        Dictionary mapping location key to display name
    """
    if config is None:
        config = load_config()

    return {
        key: loc['display_name']
        for key, loc in config.get('locations', {}).items()
    }


def get_membership_types(config: Dict[str, Any] = None) -> Dict[str, str]:
    """
    Get available membership types

    Returns:
        Dictionary mapping type key to display name
    """
    if config is None:
        config = load_config()

    return {
        key: mt['name']
        for key, mt in config.get('membership_types', {}).items()
        if mt.get('enabled', True)
    }


def create_checker(membership_type: str, location: str, format_type: str = 'old') -> RedFlagChecker:
    """
    Create a RedFlagChecker for the specified membership type and location

    Args:
        membership_type: Key from config (e.g., '1_year_paid_in_full')
        location: Key from config (e.g., 'bqe', 'greenpoint', 'lic')
        format_type: 'old' for 17-column format, 'new' for 20-column format

    Returns:
        Configured RedFlagChecker instance
    """
    return RedFlagChecker(membership_type, location, format_type=format_type)


# Backwards compatibility
def create_default_checker() -> RedFlagChecker:
    """
    Create a RedFlagChecker with default rules (1 Year Paid in Full, BQE)
    For backwards compatibility

    Returns:
        Configured RedFlagChecker instance
    """
    return RedFlagChecker('1_year_paid_in_full', 'bqe')
