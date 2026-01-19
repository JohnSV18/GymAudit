"""
Red Flag Detection Module
Defines and checks for membership data anomalies
"""

from datetime import datetime
from typing import Dict, List, Any, Tuple


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

    # Column indices (matching the CSV structure)
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

    def __init__(self, rules: Dict[str, Any]):
        """
        Initialize with red flag rules

        Args:
            rules: Dictionary containing red flag criteria
        """
        self.rules = rules

    @staticmethod
    def parse_date(date_str: str) -> datetime:
        """Parse date in M/D/YY format"""
        try:
            return datetime.strptime(date_str, '%m/%d/%y')
        except:
            return None

    @staticmethod
    def parse_currency(currency_str: str) -> float:
        """Parse currency value, handling commas and quotes"""
        try:
            cleaned = currency_str.replace(',', '').replace('"', '').strip()
            return float(cleaned)
        except:
            return None

    def check_date_difference(self, row: List[str]) -> Tuple[bool, RedFlag]:
        """
        Check if join date and expiration date are exactly one year apart

        Returns:
            (is_flagged, RedFlag or None)
        """
        join_date = self.parse_date(row[self.COL_JOIN_DATE])
        exp_date = self.parse_date(row[self.COL_EXP_DATE])

        if not join_date or not exp_date:
            return True, RedFlag(
                "date_invalid",
                "Invalid date format"
            )

        diff_days = (exp_date - join_date).days
        expected_min = self.rules.get('date_diff_min_days', 365)
        expected_max = self.rules.get('date_diff_max_days', 366)

        if not (expected_min <= diff_days <= expected_max):
            return True, RedFlag(
                "date_mismatch",
                f"Join/Exp dates not 1 year apart ({diff_days} days)",
                diff_days
            )

        return False, None

    def check_dues_amount(self, row: List[str]) -> Tuple[bool, RedFlag]:
        """
        Check if dues amount meets minimum threshold

        Returns:
            (is_flagged, RedFlag or None)
        """
        dues_amt = self.parse_currency(row[self.COL_DUES_AMT])

        if dues_amt is None:
            return True, RedFlag(
                "dues_invalid",
                "Invalid dues amount"
            )

        min_dues = self.rules.get('min_dues_amount', 600)

        if dues_amt < min_dues:
            return True, RedFlag(
                "dues_low",
                f"Dues < ${min_dues} (${dues_amt:.2f})",
                dues_amt
            )

        return False, None

    def check_pay_type(self, row: List[str]) -> Tuple[bool, RedFlag]:
        """
        Check if pay type matches expected value

        Returns:
            (is_flagged, RedFlag or None)
        """
        pay_type = row[self.COL_PAY_TYPE].strip()
        expected = self.rules.get('expected_pay_type', 'ANNUAL BILL')

        if pay_type.upper() != expected.upper():
            return True, RedFlag(
                "pay_type_wrong",
                f"Pay Type: {pay_type}",
                pay_type
            )

        return False, None

    def check_end_draft_date(self, row: List[str]) -> Tuple[bool, RedFlag]:
        """
        Check if end draft date matches expected placeholder

        Returns:
            (is_flagged, RedFlag or None)
        """
        end_draft = row[self.COL_END_DRAFT].strip()
        expected = self.rules.get('expected_end_draft', '12/31/99')

        if end_draft != expected:
            return True, RedFlag(
                "end_draft_wrong",
                f"End Draft: {end_draft}",
                end_draft
            )

        return False, None

    def check_cycle(self, row: List[str]) -> Tuple[bool, RedFlag]:
        """
        Check if cycle number matches expected value

        Returns:
            (is_flagged, RedFlag or None)
        """
        try:
            cycle = int(row[self.COL_CYCLE])
            expected = self.rules.get('expected_cycle', 1)

            if cycle != expected:
                return True, RedFlag(
                    "cycle_wrong",
                    f"Cycle: {cycle}",
                    cycle
                )
        except:
            return True, RedFlag(
                "cycle_invalid",
                "Invalid cycle value"
            )

        return False, None

    def check_balance(self, row: List[str]) -> Tuple[bool, RedFlag]:
        """
        Check if balance is exactly zero

        Returns:
            (is_flagged, RedFlag or None)
        """
        balance = self.parse_currency(row[self.COL_BALANCE])

        if balance is None:
            return True, RedFlag(
                "balance_invalid",
                "Invalid balance"
            )

        if balance != 0:
            balance_type = "credit" if balance < 0 else "debit"
            return True, RedFlag(
                f"balance_{balance_type}",
                f"Balance: ${balance:.2f} ({balance_type})",
                balance
            )

        return False, None

    def check_all(self, row: List[str]) -> List[RedFlag]:
        """
        Run all red flag checks on a row

        Args:
            row: List representing a CSV row

        Returns:
            List of RedFlag objects (empty if no flags)
        """
        red_flags = []

        # Run all checks
        checks = [
            self.check_date_difference,
            self.check_dues_amount,
            self.check_pay_type,
            self.check_end_draft_date,
            self.check_cycle,
            self.check_balance
        ]

        for check_func in checks:
            is_flagged, flag = check_func(row)
            if is_flagged:
                red_flags.append(flag)

        return red_flags

    def calculate_membership_age(self, row: List[str]) -> int:
        """Calculate days since join date"""
        join_date = self.parse_date(row[self.COL_JOIN_DATE])
        if join_date:
            return (datetime.now() - join_date).days
        return None

    def is_membership_expired(self, row: List[str]) -> bool:
        """Check if membership is expired"""
        exp_date = self.parse_date(row[self.COL_EXP_DATE])
        if exp_date:
            return datetime.now() > exp_date
        return None

    def get_financial_impact(self, row: List[str], red_flags: List[RedFlag]) -> float:
        """
        Calculate financial impact of red flags

        Returns total dollar amount at risk/missing
        """
        impact = 0.0

        for flag in red_flags:
            if flag.flag_type in ['dues_low', 'dues_invalid']:
                # Missing dues (expected - actual)
                expected = self.rules.get('min_dues_amount', 600)
                actual = self.parse_currency(row[self.COL_DUES_AMT]) or 0
                impact += (expected - actual)

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

        for flag in red_flags:
            if flag.flag_type in ['dues_low', 'dues_invalid']:
                # Missing dues (expected - actual)
                expected = self.rules.get('min_dues_amount', 600)
                actual = self.parse_currency(row[self.COL_DUES_AMT]) or 0
                dues_impact += (expected - actual)

            elif flag.flag_type.startswith('balance_'):
                # Outstanding balance
                balance = abs(flag.value) if flag.value else 0
                balance_impact += balance

        return {
            'dues_impact': dues_impact,
            'balance_impact': balance_impact,
            'total': dues_impact + balance_impact
        }


def create_default_checker() -> RedFlagChecker:
    """
    Create a RedFlagChecker with default rules for Year Paid in Full memberships

    Returns:
        Configured RedFlagChecker instance
    """
    default_rules = {
        'date_diff_min_days': 365,
        'date_diff_max_days': 366,
        'min_dues_amount': 600,
        'expected_pay_type': 'ANNUAL BILL',
        'expected_end_draft': '12/31/99',
        'expected_cycle': 1
    }

    return RedFlagChecker(default_rules)
