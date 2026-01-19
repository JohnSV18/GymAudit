"""
Statistics Module
Provides pattern detection and statistical analysis for audit results
"""

from typing import List, Dict, Any
from collections import defaultdict, Counter
from datetime import datetime


class AuditStatistics:
    """Analyzes audit results for patterns and trends"""

    def __init__(self, audit_results: List[Dict[str, Any]]):
        """
        Initialize with audit results

        Args:
            audit_results: List of audit results from audit_engine
        """
        self.audit_results = audit_results
        self.flagged_only = [r for r in audit_results if r['has_flags']]

    def get_red_flag_counts(self) -> Dict[str, int]:
        """
        Count occurrences of each red flag type

        Returns:
            Dictionary mapping flag_type to count
        """
        flag_counts = Counter()

        for result in self.audit_results:
            for flag in result['red_flags']:
                flag_counts[flag.flag_type] += 1

        return dict(flag_counts)

    def get_red_flag_combinations(self) -> Dict[str, int]:
        """
        Find which red flags appear together

        Returns:
            Dictionary mapping flag combination to count
        """
        combinations = Counter()

        for result in self.flagged_only:
            flag_types = sorted([flag.flag_type for flag in result['red_flags']])

            if len(flag_types) > 1:
                combo_key = " + ".join(flag_types)
                combinations[combo_key] += 1

        return dict(combinations)

    def get_most_common_combinations(self, top_n: int = 10) -> List[tuple]:
        """
        Get most common red flag combinations

        Args:
            top_n: Number of top combinations to return

        Returns:
            List of (combination, count) tuples
        """
        combinations = self.get_red_flag_combinations()
        sorted_combos = sorted(combinations.items(), key=lambda x: x[1], reverse=True)
        return sorted_combos[:top_n]

    def group_by_sales_rep(self) -> Dict[str, Dict[str, Any]]:
        """
        Group results by sales representative

        Returns:
            Dictionary with stats per sales rep
        """
        from core.red_flags import RedFlagChecker

        rep_stats = defaultdict(lambda: {
            'total': 0,
            'flagged': 0,
            'clean': 0,
            'flag_percentage': 0,
            'financial_impact': 0
        })

        for result in self.audit_results:
            row = result['row_data']
            sales_rep = row[RedFlagChecker.COL_SALES_REP] if len(row) > RedFlagChecker.COL_SALES_REP else 'Unknown'

            if not sales_rep or sales_rep.strip() == '':
                sales_rep = 'Not Assigned'

            rep_stats[sales_rep]['total'] += 1

            if result['has_flags']:
                rep_stats[sales_rep]['flagged'] += 1
            else:
                rep_stats[sales_rep]['clean'] += 1

            rep_stats[sales_rep]['financial_impact'] += result.get('financial_impact', 0)

        # Calculate percentages
        for rep in rep_stats:
            total = rep_stats[rep]['total']
            if total > 0:
                rep_stats[rep]['flag_percentage'] = (rep_stats[rep]['flagged'] / total) * 100

        return dict(rep_stats)

    def group_by_join_date_range(self, bin_size_months: int = 3) -> Dict[str, Dict[str, Any]]:
        """
        Group results by join date ranges

        Args:
            bin_size_months: Size of date bins in months (default 3 = quarterly)

        Returns:
            Dictionary with stats per date range
        """
        from core.red_flags import RedFlagChecker

        date_stats = defaultdict(lambda: {
            'total': 0,
            'flagged': 0,
            'clean': 0,
            'flag_percentage': 0,
            'financial_impact': 0
        })

        for result in self.audit_results:
            row = result['row_data']
            join_date_str = row[RedFlagChecker.COL_JOIN_DATE] if len(row) > RedFlagChecker.COL_JOIN_DATE else ''

            # Parse date
            try:
                join_date = datetime.strptime(join_date_str, '%m/%d/%y')
                # Create range key (e.g., "2024-Q1", "2024-Q2")
                year = join_date.year
                quarter = (join_date.month - 1) // 3 + 1
                date_key = f"{year}-Q{quarter}"

            except:
                date_key = 'Invalid Date'

            date_stats[date_key]['total'] += 1

            if result['has_flags']:
                date_stats[date_key]['flagged'] += 1
            else:
                date_stats[date_key]['clean'] += 1

            date_stats[date_key]['financial_impact'] += result.get('financial_impact', 0)

        # Calculate percentages
        for date_range in date_stats:
            total = date_stats[date_range]['total']
            if total > 0:
                date_stats[date_range]['flag_percentage'] = (date_stats[date_range]['flagged'] / total) * 100

        return dict(date_stats)

    def get_financial_summary(self) -> Dict[str, Any]:
        """
        Calculate financial impact statistics

        Returns:
            Dictionary with financial metrics
        """
        total_impact = sum(r.get('financial_impact', 0) for r in self.audit_results)
        flagged_impact = sum(r.get('financial_impact', 0) for r in self.flagged_only)

        # Break down by flag type
        impact_by_type = defaultdict(float)

        for result in self.flagged_only:
            for flag in result['red_flags']:
                # This is a simplified approach - in reality, each flag contributes to total impact
                if flag.flag_type.startswith('dues_'):
                    impact_by_type['Missing Dues'] += result.get('financial_impact', 0) / result['flag_count']
                elif flag.flag_type.startswith('balance_'):
                    impact_by_type['Outstanding Balances'] += result.get('financial_impact', 0) / result['flag_count']

        return {
            'total_impact': total_impact,
            'flagged_accounts_impact': flagged_impact,
            'average_impact_per_flagged_account': flagged_impact / len(self.flagged_only) if self.flagged_only else 0,
            'impact_by_type': dict(impact_by_type),
            'accounts_with_impact': sum(1 for r in self.audit_results if r.get('financial_impact', 0) > 0)
        }

    def get_top_impact_accounts(self, top_n: int = 20) -> List[Dict[str, Any]]:
        """
        Get accounts with highest financial impact

        Args:
            top_n: Number of top accounts to return

        Returns:
            List of account details sorted by impact
        """
        # Sort by financial impact
        sorted_results = sorted(
            self.audit_results,
            key=lambda r: r.get('financial_impact', 0),
            reverse=True
        )

        top_accounts = []
        for result in sorted_results[:top_n]:
            if result.get('financial_impact', 0) > 0:
                top_accounts.append({
                    'member_id': result['member_id'],
                    'member_name': result['member_name'],
                    'financial_impact': result['financial_impact'],
                    'flag_count': result['flag_count'],
                    'red_flags': [str(flag) for flag in result['red_flags']]
                })

        return top_accounts

    def get_expired_vs_active_stats(self) -> Dict[str, Dict[str, Any]]:
        """
        Compare stats for expired vs active memberships

        Returns:
            Dictionary with stats for expired and active
        """
        stats = {
            'active': {'total': 0, 'flagged': 0, 'flag_percentage': 0},
            'expired': {'total': 0, 'flagged': 0, 'flag_percentage': 0},
            'unknown': {'total': 0, 'flagged': 0, 'flag_percentage': 0}
        }

        for result in self.audit_results:
            is_expired = result.get('is_expired')

            if is_expired is None:
                key = 'unknown'
            elif is_expired:
                key = 'expired'
            else:
                key = 'active'

            stats[key]['total'] += 1
            if result['has_flags']:
                stats[key]['flagged'] += 1

        # Calculate percentages
        for key in stats:
            total = stats[key]['total']
            if total > 0:
                stats[key]['flag_percentage'] = (stats[key]['flagged'] / total) * 100

        return stats

    def get_summary_statistics(self) -> Dict[str, Any]:
        """
        Get comprehensive summary of all statistics

        Returns:
            Dictionary with all key metrics
        """
        total = len(self.audit_results)
        flagged = len(self.flagged_only)

        return {
            'total_records': total,
            'flagged_count': flagged,
            'clean_count': total - flagged,
            'flagged_percentage': (flagged / total * 100) if total > 0 else 0,
            'red_flag_counts': self.get_red_flag_counts(),
            'most_common_combinations': self.get_most_common_combinations(5),
            'financial_summary': self.get_financial_summary(),
            'expired_vs_active': self.get_expired_vs_active_stats()
        }

    def generate_member_id_list(self, flagged_only: bool = True) -> List[str]:
        """
        Generate list of member IDs

        Args:
            flagged_only: If True, return only flagged members

        Returns:
            List of member IDs
        """
        results = self.flagged_only if flagged_only else self.audit_results
        return [r['member_id'] for r in results if r['member_id']]
