"""
Audit Engine Module
Orchestrates the membership data audit process
Supports multiple membership types and locations
"""

from typing import List, Dict, Any, Optional
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
from collections import defaultdict
from .red_flags import RedFlagChecker, RedFlag, create_checker, load_config
from .file_handler import MembershipFileReader
from .report_generator import AuditReportGenerator


class AuditEngine:
    """Main audit orchestrator"""

    def __init__(
        self,
        membership_type: str,
        location: str,
        output_folder: str = 'outputs',
        format_type: str = 'old'
    ):
        """
        Initialize audit engine with membership type and location

        Args:
            membership_type: Key from config (e.g., '1_year_paid_in_full', '3_months_paid_in_full')
            location: Key from config (e.g., 'bqe', 'greenpoint', 'lic')
            output_folder: Directory for output reports
            format_type: 'old' for 17-column format, 'new' for 20-column format
        """
        self.membership_type = membership_type
        self.location = location
        self.format_type = format_type
        self.checker = create_checker(membership_type, location, format_type=format_type)
        self.file_reader = MembershipFileReader()
        self.report_generator = AuditReportGenerator(output_folder)

        # Load BP detection config
        config = load_config()
        self.bp_config = config.get('bp_detection', {
            'enabled': True,
            'columns': ['code', 'member_type'],
            'keywords': ['bp', 'billing'],
            'case_sensitive': False
        })

    def audit_rows(self, data_rows: List[List[str]]) -> List[Dict[str, Any]]:
        """
        Audit all data rows and collect results

        Args:
            data_rows: List of data rows to audit

        Returns:
            List of audit results, one per row
        """
        audit_results = []

        for i, row in enumerate(data_rows):
            # Get adjacent rows for payment verification
            prev_row = data_rows[i-1] if i > 0 else None
            next_row = data_rows[i+1] if i < len(data_rows) - 1 else None

            # Run standard checks
            red_flags = self.checker.check_all(row)

            # Check for charge without matching payment (new format only)
            if self.checker.format_type == 'new':
                needs_verify_flag = self.checker.check_charge_needs_verification(row, prev_row, next_row)
                if needs_verify_flag:
                    red_flags.append(needs_verify_flag)

            # Calculate additional context
            membership_age = self.checker.calculate_membership_age(row)
            is_expired = self.checker.is_membership_expired(row)
            financial_impact = self.checker.get_financial_impact(row, red_flags)
            impact_breakdown = self.checker.get_financial_impact_breakdown(row, red_flags)

            # Compile result
            result = {
                'row_data': row,
                'red_flags': red_flags,
                'has_flags': len(red_flags) > 0,
                'flag_count': len(red_flags),
                'membership_age': membership_age,
                'is_expired': is_expired,
                'financial_impact': financial_impact,
                'dues_impact': impact_breakdown['dues_impact'],
                'balance_impact': impact_breakdown['balance_impact'],
                'member_id': row[self.checker.COL_MEMBER] if len(row) > self.checker.COL_MEMBER else '',
                'member_name': f"{row[self.checker.COL_FIRST_NAME]} {row[self.checker.COL_LAST_NAME]}" if len(row) > self.checker.COL_FIRST_NAME else ''
            }

            audit_results.append(result)

        return audit_results

    def audit_file(self, file_path: str, generate_report: bool = True) -> Dict[str, Any]:
        """
        Audit a single file

        Args:
            file_path: Path to file to audit
            generate_report: Whether to generate Excel report

        Returns:
            Dictionary with audit results and statistics
        """
        # Read and validate file
        file_data = self.file_reader.read_and_validate(file_path)

        if not file_data['is_valid']:
            return {
                'success': False,
                'error': file_data['error'],
                'filename': file_data['filename']
            }

        # Update checker format if file format differs from engine default
        detected_format = file_data.get('format_type', 'old')
        if detected_format != self.format_type:
            self.checker = create_checker(self.membership_type, self.location, format_type=detected_format)
            self.format_type = detected_format

        # Audit the rows
        audit_results = self.audit_rows(file_data['data_rows'])

        # Calculate statistics
        total_records = len(audit_results)
        flagged_count = sum(1 for r in audit_results if r['has_flags'])
        clean_count = total_records - flagged_count
        flagged_percentage = (flagged_count / total_records * 100) if total_records > 0 else 0
        total_financial_impact = sum(r['financial_impact'] for r in audit_results)
        total_dues_impact = sum(r['dues_impact'] for r in audit_results)
        total_balance_impact = sum(r['balance_impact'] for r in audit_results)

        # Get flagged member IDs
        flagged_member_ids = [r['member_id'] for r in audit_results if r['has_flags']]

        # Get column mapping for BP detection
        column_mapping = self.checker.get_bp_detection_columns()

        # Generate report if requested
        report_path = None
        if generate_report:
            # Create output filename
            original_name = Path(file_data['filename']).stem
            output_filename = f"{original_name}_Audit_Report.xlsx"

            report_path = self.report_generator.create_audit_report(
                header_row=file_data['header'],
                data_rows=file_data['data_rows'],
                audit_results=audit_results,
                output_filename=output_filename,
                include_summary_sheet=True,
                column_mapping=column_mapping,
                bp_config=self.bp_config
            )

        return {
            'success': True,
            'filename': file_data['filename'],
            'format_type': detected_format,
            'total_records': total_records,
            'flagged_count': flagged_count,
            'clean_count': clean_count,
            'flagged_percentage': flagged_percentage,
            'total_financial_impact': total_financial_impact,
            'total_dues_impact': total_dues_impact,
            'total_balance_impact': total_balance_impact,
            'flagged_member_ids': flagged_member_ids,
            'audit_results': audit_results,
            'report_path': report_path
        }

    def audit_uploaded_file(self, uploaded_file, generate_report: bool = True) -> Dict[str, Any]:
        """
        Audit a file from Streamlit upload

        Args:
            uploaded_file: Streamlit UploadedFile object
            generate_report: Whether to generate Excel report

        Returns:
            Dictionary with audit results and statistics
        """
        # Read and validate file
        try:
            file_data = self.file_reader.read_and_validate_upload(uploaded_file)
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'filename': uploaded_file.name if hasattr(uploaded_file, 'name') else 'Unknown'
            }

        if not file_data['is_valid']:
            return {
                'success': False,
                'error': file_data['error'],
                'filename': file_data['filename']
            }

        # Update checker format if file format differs from engine default
        detected_format = file_data.get('format_type', 'old')
        if detected_format != self.format_type:
            self.checker = create_checker(self.membership_type, self.location, format_type=detected_format)
            self.format_type = detected_format

        # Audit the rows
        audit_results = self.audit_rows(file_data['data_rows'])

        # Calculate statistics
        total_records = len(audit_results)
        flagged_count = sum(1 for r in audit_results if r['has_flags'])
        clean_count = total_records - flagged_count
        flagged_percentage = (flagged_count / total_records * 100) if total_records > 0 else 0
        total_financial_impact = sum(r['financial_impact'] for r in audit_results)
        total_dues_impact = sum(r['dues_impact'] for r in audit_results)
        total_balance_impact = sum(r['balance_impact'] for r in audit_results)

        # Get flagged member IDs
        flagged_member_ids = [r['member_id'] for r in audit_results if r['has_flags']]

        # Get column mapping for BP detection
        column_mapping = self.checker.get_bp_detection_columns()

        # Generate report if requested
        report_path = None
        if generate_report:
            # Create output filename
            original_name = Path(file_data['filename']).stem
            output_filename = f"{original_name}_Audit_Report.xlsx"

            report_path = self.report_generator.create_audit_report(
                header_row=file_data['header'],
                data_rows=file_data['data_rows'],
                audit_results=audit_results,
                output_filename=output_filename,
                include_summary_sheet=True,
                column_mapping=column_mapping,
                bp_config=self.bp_config
            )

        return {
            'success': True,
            'filename': file_data['filename'],
            'format_type': detected_format,
            'total_records': total_records,
            'flagged_count': flagged_count,
            'clean_count': clean_count,
            'flagged_percentage': flagged_percentage,
            'total_financial_impact': total_financial_impact,
            'total_dues_impact': total_dues_impact,
            'total_balance_impact': total_balance_impact,
            'flagged_member_ids': flagged_member_ids,
            'audit_results': audit_results,
            'report_path': report_path
        }

    def audit_multiple_files(self, file_paths: List[str], generate_individual_reports: bool = True, generate_consolidated: bool = True) -> Dict[str, Any]:
        """
        Audit multiple files

        Args:
            file_paths: List of file paths to audit
            generate_individual_reports: Generate report for each file
            generate_consolidated: Generate consolidated summary report

        Returns:
            Dictionary with results for all files
        """
        all_results = []

        for file_path in file_paths:
            result = self.audit_file(file_path, generate_report=generate_individual_reports)
            all_results.append(result)

        # Generate consolidated report if requested
        consolidated_report_path = None
        if generate_consolidated:
            successful_results = [r for r in all_results if r['success']]
            if successful_results:
                consolidated_report_path = self.report_generator.create_consolidated_report(
                    file_results=successful_results
                )

        # Calculate overall statistics
        total_files = len(all_results)
        successful_files = sum(1 for r in all_results if r['success'])
        failed_files = total_files - successful_files

        total_records = sum(r.get('total_records', 0) for r in all_results if r['success'])
        total_flagged = sum(r.get('flagged_count', 0) for r in all_results if r['success'])
        total_financial_impact = sum(r.get('total_financial_impact', 0) for r in all_results if r['success'])

        return {
            'total_files': total_files,
            'successful_files': successful_files,
            'failed_files': failed_files,
            'total_records': total_records,
            'total_flagged': total_flagged,
            'total_financial_impact': total_financial_impact,
            'file_results': all_results,
            'consolidated_report_path': consolidated_report_path
        }

    def audit_multiple_uploaded_files(self, uploaded_files, generate_individual_reports: bool = True, generate_consolidated: bool = True) -> Dict[str, Any]:
        """
        Audit multiple uploaded files

        Args:
            uploaded_files: List of Streamlit UploadedFile objects
            generate_individual_reports: Generate report for each file
            generate_consolidated: Generate consolidated summary report

        Returns:
            Dictionary with results for all files
        """
        all_results = []

        for uploaded_file in uploaded_files:
            result = self.audit_uploaded_file(uploaded_file, generate_report=generate_individual_reports)
            all_results.append(result)

        # Generate consolidated report if requested
        consolidated_report_path = None
        if generate_consolidated and len(uploaded_files) > 1:
            successful_results = [r for r in all_results if r['success']]
            if successful_results:
                consolidated_report_path = self.report_generator.create_consolidated_report(
                    file_results=successful_results
                )

        # Calculate overall statistics
        total_files = len(all_results)
        successful_files = sum(1 for r in all_results if r['success'])
        failed_files = total_files - successful_files

        total_records = sum(r.get('total_records', 0) for r in all_results if r['success'])
        total_flagged = sum(r.get('flagged_count', 0) for r in all_results if r['success'])
        total_financial_impact = sum(r.get('total_financial_impact', 0) for r in all_results if r['success'])

        return {
            'total_files': total_files,
            'successful_files': successful_files,
            'failed_files': failed_files,
            'total_records': total_records,
            'total_flagged': total_flagged,
            'total_financial_impact': total_financial_impact,
            'file_results': all_results,
            'consolidated_report_path': consolidated_report_path
        }

    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse date in M/D/YY format"""
        if not date_str or not date_str.strip():
            return None
        try:
            return datetime.strptime(date_str.strip(), '%m/%d/%y')
        except:
            return None

    def _parse_currency(self, currency_str: str) -> Optional[float]:
        """Parse currency value, handling commas and quotes"""
        if not currency_str:
            return None
        try:
            cleaned = currency_str.replace(',', '').replace('"', '').replace('$', '').strip()
            return float(cleaned)
        except:
            return None

    def _get_month_key(self, date: datetime) -> str:
        """Get year-month key from date (e.g., '2026-01')"""
        return date.strftime('%Y-%m')

    def _get_member_key(self, row: List[str]) -> str:
        """
        Create unique member key from first_name + last_name + member_number.
        Uses new format column indices.
        """
        first_name = row[self.checker.NEW_FORMAT_COLUMNS['first_name']].strip().lower()
        last_name = row[self.checker.NEW_FORMAT_COLUMNS['last_name']].strip().lower()
        member_number = row[self.checker.NEW_FORMAT_COLUMNS['member_number']].strip()
        return f"{first_name}|{last_name}|{member_number}"

    def _is_mtmcore_member(self, row: List[str]) -> bool:
        """Check if member type is MTMCORE"""
        member_type_idx = self.checker.NEW_FORMAT_COLUMNS['member_type']
        if member_type_idx < len(row):
            member_type = row[member_type_idx].strip().upper()
            return member_type == 'MTMCORE'
        return False

    def audit_1mcore_transactions(
        self,
        data_rows: List[List[str]],
        expected_price: float
    ) -> Dict[str, Any]:
        """
        Audit 1-Month Paid in Full (1MCORE) transactions.

        Groups transactions by member and checks:
        1. Payment balance (charges vs payments)
        2. Price match (each charge against expected price)

        Args:
            data_rows: List of transaction rows (new format)
            expected_price: Expected price for 1MCORE at this location

        Returns:
            Dictionary with member-level audit results
        """
        from .red_flags import RedFlag

        cols = self.checker.NEW_FORMAT_COLUMNS

        # Group transactions by member_number
        member_transactions = defaultdict(list)
        for row in data_rows:
            member_number = row[cols['member_number']].strip() if cols['member_number'] < len(row) else ''
            if member_number:
                member_transactions[member_number].append(row)

        member_results = {}

        for member_number, transactions in member_transactions.items():
            first_row = transactions[0]

            # Get member info
            first_name = first_row[cols['first_name']] if cols['first_name'] < len(first_row) else ''
            last_name = first_row[cols['last_name']] if cols['last_name'] < len(first_row) else ''

            all_flags = []
            total_amount = 0.0
            price_mismatches = []
            low_amounts = []

            # Calculate 90% threshold for low amount detection
            threshold_percent = 90
            min_expected = expected_price * (threshold_percent / 100) if expected_price > 0 else 0

            # Process each transaction
            for row in transactions:
                amount_str = row[cols['amount']] if cols['amount'] < len(row) else ''
                amount = self._parse_currency(amount_str)

                if amount is not None:
                    total_amount += amount
                    abs_amount = abs(amount)

                    # Check if amount (charge or payment) is less than 90% of expected
                    if expected_price > 0 and abs_amount < min_expected:
                        low_amounts.append(amount)

                    # Check for price mismatch on charges (positive amounts)
                    if amount > 0 and expected_price > 0:
                        # Allow small tolerance (e.g., $0.01 for rounding)
                        if abs(amount - expected_price) > 0.01:
                            price_mismatches.append(amount)

            # Flag if unpaid balance (charges exceed payments)
            if total_amount > 0.01:  # Small tolerance for floating point
                all_flags.append(RedFlag(
                    "unpaid_balance",
                    f"Unpaid balance: ${total_amount:.2f} (charges exceed payments)",
                    total_amount
                ))

            # Flag low amounts (less than 90% of expected price)
            for low_amount in low_amounts:
                abs_amt = abs(low_amount)
                txn_type = "Charge" if low_amount > 0 else "Payment"
                all_flags.append(RedFlag(
                    "low_amount",
                    f"{txn_type} ${abs_amt:.2f} is less than {threshold_percent}% of expected ${expected_price:.2f} (min ${min_expected:.2f})",
                    low_amount
                ))

            # Flag price mismatches
            for mismatch_amount in price_mismatches:
                all_flags.append(RedFlag(
                    "price_mismatch",
                    f"Charge ${mismatch_amount:.2f} doesn't match expected ${expected_price:.2f} - possible wrong membership type",
                    mismatch_amount
                ))

            # Run basic date/expiration checks on first row
            basic_flags = self.checker.check_all(first_row)
            # Filter out any existing date_invalid flags since we're handling dates better now
            basic_flags = [f for f in basic_flags if f.flag_type != 'date_invalid' or self._parse_date(first_row[cols['join_date']]) is None]
            all_flags.extend(basic_flags)

            member_results[member_number] = {
                'member_number': member_number,
                'first_name': first_name,
                'last_name': last_name,
                'transactions': transactions,
                'transaction_count': len(transactions),
                'total_amount': total_amount,
                'flags': all_flags,
                'has_flags': len(all_flags) > 0,
                'flag_count': len(all_flags),
                'price_mismatches': price_mismatches,
                'low_amounts': low_amounts,
                'first_row': first_row  # Keep for report generation
            }

        return {
            'member_results': member_results,
            'total_members': len(member_results),
            'flagged_members': sum(1 for r in member_results.values() if r['has_flags']),
            'total_transactions': sum(len(r['transactions']) for r in member_results.values())
        }

    def _check_basic_mtm_rules(self, row: List[str]) -> List:
        """
        Check basic Month-to-Month validation rules on a single row.

        Rules:
        - Expiration year = 2099
        - Cycle = 1
        - Start draft date within 3 months of join date
        - End draft year = 2099
        """
        from .red_flags import RedFlag

        flags = []
        cols = self.checker.NEW_FORMAT_COLUMNS
        rules = self.checker.rules

        # Check expiration year
        exp_date_str = row[cols['expiration_date']] if cols['expiration_date'] < len(row) else ''
        exp_date = self._parse_date(exp_date_str)
        expected_exp_year = rules.get('expected_exp_year', 2099)
        if exp_date and exp_date.year != expected_exp_year:
            flags.append(RedFlag(
                "exp_year_wrong",
                f"Exp year should be {expected_exp_year} (found {exp_date.year})",
                exp_date.year
            ))

        # Check start draft date (within 3 months of join date)
        join_date_str = row[cols['join_date']] if cols['join_date'] < len(row) else ''
        start_draft_str = row[cols['start_draft']] if cols['start_draft'] < len(row) else ''
        join_date = self._parse_date(join_date_str)
        start_draft = self._parse_date(start_draft_str)

        max_months = rules.get('draft_date_max_months_from_join', 3)
        if join_date and start_draft:
            diff_days = (start_draft - join_date).days
            max_days = max_months * 31  # Approximate
            if diff_days > max_days:
                flags.append(RedFlag(
                    "draft_date_too_far",
                    f"Draft date is {diff_days} days from join date (max ~{max_days} days / {max_months} months)",
                    diff_days
                ))

        # Check end draft year
        end_draft_str = row[cols['end_draft']] if cols['end_draft'] < len(row) else ''
        end_draft = self._parse_date(end_draft_str)
        expected_end_year = rules.get('expected_end_draft_year', 2099)
        if end_draft and end_draft.year != expected_end_year:
            flags.append(RedFlag(
                "end_draft_year_wrong",
                f"End draft year should be {expected_end_year} (found {end_draft.year})",
                end_draft.year
            ))

        return flags

    def audit_month_to_month_transactions(
        self,
        data_rows: List[List[str]],
        system_date: datetime = None
    ) -> Dict[str, Any]:
        """
        Audit Month-to-Month membership transactions.

        Groups transactions by member and checks:
        1. Basic validation rules (exp year, cycle, draft dates)
        2. Monthly payment verification (after 3-month grace period)
        3. Duplicate payment detection

        Args:
            data_rows: List of transaction rows (new format)
            system_date: Current system date (defaults to today)

        Returns:
            Dictionary with member-level audit results
        """
        from .red_flags import RedFlag

        if system_date is None:
            system_date = datetime.now()

        cols = self.checker.NEW_FORMAT_COLUMNS
        min_monthly_fee = self.checker.get_min_monthly_fee()
        grace_months = self.checker.get_grace_period_months()

        # Group transactions by member
        member_transactions = defaultdict(list)
        for row in data_rows:
            if not self._is_mtmcore_member(row):
                continue
            member_key = self._get_member_key(row)
            member_transactions[member_key].append(row)

        member_results = {}

        for member_key, transactions in member_transactions.items():
            # Use first transaction for member info and basic checks
            first_row = transactions[0]

            # Get member info
            first_name = first_row[cols['first_name']]
            last_name = first_row[cols['last_name']]
            member_number = first_row[cols['member_number']]

            # Get join date (same for all transactions)
            join_date_str = first_row[cols['join_date']]
            join_date = self._parse_date(join_date_str)

            if not join_date:
                member_results[member_key] = {
                    'member_key': member_key,
                    'first_name': first_name,
                    'last_name': last_name,
                    'member_number': member_number,
                    'transactions': transactions,
                    'flags': [RedFlag("join_date_invalid", "Invalid join date")],
                    'missing_months': [],
                    'duplicate_months': [],
                    'has_flags': True
                }
                continue

            # Calculate grace period end
            grace_end = join_date + relativedelta(months=grace_months)

            # Get all required months from grace_end to system_date
            required_months = []
            current_month = datetime(grace_end.year, grace_end.month, 1)
            system_month = datetime(system_date.year, system_date.month, 1)

            while current_month <= system_month:
                required_months.append(self._get_month_key(current_month))
                current_month += relativedelta(months=1)

            # Get payment months from transactions (where amount >= min_monthly_fee)
            payment_months = defaultdict(list)  # month -> list of transaction indices
            for idx, txn in enumerate(transactions):
                txn_date_str = txn[cols['transaction_date']]
                txn_date = self._parse_date(txn_date_str)
                amount = self._parse_currency(txn[cols['amount']])

                if txn_date and amount is not None and amount >= min_monthly_fee:
                    month_key = self._get_month_key(txn_date)
                    payment_months[month_key].append(idx)

            # Check for missing payments
            missing_months = []
            for required_month in required_months:
                if required_month not in payment_months:
                    missing_months.append(required_month)

            # Check for duplicate payments
            duplicate_months = []
            for month_key, txn_indices in payment_months.items():
                if len(txn_indices) > 1:
                    duplicate_months.append(month_key)

            # Run basic validation rules on first transaction
            basic_flags = self._check_basic_mtm_rules(first_row)

            # Create flags for missing/duplicate payments
            all_flags = basic_flags.copy()

            for missing_month in missing_months:
                all_flags.append(RedFlag(
                    "missing_payment",
                    f"No qualifying payment (>= ${min_monthly_fee:.2f}) for {missing_month}",
                    missing_month
                ))

            for dup_month in duplicate_months:
                all_flags.append(RedFlag(
                    "duplicate_payment",
                    f"Multiple transactions in {dup_month}",
                    dup_month
                ))

            member_results[member_key] = {
                'member_key': member_key,
                'first_name': first_name,
                'last_name': last_name,
                'member_number': member_number,
                'join_date': join_date,
                'grace_end': grace_end,
                'transactions': transactions,
                'transaction_count': len(transactions),
                'flags': all_flags,
                'missing_months': missing_months,
                'duplicate_months': duplicate_months,
                'required_months': required_months,
                'payment_months': list(payment_months.keys()),
                'has_flags': len(all_flags) > 0,
                'flag_count': len(all_flags)
            }

        return {
            'member_results': member_results,
            'total_members': len(member_results),
            'flagged_members': sum(1 for r in member_results.values() if r['has_flags']),
            'total_transactions': sum(len(r['transactions']) for r in member_results.values())
        }

    def audit_mtm_file(self, file_path: str, generate_report: bool = True) -> Dict[str, Any]:
        """
        Audit a Month-to-Month transaction file.

        Args:
            file_path: Path to transaction file
            generate_report: Whether to generate Excel report

        Returns:
            Dictionary with audit results and statistics
        """
        # Read and validate file
        file_data = self.file_reader.read_and_validate(file_path)

        if not file_data['is_valid']:
            return {
                'success': False,
                'error': file_data['error'],
                'filename': file_data['filename']
            }

        # Ensure we're using new format
        detected_format = file_data.get('format_type', 'old')
        if detected_format != 'new':
            return {
                'success': False,
                'error': "MTM audit requires new format (20-column) transaction data",
                'filename': file_data['filename']
            }

        # Update checker to new format
        self.checker = create_checker(self.membership_type, self.location, format_type='new')
        self.format_type = 'new'

        # Run MTM transaction audit
        mtm_results = self.audit_month_to_month_transactions(file_data['data_rows'])

        # Get column mapping for BP detection
        column_mapping = self.checker.get_bp_detection_columns()

        # Generate report if requested
        report_path = None
        if generate_report:
            original_name = Path(file_data['filename']).stem
            output_filename = f"{original_name}_MTM_Audit_Report.xlsx"

            report_path = self.report_generator.create_mtm_audit_report(
                header_row=file_data['header'],
                mtm_results=mtm_results,
                output_filename=output_filename,
                column_mapping=column_mapping,
                bp_config=self.bp_config
            )

        return {
            'success': True,
            'filename': file_data['filename'],
            'format_type': detected_format,
            'total_members': mtm_results['total_members'],
            'flagged_members': mtm_results['flagged_members'],
            'total_transactions': mtm_results['total_transactions'],
            'member_results': mtm_results['member_results'],
            'report_path': report_path
        }

    def audit_mtm_uploaded_file(self, uploaded_file, generate_report: bool = True) -> Dict[str, Any]:
        """
        Audit an uploaded Month-to-Month transaction file.

        Args:
            uploaded_file: Streamlit UploadedFile object
            generate_report: Whether to generate Excel report

        Returns:
            Dictionary with audit results and statistics
        """
        try:
            file_data = self.file_reader.read_and_validate_upload(uploaded_file)
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'filename': uploaded_file.name if hasattr(uploaded_file, 'name') else 'Unknown'
            }

        if not file_data['is_valid']:
            return {
                'success': False,
                'error': file_data['error'],
                'filename': file_data['filename']
            }

        # Ensure we're using new format
        detected_format = file_data.get('format_type', 'old')
        if detected_format != 'new':
            return {
                'success': False,
                'error': "MTM audit requires new format (20-column) transaction data",
                'filename': file_data['filename']
            }

        # Update checker to new format
        self.checker = create_checker(self.membership_type, self.location, format_type='new')
        self.format_type = 'new'

        # Run MTM transaction audit
        mtm_results = self.audit_month_to_month_transactions(file_data['data_rows'])

        # Get column mapping for BP detection
        column_mapping = self.checker.get_bp_detection_columns()

        # Generate report if requested
        report_path = None
        if generate_report:
            original_name = Path(file_data['filename']).stem
            output_filename = f"{original_name}_MTM_Audit_Report.xlsx"

            report_path = self.report_generator.create_mtm_audit_report(
                header_row=file_data['header'],
                mtm_results=mtm_results,
                output_filename=output_filename,
                column_mapping=column_mapping,
                bp_config=self.bp_config
            )

        return {
            'success': True,
            'filename': file_data['filename'],
            'format_type': detected_format,
            'total_members': mtm_results['total_members'],
            'flagged_members': mtm_results['flagged_members'],
            'total_transactions': mtm_results['total_transactions'],
            'member_results': mtm_results['member_results'],
            'report_path': report_path
        }

    def audit_all_membership_types_uploaded(self, uploaded_file, generate_report: bool = True) -> Dict[str, Any]:
        """
        Audit an uploaded file containing all membership types.
        Groups data by member_type column, applies appropriate rules per type,
        and generates a multi-tab Excel report.

        Args:
            uploaded_file: Streamlit UploadedFile object
            generate_report: Whether to generate Excel report

        Returns:
            Dictionary with audit results grouped by member_type
        """
        try:
            file_data = self.file_reader.read_and_validate_upload(uploaded_file)
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'filename': uploaded_file.name if hasattr(uploaded_file, 'name') else 'Unknown'
            }

        if not file_data['is_valid']:
            return {
                'success': False,
                'error': file_data['error'],
                'filename': file_data['filename']
            }

        # This requires new format (20-column) for member_type column
        detected_format = file_data.get('format_type', 'old')
        if detected_format != 'new':
            return {
                'success': False,
                'error': "All membership types audit requires new format (20-column) data with member_type column",
                'filename': file_data['filename']
            }

        # Load config for member_type mapping
        config = load_config()
        member_type_mapping = config.get('member_type_mapping', {
            '1MCORE': '1_month_paid_in_full',
            '1YRCORE': '1_year_paid_in_full',
            '3MCORE': '3_months_paid_in_full',
            'MTMCORE': 'month_to_month'
        })

        # Get member_type column index (index 9 in new format)
        member_type_col = 9  # member_type column in new format

        # Group rows by member_type
        rows_by_type = defaultdict(list)
        for row in file_data['data_rows']:
            if len(row) > member_type_col:
                member_type = row[member_type_col].strip().upper()
                rows_by_type[member_type].append(row)
            else:
                rows_by_type['UNKNOWN'].append(row)

        # Process each member_type group
        type_results = {}
        total_records = 0
        total_flagged = 0
        total_financial_impact = 0

        for member_type, rows in rows_by_type.items():
            config_key = member_type_mapping.get(member_type)

            if config_key == '1_month_paid_in_full':
                # 1MCORE - use transaction-based audit with payment balance check
                type_checker = create_checker(config_key, self.location, format_type='new')
                self.checker = type_checker
                expected_price = type_checker.expected_dues

                onemcore_results = self.audit_1mcore_transactions(rows, expected_price)

                # Build audit_results from member_results for consistent reporting
                audit_results = []
                for member_data in onemcore_results['member_results'].values():
                    # Create an audit result for each member (using first row as representative)
                    first_row = member_data['first_row']
                    red_flags = member_data['flags']

                    result = {
                        'row_data': first_row,
                        'red_flags': red_flags,
                        'has_flags': member_data['has_flags'],
                        'flag_count': member_data['flag_count'],
                        'membership_age': type_checker.calculate_membership_age(first_row),
                        'is_expired': type_checker.is_membership_expired(first_row),
                        'financial_impact': member_data['total_amount'] if member_data['total_amount'] > 0 else 0,
                        'dues_impact': 0,
                        'balance_impact': member_data['total_amount'] if member_data['total_amount'] > 0 else 0,
                        'member_id': member_data['member_number'],
                        'member_name': f"{member_data['first_name']} {member_data['last_name']}"
                    }
                    audit_results.append(result)

                flagged_count = onemcore_results['flagged_members']
                type_financial_impact = sum(r['financial_impact'] for r in audit_results)

                type_results[member_type] = {
                    'config_key': config_key,
                    'is_known_type': True,
                    'has_rules': True,
                    'is_1mcore': True,
                    'total_records': len(rows),
                    'total_members': onemcore_results['total_members'],
                    'flagged_count': flagged_count,
                    'flagged_percentage': (flagged_count / onemcore_results['total_members'] * 100) if onemcore_results['total_members'] > 0 else 0,
                    'financial_impact': type_financial_impact,
                    'audit_results': audit_results,
                    'member_results': onemcore_results['member_results'],
                    'rows': rows
                }

                total_records += len(rows)
                total_flagged += flagged_count
                total_financial_impact += type_financial_impact

            elif config_key and config_key != 'month_to_month':
                # Known non-MTM type (not 1MCORE) - apply standard audit rules
                type_checker = create_checker(config_key, self.location, format_type='new')
                audit_results = []

                for row in rows:
                    red_flags = type_checker.check_all(row)
                    membership_age = type_checker.calculate_membership_age(row)
                    is_expired = type_checker.is_membership_expired(row)
                    financial_impact = type_checker.get_financial_impact(row, red_flags)
                    impact_breakdown = type_checker.get_financial_impact_breakdown(row, red_flags)

                    result = {
                        'row_data': row,
                        'red_flags': red_flags,
                        'has_flags': len(red_flags) > 0,
                        'flag_count': len(red_flags),
                        'membership_age': membership_age,
                        'is_expired': is_expired,
                        'financial_impact': financial_impact,
                        'dues_impact': impact_breakdown['dues_impact'],
                        'balance_impact': impact_breakdown['balance_impact'],
                        'member_id': row[type_checker.COL_MEMBER] if len(row) > type_checker.COL_MEMBER else '',
                        'member_name': f"{row[type_checker.COL_FIRST_NAME]} {row[type_checker.COL_LAST_NAME]}" if len(row) > type_checker.COL_FIRST_NAME else ''
                    }
                    audit_results.append(result)

                flagged_count = sum(1 for r in audit_results if r['has_flags'])
                type_financial_impact = sum(r['financial_impact'] for r in audit_results)

                type_results[member_type] = {
                    'config_key': config_key,
                    'is_known_type': True,
                    'has_rules': True,
                    'total_records': len(rows),
                    'flagged_count': flagged_count,
                    'flagged_percentage': (flagged_count / len(rows) * 100) if rows else 0,
                    'financial_impact': type_financial_impact,
                    'audit_results': audit_results,
                    'rows': rows
                }

                total_records += len(rows)
                total_flagged += flagged_count
                total_financial_impact += type_financial_impact

            elif config_key == 'month_to_month':
                # MTM type - use transaction-based MTM audit
                mtm_checker = create_checker('month_to_month', self.location, format_type='new')
                self.checker = mtm_checker
                mtm_results = self.audit_month_to_month_transactions(rows)

                type_results[member_type] = {
                    'config_key': config_key,
                    'is_known_type': True,
                    'has_rules': True,
                    'is_mtm': True,
                    'total_records': len(rows),
                    'total_members': mtm_results['total_members'],
                    'flagged_members': mtm_results['flagged_members'],
                    'total_transactions': mtm_results['total_transactions'],
                    'member_results': mtm_results['member_results'],
                    'rows': rows
                }

                total_records += len(rows)
                total_flagged += mtm_results['flagged_members']

            else:
                # Unknown type - just group data without rules
                type_results[member_type] = {
                    'config_key': None,
                    'is_known_type': False,
                    'has_rules': False,
                    'total_records': len(rows),
                    'flagged_count': 0,
                    'flagged_percentage': 0,
                    'financial_impact': 0,
                    'audit_results': [{'row_data': row, 'red_flags': [], 'has_flags': False} for row in rows],
                    'rows': rows
                }

                total_records += len(rows)

        # Get column mapping for BP detection
        temp_checker = create_checker('1_year_paid_in_full', self.location, format_type='new')
        column_mapping = temp_checker.get_bp_detection_columns()

        # Generate report if requested
        report_path = None
        individual_file_paths = {}
        if generate_report:
            original_name = Path(file_data['filename']).stem
            output_filename = f"{original_name}_All_Types_Audit_Report.xlsx"

            report_path = self.report_generator.create_all_types_report(
                header_row=file_data['header'],
                type_results=type_results,
                output_filename=output_filename,
                column_mapping=column_mapping,
                bp_config=self.bp_config,
                member_type_mapping=member_type_mapping
            )

            # Generate individual files for each member_type
            individual_file_paths = self.report_generator.create_individual_type_files(
                header_row=file_data['header'],
                type_results=type_results,
                base_filename=original_name,
                member_type_mapping=member_type_mapping,
                column_mapping=column_mapping,
                bp_config=self.bp_config
            )

        return {
            'success': True,
            'filename': file_data['filename'],
            'format_type': detected_format,
            'total_records': total_records,
            'total_flagged': total_flagged,
            'total_financial_impact': total_financial_impact,
            'type_results': type_results,
            'member_types_found': list(rows_by_type.keys()),
            'report_path': report_path,
            'individual_file_paths': individual_file_paths
        }

    def _clean_date_format(self, date_str: str) -> str:
        """
        Clean date string: remove timestamp and return M/D/YYYY format.
        Does NOT change the year - just cleans the format.

        Args:
            date_str: Date string in various formats

        Returns:
            Clean date string in M/D/YYYY format, or original if unparseable
        """
        if not date_str or not date_str.strip():
            return date_str

        date_str = date_str.strip()

        # Skip if it looks like 'nan' or empty
        if date_str.lower() in ('nan', 'nat', 'none', ''):
            return ''

        try:
            parsed_date = None

            # Try multiple date formats
            date_formats = [
                '%Y-%m-%d %H:%M:%S',  # 1999-12-31 00:00:00 (pandas default)
                '%Y-%m-%d',            # 1999-12-31
                '%m/%d/%Y %H:%M:%S',   # 12/31/1999 00:00:00
                '%m/%d/%Y',            # 12/31/1999
                '%m/%d/%y',            # 12/31/99
            ]

            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue

            if parsed_date is None:
                return date_str  # Couldn't parse, return original

            # Return in clean M/D/YYYY format (no timestamp)
            return f"{parsed_date.month}/{parsed_date.day}/{parsed_date.year}"

        except Exception:
            return date_str

    def _fix_1999_year_in_date(self, date_str: str) -> str:
        """
        Fix dates where year is 1999 (gym software exports 2099 as 1999).
        Also removes any timestamp component and returns clean M/D/YYYY format.

        Args:
            date_str: Date string in various formats (M/D/YY, YYYY-MM-DD HH:MM:SS, etc.)

        Returns:
            Fixed date string in M/D/YYYY format without timestamp, or original if unparseable
        """
        if not date_str or not date_str.strip():
            return date_str

        date_str = date_str.strip()

        # Skip if it looks like 'nan' or empty
        if date_str.lower() in ('nan', 'nat', 'none', ''):
            return ''

        try:
            parsed_date = None

            # Try multiple date formats
            date_formats = [
                '%Y-%m-%d %H:%M:%S',  # 1999-12-31 00:00:00 (pandas default)
                '%Y-%m-%d',            # 1999-12-31
                '%m/%d/%Y %H:%M:%S',   # 12/31/1999 00:00:00
                '%m/%d/%Y',            # 12/31/1999
                '%m/%d/%y',            # 12/31/99
            ]

            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue

            if parsed_date is None:
                return date_str  # Couldn't parse, return original

            # Fix 1999 -> 2099
            if parsed_date.year == 1999:
                parsed_date = parsed_date.replace(year=2099)

            # Return in clean M/D/YYYY format (no timestamp)
            # Use manual formatting for cross-platform compatibility (%-m doesn't work on Windows)
            return f"{parsed_date.month}/{parsed_date.day}/{parsed_date.year}"

        except Exception:
            return date_str

    def split_file_by_membership_type_uploaded(self, uploaded_file) -> Dict[str, Any]:
        """
        Split an uploaded file by member_type column into separate raw data files.
        No rules applied - just splits data and fixes 1999->2099 dates.

        Args:
            uploaded_file: Streamlit UploadedFile object

        Returns:
            Dictionary with split data, counts, and verification info
        """
        try:
            file_data = self.file_reader.read_and_validate_upload(uploaded_file)
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'filename': uploaded_file.name if hasattr(uploaded_file, 'name') else 'Unknown'
            }

        if not file_data['is_valid']:
            return {
                'success': False,
                'error': file_data['error'],
                'filename': file_data['filename']
            }

        # This requires new format (20-column) for member_type column
        detected_format = file_data.get('format_type', 'old')
        if detected_format != 'new':
            return {
                'success': False,
                'error': "Split by membership type requires new format (20-column) data with member_type column",
                'filename': file_data['filename']
            }

        # Column indices in new format
        MEMBER_TYPE_COL = 9        # member_type
        TRANSACTION_DATE_COL = 3   # transaction_date
        JOIN_DATE_COL = 7          # join_date
        EXPIRATION_DATE_COL = 8    # expiration_date
        START_DRAFT_COL = 15       # start_draft
        END_DRAFT_COL = 16         # end_draft
        CONTRACT_DATE_COL = 17     # contract_date

        # Columns that need 1999->2099 fix (future dates that got exported as 1999)
        FIX_1999_COLS = [EXPIRATION_DATE_COL, START_DRAFT_COL, END_DRAFT_COL, CONTRACT_DATE_COL]

        # All date columns that need timestamp removal
        ALL_DATE_COLS = [TRANSACTION_DATE_COL, JOIN_DATE_COL, EXPIRATION_DATE_COL, START_DRAFT_COL, END_DRAFT_COL, CONTRACT_DATE_COL]

        # Step 1: Count original rows
        data_rows = file_data['data_rows']
        original_row_count = len(data_rows)

        # Step 2: Initialize grouping dict
        rows_by_type = defaultdict(list)
        processed_count = 0

        # Step 3: Process each row
        for row in data_rows:
            # Get member_type from column 9
            if len(row) > MEMBER_TYPE_COL:
                member_type = row[MEMBER_TYPE_COL].strip().upper()
                if not member_type:
                    member_type = 'UNKNOWN'
            else:
                member_type = 'UNKNOWN'

            # Create a copy of the row for modification
            fixed_row = list(row)

            # Clean all date columns (remove timestamps)
            for col_idx in ALL_DATE_COLS:
                if len(fixed_row) > col_idx:
                    if col_idx in FIX_1999_COLS:
                        # Fix 1999->2099 AND clean timestamp
                        fixed_row[col_idx] = self._fix_1999_year_in_date(str(fixed_row[col_idx]))
                    else:
                        # Just clean timestamp
                        fixed_row[col_idx] = self._clean_date_format(str(fixed_row[col_idx]))

            # Add row to appropriate group
            rows_by_type[member_type].append(fixed_row)
            processed_count += 1

        # Step 4: Verify data integrity
        split_total = sum(len(rows) for rows in rows_by_type.values())

        if split_total != original_row_count:
            return {
                'success': False,
                'error': f"DATA INTEGRITY ERROR: Original={original_row_count}, Split total={split_total}, Missing={original_row_count - split_total} rows",
                'filename': file_data['filename'],
                'original_row_count': original_row_count,
                'split_total': split_total
            }

        # Step 5: Build counts per type for display
        type_counts = {}
        for member_type, rows in rows_by_type.items():
            type_counts[member_type] = len(rows)

        return {
            'success': True,
            'filename': file_data['filename'],
            'format_type': detected_format,
            'header_row': file_data['header'],
            'original_row_count': original_row_count,
            'split_total': split_total,
            'rows_by_type': dict(rows_by_type),  # Convert defaultdict to regular dict
            'type_counts': type_counts,
            'member_types_found': list(rows_by_type.keys()),
            'verification_passed': split_total == original_row_count
        }
