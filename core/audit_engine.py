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

    def audit_pif_grouped(self, data_rows: List[List[str]], expected_price: float = None) -> Dict[str, Any]:
        """
        Audit PIF transactions grouped by member.

        For each member_number:
        1. Collect all transactions
        2. Verify member names are consistent (flag if mismatch)
        3. Calculate net balance (sum of all amounts)
        4. Run standard checks on first row (dates, dues, etc.)
        5. Flag if net balance != 0

        Args:
            data_rows: List of transaction rows (new format)
            expected_price: Expected price for this membership type at this location

        Returns:
            Dictionary with member_results keyed by member_number
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

            # Get member info from first row
            first_name = first_row[cols['first_name']] if cols['first_name'] < len(first_row) else ''
            last_name = first_row[cols['last_name']] if cols['last_name'] < len(first_row) else ''

            all_flags = []
            net_balance = 0.0
            low_amounts = []

            # Check member name consistency across all transactions
            name_variants = set()
            for row in transactions:
                fn = row[cols['first_name']].strip() if cols['first_name'] < len(row) else ''
                ln = row[cols['last_name']].strip() if cols['last_name'] < len(row) else ''
                name_variants.add(f"{fn} {ln}")

            name_mismatch = len(name_variants) > 1
            if name_mismatch:
                all_flags.append(RedFlag(
                    "member_name_mismatch",
                    f"Different names found for member {member_number}: {', '.join(sorted(name_variants))}",
                    list(name_variants)
                ))

            # Calculate thresholds for price checking
            if expected_price and expected_price > 0:
                threshold_percent = 90
                min_expected = expected_price * (threshold_percent / 100)
            else:
                threshold_percent = 90
                min_expected = 0

            # Process each transaction
            for row in transactions:
                amount_str = row[cols['amount']] if cols['amount'] < len(row) else ''
                amount = self._parse_currency(amount_str)

                if amount is not None:
                    net_balance += amount
                    abs_amount = abs(amount)

                    # Check if amount is less than 90% of expected price
                    if expected_price and expected_price > 0 and abs_amount < min_expected:
                        low_amounts.append(amount)

            # Flag if unpaid balance (charges exceed payments)
            if net_balance > 0.01:
                all_flags.append(RedFlag(
                    "unpaid_balance",
                    f"Net balance ${net_balance:.2f} - charge without matching payment",
                    net_balance
                ))

            # Flag if overpayment (payments exceed charges)
            if net_balance < -0.01:
                all_flags.append(RedFlag(
                    "overpayment",
                    f"Net balance -${abs(net_balance):.2f} - payment exceeds charges",
                    net_balance
                ))

            # Flag low amounts
            for low_amount in low_amounts:
                abs_amt = abs(low_amount)
                txn_type = "Charge" if low_amount > 0 else "Payment"
                all_flags.append(RedFlag(
                    "low_amount",
                    f"{txn_type} ${abs_amt:.2f} is less than {threshold_percent}% of expected ${expected_price:.2f} (min ${min_expected:.2f})",
                    low_amount
                ))

            # Run basic date/expiration checks on first row
            basic_flags = self.checker.check_all(first_row)
            basic_flags = [f for f in basic_flags if f.flag_type != 'date_invalid' or self._parse_date(first_row[cols['join_date']]) is None]
            all_flags.extend(basic_flags)

            member_results[member_number] = {
                'member_number': member_number,
                'first_name': first_name,
                'last_name': last_name,
                'transactions': transactions,
                'transaction_count': len(transactions),
                'net_balance': net_balance,
                'flags': all_flags,
                'has_flags': len(all_flags) > 0,
                'flag_count': len(all_flags),
                'low_amounts': low_amounts,
                'name_mismatch': name_mismatch,
                'name_variants': list(name_variants),
                'first_row': first_row
            }

        return {
            'member_results': member_results,
            'total_members': len(member_results),
            'flagged_members': sum(1 for r in member_results.values() if r['has_flags']),
            'total_transactions': sum(len(r['transactions']) for r in member_results.values())
        }

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

        # Use grouped approach for new format, row-by-row for old format
        if detected_format == 'new':
            return self._audit_file_grouped(file_data, generate_report)

        # Old format: row-by-row audit
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

    def _audit_file_grouped(self, file_data: Dict[str, Any], generate_report: bool = True) -> Dict[str, Any]:
        """
        Audit a file using grouped member matching (new format only).

        Args:
            file_data: Validated file data dict from file_reader
            generate_report: Whether to generate Excel report

        Returns:
            Dictionary with audit results and statistics
        """
        expected_price = self.checker.expected_dues
        grouped_results = self.audit_pif_grouped(file_data['data_rows'], expected_price)

        # Build flat audit_results list from grouped results (for compatibility)
        audit_results = []
        for member_data in grouped_results['member_results'].values():
            first_row = member_data['first_row']
            red_flags = member_data['flags']

            result = {
                'row_data': first_row,
                'red_flags': red_flags,
                'has_flags': member_data['has_flags'],
                'flag_count': member_data['flag_count'],
                'membership_age': self.checker.calculate_membership_age(first_row),
                'is_expired': self.checker.is_membership_expired(first_row),
                'financial_impact': member_data['net_balance'] if member_data['net_balance'] > 0 else 0,
                'dues_impact': 0,
                'balance_impact': member_data['net_balance'] if member_data['net_balance'] > 0 else 0,
                'member_id': member_data['member_number'],
                'member_name': f"{member_data['first_name']} {member_data['last_name']}"
            }
            audit_results.append(result)

        # Calculate statistics
        total_records = len(file_data['data_rows'])
        flagged_count = grouped_results['flagged_members']
        clean_count = grouped_results['total_members'] - flagged_count
        flagged_percentage = (flagged_count / grouped_results['total_members'] * 100) if grouped_results['total_members'] > 0 else 0
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
            original_name = Path(file_data['filename']).stem
            output_filename = f"{original_name}_Audit_Report.xlsx"

            report_path = self.report_generator.create_grouped_audit_report(
                header_row=file_data['header'],
                grouped_results=grouped_results,
                output_filename=output_filename,
                column_mapping=column_mapping,
                bp_config=self.bp_config
            )

        return {
            'success': True,
            'filename': file_data['filename'],
            'format_type': 'new',
            'total_records': total_records,
            'total_members': grouped_results['total_members'],
            'flagged_count': flagged_count,
            'clean_count': clean_count,
            'flagged_percentage': flagged_percentage,
            'total_financial_impact': total_financial_impact,
            'total_dues_impact': total_dues_impact,
            'total_balance_impact': total_balance_impact,
            'flagged_member_ids': flagged_member_ids,
            'audit_results': audit_results,
            'member_results': grouped_results['member_results'],
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

        # Use grouped approach for new format, row-by-row for old format
        if detected_format == 'new':
            return self._audit_file_grouped(file_data, generate_report)

        # Old format: row-by-row audit
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
        """Parse date in multiple formats"""
        if not date_str or not date_str.strip():
            return None

        date_str = date_str.strip()

        # Skip if it looks like 'nan' or empty
        if date_str.lower() in ('nan', 'nat', 'none', ''):
            return None

        # Try multiple date formats
        date_formats = [
            '%m/%d/%y',            # 1/15/25
            '%m/%d/%Y',            # 1/15/2025
            '%Y-%m-%d %H:%M:%S',   # 2025-01-15 00:00:00
            '%Y-%m-%d',            # 2025-01-15
            '%m/%d/%Y %H:%M:%S',   # 1/15/2025 00:00:00
        ]

        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue

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

    def _detect_enrollment_fee(
        self,
        transactions: List[List[str]],
        enrollment_fee: float,
        enrollment_keyword: str
    ) -> tuple:
        """
        Check if member has enrollment fee transaction.
        Looks for transaction with amount matching enrollment_fee AND
        transaction_reference containing enrollment_keyword.

        Returns:
            Tuple of (found: bool, transaction_date: datetime or None)
        """
        cols = self.checker.NEW_FORMAT_COLUMNS
        ref_col = cols.get('transaction_reference', 19)  # transaction_reference column

        for txn in transactions:
            amount = self._parse_currency(txn[cols['amount']]) if cols['amount'] < len(txn) else None
            ref = txn[ref_col].strip().upper() if ref_col < len(txn) else ''

            # Check for enrollment fee: amount matches AND keyword in reference
            if amount is not None and abs(amount - enrollment_fee) < 0.01:
                if enrollment_keyword.upper() in ref:
                    txn_date = self._parse_date(txn[cols['transaction_date']])
                    return (True, txn_date)

        return (False, None)

    def _detect_initial_payment(
        self,
        transactions: List[List[str]],
        threshold: float
    ) -> tuple:
        """
        Detect large initial payment (prorated + 2 months).
        Looks for first transaction with amount >= threshold.

        Returns:
            Tuple of (found: bool, date: datetime or None, amount: float or None)
        """
        cols = self.checker.NEW_FORMAT_COLUMNS

        # Sort transactions by date to find the first large payment
        dated_txns = []
        for txn in transactions:
            txn_date = self._parse_date(txn[cols['transaction_date']])
            amount = self._parse_currency(txn[cols['amount']])
            if txn_date and amount is not None:
                dated_txns.append((txn_date, amount, txn))

        dated_txns.sort(key=lambda x: x[0])

        for txn_date, amount, txn in dated_txns:
            # Initial payment is typically a large negative (payment) or could be charge
            # We look for absolute amount >= threshold
            if abs(amount) >= threshold:
                return (True, txn_date, amount)

        return (False, None, None)

    def _is_annual_fee_transaction(
        self,
        row: List[str],
        keyword: str,
        min_amount: float,
        max_amount: float
    ) -> bool:
        """
        Check if transaction is an annual fee.
        Based on transaction_reference containing keyword AND amount in range.
        """
        cols = self.checker.NEW_FORMAT_COLUMNS
        ref_col = cols.get('transaction_reference', 19)

        amount = self._parse_currency(row[cols['amount']]) if cols['amount'] < len(row) else None
        ref = row[ref_col].strip().upper() if ref_col < len(row) else ''

        if amount is None:
            return False

        abs_amount = abs(amount)
        return keyword.upper() in ref and min_amount <= abs_amount <= max_amount

    def _check_mtm_charge_payment_pairs(
        self,
        transactions: List[List[str]]
    ) -> List[Dict[str, Any]]:
        """
        Verify each charge has matching payment from same member.
        Returns list of unmatched transactions needing verification.

        Charges (positive amounts) should have adjacent payment (negative amount)
        from the same member within reasonable time window.
        """
        cols = self.checker.NEW_FORMAT_COLUMNS
        unmatched = []

        # Group by member key and sort by date
        txn_data = []
        for i, txn in enumerate(transactions):
            txn_date = self._parse_date(txn[cols['transaction_date']])
            amount = self._parse_currency(txn[cols['amount']])
            if txn_date and amount is not None:
                txn_data.append({
                    'index': i,
                    'date': txn_date,
                    'amount': amount,
                    'txn': txn,
                    'matched': False
                })

        txn_data.sort(key=lambda x: x['date'])

        # Find charges without matching payments
        for i, txn in enumerate(txn_data):
            if txn['amount'] > 0:  # This is a charge
                charge_amount = txn['amount']
                charge_date = txn['date']

                # Look for matching payment (within 7 days)
                matched = False
                for j, other in enumerate(txn_data):
                    if i == j or other['matched']:
                        continue
                    if other['amount'] < 0:  # This is a payment
                        payment_amount = abs(other['amount'])
                        days_diff = abs((other['date'] - charge_date).days)

                        # Match if amounts are close and within 7 days
                        if abs(charge_amount - payment_amount) < 0.01 and days_diff <= 7:
                            txn['matched'] = True
                            other['matched'] = True
                            matched = True
                            break

                if not matched:
                    unmatched.append({
                        'transaction': txn['txn'],
                        'date': txn['date'],
                        'amount': txn['amount'],
                        'type': 'charge_without_payment'
                    })

        return unmatched

    def _calculate_monthly_coverage_start(
        self,
        join_date: datetime,
        has_initial_payment: bool,
        initial_payment_date: datetime,
        report_start: datetime,
        covers_months: int
    ) -> datetime:
        """
        Determine when monthly payment checks should begin.

        Logic:
        - If initial payment found: coverage starts after initial_payment_covers_months
        - If no initial payment: coverage starts at report_start_date or join_date, whichever is later
        """
        if has_initial_payment and initial_payment_date:
            # Coverage starts after the initial payment covers its months
            coverage_start = initial_payment_date + relativedelta(months=covers_months)
        else:
            # No initial payment detected, use join_date or report_start
            coverage_start = max(join_date, report_start) if join_date else report_start

        return coverage_start

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

            # Run basic date/expiration checks on first row
            basic_flags = self.checker.check_all(first_row)
            # Filter out any existing date_invalid flags since we're handling dates better now
            basic_flags = [f for f in basic_flags if f.flag_type != 'date_invalid' or self._parse_date(first_row[cols['join_date']]) is None]
            all_flags.extend(basic_flags)

            # Check member name consistency
            name_variants = set()
            for row in transactions:
                fn = row[cols['first_name']].strip() if cols['first_name'] < len(row) else ''
                ln = row[cols['last_name']].strip() if cols['last_name'] < len(row) else ''
                name_variants.add(f"{fn} {ln}")

            name_mismatch = len(name_variants) > 1
            if name_mismatch:
                all_flags.append(RedFlag(
                    "member_name_mismatch",
                    f"Different names found for member {member_number}: {', '.join(sorted(name_variants))}",
                    list(name_variants)
                ))

            member_results[member_number] = {
                'member_number': member_number,
                'first_name': first_name,
                'last_name': last_name,
                'transactions': transactions,
                'transaction_count': len(transactions),
                'total_amount': total_amount,
                'net_balance': total_amount,
                'flags': all_flags,
                'has_flags': len(all_flags) > 0,
                'flag_count': len(all_flags),
                'price_mismatches': price_mismatches,
                'low_amounts': low_amounts,
                'name_mismatch': name_mismatch,
                'name_variants': list(name_variants),
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
        1. Charge/payment pair matching - verifies each charge has matching payment
        2. Member type detection - New (has enrollment fee) vs Existing
        3. Missing enrollment fee - flags new members (join >= report_start) without enrollment
        4. Initial payment detection - looks for larger payment (~$180+)
        5. Monthly payment verification - from coverage start to current date
        6. Annual fee tracking - informational only

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
        rules = self.checker.rules

        # Get pricing configuration for the location
        config = load_config()
        mtm_config = config.get('membership_types', {}).get('month_to_month', {})
        pricing = mtm_config.get('pricing', {}).get(self.location, {})

        # Extract pricing values (handle both old dict format and new nested format)
        if isinstance(pricing, dict):
            monthly_rate = pricing.get('monthly_rate', 59.99)
            enrollment_fee = pricing.get('enrollment_fee', 50.00)
            annual_fee_min = pricing.get('annual_fee_min', 24.95)
            annual_fee_max = pricing.get('annual_fee_max', 39.99)
        else:
            # Fallback to old format
            monthly_rate = 59.99
            enrollment_fee = 50.00
            annual_fee_min = 24.95
            annual_fee_max = 39.99

        # Get rule parameters
        initial_payment_threshold = rules.get('initial_payment_threshold', 150.00)
        initial_payment_covers_months = rules.get('initial_payment_covers_months', 3)
        enrollment_keyword = rules.get('enrollment_keyword', 'ENROLL')
        annual_fee_keyword = rules.get('annual_fee_keyword', 'ANNUAL FEES')
        report_start_str = rules.get('report_start_date', '2025-01-01')

        # Parse report start date
        try:
            report_start = datetime.strptime(report_start_str, '%Y-%m-%d')
        except:
            report_start = datetime(2025, 1, 1)

        # Check flags
        check_charge_payment = rules.get('check_charge_payment_matching', True)
        check_monthly = rules.get('check_monthly_payments', True)
        check_enrollment = rules.get('check_enrollment_fee', True)
        check_annual = rules.get('check_annual_fee', False)

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
                    'transaction_count': len(transactions),
                    'flags': [RedFlag("join_date_invalid", "Invalid join date")],
                    'missing_months': [],
                    'has_flags': True,
                    'flag_count': 1,
                    'member_type': 'Unknown',
                    'has_enrollment_fee': False,
                    'has_initial_payment': False,
                    'has_annual_fee': False,
                    'coverage_start': None,
                    'months_paid_count': 0
                }
                continue

            all_flags = []

            # --- Step 1: Check charge/payment pairs ---
            unmatched_charges = []
            if check_charge_payment:
                unmatched_charges = self._check_mtm_charge_payment_pairs(transactions)
                for unmatched in unmatched_charges:
                    all_flags.append(RedFlag(
                        "needs_verification",
                        f"Charge ${unmatched['amount']:.2f} on {unmatched['date'].strftime('%m/%d/%y')} without matching payment",
                        unmatched['amount']
                    ))

            # --- Step 2: Detect enrollment fee (determines member type) ---
            has_enrollment, enrollment_date = self._detect_enrollment_fee(
                transactions, enrollment_fee, enrollment_keyword
            )

            # Determine member type
            is_new_member = has_enrollment
            member_type = 'New' if is_new_member else 'Existing'

            # --- Step 3: Flag missing enrollment for new members ---
            if check_enrollment:
                # Only flag if join_date is on or after report_start AND no enrollment fee
                if join_date >= report_start and not has_enrollment:
                    all_flags.append(RedFlag(
                        "missing_enrollment_fee",
                        f"New member (joined {join_date.strftime('%m/%d/%y')}) without ${enrollment_fee:.2f} enrollment fee",
                        enrollment_fee
                    ))

            # --- Step 4: Detect initial payment ---
            has_initial, initial_date, initial_amount = self._detect_initial_payment(
                transactions, initial_payment_threshold
            )

            # --- Step 5: Calculate coverage start for monthly payment checks ---
            coverage_start = self._calculate_monthly_coverage_start(
                join_date=join_date,
                has_initial_payment=has_initial,
                initial_payment_date=initial_date,
                report_start=report_start,
                covers_months=initial_payment_covers_months
            )

            # --- Step 6: Check monthly payments from coverage_start to today ---
            missing_months = []
            months_paid = []

            if check_monthly:
                # Get payment months from transactions (where absolute amount >= monthly_rate)
                payment_months = defaultdict(list)
                for idx, txn in enumerate(transactions):
                    txn_date_str = txn[cols['transaction_date']]
                    txn_date = self._parse_date(txn_date_str)
                    amount = self._parse_currency(txn[cols['amount']])

                    if txn_date and amount is not None:
                        # Consider both charges and payments that could qualify as monthly payment
                        # Usually payments are negative, charges are positive
                        if abs(amount) >= monthly_rate * 0.9:  # 90% tolerance
                            month_key = self._get_month_key(txn_date)
                            payment_months[month_key].append(idx)

                months_paid = list(payment_months.keys())

                # Build required months from coverage_start to system_date
                required_months = []
                if coverage_start:
                    current_month = datetime(coverage_start.year, coverage_start.month, 1)
                    system_month = datetime(system_date.year, system_date.month, 1)

                    while current_month <= system_month:
                        required_months.append(self._get_month_key(current_month))
                        current_month += relativedelta(months=1)

                # Check for missing payments
                for required_month in required_months:
                    if required_month not in payment_months:
                        missing_months.append(required_month)
                        all_flags.append(RedFlag(
                            "missing_monthly_payment",
                            f"No qualifying payment (>= ${monthly_rate:.2f}) for {required_month}",
                            required_month
                        ))

            # --- Step 7: Track annual fee (informational) ---
            has_annual_fee = False
            if check_annual:
                for txn in transactions:
                    if self._is_annual_fee_transaction(txn, annual_fee_keyword, annual_fee_min, annual_fee_max):
                        has_annual_fee = True
                        break

                if not has_annual_fee:
                    all_flags.append(RedFlag(
                        "missing_annual_fee",
                        f"No annual fee transaction ({annual_fee_keyword}) found",
                        0
                    ))

            # --- Check member name consistency ---
            name_variants = set()
            for txn in transactions:
                fn = txn[cols['first_name']].strip() if cols['first_name'] < len(txn) else ''
                ln = txn[cols['last_name']].strip() if cols['last_name'] < len(txn) else ''
                name_variants.add(f"{fn} {ln}")

            name_mismatch = len(name_variants) > 1
            if name_mismatch:
                all_flags.append(RedFlag(
                    "member_name_mismatch",
                    f"Different names found for member {member_number}: {', '.join(sorted(name_variants))}",
                    list(name_variants)
                ))

            # --- Calculate net balance ---
            net_balance = 0.0
            for txn in transactions:
                amount = self._parse_currency(txn[cols['amount']]) if cols['amount'] < len(txn) else None
                if amount is not None:
                    net_balance += amount

            if net_balance > 0.01:
                all_flags.append(RedFlag(
                    "unpaid_balance",
                    f"Net balance ${net_balance:.2f} - charge without matching payment",
                    net_balance
                ))
            elif net_balance < -0.01:
                all_flags.append(RedFlag(
                    "overpayment",
                    f"Net balance -${abs(net_balance):.2f} - payment exceeds charges",
                    net_balance
                ))

            # --- Run basic validation rules (exp year, draft dates) ---
            basic_flags = self._check_basic_mtm_rules(first_row)
            all_flags.extend(basic_flags)

            # Build result
            member_results[member_key] = {
                'member_key': member_key,
                'first_name': first_name,
                'last_name': last_name,
                'member_number': member_number,
                'join_date': join_date,
                'transactions': transactions,
                'transaction_count': len(transactions),
                'net_balance': net_balance,
                'flags': all_flags,
                'has_flags': len(all_flags) > 0,
                'flag_count': len(all_flags),
                'name_mismatch': name_mismatch,
                'name_variants': list(name_variants),
                # New fields
                'member_type': member_type,
                'has_enrollment_fee': has_enrollment,
                'enrollment_date': enrollment_date,
                'has_initial_payment': has_initial,
                'initial_payment_date': initial_date,
                'initial_payment_amount': initial_amount,
                'coverage_start': coverage_start,
                'missing_months': missing_months,
                'months_paid': months_paid,
                'months_paid_count': len(months_paid),
                'has_annual_fee': has_annual_fee,
                'unmatched_charges': len(unmatched_charges)
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
                    'is_grouped': True,
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
                # Known non-MTM type (not 1MCORE) - use grouped member matching
                type_checker = create_checker(config_key, self.location, format_type='new')
                self.checker = type_checker
                expected_price = type_checker.expected_dues

                grouped_results = self.audit_pif_grouped(rows, expected_price)

                # Build audit_results from member_results for consistent reporting
                audit_results = []
                for member_data in grouped_results['member_results'].values():
                    first_row = member_data['first_row']
                    red_flags = member_data['flags']

                    result = {
                        'row_data': first_row,
                        'red_flags': red_flags,
                        'has_flags': member_data['has_flags'],
                        'flag_count': member_data['flag_count'],
                        'membership_age': type_checker.calculate_membership_age(first_row),
                        'is_expired': type_checker.is_membership_expired(first_row),
                        'financial_impact': member_data['net_balance'] if member_data['net_balance'] > 0 else 0,
                        'dues_impact': 0,
                        'balance_impact': member_data['net_balance'] if member_data['net_balance'] > 0 else 0,
                        'member_id': member_data['member_number'],
                        'member_name': f"{member_data['first_name']} {member_data['last_name']}"
                    }
                    audit_results.append(result)

                flagged_count = grouped_results['flagged_members']
                type_financial_impact = sum(r['financial_impact'] for r in audit_results)

                type_results[member_type] = {
                    'config_key': config_key,
                    'is_known_type': True,
                    'has_rules': True,
                    'is_grouped': True,
                    'total_records': len(rows),
                    'total_members': grouped_results['total_members'],
                    'flagged_count': flagged_count,
                    'flagged_percentage': (flagged_count / grouped_results['total_members'] * 100) if grouped_results['total_members'] > 0 else 0,
                    'financial_impact': type_financial_impact,
                    'audit_results': audit_results,
                    'member_results': grouped_results['member_results'],
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
