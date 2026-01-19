"""
Audit Engine Module
Orchestrates the membership data audit process
Supports multiple membership types and locations
"""

from typing import List, Dict, Any, Optional
from pathlib import Path
from .red_flags import RedFlagChecker, create_checker
from .file_handler import MembershipFileReader
from .report_generator import AuditReportGenerator


class AuditEngine:
    """Main audit orchestrator"""

    def __init__(
        self,
        membership_type: str,
        location: str,
        output_folder: str = 'outputs'
    ):
        """
        Initialize audit engine with membership type and location

        Args:
            membership_type: Key from config (e.g., '1_year_paid_in_full', '3_months_paid_in_full')
            location: Key from config (e.g., 'bqe', 'greenpoint', 'lic')
            output_folder: Directory for output reports
        """
        self.membership_type = membership_type
        self.location = location
        self.checker = create_checker(membership_type, location)
        self.file_reader = MembershipFileReader()
        self.report_generator = AuditReportGenerator(output_folder)

    def audit_rows(self, data_rows: List[List[str]]) -> List[Dict[str, Any]]:
        """
        Audit all data rows and collect results

        Args:
            data_rows: List of data rows to audit

        Returns:
            List of audit results, one per row
        """
        audit_results = []

        for row in data_rows:
            # Check for red flags
            red_flags = self.checker.check_all(row)

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
                include_summary_sheet=True
            )

        return {
            'success': True,
            'filename': file_data['filename'],
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
                include_summary_sheet=True
            )

        return {
            'success': True,
            'filename': file_data['filename'],
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
