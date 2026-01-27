"""
Report Generator Module
Creates Excel audit reports with highlighting and formatting
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from typing import List, Dict, Any
from pathlib import Path


class AuditReportGenerator:
    """Generates formatted Excel audit reports"""

    # Styling
    HIGHLIGHT_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # Yellow (flagged)
    BP_FILL = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')  # Orange (BP accounts)
    HEADER_FILL = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')  # Gray (column headers)
    SECTION_FILL = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')  # Blue (section headers)
    BOLD_FONT = Font(bold=True)
    WHITE_FONT = Font(bold=True, color='FFFFFF')
    CENTER_ALIGN = Alignment(horizontal='center', vertical='center')

    def __init__(self, output_folder: str = 'outputs'):
        """
        Initialize report generator

        Args:
            output_folder: Directory to save reports
        """
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(exist_ok=True)

    def _is_bp_member(self, row: List[str], column_indices: List[int], bp_config: Dict[str, Any] = None) -> bool:
        """
        Check if row has BP/Billing indicators based on config.

        Args:
            row: Data row
            column_indices: List of column indices to check
            bp_config: BP detection configuration with 'keywords' and 'case_sensitive'

        Returns:
            True if member has BP/Billing indicator
        """
        if bp_config and not bp_config.get('enabled', True):
            return False

        keywords = ['bp', 'billing']
        case_sensitive = False

        if bp_config:
            keywords = bp_config.get('keywords', keywords)
            case_sensitive = bp_config.get('case_sensitive', False)

        for idx in column_indices:
            if idx is not None and idx >= 0 and idx < len(row):
                value = str(row[idx])
                if not case_sensitive:
                    value = value.lower()
                    check_keywords = [k.lower() for k in keywords]
                else:
                    check_keywords = keywords

                for keyword in check_keywords:
                    if keyword in value:
                        return True
        return False

    def _categorize_rows(
        self,
        data_rows: List[List[str]],
        audit_results: List[Dict[str, Any]],
        column_indices: List[int],
        bp_config: Dict[str, Any] = None
    ) -> tuple:
        """
        Split rows into flagged, BP, and valid categories.

        Args:
            data_rows: Original data rows
            audit_results: Audit results for each row
            column_indices: List of column indices to check for BP detection
            bp_config: BP detection configuration

        Returns:
            Tuple of (flagged, bp, valid) lists, each containing (row, result) tuples
        """
        flagged, bp, valid = [], [], []

        for row, result in zip(data_rows, audit_results):
            is_bp = self._is_bp_member(row, column_indices, bp_config)
            has_flags = bool(result.get('red_flags'))

            if is_bp:
                bp.append((row, result))  # BP takes priority
            elif has_flags:
                flagged.append((row, result))
            else:
                valid.append((row, result))

        return flagged, bp, valid

    def _write_section_header(self, sheet, row_num: int, header_text: str, col_count: int) -> int:
        """
        Write a section header row.

        Args:
            sheet: Excel worksheet
            row_num: Row number to write at
            header_text: Text for the section header
            col_count: Number of columns to merge

        Returns:
            Next available row number
        """
        cell = sheet.cell(row=row_num, column=1, value=header_text)
        cell.font = self.WHITE_FONT
        cell.fill = self.SECTION_FILL
        sheet.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=col_count)
        return row_num + 1

    def create_audit_report(
        self,
        header_row: List[str],
        data_rows: List[List[str]],
        audit_results: List[Dict[str, Any]],
        output_filename: str,
        include_summary_sheet: bool = True,
        column_mapping: Dict[str, int] = None,
        bp_config: Dict[str, Any] = None
    ) -> str:
        """
        Create Excel audit report with highlighted red flags organized into sections.

        Sections:
        1. FLAGGED ACCOUNTS - Members with red flags (excluding BP members)
        2. BILLING PROBLEM ACCOUNTS - Members with BP/Billing codes
        3. VALID ACCOUNTS - Members with no flags and no BP codes

        Args:
            header_row: Column headers
            data_rows: Original data rows
            audit_results: List of audit results for each row (from audit_engine)
            output_filename: Name for output file
            include_summary_sheet: Whether to add a summary sheet
            column_mapping: Dictionary with column name to index mappings
            bp_config: BP detection configuration with 'enabled', 'columns', 'keywords', 'case_sensitive'

        Returns:
            Full path to generated report file
        """
        wb = Workbook()

        # Create main audit sheet
        audit_sheet = wb.active
        audit_sheet.title = "Audit Report"

        # Add only "Notes" column to header (keep original columns intact)
        enhanced_header = header_row + ["Notes"]
        col_count = len(enhanced_header)

        # Write header row with formatting
        for col_idx, header_text in enumerate(enhanced_header, start=1):
            cell = audit_sheet.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.BOLD_FONT
            cell.fill = self.HEADER_FILL
            cell.alignment = self.CENTER_ALIGN

        # Get column indices for BP detection
        column_indices = []
        if column_mapping:
            # Use column_mapping to get indices for configured columns
            bp_columns = bp_config.get('columns', ['code', 'member_type']) if bp_config else ['code', 'member_type']
            for col_name in bp_columns:
                idx = column_mapping.get(col_name)
                if idx is not None and idx >= 0:
                    column_indices.append(idx)
        if not column_indices:
            # Default to old format indices for code (7) and member_type (5)
            column_indices = [7, 5]

        # Categorize rows into sections
        flagged_rows, bp_rows, valid_rows = self._categorize_rows(
            data_rows, audit_results, column_indices, bp_config
        )

        excel_row = 2
        flagged_count = 0
        bp_count = 0

        # Section 1: FLAGGED ACCOUNTS
        if flagged_rows:
            excel_row = self._write_section_header(
                audit_sheet, excel_row,
                f"FLAGGED ACCOUNTS ({len(flagged_rows)} records)",
                col_count
            )

            for data_row, audit_result in flagged_rows:
                red_flags = audit_result.get('red_flags', [])
                notes = " | ".join([str(flag) for flag in red_flags]) if red_flags else ""
                enhanced_row = data_row + [notes]

                for col_idx, value in enumerate(enhanced_row, start=1):
                    cell = audit_sheet.cell(row=excel_row, column=col_idx, value=value)
                    cell.fill = self.HIGHLIGHT_FILL  # Yellow

                flagged_count += 1
                excel_row += 1

            # Blank row after section
            excel_row += 1

        # Section 2: BILLING PROBLEM ACCOUNTS
        if bp_rows:
            excel_row = self._write_section_header(
                audit_sheet, excel_row,
                f"BILLING PROBLEM ACCOUNTS ({len(bp_rows)} records)",
                col_count
            )

            for data_row, audit_result in bp_rows:
                red_flags = audit_result.get('red_flags', [])
                notes = " | ".join([str(flag) for flag in red_flags]) if red_flags else ""
                enhanced_row = data_row + [notes]

                for col_idx, value in enumerate(enhanced_row, start=1):
                    cell = audit_sheet.cell(row=excel_row, column=col_idx, value=value)
                    cell.fill = self.BP_FILL  # Orange

                bp_count += 1
                excel_row += 1

            # Blank row after section
            excel_row += 1

        # Section 3: VALID ACCOUNTS
        if valid_rows:
            excel_row = self._write_section_header(
                audit_sheet, excel_row,
                f"VALID ACCOUNTS ({len(valid_rows)} records)",
                col_count
            )

            for data_row, audit_result in valid_rows:
                enhanced_row = data_row + [""]  # No notes for valid accounts

                for col_idx, value in enumerate(enhanced_row, start=1):
                    cell = audit_sheet.cell(row=excel_row, column=col_idx, value=value)
                    # No fill for valid accounts

                excel_row += 1

        # Auto-adjust column widths
        self._auto_adjust_column_widths(audit_sheet)

        # Add summary sheet if requested
        if include_summary_sheet:
            self._add_summary_sheet(wb, audit_results, flagged_count, len(data_rows), bp_count)

        # Save workbook
        output_path = self.output_folder / output_filename
        wb.save(output_path)

        return str(output_path)

    def _auto_adjust_column_widths(self, sheet):
        """Auto-adjust column widths based on content"""
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

            # Set width (max 50 for readability)
            adjusted_width = min(max_length + 2, 50)
            sheet.column_dimensions[column_letter].width = adjusted_width

    def _add_summary_sheet(self, workbook, audit_results: List[Dict[str, Any]], flagged_count: int, total_count: int, bp_count: int = 0):
        """Add a summary sheet with statistics"""
        summary_sheet = workbook.create_sheet("Summary", 0)  # Insert at beginning

        # Title
        summary_sheet['A1'] = "AUDIT SUMMARY"
        summary_sheet['A1'].font = Font(bold=True, size=16)

        # Overall stats
        row = 3
        summary_sheet[f'A{row}'] = "Total Records:"
        summary_sheet[f'B{row}'] = total_count
        summary_sheet[f'A{row}'].font = self.BOLD_FONT

        row += 1
        summary_sheet[f'A{row}'] = "Flagged Records:"
        summary_sheet[f'B{row}'] = flagged_count
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'].fill = self.HIGHLIGHT_FILL

        row += 1
        summary_sheet[f'A{row}'] = "Billing Problem Records:"
        summary_sheet[f'B{row}'] = bp_count
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'].fill = self.BP_FILL

        row += 1
        valid_count = total_count - flagged_count - bp_count
        summary_sheet[f'A{row}'] = "Valid Records:"
        summary_sheet[f'B{row}'] = valid_count
        summary_sheet[f'A{row}'].font = self.BOLD_FONT

        row += 1
        flagged_percentage = (flagged_count / total_count * 100) if total_count > 0 else 0
        summary_sheet[f'A{row}'] = "Flagged Percentage:"
        summary_sheet[f'B{row}'] = f"{flagged_percentage:.1f}%"
        summary_sheet[f'A{row}'].font = self.BOLD_FONT

        # Red flag breakdown
        row += 3
        summary_sheet[f'A{row}'] = "RED FLAG BREAKDOWN"
        summary_sheet[f'A{row}'].font = Font(bold=True, size=14)

        # Count red flags by type
        flag_counts = {}
        for result in audit_results:
            for flag in result.get('red_flags', []):
                flag_type = flag.flag_type
                flag_counts[flag_type] = flag_counts.get(flag_type, 0) + 1

        # Sort by count (descending)
        sorted_flags = sorted(flag_counts.items(), key=lambda x: x[1], reverse=True)

        row += 2
        summary_sheet[f'A{row}'] = "Red Flag Type"
        summary_sheet[f'B{row}'] = "Count"
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'].font = self.BOLD_FONT

        for flag_type, count in sorted_flags:
            row += 1
            summary_sheet[f'A{row}'] = self._format_flag_type(flag_type)
            summary_sheet[f'B{row}'] = count

        # Financial impact
        row += 3
        summary_sheet[f'A{row}'] = "FINANCIAL IMPACT"
        summary_sheet[f'A{row}'].font = Font(bold=True, size=14)

        total_financial_impact = sum(result.get('financial_impact', 0) for result in audit_results)

        row += 2
        summary_sheet[f'A{row}'] = "Total Potential Revenue at Risk:"
        summary_sheet[f'B{row}'] = f"${total_financial_impact:,.2f}"
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'].font = Font(bold=True, color="FF0000")

        # Auto-adjust widths
        summary_sheet.column_dimensions['A'].width = 35
        summary_sheet.column_dimensions['B'].width = 20

    def _format_flag_type(self, flag_type: str) -> str:
        """Format flag type for display"""
        # Convert snake_case to Title Case
        formatted = flag_type.replace('_', ' ').title()

        # Specific formatting
        replacements = {
            'Dues Low': 'Dues Below Minimum',
            'Dues Invalid': 'Invalid Dues Amount',
            'Date Mismatch': 'Join/Exp Date Mismatch',
            'Date Invalid': 'Invalid Date Format',
            'Pay Type Wrong': 'Incorrect Pay Type',
            'End Draft Wrong': 'Incorrect End Draft Date',
            'Cycle Wrong': 'Incorrect Cycle Number',
            'Cycle Invalid': 'Invalid Cycle Value',
            'Balance Debit': 'Outstanding Balance (Owed)',
            'Balance Credit': 'Credit Balance (Refund Due)',
            'Balance Invalid': 'Invalid Balance Value'
        }

        return replacements.get(formatted, formatted)

    def create_consolidated_report(
        self,
        file_results: List[Dict[str, Any]],
        output_filename: str = "Consolidated_Audit_Report.xlsx"
    ) -> str:
        """
        Create a consolidated report across multiple files

        Args:
            file_results: List of file audit results
            output_filename: Name for output file

        Returns:
            Full path to generated report file
        """
        wb = Workbook()
        summary_sheet = wb.active
        summary_sheet.title = "Overview"

        # Title
        summary_sheet['A1'] = "CONSOLIDATED AUDIT REPORT"
        summary_sheet['A1'].font = Font(bold=True, size=16)

        # Per-file summary
        row = 3
        summary_sheet['A3'] = "Filename"
        summary_sheet['B3'] = "Total Records"
        summary_sheet['C3'] = "Flagged"
        summary_sheet['D3'] = "Clean"
        summary_sheet['E3'] = "Flag %"
        summary_sheet['F3'] = "Financial Impact"

        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            summary_sheet[f'{col}3'].font = self.BOLD_FONT
            summary_sheet[f'{col}3'].fill = self.HEADER_FILL

        row = 4
        total_records = 0
        total_flagged = 0
        total_impact = 0

        for file_result in file_results:
            filename = file_result['filename']
            records = file_result['total_records']
            flagged = file_result['flagged_count']
            clean = records - flagged
            flag_pct = (flagged / records * 100) if records > 0 else 0
            impact = file_result.get('total_financial_impact', 0)

            summary_sheet[f'A{row}'] = filename
            summary_sheet[f'B{row}'] = records
            summary_sheet[f'C{row}'] = flagged
            summary_sheet[f'D{row}'] = clean
            summary_sheet[f'E{row}'] = f"{flag_pct:.1f}%"
            summary_sheet[f'F{row}'] = f"${impact:,.2f}"

            if flagged > 0:
                for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                    summary_sheet[f'{col}{row}'].fill = self.HIGHLIGHT_FILL

            total_records += records
            total_flagged += flagged
            total_impact += impact

            row += 1

        # Totals row
        row += 1
        summary_sheet[f'A{row}'] = "TOTALS"
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'] = total_records
        summary_sheet[f'C{row}'] = total_flagged
        summary_sheet[f'D{row}'] = total_records - total_flagged
        total_flag_pct = (total_flagged / total_records * 100) if total_records > 0 else 0
        summary_sheet[f'E{row}'] = f"{total_flag_pct:.1f}%"
        summary_sheet[f'F{row}'] = f"${total_impact:,.2f}"

        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            summary_sheet[f'{col}{row}'].font = self.BOLD_FONT

        # Auto-adjust widths
        summary_sheet.column_dimensions['A'].width = 40
        summary_sheet.column_dimensions['B'].width = 15
        summary_sheet.column_dimensions['C'].width = 15
        summary_sheet.column_dimensions['D'].width = 15
        summary_sheet.column_dimensions['E'].width = 12
        summary_sheet.column_dimensions['F'].width = 20

        # Save
        output_path = self.output_folder / output_filename
        wb.save(output_path)

        return str(output_path)

    def create_mtm_audit_report(
        self,
        header_row: List[str],
        mtm_results: Dict[str, Any],
        output_filename: str,
        column_mapping: Dict[str, int] = None,
        bp_config: Dict[str, Any] = None
    ) -> str:
        """
        Create Excel audit report for Month-to-Month transaction audits.

        Organized by member with sections:
        1. FLAGGED MEMBERS - Members with missing payments or other flags
        2. BILLING PROBLEM MEMBERS - Members with BP indicators
        3. VALID MEMBERS - Members with no issues

        Args:
            header_row: Column headers from original file
            mtm_results: Results from audit_month_to_month_transactions()
            output_filename: Name for output file
            column_mapping: Dictionary with column name to index mappings
            bp_config: BP detection configuration

        Returns:
            Full path to generated report file
        """
        wb = Workbook()

        # Create main audit sheet
        audit_sheet = wb.active
        audit_sheet.title = "MTM Audit Report"

        # Create member-level header (different from transaction header)
        member_header = [
            "First Name", "Last Name", "Member #", "Join Date",
            "Transaction Count", "Missing Months", "Duplicate Months", "Notes"
        ]
        col_count = len(member_header)

        # Write header row
        for col_idx, header_text in enumerate(member_header, start=1):
            cell = audit_sheet.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.BOLD_FONT
            cell.fill = self.HEADER_FILL
            cell.alignment = self.CENTER_ALIGN

        # Get column indices for BP detection
        column_indices = []
        if column_mapping:
            bp_columns = bp_config.get('columns', ['code', 'member_type', 'member_group']) if bp_config else ['code', 'member_type', 'member_group']
            for col_name in bp_columns:
                idx = column_mapping.get(col_name)
                if idx is not None and idx >= 0:
                    column_indices.append(idx)

        # Categorize members into flagged, BP, and valid
        member_results = mtm_results.get('member_results', {})
        flagged_members = []
        bp_members = []
        valid_members = []

        for member_key, result in member_results.items():
            # Check if any transaction has BP indicator
            is_bp = False
            for txn in result.get('transactions', []):
                if self._is_bp_member(txn, column_indices, bp_config):
                    is_bp = True
                    break

            if is_bp:
                bp_members.append(result)
            elif result.get('has_flags', False):
                flagged_members.append(result)
            else:
                valid_members.append(result)

        excel_row = 2
        flagged_count = 0
        bp_count = 0

        # Section 1: FLAGGED MEMBERS
        if flagged_members:
            excel_row = self._write_section_header(
                audit_sheet, excel_row,
                f"FLAGGED MEMBERS ({len(flagged_members)} members)",
                col_count
            )

            for result in flagged_members:
                # Format missing months
                missing_months = result.get('missing_months', [])
                missing_str = ', '.join(missing_months) if missing_months else ''

                # Format duplicate months
                duplicate_months = result.get('duplicate_months', [])
                duplicate_str = ', '.join(duplicate_months) if duplicate_months else ''

                # Format notes from flags
                flags = result.get('flags', [])
                notes = ' | '.join([str(flag) for flag in flags]) if flags else ''

                # Format join date
                join_date = result.get('join_date')
                join_date_str = join_date.strftime('%m/%d/%y') if join_date else ''

                row_data = [
                    result.get('first_name', ''),
                    result.get('last_name', ''),
                    result.get('member_number', ''),
                    join_date_str,
                    result.get('transaction_count', 0),
                    missing_str,
                    duplicate_str,
                    notes
                ]

                for col_idx, value in enumerate(row_data, start=1):
                    cell = audit_sheet.cell(row=excel_row, column=col_idx, value=value)
                    cell.fill = self.HIGHLIGHT_FILL

                flagged_count += 1
                excel_row += 1

            excel_row += 1  # Blank row after section

        # Section 2: BILLING PROBLEM MEMBERS
        if bp_members:
            excel_row = self._write_section_header(
                audit_sheet, excel_row,
                f"BILLING PROBLEM MEMBERS ({len(bp_members)} members)",
                col_count
            )

            for result in bp_members:
                missing_months = result.get('missing_months', [])
                missing_str = ', '.join(missing_months) if missing_months else ''

                duplicate_months = result.get('duplicate_months', [])
                duplicate_str = ', '.join(duplicate_months) if duplicate_months else ''

                flags = result.get('flags', [])
                notes = ' | '.join([str(flag) for flag in flags]) if flags else ''

                join_date = result.get('join_date')
                join_date_str = join_date.strftime('%m/%d/%y') if join_date else ''

                row_data = [
                    result.get('first_name', ''),
                    result.get('last_name', ''),
                    result.get('member_number', ''),
                    join_date_str,
                    result.get('transaction_count', 0),
                    missing_str,
                    duplicate_str,
                    notes
                ]

                for col_idx, value in enumerate(row_data, start=1):
                    cell = audit_sheet.cell(row=excel_row, column=col_idx, value=value)
                    cell.fill = self.BP_FILL

                bp_count += 1
                excel_row += 1

            excel_row += 1

        # Section 3: VALID MEMBERS
        if valid_members:
            excel_row = self._write_section_header(
                audit_sheet, excel_row,
                f"VALID MEMBERS ({len(valid_members)} members)",
                col_count
            )

            for result in valid_members:
                join_date = result.get('join_date')
                join_date_str = join_date.strftime('%m/%d/%y') if join_date else ''

                row_data = [
                    result.get('first_name', ''),
                    result.get('last_name', ''),
                    result.get('member_number', ''),
                    join_date_str,
                    result.get('transaction_count', 0),
                    '',  # No missing months
                    '',  # No duplicate months
                    ''   # No notes
                ]

                for col_idx, value in enumerate(row_data, start=1):
                    cell = audit_sheet.cell(row=excel_row, column=col_idx, value=value)

                excel_row += 1

        # Auto-adjust column widths
        self._auto_adjust_column_widths(audit_sheet)

        # Add summary sheet
        self._add_mtm_summary_sheet(
            wb, mtm_results, flagged_count, bp_count, len(valid_members)
        )

        # Add transactions detail sheet
        self._add_transactions_sheet(wb, header_row, mtm_results, column_indices, bp_config)

        # Save workbook
        output_path = self.output_folder / output_filename
        wb.save(output_path)

        return str(output_path)

    def _add_mtm_summary_sheet(
        self,
        workbook,
        mtm_results: Dict[str, Any],
        flagged_count: int,
        bp_count: int,
        valid_count: int
    ):
        """Add summary sheet for MTM audit"""
        summary_sheet = workbook.create_sheet("Summary", 0)

        # Title
        summary_sheet['A1'] = "MTM AUDIT SUMMARY"
        summary_sheet['A1'].font = Font(bold=True, size=16)

        # Overall stats
        row = 3
        total_members = mtm_results.get('total_members', 0)
        total_transactions = mtm_results.get('total_transactions', 0)

        summary_sheet[f'A{row}'] = "Total Members:"
        summary_sheet[f'B{row}'] = total_members
        summary_sheet[f'A{row}'].font = self.BOLD_FONT

        row += 1
        summary_sheet[f'A{row}'] = "Total Transactions:"
        summary_sheet[f'B{row}'] = total_transactions
        summary_sheet[f'A{row}'].font = self.BOLD_FONT

        row += 2
        summary_sheet[f'A{row}'] = "Flagged Members:"
        summary_sheet[f'B{row}'] = flagged_count
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'].fill = self.HIGHLIGHT_FILL

        row += 1
        summary_sheet[f'A{row}'] = "Billing Problem Members:"
        summary_sheet[f'B{row}'] = bp_count
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'].fill = self.BP_FILL

        row += 1
        summary_sheet[f'A{row}'] = "Valid Members:"
        summary_sheet[f'B{row}'] = valid_count
        summary_sheet[f'A{row}'].font = self.BOLD_FONT

        row += 1
        flagged_percentage = (flagged_count / total_members * 100) if total_members > 0 else 0
        summary_sheet[f'A{row}'] = "Flagged Percentage:"
        summary_sheet[f'B{row}'] = f"{flagged_percentage:.1f}%"
        summary_sheet[f'A{row}'].font = self.BOLD_FONT

        # Red flag breakdown
        row += 3
        summary_sheet[f'A{row}'] = "RED FLAG BREAKDOWN"
        summary_sheet[f'A{row}'].font = Font(bold=True, size=14)

        # Count red flags by type
        flag_counts = {}
        member_results = mtm_results.get('member_results', {})
        for result in member_results.values():
            for flag in result.get('flags', []):
                flag_type = flag.flag_type
                flag_counts[flag_type] = flag_counts.get(flag_type, 0) + 1

        sorted_flags = sorted(flag_counts.items(), key=lambda x: x[1], reverse=True)

        row += 2
        summary_sheet[f'A{row}'] = "Red Flag Type"
        summary_sheet[f'B{row}'] = "Count"
        summary_sheet[f'A{row}'].font = self.BOLD_FONT
        summary_sheet[f'B{row}'].font = self.BOLD_FONT

        for flag_type, count in sorted_flags:
            row += 1
            summary_sheet[f'A{row}'] = self._format_flag_type(flag_type)
            summary_sheet[f'B{row}'] = count

        # Auto-adjust widths
        summary_sheet.column_dimensions['A'].width = 35
        summary_sheet.column_dimensions['B'].width = 20

    def _add_transactions_sheet(
        self,
        workbook,
        header_row: List[str],
        mtm_results: Dict[str, Any],
        column_indices: List[int],
        bp_config: Dict[str, Any]
    ):
        """Add detailed transactions sheet"""
        txn_sheet = workbook.create_sheet("All Transactions")

        # Write header
        enhanced_header = header_row + ["Member Status", "Flags"]
        for col_idx, header_text in enumerate(enhanced_header, start=1):
            cell = txn_sheet.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.BOLD_FONT
            cell.fill = self.HEADER_FILL

        excel_row = 2
        member_results = mtm_results.get('member_results', {})

        for member_key, result in member_results.items():
            # Determine member status
            is_bp = False
            for txn in result.get('transactions', []):
                if self._is_bp_member(txn, column_indices, bp_config):
                    is_bp = True
                    break

            has_flags = result.get('has_flags', False)
            if is_bp:
                status = "BP"
                fill = self.BP_FILL
            elif has_flags:
                status = "Flagged"
                fill = self.HIGHLIGHT_FILL
            else:
                status = "Valid"
                fill = None

            # Format flags
            flags = result.get('flags', [])
            flags_str = ' | '.join([str(flag) for flag in flags]) if flags else ''

            # Write each transaction
            for txn in result.get('transactions', []):
                row_data = txn + [status, flags_str]
                for col_idx, value in enumerate(row_data, start=1):
                    cell = txn_sheet.cell(row=excel_row, column=col_idx, value=value)
                    if fill:
                        cell.fill = fill

                excel_row += 1

        # Auto-adjust widths
        self._auto_adjust_column_widths(txn_sheet)

    def create_all_types_report(
        self,
        header_row: List[str],
        type_results: Dict[str, Any],
        output_filename: str,
        column_mapping: Dict[str, int] = None,
        bp_config: Dict[str, Any] = None,
        member_type_mapping: Dict[str, str] = None
    ) -> str:
        """
        Create Excel report with tabs for each member_type plus a summary tab.

        Args:
            header_row: Column headers from original file
            type_results: Dictionary with results per member_type
            output_filename: Name for output file
            column_mapping: Dictionary with column name to index mappings
            bp_config: BP detection configuration
            member_type_mapping: Mapping of member_type codes to config keys

        Returns:
            Full path to generated report file
        """
        wb = Workbook()

        # Remove default sheet, we'll add our own
        wb.remove(wb.active)

        # Create Summary sheet first
        summary_sheet = wb.create_sheet("Summary", 0)
        self._add_all_types_summary_sheet(summary_sheet, type_results, member_type_mapping)

        # Get column indices for BP detection
        column_indices = []
        if column_mapping:
            bp_columns = bp_config.get('columns', ['code', 'member_type', 'member_group']) if bp_config else ['code', 'member_type', 'member_group']
            for col_name in bp_columns:
                idx = column_mapping.get(col_name)
                if idx is not None and idx >= 0:
                    column_indices.append(idx)
        if not column_indices:
            column_indices = [11, 9, 10]  # code, member_type, member_group in new format

        # Create a tab for each member_type
        for member_type, results in type_results.items():
            # Sanitize sheet name (max 31 chars, no special chars)
            sheet_name = member_type[:31].replace('/', '-').replace('\\', '-').replace('*', '').replace('?', '').replace('[', '').replace(']', '')
            if not sheet_name:
                sheet_name = "Unknown"

            type_sheet = wb.create_sheet(sheet_name)

            if results.get('is_mtm'):
                # MTM type - use member-based format
                self._write_mtm_type_sheet(type_sheet, header_row, results, column_indices, bp_config)
            else:
                # Standard type or unknown - use row-based format
                self._write_standard_type_sheet(type_sheet, header_row, results, column_indices, bp_config)

        # Save workbook
        output_path = self.output_folder / output_filename
        wb.save(output_path)

        return str(output_path)

    def _add_all_types_summary_sheet(
        self,
        sheet,
        type_results: Dict[str, Any],
        member_type_mapping: Dict[str, str] = None
    ):
        """Add summary sheet for all types report"""

        # Title
        sheet['A1'] = "ALL MEMBERSHIP TYPES AUDIT SUMMARY"
        sheet['A1'].font = Font(bold=True, size=16)

        # Overall totals
        row = 3
        total_records = sum(r.get('total_records', 0) for r in type_results.values())
        total_flagged = 0
        total_financial_impact = 0

        for results in type_results.values():
            if results.get('is_mtm'):
                total_flagged += results.get('flagged_members', 0)
            else:
                total_flagged += results.get('flagged_count', 0)
            total_financial_impact += results.get('financial_impact', 0)

        sheet[f'A{row}'] = "Total Records:"
        sheet[f'B{row}'] = total_records
        sheet[f'A{row}'].font = self.BOLD_FONT

        row += 1
        sheet[f'A{row}'] = "Total Flagged:"
        sheet[f'B{row}'] = total_flagged
        sheet[f'A{row}'].font = self.BOLD_FONT
        sheet[f'B{row}'].fill = self.HIGHLIGHT_FILL

        row += 1
        flagged_pct = (total_flagged / total_records * 100) if total_records > 0 else 0
        sheet[f'A{row}'] = "Flagged Percentage:"
        sheet[f'B{row}'] = f"{flagged_pct:.1f}%"
        sheet[f'A{row}'].font = self.BOLD_FONT

        row += 1
        sheet[f'A{row}'] = "Total Financial Impact:"
        sheet[f'B{row}'] = f"${total_financial_impact:,.2f}"
        sheet[f'A{row}'].font = self.BOLD_FONT
        sheet[f'B{row}'].font = Font(bold=True, color="FF0000")

        # Breakdown by member type
        row += 3
        sheet[f'A{row}'] = "BREAKDOWN BY MEMBER TYPE"
        sheet[f'A{row}'].font = Font(bold=True, size=14)

        row += 2
        # Header row for breakdown table
        headers = ["Member Type", "Config", "Records", "Flagged", "Flag %", "Financial Impact", "Rules Applied"]
        for col_idx, header_text in enumerate(headers, start=1):
            cell = sheet.cell(row=row, column=col_idx, value=header_text)
            cell.font = self.BOLD_FONT
            cell.fill = self.HEADER_FILL

        row += 1

        # Sort by record count descending
        sorted_types = sorted(type_results.items(), key=lambda x: x[1].get('total_records', 0), reverse=True)

        for member_type, results in sorted_types:
            config_key = results.get('config_key', 'N/A')
            config_name = config_key if config_key else 'Unknown'

            records = results.get('total_records', 0)

            if results.get('is_mtm'):
                flagged = results.get('flagged_members', 0)
            else:
                flagged = results.get('flagged_count', 0)

            flag_pct = (flagged / records * 100) if records > 0 else 0
            financial_impact = results.get('financial_impact', 0)
            has_rules = "Yes" if results.get('has_rules') else "No (raw data)"

            sheet.cell(row=row, column=1, value=member_type)
            sheet.cell(row=row, column=2, value=config_name)
            sheet.cell(row=row, column=3, value=records)

            flagged_cell = sheet.cell(row=row, column=4, value=flagged)
            if flagged > 0:
                flagged_cell.fill = self.HIGHLIGHT_FILL

            sheet.cell(row=row, column=5, value=f"{flag_pct:.1f}%")
            sheet.cell(row=row, column=6, value=f"${financial_impact:,.2f}")
            sheet.cell(row=row, column=7, value=has_rules)

            row += 1

        # Legend
        row += 2
        sheet[f'A{row}'] = "LEGEND"
        sheet[f'A{row}'].font = Font(bold=True, size=12)

        row += 1
        sheet[f'A{row}'] = "Known Types:"
        sheet[f'B{row}'] = "1MCORE (1 Month PIF), 1YRCORE (1 Year PIF), 3MCORE (3 Month PIF), MTMCORE (Month to Month)"

        row += 1
        sheet[f'A{row}'] = "Unknown Types:"
        sheet[f'B{row}'] = "Raw data only - no red flag rules applied. Review data to define rules."

        row += 2
        yellow_cell = sheet.cell(row=row, column=1, value="Yellow")
        yellow_cell.fill = self.HIGHLIGHT_FILL
        sheet[f'B{row}'] = "= Flagged accounts (red flags detected)"

        row += 1
        orange_cell = sheet.cell(row=row, column=1, value="Orange")
        orange_cell.fill = self.BP_FILL
        sheet[f'B{row}'] = "= Billing Problem accounts (BP indicator in code/member_type)"

        # Auto-adjust widths
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 25
        sheet.column_dimensions['C'].width = 12
        sheet.column_dimensions['D'].width = 12
        sheet.column_dimensions['E'].width = 12
        sheet.column_dimensions['F'].width = 18
        sheet.column_dimensions['G'].width = 18

    def _write_standard_type_sheet(
        self,
        sheet,
        header_row: List[str],
        results: Dict[str, Any],
        column_indices: List[int],
        bp_config: Dict[str, Any]
    ):
        """Write a sheet for a standard (non-MTM) member type"""

        # Add Notes column to header
        enhanced_header = header_row + ["Notes"]
        col_count = len(enhanced_header)

        # Write header row
        for col_idx, header_text in enumerate(enhanced_header, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.BOLD_FONT
            cell.fill = self.HEADER_FILL
            cell.alignment = self.CENTER_ALIGN

        audit_results = results.get('audit_results', [])
        rows = results.get('rows', [])

        # If we have audit_results with row_data, use those
        if audit_results and 'row_data' in audit_results[0]:
            data_pairs = [(r.get('row_data', []), r) for r in audit_results]
        else:
            # Pair rows with empty results
            data_pairs = [(row, {'red_flags': [], 'has_flags': False}) for row in rows]

        # Categorize rows
        flagged_rows, bp_rows, valid_rows = self._categorize_rows(
            [pair[0] for pair in data_pairs],
            [pair[1] for pair in data_pairs],
            column_indices,
            bp_config
        )

        excel_row = 2

        # Section 1: FLAGGED ACCOUNTS
        if flagged_rows:
            excel_row = self._write_section_header(
                sheet, excel_row,
                f"FLAGGED ACCOUNTS ({len(flagged_rows)} records)",
                col_count
            )

            for data_row, audit_result in flagged_rows:
                red_flags = audit_result.get('red_flags', [])
                notes = " | ".join([str(flag) for flag in red_flags]) if red_flags else ""
                enhanced_row = list(data_row) + [notes]

                for col_idx, value in enumerate(enhanced_row, start=1):
                    cell = sheet.cell(row=excel_row, column=col_idx, value=value)
                    cell.fill = self.HIGHLIGHT_FILL

                excel_row += 1

            excel_row += 1

        # Section 2: BILLING PROBLEM ACCOUNTS
        if bp_rows:
            excel_row = self._write_section_header(
                sheet, excel_row,
                f"BILLING PROBLEM ACCOUNTS ({len(bp_rows)} records)",
                col_count
            )

            for data_row, audit_result in bp_rows:
                red_flags = audit_result.get('red_flags', [])
                notes = " | ".join([str(flag) for flag in red_flags]) if red_flags else ""
                enhanced_row = list(data_row) + [notes]

                for col_idx, value in enumerate(enhanced_row, start=1):
                    cell = sheet.cell(row=excel_row, column=col_idx, value=value)
                    cell.fill = self.BP_FILL

                excel_row += 1

            excel_row += 1

        # Section 3: VALID ACCOUNTS
        if valid_rows:
            excel_row = self._write_section_header(
                sheet, excel_row,
                f"VALID ACCOUNTS ({len(valid_rows)} records)",
                col_count
            )

            for data_row, audit_result in valid_rows:
                enhanced_row = list(data_row) + [""]

                for col_idx, value in enumerate(enhanced_row, start=1):
                    sheet.cell(row=excel_row, column=col_idx, value=value)

                excel_row += 1

        # Auto-adjust column widths
        self._auto_adjust_column_widths(sheet)

    def _write_mtm_type_sheet(
        self,
        sheet,
        header_row: List[str],
        results: Dict[str, Any],
        column_indices: List[int],
        bp_config: Dict[str, Any]
    ):
        """Write a sheet for MTM member type (member-based grouping)"""

        # Create member-level header
        member_header = [
            "First Name", "Last Name", "Member #", "Join Date",
            "Transaction Count", "Missing Months", "Duplicate Months", "Notes"
        ]
        col_count = len(member_header)

        # Write header row
        for col_idx, header_text in enumerate(member_header, start=1):
            cell = sheet.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.BOLD_FONT
            cell.fill = self.HEADER_FILL
            cell.alignment = self.CENTER_ALIGN

        member_results = results.get('member_results', {})

        # Categorize members
        flagged_members = []
        bp_members = []
        valid_members = []

        for member_key, member_result in member_results.items():
            # Check if any transaction has BP indicator
            is_bp = False
            for txn in member_result.get('transactions', []):
                if self._is_bp_member(txn, column_indices, bp_config):
                    is_bp = True
                    break

            if is_bp:
                bp_members.append(member_result)
            elif member_result.get('has_flags', False):
                flagged_members.append(member_result)
            else:
                valid_members.append(member_result)

        excel_row = 2

        # Section 1: FLAGGED MEMBERS
        if flagged_members:
            excel_row = self._write_section_header(
                sheet, excel_row,
                f"FLAGGED MEMBERS ({len(flagged_members)} members)",
                col_count
            )

            for member_result in flagged_members:
                excel_row = self._write_mtm_member_row(sheet, excel_row, member_result, self.HIGHLIGHT_FILL)

            excel_row += 1

        # Section 2: BILLING PROBLEM MEMBERS
        if bp_members:
            excel_row = self._write_section_header(
                sheet, excel_row,
                f"BILLING PROBLEM MEMBERS ({len(bp_members)} members)",
                col_count
            )

            for member_result in bp_members:
                excel_row = self._write_mtm_member_row(sheet, excel_row, member_result, self.BP_FILL)

            excel_row += 1

        # Section 3: VALID MEMBERS
        if valid_members:
            excel_row = self._write_section_header(
                sheet, excel_row,
                f"VALID MEMBERS ({len(valid_members)} members)",
                col_count
            )

            for member_result in valid_members:
                excel_row = self._write_mtm_member_row(sheet, excel_row, member_result, None)

        # Auto-adjust column widths
        self._auto_adjust_column_widths(sheet)

    def _write_mtm_member_row(self, sheet, row_num: int, member_result: Dict[str, Any], fill) -> int:
        """Write a single MTM member row and return next row number"""

        missing_months = member_result.get('missing_months', [])
        missing_str = ', '.join(missing_months) if missing_months else ''

        duplicate_months = member_result.get('duplicate_months', [])
        duplicate_str = ', '.join(duplicate_months) if duplicate_months else ''

        flags = member_result.get('flags', [])
        notes = ' | '.join([str(flag) for flag in flags]) if flags else ''

        join_date = member_result.get('join_date')
        join_date_str = join_date.strftime('%m/%d/%y') if join_date else ''

        row_data = [
            member_result.get('first_name', ''),
            member_result.get('last_name', ''),
            member_result.get('member_number', ''),
            join_date_str,
            member_result.get('transaction_count', 0),
            missing_str,
            duplicate_str,
            notes
        ]

        for col_idx, value in enumerate(row_data, start=1):
            cell = sheet.cell(row=row_num, column=col_idx, value=value)
            if fill:
                cell.fill = fill

        return row_num + 1

    def create_individual_type_files(
        self,
        header_row: List[str],
        type_results: Dict[str, Any],
        base_filename: str,
        member_type_mapping: Dict[str, str],
        column_mapping: Dict[str, int] = None,
        bp_config: Dict[str, Any] = None
    ) -> Dict[str, Dict[str, str]]:
        """
        Create separate Excel files for each member_type.

        For known types (1MCORE, 1YRCORE, 3MCORE, MTMCORE):
            - Audited file with Notes column, organized as:
              Flagged (yellow) → Valid → BP (orange at bottom)
            - Raw file without Notes (for re-verification)
        For unknown types: Raw data only (original columns, no Notes)

        Args:
            header_row: Column headers from original file
            type_results: Dictionary with results per member_type
            base_filename: Base name for output files (without extension)
            member_type_mapping: Mapping of member_type codes to config keys
            column_mapping: Dictionary with column name to index mappings
            bp_config: BP detection configuration

        Returns:
            Dict mapping member_type to dict with 'audited' and/or 'raw' file paths
        """
        individual_files = {}

        # Get column indices for BP detection
        column_indices = []
        if column_mapping:
            bp_columns = bp_config.get('columns', ['code', 'member_type', 'member_group']) if bp_config else ['code', 'member_type', 'member_group']
            for col_name in bp_columns:
                idx = column_mapping.get(col_name)
                if idx is not None and idx >= 0:
                    column_indices.append(idx)
        if not column_indices:
            column_indices = [11, 9, 10]  # code, member_type, member_group in new format

        for member_type, results in type_results.items():
            is_known_type = results.get('has_rules', False)
            rows = results.get('rows', [])

            if not rows:
                continue

            # Sanitize member_type for filename
            safe_type = member_type.replace('/', '-').replace('\\', '-').replace('*', '').replace('?', '').replace('[', '').replace(']', '')

            individual_files[member_type] = {}

            if is_known_type:
                # Known type: generate BOTH audited file (with Notes) AND raw file (for re-upload)

                # 1. Create audited file with Notes column and BP separation
                wb_audited = Workbook()
                sheet_audited = wb_audited.active
                sheet_audited.title = "Data"

                enhanced_header = header_row + ["Notes"]
                col_count = len(enhanced_header)

                # Write header row
                for col_idx, header_text in enumerate(enhanced_header, start=1):
                    cell = sheet_audited.cell(row=1, column=col_idx, value=header_text)
                    cell.font = self.BOLD_FONT
                    cell.fill = self.HEADER_FILL
                    cell.alignment = self.CENTER_ALIGN

                # Get audit results for notes
                audit_results = results.get('audit_results', [])

                # Pair rows with audit results
                if audit_results and len(audit_results) > 0:
                    data_pairs = list(zip(rows, audit_results))
                else:
                    # No audit results - create empty results
                    data_pairs = [(row, {'red_flags': [], 'has_flags': False}) for row in rows]

                # Categorize rows: flagged, BP, valid
                flagged_rows, bp_rows, valid_rows = self._categorize_rows(
                    [pair[0] for pair in data_pairs],
                    [pair[1] for pair in data_pairs],
                    column_indices,
                    bp_config
                )

                excel_row = 2

                # Section 1: FLAGGED ACCOUNTS (yellow)
                if flagged_rows:
                    excel_row = self._write_section_header(
                        sheet_audited, excel_row,
                        f"FLAGGED ACCOUNTS ({len(flagged_rows)} records)",
                        col_count
                    )

                    for data_row, audit_result in flagged_rows:
                        red_flags = audit_result.get('red_flags', [])
                        notes = " | ".join([str(flag) for flag in red_flags]) if red_flags else ""
                        enhanced_row = list(data_row) + [notes]

                        for col_idx, value in enumerate(enhanced_row, start=1):
                            cell = sheet_audited.cell(row=excel_row, column=col_idx, value=value)
                            cell.fill = self.HIGHLIGHT_FILL  # Yellow

                        excel_row += 1

                    excel_row += 1  # Blank row after section

                # Section 2: VALID ACCOUNTS (no highlight)
                if valid_rows:
                    excel_row = self._write_section_header(
                        sheet_audited, excel_row,
                        f"VALID ACCOUNTS ({len(valid_rows)} records)",
                        col_count
                    )

                    for data_row, audit_result in valid_rows:
                        enhanced_row = list(data_row) + [""]

                        for col_idx, value in enumerate(enhanced_row, start=1):
                            sheet_audited.cell(row=excel_row, column=col_idx, value=value)

                        excel_row += 1

                    excel_row += 1  # Blank row after section

                # Section 3: BILLING PROBLEM ACCOUNTS (orange at bottom)
                if bp_rows:
                    excel_row = self._write_section_header(
                        sheet_audited, excel_row,
                        f"BILLING PROBLEM ACCOUNTS ({len(bp_rows)} records)",
                        col_count
                    )

                    for data_row, audit_result in bp_rows:
                        red_flags = audit_result.get('red_flags', [])
                        notes = " | ".join([str(flag) for flag in red_flags]) if red_flags else ""
                        enhanced_row = list(data_row) + [notes]

                        for col_idx, value in enumerate(enhanced_row, start=1):
                            cell = sheet_audited.cell(row=excel_row, column=col_idx, value=value)
                            cell.fill = self.BP_FILL  # Orange

                        excel_row += 1

                self._auto_adjust_column_widths(sheet_audited)

                audited_filename = f"{base_filename}_{safe_type}.xlsx"
                audited_path = self.output_folder / audited_filename
                wb_audited.save(audited_path)
                individual_files[member_type]['audited'] = str(audited_path)

                # 2. Create raw file without Notes (for re-verification)
                wb_raw = Workbook()
                sheet_raw = wb_raw.active
                sheet_raw.title = "Data"

                # Write header row (no Notes)
                for col_idx, header_text in enumerate(header_row, start=1):
                    cell = sheet_raw.cell(row=1, column=col_idx, value=header_text)
                    cell.font = self.BOLD_FONT
                    cell.fill = self.HEADER_FILL
                    cell.alignment = self.CENTER_ALIGN

                # Write data rows
                excel_row = 2
                for row in rows:
                    for col_idx, value in enumerate(row, start=1):
                        sheet_raw.cell(row=excel_row, column=col_idx, value=value)
                    excel_row += 1

                self._auto_adjust_column_widths(sheet_raw)

                raw_filename = f"{base_filename}_{safe_type}_raw.xlsx"
                raw_path = self.output_folder / raw_filename
                wb_raw.save(raw_path)
                individual_files[member_type]['raw'] = str(raw_path)

            else:
                # Unknown type: raw data only, no Notes column
                wb = Workbook()
                sheet = wb.active
                sheet.title = "Data"

                # Write header row
                for col_idx, header_text in enumerate(header_row, start=1):
                    cell = sheet.cell(row=1, column=col_idx, value=header_text)
                    cell.font = self.BOLD_FONT
                    cell.fill = self.HEADER_FILL
                    cell.alignment = self.CENTER_ALIGN

                # Write data rows
                excel_row = 2
                for row in rows:
                    for col_idx, value in enumerate(row, start=1):
                        sheet.cell(row=excel_row, column=col_idx, value=value)
                    excel_row += 1

                self._auto_adjust_column_widths(sheet)

                output_filename = f"{base_filename}_{safe_type}.xlsx"
                output_path = self.output_folder / output_filename
                wb.save(output_path)
                individual_files[member_type]['raw'] = str(output_path)

        return individual_files

    def create_split_type_files(
        self,
        header_row: List[str],
        rows_by_type: Dict[str, List[List[str]]],
        base_filename: str
    ) -> Dict[str, Dict[str, Any]]:
        """
        Create separate raw Excel files for each member_type.
        NO rules applied, NO highlighting, NO Notes column - pure raw data output.

        Args:
            header_row: Column headers from original file
            rows_by_type: Dictionary mapping member_type to list of rows
            base_filename: Base name for output files (without extension)

        Returns:
            Dict mapping member_type to dict with 'file_path', 'row_count', 'file_size'
        """
        split_files = {}

        for member_type, rows in rows_by_type.items():
            if not rows:
                continue

            # Sanitize member_type for filename
            safe_type = member_type.replace('/', '-').replace('\\', '-').replace('*', '').replace('?', '').replace('[', '').replace(']', '').replace(':', '-')
            if not safe_type:
                safe_type = 'UNKNOWN'

            # Create workbook
            wb = Workbook()
            sheet = wb.active
            sheet.title = "Data"

            # Write header row with basic formatting
            for col_idx, header_text in enumerate(header_row, start=1):
                cell = sheet.cell(row=1, column=col_idx, value=header_text)
                cell.font = self.BOLD_FONT
                cell.fill = self.HEADER_FILL

            # Write data rows - no highlighting, no modifications
            excel_row = 2
            for row in rows:
                for col_idx, value in enumerate(row, start=1):
                    sheet.cell(row=excel_row, column=col_idx, value=value)
                excel_row += 1

            # Auto-adjust column widths
            self._auto_adjust_column_widths(sheet)

            # Generate output filename
            output_filename = f"{base_filename}_{safe_type}.xlsx"
            output_path = self.output_folder / output_filename
            wb.save(output_path)

            # Get file size
            file_size = output_path.stat().st_size

            split_files[member_type] = {
                'file_path': str(output_path),
                'row_count': len(rows),
                'file_size': file_size,
                'filename': output_filename
            }

        return split_files
