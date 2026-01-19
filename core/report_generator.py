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
    HIGHLIGHT_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # Yellow
    HEADER_FILL = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')  # Gray
    BOLD_FONT = Font(bold=True)
    CENTER_ALIGN = Alignment(horizontal='center', vertical='center')

    def __init__(self, output_folder: str = 'outputs'):
        """
        Initialize report generator

        Args:
            output_folder: Directory to save reports
        """
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(exist_ok=True)

    def create_audit_report(
        self,
        header_row: List[str],
        data_rows: List[List[str]],
        audit_results: List[Dict[str, Any]],
        output_filename: str,
        include_summary_sheet: bool = True
    ) -> str:
        """
        Create Excel audit report with highlighted red flags

        Args:
            header_row: Column headers
            data_rows: Original data rows
            audit_results: List of audit results for each row (from audit_engine)
            output_filename: Name for output file
            include_summary_sheet: Whether to add a summary sheet

        Returns:
            Full path to generated report file
        """
        wb = Workbook()

        # Create main audit sheet
        audit_sheet = wb.active
        audit_sheet.title = "Audit Report"

        # Add only "Notes" column to header (keep original columns intact)
        enhanced_header = header_row + ["Notes"]

        # Write header row with formatting
        for col_idx, header_text in enumerate(enhanced_header, start=1):
            cell = audit_sheet.cell(row=1, column=col_idx, value=header_text)
            cell.font = self.BOLD_FONT
            cell.fill = self.HEADER_FILL
            cell.alignment = self.CENTER_ALIGN

        # Write data rows with highlighting
        excel_row = 2
        flagged_count = 0

        for idx, (data_row, audit_result) in enumerate(zip(data_rows, audit_results)):
            red_flags = audit_result.get('red_flags', [])
            has_flags = len(red_flags) > 0

            # Build notes column - only populated if there are red flags
            notes = " | ".join([str(flag) for flag in red_flags]) if red_flags else ""

            # Combine original row data with just the Notes column
            enhanced_row = data_row + [notes]

            # Write row
            for col_idx, value in enumerate(enhanced_row, start=1):
                cell = audit_sheet.cell(row=excel_row, column=col_idx, value=value)

                # Apply yellow highlighting ONLY if flagged
                if has_flags:
                    cell.fill = self.HIGHLIGHT_FILL

            if has_flags:
                flagged_count += 1

            excel_row += 1

        # Auto-adjust column widths
        self._auto_adjust_column_widths(audit_sheet)

        # Add summary sheet if requested
        if include_summary_sheet:
            self._add_summary_sheet(wb, audit_results, flagged_count, len(data_rows))

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

    def _add_summary_sheet(self, workbook, audit_results: List[Dict[str, Any]], flagged_count: int, total_count: int):
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
        summary_sheet[f'A{row}'] = "Clean Records:"
        summary_sheet[f'B{row}'] = total_count - flagged_count
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
