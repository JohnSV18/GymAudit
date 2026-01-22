"""
File Handler Module
Handles reading CSV and Excel files for membership data
"""

import csv
import pandas as pd
from pathlib import Path
from typing import List, Tuple, Dict, Any
from io import BytesIO


class FileReadError(Exception):
    """Custom exception for file reading errors"""
    pass


class MembershipFileReader:
    """Reads membership data from CSV and Excel files"""

    # Old format (17 columns)
    REQUIRED_COLUMNS = [
        'Last Name',
        'First Name',
        'Member #',
        'Join Date',
        'Exp Date',
        'Type',
        'Pay Type',
        'Dues Amt',
        'Cycle',
        'Balance',
        'End Draft'
    ]

    # New format (20 columns) - required columns for validation
    REQUIRED_COLUMNS_NEW = [
        'last_name',
        'first_name',
        'member_number',
        'join_date',
        'expiration_date',
        'member_type',
        'member_group',
        'code',
        'payment_method',
        'dues_amount',
        'balance',
        'start_draft',
        'end_draft'
    ]

    # Unique indicators for new format detection
    NEW_FORMAT_INDICATORS = ['transaction_date', 'receipt', 'site_number', 'postedby']

    def __init__(self):
        self.supported_extensions = ['.csv', '.xlsx', '.xls']

    def detect_format(self, header_row: List[str]) -> str:
        """
        Detect if this is old or new format based on column names.

        Args:
            header_row: List of column headers

        Returns:
            'new' for 20-column format, 'old' for 17-column format
        """
        header_lower = [col.lower().strip() for col in header_row]

        # New format has these unique columns
        for indicator in self.NEW_FORMAT_INDICATORS:
            if any(indicator in col for col in header_lower):
                return 'new'
        return 'old'

    def is_supported_file(self, filename: str) -> bool:
        """Check if file extension is supported"""
        ext = Path(filename).suffix.lower()
        return ext in self.supported_extensions

    def read_csv_file(self, file_path: str) -> Tuple[List[List[str]], str]:
        """
        Read CSV file and return rows

        Args:
            file_path: Path to CSV file

        Returns:
            Tuple of (rows, filename)

        Raises:
            FileReadError: If file cannot be read
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)

            filename = Path(file_path).name
            return rows, filename

        except Exception as e:
            raise FileReadError(f"Error reading CSV file: {str(e)}")

    def read_excel_file(self, file_path: str) -> Tuple[List[List[str]], str]:
        """
        Read Excel file and return rows

        Args:
            file_path: Path to Excel file

        Returns:
            Tuple of (rows, filename)

        Raises:
            FileReadError: If file cannot be read
        """
        try:
            df = pd.read_excel(file_path)

            # Convert DataFrame to list of lists (similar to CSV format)
            rows = []

            # Add header row
            headers = df.columns.tolist()
            rows.append(headers)

            # Add data rows, converting to strings
            for _, row in df.iterrows():
                row_data = [str(val) if pd.notna(val) else '' for val in row]
                rows.append(row_data)

            filename = Path(file_path).name
            return rows, filename

        except Exception as e:
            raise FileReadError(f"Error reading Excel file: {str(e)}")

    def read_file_from_upload(self, uploaded_file) -> Tuple[List[List[str]], str]:
        """
        Read file from Streamlit uploaded file object

        Args:
            uploaded_file: Streamlit UploadedFile object

        Returns:
            Tuple of (rows, filename)

        Raises:
            FileReadError: If file cannot be read
        """
        filename = uploaded_file.name
        ext = Path(filename).suffix.lower()

        try:
            if ext == '.csv':
                # Read CSV from BytesIO
                content = uploaded_file.read().decode('utf-8')
                reader = csv.reader(content.splitlines())
                rows = list(reader)

            elif ext in ['.xlsx', '.xls']:
                # Read Excel from BytesIO
                df = pd.read_excel(BytesIO(uploaded_file.read()))

                # Convert to list of lists
                rows = []
                headers = df.columns.tolist()
                rows.append(headers)

                for _, row in df.iterrows():
                    row_data = [str(val) if pd.notna(val) else '' for val in row]
                    rows.append(row_data)

            else:
                raise FileReadError(f"Unsupported file type: {ext}")

            return rows, filename

        except Exception as e:
            raise FileReadError(f"Error reading uploaded file '{filename}': {str(e)}")

    def read_file(self, file_path: str) -> Tuple[List[List[str]], str]:
        """
        Read file (auto-detect CSV or Excel)

        Args:
            file_path: Path to file

        Returns:
            Tuple of (rows, filename)

        Raises:
            FileReadError: If file cannot be read or is unsupported
        """
        ext = Path(file_path).suffix.lower()

        if ext == '.csv':
            return self.read_csv_file(file_path)
        elif ext in ['.xlsx', '.xls']:
            return self.read_excel_file(file_path)
        else:
            raise FileReadError(f"Unsupported file type: {ext}. Supported: {', '.join(self.supported_extensions)}")

    def validate_structure(self, rows: List[List[str]]) -> Tuple[bool, str, int, str]:
        """
        Validate that file has expected structure

        Args:
            rows: List of rows from file

        Returns:
            Tuple of (is_valid, error_message, header_row_index, format_type)
        """
        if not rows or len(rows) < 2:
            return False, "File is empty or has insufficient rows", -1, 'unknown'

        # Check if first row might be a title row (like "Table 1")
        # If so, consider second row as header
        first_row = rows[0]
        second_row = rows[1] if len(rows) > 1 else None

        header_row_index = 0

        # Heuristic: if first row has only 1-2 columns with data, it might be a title
        if len([cell for cell in first_row if cell.strip()]) <= 2 and second_row:
            header_row_index = 1

        header_row = rows[header_row_index]

        # Detect format type
        format_type = self.detect_format(header_row)

        # Check for required columns based on format (case-insensitive, partial match)
        header_lower = [col.lower().strip() for col in header_row]

        if format_type == 'new':
            required_columns = self.REQUIRED_COLUMNS_NEW
        else:
            required_columns = self.REQUIRED_COLUMNS

        missing_columns = []
        for required_col in required_columns:
            found = False
            for header_col in header_lower:
                if required_col.lower() in header_col:
                    found = True
                    break
            if not found:
                missing_columns.append(required_col)

        if missing_columns:
            return False, f"Missing required columns: {', '.join(missing_columns)}", header_row_index, format_type

        return True, "", header_row_index, format_type

    def get_data_rows(self, rows: List[List[str]], header_row_index: int) -> List[List[str]]:
        """
        Extract data rows (skip title and header rows)

        Args:
            rows: All rows from file
            header_row_index: Index of header row

        Returns:
            List of data rows only
        """
        # Data starts after header row
        return rows[header_row_index + 1:]

    def get_header_row(self, rows: List[List[str]], header_row_index: int) -> List[str]:
        """Get the header row"""
        return rows[header_row_index]

    def read_and_validate(self, file_path: str) -> Dict[str, Any]:
        """
        Read file and validate structure in one operation

        Args:
            file_path: Path to file

        Returns:
            Dictionary with:
                - rows: All rows
                - filename: Original filename
                - header_row_index: Index of header row
                - header: Header row
                - data_rows: Data rows only
                - is_valid: Whether structure is valid
                - error: Error message if invalid
                - format_type: 'old' or 'new' format

        Raises:
            FileReadError: If file cannot be read
        """
        rows, filename = self.read_file(file_path)
        is_valid, error, header_row_index, format_type = self.validate_structure(rows)

        result = {
            'rows': rows,
            'filename': filename,
            'header_row_index': header_row_index,
            'is_valid': is_valid,
            'error': error,
            'format_type': format_type
        }

        if is_valid:
            result['header'] = self.get_header_row(rows, header_row_index)
            result['data_rows'] = self.get_data_rows(rows, header_row_index)
            result['total_records'] = len(result['data_rows'])

        return result

    def read_and_validate_upload(self, uploaded_file) -> Dict[str, Any]:
        """
        Read uploaded file and validate structure

        Args:
            uploaded_file: Streamlit UploadedFile object

        Returns:
            Dictionary with file data and validation results
        """
        rows, filename = self.read_file_from_upload(uploaded_file)
        is_valid, error, header_row_index, format_type = self.validate_structure(rows)

        result = {
            'rows': rows,
            'filename': filename,
            'header_row_index': header_row_index,
            'is_valid': is_valid,
            'error': error,
            'format_type': format_type
        }

        if is_valid:
            result['header'] = self.get_header_row(rows, header_row_index)
            result['data_rows'] = self.get_data_rows(rows, header_row_index)
            result['total_records'] = len(result['data_rows'])

        return result
