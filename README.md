# Gym Membership Audit Tool

A professional, user-friendly web application for automatically auditing gym membership data and detecting anomalies.

---

## 🚀 First Time? Start Here → [SETUP_GUIDE.md](SETUP_GUIDE.md)

**New to this tool?** Follow the step-by-step installation instructions for Windows and Mac.

---

## 🎯 What This Tool Does

- **Automated Auditing**: Checks membership records against 6 red flag criteria
- **Batch Processing**: Upload and process multiple files at once
- **Visual Reports**: Generates Excel reports with highlighted problem accounts
- **Pattern Detection**: Identifies trends and common issues across your data
- **Financial Analysis**: Calculates revenue impact of data errors
- **User-Friendly**: Simple drag-and-drop interface anyone can use

## ✅ Project Status

**Version 1.0.0 - Production Ready**

All planned features for Phase 1 are complete and tested:
- ✅ Core audit engine
- ✅ Web-based user interface
- ✅ Multiple file support
- ✅ Pattern analysis
- ✅ Financial summaries
- ✅ Highlighted Excel reports
- ✅ Tested with real data (1,174 records)

## 🚀 Quick Start

### 1. Install Dependencies

```bash
pip3 install -r requirements.txt
```

### 2. Run the Application

```bash
python3 audit_app.py
```

Or use Streamlit directly:

```bash
streamlit run audit_app.py
```

### 3. Access in Browser

The tool will automatically open at: `http://localhost:8501`

## 📁 Project Structure

```
gymAudits/
├── audit_app.py                      # Main Streamlit application
├── core/                             # Core audit logic
│   ├── __init__.py
│   ├── audit_engine.py               # Orchestrates audit process
│   ├── file_handler.py               # CSV/Excel file reading
│   ├── red_flags.py                  # Red flag detection logic
│   └── report_generator.py           # Excel report creation
├── utils/                            # Analysis utilities
│   ├── __init__.py
│   └── statistics.py                 # Pattern detection & statistics
├── config/                           # Configuration files
│   ├── red_flag_rules.json           # Red flag criteria (extensible)
│   └── settings.json                 # Application settings
├── outputs/                          # Generated audit reports (auto-created)
├── requirements.txt                  # Python dependencies
├── README.md                         # This file
├── README_STAFF.md                   # Simple guide for non-technical users
├── FUTURE_ENHANCEMENTS.md            # Roadmap for future features
└── HANDOFF_DOCUMENT.md               # Original context from Claude.ai

Legacy files (from original manual audit):
├── fill_names.py                     # Name generation script
├── audit_memberships.py              # Original audit script
├── Sample_Year_Paid_In_Full_Data_With_Names.csv
└── Year_Paid_In_Full_Audit_Report.xlsx
```

## 🔍 Red Flag Criteria (Current)

The tool currently uses "Year Paid in Full" membership rules:

1. **Date Mismatch**: Join and Expiration dates not exactly 365-366 days apart
2. **Low Dues**: Dues amount less than $600
3. **Wrong Pay Type**: Pay Type is not "Annual Bill"
4. **End Draft Error**: End Draft date is not 12/31/99
5. **Cycle Error**: Cycle number is not 1
6. **Balance Issue**: Balance is not exactly $0.00 (credits or debits found)

## 📊 Test Results

Tested with your existing dataset:
- **Total Records**: 1,174
- **Flagged**: 282 (24.0%)
- **Clean**: 892 (76.0%)
- **Financial Impact**: $37,562.27

✅ Matches expected results from manual audit

## 🎨 Features

### Upload & Audit Tab
- Drag-and-drop file upload
- Multiple file processing
- Summary statistics
- Flagged member ID lists
- Download individual audit reports
- Download consolidated report (for multiple files)

### Pattern Analysis Tab
- Red flag distribution charts
- Common flag combinations
- Trends by join date period
- Performance by sales representative

### Financial Summary Tab
- Total financial impact
- Top accounts by impact
- Active vs expired membership comparison
- Average impact per flagged account

## 📖 User Documentation

- **For Staff**: See [README_STAFF.md](README_STAFF.md) for simple, step-by-step instructions
- **For Admins**: See sections below for configuration and deployment

## ⚙️ Configuration

### Red Flag Rules

Edit `config/red_flag_rules.json` to:
- Modify existing rules (change thresholds, expected values)
- Add rules for new membership types (monthly, quarterly, etc.)

### Application Settings

Edit `config/settings.json` to:
- Change output folder location
- Adjust display settings
- Configure server options

## 🌐 Network Deployment

To allow multiple staff members to access the tool:

```bash
streamlit run audit_app.py --server.address=0.0.0.0
```

Staff can then access at: `http://[your-computer-ip]:8501`

## 🔮 Future Enhancements

See [FUTURE_ENHANCEMENTS.md](FUTURE_ENHANCEMENTS.md) for the complete roadmap, including:
- Different rules per membership type
- Continuous monitoring & alerts
- Correction tracking
- Integration with Square POS
- Advanced financial reporting

## 🛠️ Technical Details

### Built With
- **Streamlit**: Web application framework
- **pandas**: Data processing
- **openpyxl**: Excel file generation
- **plotly**: Interactive charts

### Python Version
- Python 3.9+

### Architecture
- **Modular design**: Separated concerns (audit logic, file handling, reporting)
- **Extensible**: Easy to add new membership types and red flag criteria
- **Cached resources**: Fast performance with Streamlit caching

## 📝 How to Add New Membership Types

1. Edit `config/red_flag_rules.json`
2. Add new membership type under `"membership_types"`
3. Define red flag criteria for that type
4. Use the AskUserQuestion approach (as documented in FUTURE_ENHANCEMENTS.md)

Example:
```json
"monthly": {
  "name": "Monthly Membership",
  "rules": {
    "date_diff_min_days": 28,
    "date_diff_max_days": 31,
    "min_dues_amount": 50,
    "expected_pay_type": "MONTHLY DRAFT",
    "expected_end_draft": null,
    "expected_cycle": 12
  },
  "enabled": true
}
```

## 🐛 Troubleshooting

### "Module not found" errors
```bash
pip3 install -r requirements.txt
```

### Streamlit not found
```bash
# Add to PATH or use full path
python3 -m streamlit run audit_app.py
```

### Port already in use
```bash
streamlit run audit_app.py --server.port=8502
```

### File validation errors
- Ensure CSV/Excel has required columns: Last Name, First Name, Member #, Join Date, Exp Date, etc.
- Check that file is not corrupted
- Verify export format from gym software

## 📧 Support

For issues or questions:
1. Check [README_STAFF.md](README_STAFF.md) for common problems
2. Review [FUTURE_ENHANCEMENTS.md](FUTURE_ENHANCEMENTS.md) for planned features
3. Check `outputs/` folder for generated reports

## 📄 License

Internal tool for gym operations. Not for redistribution.

---

**Created**: January 18, 2026
**Version**: 1.0.0
**Status**: Production Ready
**Last Tested**: January 18, 2026 with 1,174 records
