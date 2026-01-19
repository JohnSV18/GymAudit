# Gym Membership Audit Tool - User Guide

## Quick Start Guide for Staff

### What This Tool Does
This tool automatically checks gym membership spreadsheets for errors and problems. When you upload a file, it:
- ✅ Checks every membership for common issues
- ⚠️ Highlights accounts that need attention
- 📊 Shows you statistics about what's wrong
- 📥 Creates a downloadable report with all problem accounts highlighted in yellow

---

## How to Use the Tool

### Step 1: Start the Application

**Option A - If you're the admin:**
1. Open Terminal (Mac) or Command Prompt (Windows)
2. Navigate to the gymAudits folder
3. Type: `streamlit run audit_app.py`
4. Press Enter

**Option B - If someone else is running it:**
1. Open your web browser (Chrome, Safari, Firefox, etc.)
2. Go to: `http://[computer-name]:8501`
   - Your manager will give you the exact address
3. The audit tool will appear in your browser

### Step 2: Upload Your Files

1. Look for the **"Choose CSV or Excel files"** section
2. Click the **"Browse files"** button
3. Select one or more membership spreadsheets from your computer
   - You can select multiple files at once by holding Ctrl (Windows) or Cmd (Mac)
   - Supported formats: `.csv`, `.xlsx`, `.xls`
4. After selecting, you'll see a confirmation showing how many files were uploaded

### Step 3: Process the Files

1. Click the big blue **"🚀 Process Files"** button
2. Wait while the tool analyzes your files (this usually takes a few seconds)
3. A progress bar will show you the status

### Step 4: Review the Results

You'll see several sections:

#### Overall Summary (at the top)
- **Total Files**: How many files you uploaded
- **Total Records**: How many memberships were checked
- **⚠️ Flagged**: How many accounts have issues (and what percentage)
- **Financial Impact**: Total money at risk from errors

#### Individual File Results
For each file, you'll see:
- Number of records checked
- How many are flagged
- A list of Member IDs that have problems
- A **"📥 Download Audit Report"** button

#### What to Do with Flagged Accounts
Accounts are flagged when they have one or more of these problems:
- ⚠️ Join and expiration dates aren't exactly 1 year apart
- ⚠️ Dues amount is less than $600
- ⚠️ Pay type doesn't say "Annual Bill"
- ⚠️ End draft date isn't 12/31/99
- ⚠️ Cycle number isn't 1
- ⚠️ Balance isn't exactly $0.00

**When you see a yellow highlighted account, it means something needs to be checked!**

### Step 5: Download Reports

1. Click the **"📥 Download Audit Report"** button for each file
2. The report will download as an Excel file
3. Open it in Excel or Numbers
4. All problem accounts are **highlighted in yellow**
5. Check the **"Notes"** column to see exactly what's wrong with each account

---

## Understanding the Reports

### What You'll See in the Excel Report

1. **All your original data** - Nothing is changed, just highlighted
2. **Yellow rows** = Accounts with problems
3. **Notes column** - Explains what's wrong with each flagged account
4. **Status column** - Shows ⚠️ FLAGGED or ✅ OK
5. **Summary sheet** - Overview of all issues found

### Example Notes

Here's what you might see in the Notes column:
- `Dues < $600 ($0.00) | Pay Type: NO PAY`
  - This means the account has $0 dues AND wrong pay type
- `Balance: $50.00 (debit)`
  - This member owes $50
- `Join/Exp dates not 1 year apart`
  - The membership dates don't add up correctly

---

## Additional Features

### Pattern Analysis Tab

Click this tab to see:
- **Charts** showing which types of errors are most common
- **Trends** by join date (are certain time periods worse?)
- **Sales rep performance** (which staff create the cleanest accounts?)

This helps identify if errors are random or if there's a pattern.

### Financial Summary Tab

Click this tab to see:
- **Total money at risk** from all errors
- **Top accounts** with the biggest financial impact
- **Active vs expired** membership comparison

This helps prioritize which accounts to fix first.

---

## Tips for Best Results

### ✅ Do This:
- Upload membership export files directly from your gym software
- Process files regularly (weekly or monthly) to catch issues early
- Download and save audit reports for your records
- Share flagged member IDs with your manager for follow-up

### ❌ Don't Do This:
- Don't modify the spreadsheet before uploading (the tool needs the original columns)
- Don't ignore yellow highlighted accounts - they need attention
- Don't upload files that aren't membership data

---

## Troubleshooting

### Problem: File won't upload
**Solution:**
- Make sure it's a CSV or Excel file (.csv, .xlsx, .xls)
- Check that the file isn't corrupted
- Make sure the file has the correct columns (Last Name, First Name, Member #, etc.)

### Problem: Tool says "Missing required columns"
**Solution:**
- The file needs specific columns to work
- Make sure you're exporting the full membership report from your gym software
- Ask your manager if you're using the right export template

### Problem: Can't access the tool in browser
**Solution:**
- Make sure the person running the server has started the application
- Check that you're using the correct URL
- Try refreshing the page
- Ask your manager for help

### Problem: Download button doesn't work
**Solution:**
- Check that your browser allows downloads
- Try a different browser
- The file might be too large - ask your manager

---

## Getting Help

If you need assistance:
1. **Ask your manager** - They can help with most issues
2. **Check this guide** - The answer might be here
3. **Note the error message** - Write down exactly what you see so your manager can help

---

## For Managers/Admins

### Starting the Tool
```bash
# Navigate to the project folder
cd /path/to/gymAudits

# Start the application
streamlit run audit_app.py
```

The tool will start and automatically open in your browser at `http://localhost:8501`

### For Network Access (Multiple Users)
If you want other staff to access it:
```bash
streamlit run audit_app.py --server.address=0.0.0.0
```

Then staff can access it at: `http://[your-computer-ip]:8501`

### Where Reports Are Saved
All generated audit reports are saved in the `outputs/` folder in the gymAudits directory.

### Updating Red Flag Rules
To change what gets flagged, edit `config/red_flag_rules.json`

---

**Remember: Yellow = Problem. Always review highlighted accounts!**

Last Updated: January 18, 2026
