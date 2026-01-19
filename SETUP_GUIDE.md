# Setup Guide - Gym Membership Audit Tool

This guide explains how to install and run the Gym Membership Audit Tool on your computer.

**Important:** This is a standalone application. You do NOT need Claude or any AI service to run it. It's just a Python program that runs locally on your computer.

---

## Quick Start (If You Already Have Python)

```bash
# 1. Open Command Prompt (Windows) or Terminal (Mac)
# 2. Navigate to the project folder
cd path/to/gymAudits

# 3. Install dependencies (one-time only)
pip install -r requirements.txt

# 4. Run the application
streamlit run audit_app.py

# 5. Open your browser to: http://localhost:8501
```

---

## Detailed Setup Instructions

### For Windows Users

#### Step 1: Install Python

1. **Check if Python is already installed:**
   - Press `Windows + R`, type `cmd`, press Enter
   - Type `python --version` and press Enter
   - If you see something like `Python 3.9.0`, skip to Step 2
   - If you see an error, continue below

2. **Download Python:**
   - Go to: https://www.python.org/downloads/
   - Click the big yellow "Download Python 3.x.x" button
   - Run the downloaded installer

3. **Install Python:**
   - **IMPORTANT:** Check the box that says **"Add Python to PATH"** at the bottom of the installer
   - Click "Install Now"
   - Wait for installation to complete
   - Click "Close"

4. **Verify installation:**
   - Open a NEW Command Prompt window
   - Type `python --version`
   - You should see the Python version number

#### Step 2: Download the Project

**Option A: Download ZIP (Easiest)**
1. Go to the GitHub repository page
2. Click the green "Code" button
3. Click "Download ZIP"
4. Extract the ZIP file to a folder (e.g., `C:\Users\YourName\Desktop\gymAudits`)

**Option B: Using Git**
```bash
git clone https://github.com/YOUR_USERNAME/gymAudits.git
```

#### Step 3: Install Dependencies

1. Open Command Prompt
2. Navigate to the project folder:
   ```bash
   cd C:\Users\YourName\Desktop\gymAudits
   ```
   (Replace with your actual folder path)

3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

4. Wait for installation to complete (may take 1-2 minutes)

#### Step 4: Run the Application

1. In the same Command Prompt window, type:
   ```bash
   streamlit run audit_app.py
   ```

2. You should see something like:
   ```
   You can now view your Streamlit app in your browser.
   Local URL: http://localhost:8501
   ```

3. Your default browser should open automatically. If not, open your browser and go to:
   ```
   http://localhost:8501
   ```

4. You should see the Gym Membership Audit Tool!

#### Step 5: Using the Tool

1. Click "Browse files" or drag and drop your membership spreadsheet
2. Click "Process Files"
3. Review the results
4. Click "Download Audit Report" to get your highlighted Excel file

#### Step 6: Stopping the Application

- Go back to the Command Prompt window
- Press `Ctrl + C`
- The application will stop

---

### For Mac Users

#### Step 1: Install Python

1. **Check if Python is already installed:**
   - Open Terminal (press `Cmd + Space`, type "Terminal", press Enter)
   - Type `python3 --version` and press Enter
   - If you see `Python 3.x.x`, skip to Step 2

2. **Install Python (if needed):**
   - Go to: https://www.python.org/downloads/
   - Download and install the latest Python 3 version
   - Or use Homebrew: `brew install python3`

#### Step 2: Download and Setup

1. Download the project from GitHub (ZIP or git clone)

2. Open Terminal and navigate to the folder:
   ```bash
   cd ~/Desktop/gymAudits
   ```

3. Install dependencies:
   ```bash
   pip3 install -r requirements.txt
   ```

4. Run the application:
   ```bash
   streamlit run audit_app.py
   ```

   Or if that doesn't work:
   ```bash
   python3 -m streamlit run audit_app.py
   ```

5. Open browser to: http://localhost:8501

---

## Running the Tool Daily

Once everything is installed, you only need to do this each time:

### Windows:
```bash
# Open Command Prompt, then:
cd C:\Users\YourName\Desktop\gymAudits
streamlit run audit_app.py
```

### Mac:
```bash
# Open Terminal, then:
cd ~/Desktop/gymAudits
streamlit run audit_app.py
```

**Tip:** You can create a shortcut/batch file to make this easier (see "Creating a Shortcut" section below).

---

## Creating a Shortcut (Windows)

To make launching easier, create a batch file:

1. Open Notepad
2. Paste this (edit the path to match your folder):
   ```batch
   @echo off
   cd C:\Users\YourName\Desktop\gymAudits
   streamlit run audit_app.py
   pause
   ```
3. Save as `Run_Audit_Tool.bat` on your Desktop
4. Double-click the file anytime to launch the tool!

---

## Troubleshooting

### "python is not recognized as an internal or external command"
- Python wasn't added to PATH during installation
- **Fix:** Reinstall Python and make sure to check "Add Python to PATH"

### "pip is not recognized"
- Try using `python -m pip` instead of just `pip`
- Example: `python -m pip install -r requirements.txt`

### "streamlit is not recognized"
- Try running with Python module syntax:
  ```bash
  python -m streamlit run audit_app.py
  ```

### "No module named streamlit" or similar
- Dependencies weren't installed properly
- Run: `pip install -r requirements.txt` again

### "Address already in use" or "Port 8501 is in use"
- Another instance is already running, or another program is using that port
- **Fix:** Either close the other instance, or run on a different port:
  ```bash
  streamlit run audit_app.py --server.port 8502
  ```
  Then go to http://localhost:8502

### The browser doesn't open automatically
- Manually open your browser and go to: http://localhost:8501

### "Permission denied" errors
- On Windows: Run Command Prompt as Administrator
- On Mac: Try prefixing commands with `sudo` (e.g., `sudo pip3 install -r requirements.txt`)

### Excel file won't open / is corrupted
- Make sure you're uploading a valid .csv or .xlsx file
- The file must have the expected columns (Last Name, First Name, Member #, etc.)

---

## Frequently Asked Questions

### Do I need internet to run this?
**For installation:** Yes, to download Python and the dependencies.
**For daily use:** No, it runs completely offline on your computer.

### Do I need Claude or AI to run this?
**No.** Claude helped create this tool, but it runs completely independently. It's just a Python program.

### Can multiple people use it at once?
Yes, each person runs it on their own computer. Or one person can run it and others access it via the network URL (see README.md for network setup).

### Where are the audit reports saved?
In the `outputs` folder inside the project directory. You can also download them directly from the web interface.

### What file formats are supported?
- CSV (.csv)
- Excel (.xlsx, .xls)

### Can I change the red flag rules?
Yes, edit the `config/red_flag_rules.json` file. See FUTURE_ENHANCEMENTS.md for details.

---

## System Requirements

- **Operating System:** Windows 10/11, macOS 10.14+, or Linux
- **Python:** Version 3.8 or higher
- **RAM:** 4GB minimum (8GB recommended)
- **Disk Space:** ~500MB for Python and dependencies
- **Browser:** Chrome, Firefox, Safari, or Edge (any modern browser)

---

## Getting Help

1. **Check this guide** for common issues
2. **Read README_STAFF.md** for usage instructions
3. **Contact your manager** for technical support

---

## Summary

| Step | Windows | Mac |
|------|---------|-----|
| 1. Install Python | python.org (check "Add to PATH") | python.org or `brew install python3` |
| 2. Open terminal | Command Prompt | Terminal |
| 3. Go to folder | `cd C:\path\to\gymAudits` | `cd ~/path/to/gymAudits` |
| 4. Install deps | `pip install -r requirements.txt` | `pip3 install -r requirements.txt` |
| 5. Run app | `streamlit run audit_app.py` | `streamlit run audit_app.py` |
| 6. Open browser | http://localhost:8501 | http://localhost:8501 |

**That's it! No AI, no cloud services, no subscriptions - just Python running locally on your computer.**

---

Last Updated: January 2026
