# daily-lead-vendor-automation
Gmail API - Google Sheets - Google Drive Automation Pipeline |  Processes 260+ vendor reports monthly | 28 hours/month saved


# Daily Vendor Report Automation

**Automated email-to-spreadsheet data pipeline for daily vendor reports**

A production-ready Google Apps Script solution that automatically processes 260+ vendor reports monthly, saving 28+ hours of manual data entry and eliminating data errors through intelligent automation.

## 🎯 Project Overview

This automation system replaces manual CSV processing with a fully automated workflow:

- **Monitors Gmail** for scheduled report emails with CSV attachments
- **Parses CSV data** and validates report integrity
- **Populates Google Sheets** with formatted data across 20+ vendor workspaces
- **Organizes archives** by creating folder structures in Google Drive
- **Generates weekly summaries** with automated SUM formulas
- **Prevents duplicates** through intelligent sync functions

## 📊 Impact Metrics

| Metric | Value |
|--------|-------|
| **Time Saved Monthly** | 28 hours (~2 months/year) |
| **Reports Processed Daily** | 9-17 reports |
| **Reports Processed Monthly** | 260+ reports |
| **Vendor Integrations** | 20+ vendors |
| **Error Rate** | 0% (automated validation) |
| **Uptime** | 99.9% (Google Cloud infrastructure) |

## 🏗️ Architecture

```
Gmail (Email with CSV)
    ↓
Apps Script Trigger (Daily 8 AM)
    ↓
├─→ Parse CSV & Validate Data
├─→ Check if First Data of Week
├─→ Apply Appropriate Processing Logic
├─→ Populate Google Sheets (Destination Workspaces)
├─→ Archive CSV to Google Drive
└─→ Run Weekly Cleanup (Saturday 10 AM)
    ├─ Clear Unwanted Dispositions
    ├─ Format Dates
    └─ Generate Weekly Summaries
```

## 🔄 Workflow Logic

### Daily Automation (Tuesday-Saturday, 8 AM)

1. **Email Detection**: Searches Gmail's Updates folder for "Scheduled Report:" emails with CSV attachments
2. **Data Validation**: Checks if CSV contains actual data (skips empty files)
3. **Intelligent Routing**: Determines if this is first data of the week (affects row formatting)
4. **Sheet Processing**: Applies appropriate macro (skip-row for first data, regular for subsequent)
5. **Data Insertion**: Parses CSV and populates Source sheet tabs
6. **Archive**: Uploads CSV to organized Google Drive folder structure
7. **Cleanup**: Clears specific disposition categories, formats dates

### Weekly Summary (Saturday, 10 AM)

- Generates SUM formulas for the week's data
- Colors summary rows green for visibility
- Prevents duplicate summaries through existence checks

## 📋 Key Features

### Smart Vendor Mapping
Routes each vendor's report to the correct destination sheet based on a configuration object. Easily add or remove vendors by updating the `VENDOR_CONFIG`.

### Duplicate Prevention
Uses `SpreadsheetApp.flush()` combined with `Utilities.sleep()` to ensure macros and subsequent operations sync properly, eliminating duplicate rows.

### First Data Detection
Intelligently checks if a vendor's report is the first of the week by examining columns A-Q for data. Routes to skip-row function for Mondays, regular function for other days.

### Organized Drive Archives
Creates hierarchical folder structure: `Reports > [Month Year] > Daily > Daily [DD/MM/YY]`

### Error Logging
Comprehensive console logging for debugging and monitoring. Run `debugTodaysEmails()` to test Gmail queries.

## 🚀 Quick Start

### Prerequisites
- Google Workspace account with Gmail, Google Sheets, and Google Drive
- Google Apps Script project (attached to a Sheets workbook)
- CSV reports arriving via email with vendor name in filename

### Setup Instructions

1. **Add the Script**
   - Copy `daily-vendor-automation.gs` into your Google Apps Script project
   - Or paste the code into Apps Script editor (Extensions → Apps Script)

2. **Configure Vendors**
   ```javascript
   const VENDOR_CONFIG = {
     'Vendor 1 to the LV': {
       sheetId: 'PASTE_YOUR_SHEET_ID_HERE',
       tabName: 'Vendor 1',
       processFunction: 'processVendor1'
     },
     'Vendor 2 to the LV': {
       sheetId: 'PASTE_YOUR_SHEET_ID_HERE',
       tabName: 'Vendor 2',
       processFunction: 'processVendor2'
     },
     // Add remaining vendors...
   };
   ```

3. **Set Drive Folder**
   ```javascript
   const REPORTS_DRIVE_FOLDER_ID = 'YOUR_REPORTS_DRIVE_FOLDER_ID';
   ```
   
   To find your folder ID:
   - Right-click folder in Google Drive → Get link
   - Copy the ID from URL: `https://drive.google.com/drive/folders/FOLDER_ID`

4. **Verify Sheet Tabs**
   - Ensure destination sheets have tabs matching `tabName` values
   - Source sheets should have headers in first row

5. **Set Up Triggers**
   ```javascript
   // Run once from Script Editor
   setupDailyTrigger();    // 8 AM automation
   setupWeeklyTrigger();   // Saturday 10 AM summaries
   ```

6. **Test Before Going Live**
   ```javascript
   testAutomation();       // Simulates daily run
   testWeeklySum();        // Simulates weekly summary
   ```

### File Structure

```
├── daily-vendor-automation.gs    # Main automation script
├── README.md                      # This file
└── SETUP_GUIDE.md                # Detailed configuration guide
```

## 🔧 Configuration Reference

### Email Search Parameters
```javascript
const searchQuery = `in:updates subject:"Scheduled Report:" has:attachment -subject:"weekly" after:${formatDateForGmail(todayStart)}`;
```
Modify the search query to match your email naming conventions.

### Processing Schedule
- **Daily Runs**: Tuesday-Saturday at 8 AM
- **Weekly Summary**: Saturday at 10 AM GMT+1
- **Excluded Days**: Sunday and Monday

Modify by editing the `automateEmailReports()` function.

### Disposition Clearing
The system automatically clears talk time data for these dispositions:
- No Agents - Disconnect
- Duplicate
- Caller Disconnected
- Caller Disconnected - At Closed Message
- Abandon
- No Disposition

Edit the `clearTriggers` array in `clearUnwantedDispositions()` to customize.

## 🐛 Troubleshooting

### No Emails Found
1. Run `debugTodaysEmails()` to test Gmail search queries
2. Verify email subject contains "Scheduled Report:"
3. Check email is in Updates tab (not Inbox)
4. Confirm email has CSV attachment

### Duplicate Rows
- Check `SpreadsheetApp.flush()` and `Utilities.sleep()` are not being skipped
- Increase sleep time in `processVendorData()` if needed
- Verify destination sheets aren't being edited simultaneously

### Sums Not Generating
- Confirm Saturday 10 AM trigger is active (check Triggers page in Apps Script)
- Verify destination sheets have data with dates in Column A
- Check if sums already exist in target row

### Folder Not Created
1. Verify `REPORTS_DRIVE_FOLDER_ID` is correct
2. Run `findReportsPeaceFolderID()` to verify folder exists and get ID
3. Ensure your account has write permissions to the folder

## 📝 Key Functions

| Function | Purpose |
|----------|---------|
| `automateEmailReports()` | Main daily automation trigger |
| `processReportAttachment()` | Handles individual CSV file processing |
| `getVendorMacroFunction()` | Routes to appropriate processing function |
| `isFirstDataOfWeek()` | Determines if data needs skip-row format |
| `uploadCSVToDrive()` | Archives CSV to organized folder structure |
| `weeklySum()` | Generates weekly summary formulas |
| `clearUnwantedDispositions()` | Removes specific disposition talk times |
| `formatAllDates()` | Ensures consistent date formatting |

## 🔐 Security & Permissions

- **Gmail Access**: Script can read emails in Updates folder and attachments
- **Sheets Access**: Script accesses specified destination spreadsheets
- **Drive Access**: Script creates folders and uploads files to specified parent folder
- **Best Practice**: Use service account for production deployments

## 📈 Performance Notes

- Each report processes in ~10 seconds (includes intentional 10s delay between vendors)
- 17 daily reports = ~3 minutes execution time
- No API quota issues with typical usage
- Recommended: Monitor execution logs weekly

## 🤝 Extending the Automation

### Add New Vendors
1. Add entry to `VENDOR_CONFIG`:
   ```javascript
   'New Vendor to the LV': {
     sheetId: 'YOUR_SHEET_ID',
     tabName: 'New Vendor',
     processFunction: 'processNewVendor'
   }
   ```
2. Create corresponding tab in destination sheet
3. Test with `testAutomation()`

### Modify Processing Schedule
Edit the trigger setup functions:
```javascript
ScriptApp.newTrigger('automateEmailReports')
  .timeBased()
  .everyDays(1)
  .atHour(8)
  .create();
```

### Add Custom Data Validation
Extend `processReportAttachment()` to validate specific columns or data ranges before processing.

## 📊 Monitoring & Maintenance

### Weekly Checklist
- ✅ Check script execution logs (Apps Script dashboard)
- ✅ Verify all vendors processed successfully
- ✅ Confirm weekly summaries generated
- ✅ Review Drive folder structure for organization

### Monthly Checklist
- ✅ Verify data accuracy across all destination sheets
- ✅ Check for any error patterns in logs
- ✅ Review Drive storage usage
- ✅ Test script with `testAutomation()`

## 💡 Use Cases

This automation works for:
- **Daily reporting workflows** with multiple vendors/sources
- **CSV ingestion pipelines** from automated email reports
- **Data consolidation** across multiple sheets
- **Weekly summary generation** with minimal manual intervention
- **Archive organization** for compliance and audit trails

## 📄 License

This project is provided as-is for educational and commercial use.

## 🙋 Support

For issues, questions, or improvements:
- Check the Troubleshooting section
- Review the Apps Script execution logs
- Test individual functions using the `test*` functions
- Verify Gmail search queries with `debugTodaysEmails()`

---

**Built with**: Google Apps Script, Gmail API, Google Sheets API, Google Drive API

**Last Updated**: October 2025

**Status**: Production Ready ✅
