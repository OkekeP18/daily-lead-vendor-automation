# Quick Reference Guide

Fast lookup for common tasks and key functions.

## Getting Started (this is copy-Paste ready, you're welcome)

### 1. Set Configuration
```javascript
// In daily-vendor-automation.gs, line ~15
const VENDOR_CONFIG = {
  'Vendor 1 to the LV': {
    sheetId: 'YOUR_SHEET_ID_1',
    tabName: 'Vendor 1'
  },
  'Vendor 2 to the LV': {
    sheetId: 'YOUR_SHEET_ID_2',
    tabName: 'Vendor 2'
  },
  // Repeat for each vendor
};

const REPORTS_DRIVE_FOLDER_ID = 'YOUR_FOLDER_ID';
```

### 2. Enable Automation
```javascript
// Run these ONCE from Apps Script editor
setupDailyTrigger();    // 8 AM automation
setupWeeklyTrigger();   // Saturday 10 AM summaries
```

### 3. Test Before Going Live
```javascript
// Run to verify everything works
testAutomation();       // Simulates daily run
testWeeklySum();        // Simulates weekly summary
```

## Key Functions at a Glance

| Function | What It Does | Run When |
|----------|-------------|----------|
| `automateEmailReports()` | Main daily automation | Automatic (8 AM Tue-Sat) |
| `weeklySum()` | Generates weekly summaries | Automatic (Sat 10 AM) |
| `testAutomation()` | Test run without triggers | Before going live |
| `testWeeklySum()` | Test weekly summary | Before going live |
| `setupDailyTrigger()` | Create 8 AM trigger | Once (initial setup) |
| `setupWeeklyTrigger()` | Create Saturday trigger | Once (initial setup) |
| `debugTodaysEmails()` | Test Gmail search | When emails not found |
| `findReportsPeaceFolderID()` | Find Drive folder ID | When finding folder |

## Email Search Customization

### Current Search Query
```javascript
in:updates subject:"Scheduled Report:" has:attachment -subject:"weekly"
```

### Modify For Your Needs

**Search in Inbox instead of Updates:**
```javascript
in:inbox subject:"Scheduled Report:" has:attachment
```

**Different subject line:**
```javascript
in:updates subject:"Daily Report:" has:attachment
```

**Search for specific vendor in subject:**
```javascript
in:updates subject:"Scheduled Report:" subject:"Vendor Name" has:attachment
```

**Search by sender:**
```javascript
in:updates from:"reports@example.com" has:attachment
```

## Troubleshooting Commands

### Check if emails are being found
```javascript
debugTodaysEmails();  // Run this if no emails found
```

### Check Gmail search variations
```javascript
debugGmailSearch();   // Test different search patterns
```

### Find your Drive folder ID
```javascript
findReportsPeaceFolderID();  // Outputs folder ID to logs
```

### Check last 10 execution logs
```
View → Logs (in Apps Script)
```

## Common Customizations

### Change automation time from 8 AM to 10 AM
Find `setupDailyTrigger()` and change:
```javascript
.atHour(8)  // Change to .atHour(10)
```

### Change weekly summary from 10 AM to 3 PM (15:00)
Find `setupWeeklyTrigger()` and change:
```javascript
.atHour(10)  // Change to .atHour(15)
```

### Exclude different days (e.g., exclude Friday)
Find `automateEmailReports()` and change:
```javascript
if (today === 0 || today === 1) {  // 0=Sun, 1=Mon
```

Day mapping: 0=Sunday, 1=Monday, 2=Tuesday, 3=Wednesday, 4=Thursday, 5=Friday, 6=Saturday

### Add more disposition triggers to clear
Find `clearUnwantedDispositions()`:
```javascript
var clearTriggers = [
  'No Agents - Disconnect',
  'Duplicate',
  'Your New Disposition',  // Add here
];
```

### Add more tabs to disposition clearing
Find `clearUnwantedDispositions()`:
```javascript
var targetTabs = ['Vendor 1', 'Vendor 2', 'Vendor 3'];  // Add your tabs
```

## Configuration Reference

### VENDOR_CONFIG Structure
```javascript
const VENDOR_CONFIG = {
  'CSV Filename (no extension)': {
    sheetId: 'Google Sheet ID',
    tabName: 'Sheet tab name'
  }
};
```

**Example:**
- Email contains: `Vendor 1 to the LV.csv`
- Sheet URL: `https://docs.google.com/spreadsheets/d/1ABC123/edit`
- Destination tab in sheet: "Vendor 1"

```javascript
'Vendor 1 to the LV': {
  sheetId: '1ABC123',
  tabName: 'Vendor 1'
}
```

## Execution Schedule

### What Runs When

**Every Tuesday-Saturday at 8:00 AM:**
- Monitors Gmail for reports
- Parses CSV files
- Populates destination sheets
- Uploads CSVs to Drive
- Clears unwanted dispositions
- Formats dates

**Every Saturday at 10:00 AM:**
- Finds last 5 data rows
- Creates SUM formulas for columns R & S
- Colors summary rows green

**Never Runs:**
- Sundays (no automation)
- Mondays (no automation, but reports may arrive)

## Performance Tips

### If script runs too slow
1. Reduce number of vendors temporarily
2. Increase sleep time: `Utilities.sleep(15000)` (15 seconds)
3. Check for heavy formula calculations in destination sheets

### If you get API quota errors
1. Reduce frequency (less common, but possible with 20+ vendors)
2. Reduce number of concurrent vendors
3. Wait 1 hour and try again

### For optimal performance
- 17 daily reports = ~3 minutes execution
- 10 second delay between each vendor is intentional
- Weekly summary = < 1 minute for all vendors

## Logs & Monitoring

### Check Execution Logs
1. Open Apps Script project
2. Click "View" → "Logs"
3. Logs show:
   - Emails found
   - Vendors processed
   - Errors (if any)
   - File uploads
   - Data population

### What to Look For
✅ `Found X email threads` - Emails detected
✅ `Processing: Vendor X` - Vendor being processed
✅ `Successfully processed` - Completed without error
✅ `Uploaded: filename.csv` - File archived
❌ `Error processing` - Something failed
❌ `No sheet ID configured` - Configuration missing

### Set Reminders
- **Weekly**: Check logs every Monday
- **Monthly**: Run `testAutomation()` to verify still working
- **Quarterly**: Review logs for error patterns

## Security Notes

### What This Script Accesses
- ✅ Read emails from Gmail (only with "Scheduled Report:" subject)
- ✅ Read/Write to specified Google Sheets
- ✅ Create/Upload files to specified Drive folder
- ✅ Create time-based triggers

### What It Does NOT Access
- ❌ Emails older than today
- ❌ Other emails in your inbox
- ❌ Other Google Drive files outside specified folder
- ❌ Other people's accounts

### Best Practices
1. Use service account for production
2. Limit folder permissions to script user
3. Review logs monthly for anomalies
4. Keep sheet IDs private (don't share code with IDs)

## Quick Fixes

### "Sheet not found" error
- Check tab name matches exactly (case-sensitive)
- Verify tab exists in destination sheet
- Confirm sheet ID is correct

### "No emails found"
- Run `debugTodaysEmails()` to test search
- Verify email subject matches search query
- Check email is in Updates folder
- Confirm attachment is .csv file

### Duplicate rows appearing
- Increase sleep time in `processVendorData()`
- Change: `Utilities.sleep(3000)` → `Utilities.sleep(5000)`
- Verify no manual edits during automation

### Weekly sums not generating
- Check destination sheet has Column A dates
- Verify Saturday 10 AM trigger is active
- Check logs for errors
- Confirm data exists in previous rows

## Testing Checklist

- ✅ Configuration saved
- ✅ All vendors added to VENDOR_CONFIG
- ✅ Folder ID set correctly
- ✅ Permissions granted
- ✅ `testAutomation()` ran successfully
- ✅ `testWeeklySum()` ran successfully
- ✅ No errors in logs
- ✅ `setupDailyTrigger()` completed
- ✅ `setupWeeklyTrigger()` completed
- ✅ Triggers visible in Triggers page
- ✅ Ready to go live!

---

**Need help?** See README.md for detailed architecture or SETUP_GUIDE.md for step-by-step instructions.
