# Setup Guide - Daily Vendor Report Automation

Complete step-by-step instructions to set up the automation in your Google Workspace.

## Prerequisites Checklist

- ✅ Google Workspace account with Gmail, Sheets, and Drive
- ✅ Access to create Google Apps Script projects
- ✅ CSV reports arriving via email (with vendor name in filename)
- ✅ Destination Google Sheets already created (one per vendor)
- ✅ Base folder created in Google Drive for report archiving

## Phase 1: Preparation (15 minutes)

### Step 1: Identify Your Vendor Information

Create a spreadsheet with this information for each vendor:

| Vendor Name | CSV Filename Pattern | Destination Sheet ID | Tab Name |
|------------|---------------------|----------------------|----------|
| Vendor 1 | Vendor 1 to the LV.csv | sheet_id_1234... | Vendor 1 |
| Vendor 2 | Vendor 2 to the LV.csv | sheet_id_5678... | Vendor 2 |

**How to get Sheet ID:**
1. Open the destination Google Sheet
2. Copy from URL: `https://docs.google.com/spreadsheets/d/SHEET_ID_HERE/edit`
3. SHEET_ID_HERE is what you need

### Step 2: Find Your Google Drive Folder ID

This is where processed CSVs will be archived.

1. Create a folder in Google Drive called "Reports" (or use existing)
2. Open the folder
3. Copy from URL: `https://drive.google.com/drive/folders/FOLDER_ID_HERE`
4. Note the FOLDER_ID_HERE

### Step 3: Set Up Destination Sheets

For each vendor, ensure the destination sheet has:
- A tab named exactly as specified (e.g., "Vendor 1")
- First row contains headers (columns A through S, 19 columns)
- Column A = Date
- Columns R & S = Summary columns (where SUM formulas go)

## Phase 2: Script Setup (10 minutes)

### Step 1: Access Google Apps Script

Option A (Recommended):
1. Open a Google Sheet (can be any sheet)
2. Go to Extensions → Apps Script
3. This creates a new script bound to that sheet

Option B:
1. Go to script.google.com
2. Click "New Project"
3. Later, authorize it to access your Sheets and Drive

### Step 2: Add the Automation Code

1. Delete the default `myFunction()` code
2. Copy all of `daily-vendor-automation.gs`
3. Paste into the Apps Script editor
4. Click Save

### Step 3: Configure Vendors

At the top of the script, find:

```javascript
const VENDOR_CONFIG = {
  'Vendor 1 to the LV': {
    sheetId: 'YOUR_SHEET_ID_HERE',
    tabName: 'Vendor 1',
    processFunction: 'processVendor1'
  },
  'Vendor 2 to the LV': {
    sheetId: 'YOUR_SHEET_ID_HERE',
    tabName: 'Vendor 2',
    processFunction: 'processVendor2'
  },
  // Add more vendors...
};
```

**Fill in for EACH vendor:**
- `'Vendor X to the LV'` = CSV filename (without .csv extension)
- `sheetId: 'YOUR_SHEET_ID_HERE'` = Paste the Sheet ID from Step 1
- `tabName: 'Vendor X'` = Tab name in destination sheet
- Keep `processFunction` as-is

### Step 4: Set Drive Folder ID

Find this line:
```javascript
const REPORTS_DRIVE_FOLDER_ID = 'YOUR_REPORTS_DRIVE_FOLDER_ID';
```

Replace with your folder ID from Phase 1, Step 2.

### Step 5: Test Configuration

Run the test function to check for configuration errors:
1. In Apps Script, select `testAutomation` from the dropdown
2. Click the Play button (▶)
3. Check logs (View → Logs) for any errors
4. If no errors, you're ready for triggers

## Phase 3: Authorization (5 minutes)

### Step 1: Grant Permissions

1. Run any function (e.g., `testAutomation`)
2. Apps Script will ask for permissions
3. Click "Review Permissions"
4. Select your Google account
5. Click "Allow" for each permission request

**This grants access to:**
- Read Gmail messages and attachments
- Read/write Google Sheets
- Create/read files in Google Drive

### Step 2: Verify Authorization

1. Go to Apps Script Settings
2. Look for your project in "Connected Projects"
3. You should see your Google account listed

## Phase 4: Enable Triggers (5 minutes)

### Step 1: Set Up Daily Automation

1. In Apps Script, select `setupDailyTrigger` from dropdown
2. Click Play (▶)
3. Check logs for: "Daily trigger set for 8 AM (Tue-Sat only)"
4. Go to Triggers (clock icon on left) and verify trigger exists

### Step 2: Set Up Weekly Summary

1. In Apps Script, select `setupWeeklyTrigger` from dropdown
2. Click Play (▶)
3. Check logs for: "Weekly trigger set for Saturday 10 AM"
4. Go to Triggers and verify trigger exists

**Trigger Settings Should Show:**
- `automateEmailReports`: Every day at 8 AM
- `weeklySum`: Every Saturday at 10 AM

## Phase 5: Testing (15 minutes)

### Test 1: Manual Automation Run

1. In Apps Script, select `testAutomation`
2. Click Play (▶)
3. Check logs (View → Logs) for:
   - Email search results
   - CSV parsing
   - Sheets population
   - No errors

### Test 2: Manual Weekly Summary

1. Select `testWeeklySum`
2. Click Play (▶)
3. Check logs for:
   - All vendors processed
   - SUM formulas added
   - No errors

### Test 3: Debug Gmail Search (Optional)

If no emails found in Test 1:
1. Select `debugTodaysEmails`
2. Click Play (▶)
3. Review logs to troubleshoot email search

## Phase 6: Go Live

### Verify Checklist

- ✅ All vendors configured in VENDOR_CONFIG
- ✅ REPORTS_DRIVE_FOLDER_ID set correctly
- ✅ Permissions granted to all accounts
- ✅ Daily trigger active (8 AM)
- ✅ Weekly trigger active (Saturday 10 AM)
- ✅ Test functions ran without errors
- ✅ Destination sheets have correct tabs

### First Live Run

1. Wait for 8 AM on a Tuesday
2. Monitor execution logs (View → Logs)
3. Check destination sheets for data
4. Verify files uploaded to Google Drive

### Post-Launch Monitoring

**First Week:**
- Check logs daily for errors
- Verify data accuracy in sheets
- Confirm Drive folder organization

**Ongoing:**
- Monitor logs weekly
- Review data quality
- Check for any error patterns

## Customization

### Change Automation Time

Edit `setupDailyTrigger()`:
```javascript
.atHour(8)  // Change 8 to desired hour (0-23)
```

### Change Weekly Summary Time

Edit `setupWeeklyTrigger()`:
```javascript
.atHour(10)  // Change 10 to desired hour
```

### Exclude Different Days

Edit `automateEmailReports()`:
```javascript
if (today === 0 || today === 1) {  // 0=Sunday, 1=Monday
  return;
}
```

### Modify Email Search Query

Edit the searchQuery in `automateEmailReports()`:
```javascript
const searchQuery = `in:updates subject:"Scheduled Report:" has:attachment -subject:"weekly" after:${formatDateForGmail(todayStart)}`;
```

Modify:
- `subject:"Scheduled Report:"` - Match your email subject
- `in:updates` - Change to `in:inbox` or other label
- `-subject:"weekly"` - Exclude emails matching pattern

### Add Vendor Disposition Clearing

Edit `clearUnwantedDispositions()`:
```javascript
var targetTabs = ['Vendor 1', 'Vendor 2', 'Vendor 3']; // Add your tabs

var clearTriggers = [
  'No Agents - Disconnect',  // Add or remove as needed
  'Duplicate', 
  // ... more dispositions
];
```

## Troubleshooting

### Issue: No Emails Found

**Check:**
1. Run `debugTodaysEmails()` to test Gmail search
2. Verify email subject contains "Scheduled Report:"
3. Confirm email is in Updates folder
4. Check attachment is actually a .csv file

### Issue: Duplicate Rows

**Check:**
1. Verify `SpreadsheetApp.flush()` is not commented out
2. Increase sleep time in `processVendorData()`: `Utilities.sleep(5000)` (5 seconds)
3. Check if destination sheet has formulas that auto-copy

### Issue: Weekly Sums Not Generated

**Check:**
1. Verify trigger is active (go to Triggers page)
2. Confirm Saturday 10 AM trigger exists
3. Check destination sheet has rows with dates in Column A
4. Look at logs for errors

### Issue: Files Not Uploading to Drive

**Check:**
1. Verify REPORTS_DRIVE_FOLDER_ID is correct
2. Run `findReportsPeaceFolderID()` to get correct ID
3. Check account has write permissions to folder
4. Verify folder structure exists

### Issue: Configuration Errors

**Check:**
1. Verify sheet IDs are exactly correct (no spaces)
2. Confirm tab names match exactly (case-sensitive)
3. Check vendor names match CSV filename (without .csv)
4. Ensure no typos in VENDOR_CONFIG

## Support

For detailed troubleshooting:
1. Check Apps Script execution logs (View → Logs)
2. Run debug functions (`debugTodaysEmails()`, `debugGmailSearch()`)
3. Check Google Cloud Console for API errors
4. Review the main README.md for architecture overview

## Next Steps

1. ✅ Set up triggers
2. ✅ Run test functions
3. ✅ Monitor first week
4. ✅ Adjust based on results
5. ✅ Document any custom modifications

---

**Questions?** Review the README.md or check your execution logs (View → Logs) in Apps Script for detailed error messages.
