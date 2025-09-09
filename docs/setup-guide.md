# Setup Guide - Google Drive Inventory Tools

This guide will walk you through setting up and running the Google Drive Inventory Tools to analyze your Google Drive files.

## Prerequisites

- Google Account with Google Drive access
- Google Apps Script access (script.google.com)
- Basic familiarity with Google Apps Script (helpful but not required)

## Quick Start (5 minutes)

### 1. Create a New Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click **"New Project"**
3. Give your project a name (e.g., "Drive Inventory")

### 2. Choose Your Script

Select the appropriate script for your needs:

- **Complete Analysis**: Use `src/core/drive-inventory-complete.js` for comprehensive Drive analysis
- **Images Only**: Use `src/specialized/image-inventory.js` for photo and graphics analysis
- **Documents Only**: Use `src/specialized/document-inventory.js` for text files, PDFs, etc.
- **Large Files**: Use `src/specialized/large-files-finder.js` for storage optimization
- **Security Audit**: Use `src/specialized/shared-files-audit.js` for permission analysis

### 3. Copy the Script

1. Copy the entire content of your chosen script
2. Replace the default `Code.gs` content with the copied script
3. Save the project (Ctrl+S or Cmd+S)

### 4. Run the Script

1. Select the main function (usually `inventoryDrive()` or similar)
2. Click the **"Run"** button (▶️)
3. **Authorize permissions** when prompted:
   - Click "Review permissions"
   - Choose your Google account
   - Click "Advanced" → "Go to [Project Name] (unsafe)"
   - Click "Allow"

### 5. View Results

- The script will create a new Google Spreadsheet with your results
- Check the Google Apps Script logs for the spreadsheet URL
- The spreadsheet will contain multiple sheets with different analyses

## Detailed Configuration

### Basic Settings

All scripts include a `CONFIG` object at the top that you can customize:

```javascript
const CONFIG = {
  // Process this many files per run
  BATCH_SIZE: 100,
  
  // Name of the results spreadsheet
  INVENTORY_SPREADSHEET_NAME: "📊 Drive Inventory Report",
  
  // Size threshold for "large" files (in MB)
  LARGE_FILE_THRESHOLD_MB: 50,
  
  // Age threshold for "old" files (in days)
  OLD_FILE_THRESHOLD_DAYS: 365,
  
  // Include Google Workspace files (Docs, Sheets, etc.)
  INCLUDE_GOOGLE_FILES: true,
  
  // Include deleted/trashed files
  INCLUDE_TRASHED: false,
  
  // Analyze sharing permissions (may slow down processing)
  TRACK_PERMISSIONS: true
};
```

### Advanced Options

For large Google Drives (10,000+ files):

```javascript
// Use smaller batch sizes to avoid timeouts
BATCH_SIZE: 50,

// Enable automatic continuation
function setupAutomaticInventory() {
  startAutomaticInventory(); // This will run continuously
}
```

## Script-Specific Setup

### Complete Inventory Script

**Functions to run:**
- `runCompleteInventory()` - One-click full analysis
- `testInventory()` - Test with 10 files first
- `getQuickStats()` - Quick overview without full scan

**Reports generated:**
- Overview with summary statistics
- Complete file list
- Large files analysis
- Old files report
- Potential duplicates
- Shared files audit
- File types breakdown

### Image Inventory Script

**Functions to run:**
- `startAutomaticImageInventory()` - Analyze all images
- `getQuickImageStats()` - Quick image overview

**Reports generated:**
- Image files list with formats and dimensions
- Large images report
- Images by format analysis
- Old images cleanup candidates

### Document Inventory Script

**Functions to run:**
- `startAutomaticDocumentInventory()` - Analyze all documents
- `getQuickDocumentStats()` - Quick document overview

**Reports generated:**
- All documents with types and sharing info
- Google Workspace vs other documents
- Document sharing analysis
- Large documents report

### Large Files Finder

**Functions to run:**
- `startAutomaticLargeFilesAnalysis()` - Find space-consuming files
- `getQuickLargeFilesStats()` - Quick large files overview

**Reports generated:**
- Large files by size category
- Cleanup candidates with scores
- Storage optimization recommendations
- Files by folder analysis

### Shared Files Audit

**Functions to run:**
- `startAutomaticSharedFilesAudit()` - Security analysis
- `getQuickSecurityCheck()` - Quick security overview

**Reports generated:**
- All shared files with permission details
- Public files (security risk)
- External sharing analysis
- High-risk files with recommendations
- Security recommendations

## Troubleshooting

### Common Issues

**"Script timeout" error:**
- Use automatic mode: `startAutomaticInventory()`
- Reduce `BATCH_SIZE` to 25-50
- Use specialized scripts instead of complete inventory

**"Permission denied" errors:**
- Re-run the authorization process
- Make sure you're the owner of files being analyzed
- Some enterprise accounts have restrictions

**"Quota exceeded" errors:**
- The script is making too many API calls
- Wait 1 hour and try again
- Use smaller batch sizes

### Performance Tips

**For small drives (< 1,000 files):**
- Use `BATCH_SIZE: 200`
- Run complete inventory in single session

**For medium drives (1,000-10,000 files):**
- Use `BATCH_SIZE: 100`
- Use automatic mode for unattended processing

**For large drives (10,000+ files):**
- Use `BATCH_SIZE: 50`
- Set up hourly scheduled runs: `setupHourlyInventory()`
- Consider using specialized scripts for specific analysis

### Getting Help

1. Check the logs in Google Apps Script for error messages
2. Review the generated spreadsheet for partial results
3. Try the test functions first (`testInventory()`, quick stats functions)
4. Use smaller batch sizes if experiencing timeouts

## Security and Privacy

- **No data leaves Google**: All processing happens in Google Apps Script
- **Your data stays yours**: Reports are saved to your own Google Drive
- **Permissions**: Scripts only request Drive read access
- **Open source**: All code is visible and auditable

## Next Steps

After your first successful run:

1. **Review the results** in your generated spreadsheet
2. **Customize the configuration** for your specific needs
3. **Set up automation** for regular analysis
4. **Try specialized scripts** for focused analysis
5. **Share insights** with your team or organization

## Automation Options

### Schedule Regular Runs

Set up automatic inventory runs:

```javascript
// Run every hour until complete
function setupHourlyInventory() {
  setupHourlyInventory();
}

// Run continuously (handles Google Apps Script time limits)
function setupContinuousInventory() {
  startAutomaticInventory();
}
```

### Email Notifications

Add email notifications when inventory completes:

```javascript
function sendCompletionEmail() {
  const email = Session.getActiveUser().getEmail();
  const subject = "Drive Inventory Complete";
  const body = `Your Google Drive inventory has finished processing. 
                View results at: ${spreadsheet.getUrl()}`;
  
  GmailApp.sendEmail(email, subject, body);
}
```

This setup guide should get you up and running quickly. For more advanced usage and customization options, check out the other documentation files in this repository.