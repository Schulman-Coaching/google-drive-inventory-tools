/**
 * Basic Setup Example - Google Drive Inventory Tools
 * 
 * This is the simplest way to get started with the Google Drive inventory tools.
 * Copy this code into a Google Apps Script project to begin analyzing your Drive.
 */

// Basic configuration - adjust these values for your needs
const CONFIG = {
  BATCH_SIZE: 50,                                    // Process 50 files at a time
  INVENTORY_SPREADSHEET_NAME: "My Drive Inventory",   // Name of the report spreadsheet
  LARGE_FILE_THRESHOLD_MB: 25,                       // Files larger than 25MB are "large"
  OLD_FILE_THRESHOLD_DAYS: 180,                      // Files older than 180 days are "old"
  INCLUDE_GOOGLE_FILES: true,                        // Include Docs, Sheets, Slides, etc.
  INCLUDE_TRASHED: false                             // Don't include deleted files
};

/**
 * STEP 1: Run this function to start your inventory
 * This will analyze your entire Google Drive and create a comprehensive report
 */
function startBasicInventory() {
  console.log("Starting basic Drive inventory...");
  
  // You can copy the main inventory function from drive-inventory-complete.js
  // For this example, we'll show the setup process
  
  console.log("Configuration:");
  console.log(`- Batch size: ${CONFIG.BATCH_SIZE} files`);
  console.log(`- Report name: ${CONFIG.INVENTORY_SPREADSHEET_NAME}`);
  console.log(`- Large file threshold: ${CONFIG.LARGE_FILE_THRESHOLD_MB}MB`);
  console.log(`- Old file threshold: ${CONFIG.OLD_FILE_THRESHOLD_DAYS} days`);
  
  // Create the spreadsheet
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  console.log(`Report spreadsheet: ${spreadsheet.getUrl()}`);
  
  // For the actual inventory, you would call:
  // runCompleteInventory(); // From the main script
}

/**
 * STEP 2: Check the status of your inventory
 */
function checkInventoryStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const stats = JSON.parse(scriptProperties.getProperty('inventoryStats') || '{}');
  
  if (stats.totalFiles) {
    console.log(`Inventory in progress: ${stats.totalFiles} files processed`);
    console.log(`Total size analyzed: ${formatBytes(stats.totalSize)}`);
  } else {
    console.log("No inventory in progress. Run startBasicInventory() to begin.");
  }
}

/**
 * STEP 3: Get quick statistics without full inventory
 */
function getQuickDriveStats() {
  console.log("Getting quick Drive statistics...");
  
  let fileCount = 0;
  let totalSize = 0;
  let largeFiles = 0;
  let oldFiles = 0;
  
  const files = DriveApp.searchFiles('trashed = false');
  const maxCheck = 200; // Limit for quick stats
  
  console.log("Analyzing first 200 files...");
  
  while (files.hasNext() && fileCount < maxCheck) {
    const file = files.next();
    fileCount++;
    
    const size = file.getSize();
    totalSize += size;
    
    // Check for large files
    if (size > CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      largeFiles++;
    }
    
    // Check for old files
    const age = (new Date() - file.getLastUpdated()) / (1000 * 60 * 60 * 24);
    if (age > CONFIG.OLD_FILE_THRESHOLD_DAYS) {
      oldFiles++;
    }
  }
  
  console.log("Quick Statistics:");
  console.log(`Files checked: ${fileCount}`);
  console.log(`Total size: ${formatBytes(totalSize)}`);
  console.log(`Average size: ${formatBytes(totalSize / fileCount)}`);
  console.log(`Large files (>${CONFIG.LARGE_FILE_THRESHOLD_MB}MB): ${largeFiles}`);
  console.log(`Old files (>${CONFIG.OLD_FILE_THRESHOLD_DAYS} days): ${oldFiles}`);
}

/**
 * Helper function to create or get existing spreadsheet
 */
function getOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    const spreadsheet = SpreadsheetApp.create(name);
    console.log(`Created new spreadsheet: ${name}`);
    return spreadsheet;
  }
}

/**
 * Helper function to format bytes in human-readable format
 */
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * NEXT STEPS:
 * 
 * 1. Copy the complete inventory script from src/core/drive-inventory-complete.js
 * 2. Paste it below this configuration
 * 3. Run startBasicInventory() to begin your analysis
 * 4. Check the generated spreadsheet for your results
 * 
 * For more advanced features, check out the specialized scripts in src/specialized/
 */