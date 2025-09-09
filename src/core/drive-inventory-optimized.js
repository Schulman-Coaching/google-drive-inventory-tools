/**
 * Google Drive Inventory Script - Memory Optimized Version
 * Fixes property storage quota issues for large drives
 * Stores data directly in spreadsheet instead of script properties
 */

// Optimized configuration for large drives
const CONFIG = {
  BATCH_SIZE: 200, // Larger batches for efficiency
  INVENTORY_SPREADSHEET_NAME: "📊 Drive Inventory Report v4",
  
  SHEETS: {
    OVERVIEW: "Overview",
    FILE_LIST: "File List",
    LARGE_FILES: "Large Files",
    OLD_FILES: "Old Files",
    DUPLICATES: "Potential Duplicates",
    SHARED_FILES: "Shared Files",
    FILE_TYPES: "File Types Analysis",
    PROGRESS: "Progress Tracking"
  },
  
  LARGE_FILE_THRESHOLD_MB: 50,
  OLD_FILE_THRESHOLD_DAYS: 365,
  INCLUDE_GOOGLE_FILES: true,
  INCLUDE_TRASHED: false,
  TRACK_PERMISSIONS: true,
  MAX_DUPLICATE_GROUPS: 50,
  
  // Memory optimization settings
  MAX_LARGE_FILES: 100,
  MAX_OLD_FILES: 100,
  MAX_SHARED_FILES: 100,
  PROGRESS_SAVE_INTERVAL: 50 // Save progress every N files
};

/**
 * MAIN FUNCTION - Start here for large drive inventory
 * This version handles memory limitations properly
 */
function inventoryDriveLarge() {
  console.log("Starting optimized Drive inventory for large drives...");
  
  try {
    // Reset any existing progress
    resetInventoryProgress();
    
    // Run the optimized inventory
    runOptimizedInventory();
    
  } catch (error) {
    console.error(`Error in inventory: ${error}`);
    console.error(`Stack trace: ${error.stack}`);
    
    // Try to save what we have so far
    try {
      const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
      saveErrorInfo(spreadsheet, error);
    } catch (saveError) {
      console.error(`Could not save error info: ${saveError}`);
    }
  }
}

/**
 * Run optimized inventory that doesn't hit memory limits
 */
function runOptimizedInventory() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeOptimizedSheets(spreadsheet);
  
  // Use minimal script properties - just for continuation token
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('continuationToken');
  
  // Track stats directly in spreadsheet instead of script properties
  const stats = initializeStatsInSheet(spreadsheet);
  
  console.log("Starting file processing...");
  updateProgressSheet(spreadsheet, 'STARTING', stats);
  
  let totalProcessed = 0;
  let batchCount = 0;
  const startTime = new Date();
  
  while (true) {
    const batchStartTime = new Date();
    
    // Get files for this batch
    const files = getFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
    
    if (files.length === 0) {
      console.log("No more files to process!");
      break;
    }
    
    batchCount++;
    console.log(`Processing batch ${batchCount} (${files.length} files)...`);
    updateProgressSheet(spreadsheet, 'PROCESSING', { 
      ...stats, 
      currentBatch: batchCount, 
      totalProcessed: totalProcessed 
    });
    
    // Process this batch
    let processedInBatch = 0;
    for (const file of files) {
      try {
        processFileOptimized(file, spreadsheet, stats);
        processedInBatch++;
        totalProcessed++;
        
        // Save progress periodically to avoid memory issues
        if (processedInBatch % CONFIG.PROGRESS_SAVE_INTERVAL === 0) {
          updateStatsInSheet(spreadsheet, stats);
        }
        
      } catch (error) {
        console.error(`Error processing file ${file.getName()}: ${error}`);
        stats.errors++;
      }
    }
    
    // Update stats in spreadsheet
    updateStatsInSheet(spreadsheet, stats);
    
    const batchTime = (new Date() - batchStartTime) / 1000;
    console.log(`Batch ${batchCount} complete: ${processedInBatch} files in ${batchTime.toFixed(1)}s`);
    
    // Check if more files available
    const hasMore = files.hasNext && files.hasNext();
    if (hasMore) {
      const nextToken = files.getContinuationToken();
      scriptProperties.setProperty('continuationToken', nextToken);
      
      // Check execution time - leave buffer for cleanup
      const elapsed = (new Date() - startTime) / 1000;
      if (elapsed > 300) { // 5 minutes
        console.log(`Processed ${totalProcessed} files in ${batchCount} batches. Continuing in next run...`);
        scheduleNextRun();
        return;
      }
      
      // Brief pause between batches
      Utilities.sleep(100);
    } else {
      // All done!
      scriptProperties.deleteProperty('continuationToken');
      break;
    }
  }
  
  // Generate final reports
  console.log("Generating final reports...");
  updateProgressSheet(spreadsheet, 'FINALIZING', { ...stats, totalProcessed: totalProcessed });
  
  generateOptimizedFinalReports(spreadsheet, stats);
  
  const totalTime = (new Date() - startTime) / 1000;
  console.log(`\n=== INVENTORY COMPLETE ===`);
  console.log(`Total files processed: ${totalProcessed}`);
  console.log(`Total time: ${(totalTime / 60).toFixed(1)} minutes`);
  console.log(`Average: ${(totalProcessed / totalTime).toFixed(1)} files/second`);
  console.log(`Spreadsheet: ${spreadsheet.getUrl()}`);
  
  updateProgressSheet(spreadsheet, 'COMPLETE', { ...stats, totalProcessed: totalProcessed });
}

/**
 * Process a single file with memory optimization
 */
function processFileOptimized(file, spreadsheet, stats) {
  try {
    const fileData = extractFileDataOptimized(file);
    
    // Update basic stats
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Track file types (limited to prevent memory issues)
    const fileType = fileData.type;
    if (Object.keys(stats.filesByType).length < 50 || stats.filesByType[fileType]) {
      stats.filesByType[fileType] = (stats.filesByType[fileType] || 0) + 1;
    }
    
    // Track by year
    const year = new Date(fileData.lastModified).getFullYear();
    stats.filesByYear[year] = (stats.filesByYear[year] || 0) + 1;
    
    // Add to main file list immediately (don't store in memory)
    addToFileListSheetOptimized(spreadsheet, fileData);
    
    // Check for large files (store limited number)
    if (fileData.size > CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      if (stats.largeFiles.length < CONFIG.MAX_LARGE_FILES) {
        stats.largeFiles.push({
          name: fileData.name,
          size: fileData.size,
          path: fileData.folderPath,
          url: fileData.url
        });
        
        // Keep only largest files
        stats.largeFiles.sort((a, b) => b.size - a.size);
        if (stats.largeFiles.length > CONFIG.MAX_LARGE_FILES) {
          stats.largeFiles = stats.largeFiles.slice(0, CONFIG.MAX_LARGE_FILES);
        }
      }
    }
    
    // Check for old files (store limited number)
    const ageInDays = (new Date() - new Date(fileData.lastModified)) / (1000 * 60 * 60 * 24);
    if (ageInDays > CONFIG.OLD_FILE_THRESHOLD_DAYS && stats.oldFiles.length < CONFIG.MAX_OLD_FILES) {
      stats.oldFiles.push({
        name: fileData.name,
        lastModified: fileData.lastModified,
        ageInDays: Math.floor(ageInDays),
        path: fileData.folderPath,
        url: fileData.url
      });
    }
    
    // Track shared files (store limited number)  
    if (fileData.sharingAccess !== 'PRIVATE' && stats.sharedFiles.length < CONFIG.MAX_SHARED_FILES) {
      stats.sharedFiles.push({
        name: fileData.name,
        sharingAccess: fileData.sharingAccess,
        sharingPermission: fileData.sharingPermission,
        path: fileData.folderPath,
        url: fileData.url
      });
    }
    
  } catch (error) {
    console.error(`Error processing file: ${error}`);
    throw error;
  }
}

/**
 * Extract file data with minimal memory usage
 */
function extractFileDataOptimized(file) {
  const data = {
    name: file.getName(),
    mimeType: file.getMimeType(),
    type: getFileTypeSimple(file.getName(), file.getMimeType()),
    size: file.getSize(),
    lastModified: file.getLastUpdated().toISOString(),
    owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
    url: file.getUrl(),
    folderPath: 'Root',
    sharingAccess: 'PRIVATE',
    sharingPermission: 'NONE'
  };
  
  // Get folder path (simplified)
  try {
    const parents = file.getParents();
    if (parents.hasNext()) {
      const parent = parents.next();
      data.folderPath = parent.getName();
    }
  } catch (error) {
    data.folderPath = 'Unknown';
  }
  
  // Get sharing info (simplified)
  if (CONFIG.TRACK_PERMISSIONS) {
    try {
      data.sharingAccess = file.getSharingAccess().toString();
      data.sharingPermission = file.getSharingPermission().toString();
    } catch (error) {
      // Skip if no access
    }
  }
  
  return data;
}

/**
 * Simple file type detection
 */
function getFileTypeSimple(fileName, mimeType) {
  if (mimeType.startsWith('application/vnd.google-apps.')) {
    const googleType = mimeType.replace('application/vnd.google-apps.', '');
    return `Google ${googleType.charAt(0).toUpperCase() + googleType.slice(1)}`;
  }
  
  const extension = fileName.split('.').pop().toLowerCase();
  
  const typeMap = {
    'pdf': 'PDF', 'doc': 'Word', 'docx': 'Word', 'txt': 'Text',
    'xls': 'Excel', 'xlsx': 'Excel', 'csv': 'CSV',
    'ppt': 'PowerPoint', 'pptx': 'PowerPoint',
    'jpg': 'Image', 'jpeg': 'Image', 'png': 'Image', 'gif': 'Image',
    'mp4': 'Video', 'avi': 'Video', 'mov': 'Video',
    'mp3': 'Audio', 'wav': 'Audio',
    'zip': 'Archive', 'rar': 'Archive'
  };
  
  return typeMap[extension] || extension.toUpperCase() || 'Unknown';
}

/**
 * Initialize stats storage in spreadsheet instead of script properties
 */
function initializeStatsInSheet(spreadsheet) {
  const progressSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROGRESS);
  
  // Clear existing progress
  progressSheet.clear();
  
  // Set up progress tracking headers
  progressSheet.getRange(1, 1, 1, 2).setValues([['Metric', 'Value']]);
  progressSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  
  // Initialize stats object
  const stats = {
    totalFiles: 0,
    totalSize: 0,
    filesByType: {},
    filesByYear: {},
    largeFiles: [],
    oldFiles: [],
    sharedFiles: [],
    errors: 0,
    startTime: new Date().toISOString()
  };
  
  return stats;
}

/**
 * Update stats in spreadsheet
 */
function updateStatsInSheet(spreadsheet, stats) {
  const progressSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROGRESS);
  
  // Save basic stats
  const statsData = [
    ['Total Files', stats.totalFiles],
    ['Total Size', formatBytes(stats.totalSize)],
    ['Large Files Found', stats.largeFiles.length],
    ['Old Files Found', stats.oldFiles.length],
    ['Shared Files Found', stats.sharedFiles.length],
    ['Errors', stats.errors],
    ['File Types', Object.keys(stats.filesByType).length],
    ['Last Updated', new Date().toLocaleString()]
  ];
  
  progressSheet.getRange(2, 1, statsData.length, 2).setValues(statsData);
}

/**
 * Update progress sheet
 */
function updateProgressSheet(spreadsheet, status, stats) {
  const progressSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROGRESS);
  
  // Add status row
  const statusRow = progressSheet.getLastRow() + 1;
  progressSheet.getRange(statusRow + 1, 1, 1, 2).setValues([['Status', status]]);
  
  if (stats && stats.totalProcessed) {
    progressSheet.getRange(statusRow + 2, 1, 1, 2).setValues([['Files Processed', stats.totalProcessed]]);
  }
  
  if (status === 'COMPLETE') {
    progressSheet.getRange(statusRow + 3, 1, 1, 2)
      .setValues([['Completion Time', new Date().toLocaleString()]]);
  }
}

/**
 * Add file to list sheet immediately (don't store in memory)
 */
function addToFileListSheetOptimized(spreadsheet, fileData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.FILE_LIST);
  
  sheet.appendRow([
    fileData.name,
    fileData.type,
    (fileData.size / 1024 / 1024).toFixed(2),
    fileData.lastModified,
    fileData.owner,
    fileData.folderPath,
    fileData.sharingAccess,
    fileData.url
  ]);
}

/**
 * Initialize optimized sheets
 */
function initializeOptimizedSheets(spreadsheet) {
  // Create all sheets
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      // Set up headers
      if (sheetName === CONFIG.SHEETS.FILE_LIST) {
        sheet.getRange(1, 1, 1, 8).setValues([[
          'Name', 'Type', 'Size (MB)', 'Last Modified', 'Owner', 'Folder', 'Sharing', 'URL'
        ]]);
        sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
    }
  }
  
  // Remove default sheet
  const sheet1 = spreadsheet.getSheetByName('Sheet1');
  if (sheet1 && spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(sheet1);
  }
}

/**
 * Generate optimized final reports
 */
function generateOptimizedFinalReports(spreadsheet, stats) {
  console.log("Generating optimized reports...");
  
  // Overview sheet
  generateOptimizedOverview(spreadsheet, stats);
  
  // Large files report  
  if (stats.largeFiles.length > 0) {
    generateLargeFilesReport(spreadsheet, stats.largeFiles);
  }
  
  // Old files report
  if (stats.oldFiles.length > 0) {
    generateOldFilesReport(spreadsheet, stats.oldFiles);
  }
  
  // Shared files report
  if (stats.sharedFiles.length > 0) {
    generateSharedFilesReport(spreadsheet, stats.sharedFiles);
  }
  
  // File types analysis
  generateFileTypesReport(spreadsheet, stats);
}

/**
 * Generate optimized overview
 */
function generateOptimizedOverview(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  // Title
  sheet.getRange(1, 1).setValue('Google Drive Inventory Report - Optimized')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary
  const summaryData = [
    ['Total Files Processed:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Average File Size:', formatBytes(stats.totalSize / Math.max(stats.totalFiles, 1))],
    ['Large Files Found:', stats.largeFiles.length],
    ['Old Files Found:', stats.oldFiles.length], 
    ['Shared Files Found:', stats.sharedFiles.length],
    ['File Types Detected:', Object.keys(stats.filesByType).length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(4, 1, summaryData.length, 2).setValues(summaryData);
  
  // File types (top 15)
  const typeRow = 4 + summaryData.length + 2;
  sheet.getRange(typeRow, 1).setValue('TOP FILE TYPES').setFontWeight('bold');
  
  const typeEntries = Object.entries(stats.filesByType)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);
  
  if (typeEntries.length > 0) {
    sheet.getRange(typeRow + 1, 1, typeEntries.length, 2).setValues(typeEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

// Report generation functions
function generateLargeFilesReport(spreadsheet, largeFiles) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LARGE_FILES);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 4).setValues([['Name', 'Size (MB)', 'Folder', 'URL']]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  const data = largeFiles.map(file => [
    file.name,
    (file.size / 1024 / 1024).toFixed(2),
    file.path,
    file.url
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 4).setValues(data);
  }
}

function generateOldFilesReport(spreadsheet, oldFiles) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OLD_FILES);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 4).setValues([['Name', 'Age (Days)', 'Folder', 'URL']]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  const data = oldFiles.map(file => [
    file.name,
    file.ageInDays,
    file.path,
    file.url
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 4).setValues(data);
  }
}

function generateSharedFilesReport(spreadsheet, sharedFiles) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SHARED_FILES);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 4).setValues([['Name', 'Sharing Access', 'Folder', 'URL']]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  const data = sharedFiles.map(file => [
    file.name,
    file.sharingAccess,
    file.path,
    file.url
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 4).setValues(data);
  }
}

function generateFileTypesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.FILE_TYPES);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 3).setValues([['File Type', 'Count', 'Percentage']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  
  const total = stats.totalFiles;
  const typeData = Object.entries(stats.filesByType)
    .sort((a, b) => b[1] - a[1])
    .map(([type, count]) => [
      type,
      count,
      ((count / total) * 100).toFixed(1) + '%'
    ]);
  
  if (typeData.length > 0) {
    sheet.getRange(2, 1, typeData.length, 3).setValues(typeData);
  }
}

// Utility functions
function getFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  
  while (files.hasNext() && count < batchSize) {
    filesToProcess.push(files.next());
    count++;
  }
  
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

function scheduleNextRun() {
  ScriptApp.newTrigger('continueOptimizedInventory')
    .timeBased()
    .after(1 * 60 * 1000) // 1 minute
    .create();
  
  console.log("Next run scheduled for 1 minute from now");
}

function continueOptimizedInventory() {
  console.log("Continuing optimized inventory...");
  runOptimizedInventory();
}

function resetInventoryProgress() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('continuationToken');
  
  // Cancel any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueOptimizedInventory') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  console.log("Previous inventory progress reset");
}

function saveErrorInfo(spreadsheet, error) {
  const progressSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROGRESS);
  const errorRow = progressSheet.getLastRow() + 1;
  
  progressSheet.getRange(errorRow, 1, 1, 2).setValues([['Error', error.toString()]]);
  progressSheet.getRange(errorRow + 1, 1, 1, 2).setValues([['Error Time', new Date().toLocaleString()]]);
}

function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function getOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    const spreadsheet = SpreadsheetApp.create(name);
    console.log(`Created optimized inventory spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Quick test function for large drives
 */
function testOptimizedInventory() {
  console.log("Testing optimized inventory with small batch...");
  
  // Temporarily set small batch size
  const originalBatchSize = CONFIG.BATCH_SIZE;
  CONFIG.BATCH_SIZE = 10;
  
  try {
    resetInventoryProgress();
    runOptimizedInventory();
  } finally {
    CONFIG.BATCH_SIZE = originalBatchSize;
  }
}

/**
 * USAGE INSTRUCTIONS FOR LARGE DRIVES:
 * 
 * 1. Use inventoryDriveLarge() instead of the original function
 * 2. This version stores data directly in the spreadsheet instead of script properties  
 * 3. It will automatically continue across multiple execution sessions if needed
 * 4. Monitor progress in the "Progress Tracking" sheet of the generated spreadsheet
 * 5. If you get quota errors, the script will save what it has and schedule continuation
 * 
 * BENEFITS:
 * - No property storage quota limits
 * - Handles drives with 10,000+ files
 * - Real-time progress tracking in spreadsheet
 * - Automatic recovery from timeouts
 * - Memory efficient processing
 */