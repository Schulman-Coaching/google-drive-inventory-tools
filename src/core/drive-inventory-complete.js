/**
 * Google Drive Inventory Script - Complete Version
 * Takes a comprehensive inventory of files in Google Drive
 * Processes files incrementally and generates detailed reports
 */

// Configuration
const CONFIG = {
  // Process this many files per run (adjust based on your needs)
  BATCH_SIZE: 100,
  
  // Spreadsheet name for inventory results
  INVENTORY_SPREADSHEET_NAME: "📊 Drive Inventory Report v3",
  
  // Sheet names for different reports
  SHEETS: {
    OVERVIEW: "Overview",
    FILE_LIST: "File List",
    LARGE_FILES: "Large Files",
    OLD_FILES: "Old Files",
    DUPLICATES: "Potential Duplicates",
    SHARED_FILES: "Shared Files",
    FILE_TYPES: "File Types Analysis",
    FOLDER_STRUCTURE: "Folder Structure"
  },
  
  // File size thresholds (in MB)
  LARGE_FILE_THRESHOLD_MB: 50,
  
  // Old file threshold (in days)
  OLD_FILE_THRESHOLD_DAYS: 365,
  
  // Include Google native files (Docs, Sheets, etc.)
  INCLUDE_GOOGLE_FILES: true,
  
  // Include trashed files in inventory
  INCLUDE_TRASHED: false,
  
  // Track file permissions and sharing
  TRACK_PERMISSIONS: true,
  
  // Maximum number of duplicate groups to show
  MAX_DUPLICATE_GROUPS: 100
};

/**
 * Main function to inventory Drive files - runs continuously until complete
 */
function inventoryDrive() {
  console.log("Starting Drive inventory...");
  
  // Check if automatic mode is enabled
  const scriptProperties = PropertiesService.getScriptProperties();
  const isAutoMode = scriptProperties.getProperty('autoMode') === 'true';
  
  if (!isAutoMode) {
    // Run single batch
    runInventoryBatch();
  } else {
    // Run continuously until time limit
    runInventoryContinuously();
  }
}

/**
 * Run inventory continuously until completion or time limit
 */
function runInventoryContinuously() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 5 * 60 * 1000; // 5 minutes (leaving 1 minute buffer for 6-minute limit)
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let batchCount = 0;
  let totalProcessed = 0;
  let hasMoreFiles = true;
  
  console.log("Running in continuous mode...");
  
  while (hasMoreFiles) {
    // Check execution time
    const currentTime = new Date().getTime();
    const elapsedTime = currentTime - startTime;
    
    if (elapsedTime > MAX_RUNTIME_MS) {
      console.log(`Approaching time limit after ${(elapsedTime/1000).toFixed(0)} seconds`);
      
      // Schedule next run
      scheduleNextRun();
      
      console.log(`Processed ${totalProcessed} files in ${batchCount} batches. Scheduled next run.`);
      return;
    }
    
    // Run a batch
    const result = runInventoryBatch();
    batchCount++;
    totalProcessed += result.processedCount;
    
    if (!result.hasMore) {
      hasMoreFiles = false;
      console.log(`Inventory complete! Processed ${totalProcessed} files in ${batchCount} batches.`);
      
      // Cancel any scheduled runs
      cancelScheduledRuns();
    } else {
      // Brief pause between batches to avoid rate limits
      Utilities.sleep(100);
    }
  }
}

/**
 * Run a single batch of inventory processing
 */
function runInventoryBatch() {
  // Get or create the inventory spreadsheet
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  
  // Initialize sheets if needed
  initializeSheets(spreadsheet);
  
  // Get the continuation token from previous run
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('continuationToken');
  
  // Get current stats
  let stats = JSON.parse(scriptProperties.getProperty('inventoryStats') || '{}');
  stats = initializeStats(stats);
  
  // Update status
  updateInventoryStatus(spreadsheet, 'RUNNING', stats);
  
  // Get files to process
  const files = getFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more files to process!");
    
    // Generate final reports
    generateFinalReports(spreadsheet, stats);
    
    // Update status
    updateInventoryStatus(spreadsheet, 'COMPLETE', stats);
    
    // Clear properties
    scriptProperties.deleteProperty('continuationToken');
    scriptProperties.deleteProperty('inventoryStats');
    scriptProperties.deleteProperty('autoMode');
    
    // Cancel any scheduled runs
    cancelScheduledRuns();
    
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing batch of ${files.length} files...`);
  
  // Process files and update stats
  let processedCount = 0;
  for (const file of files) {
    try {
      processFileForInventory(file, spreadsheet, stats);
      processedCount++;
    } catch (error) {
      console.error(`Error processing file: ${error}`);
      stats.errors++;
    }
  }
  
  // Save stats
  scriptProperties.setProperty('inventoryStats', JSON.stringify(stats));
  
  // Update overview sheet with current progress
  updateOverviewSheet(spreadsheet, stats, false);
  
  console.log(`Processed ${processedCount} files. Total so far: ${stats.totalFiles}`);
  
  // Check if there are more files
  const hasMore = files.hasNext && files.hasNext();
  
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('continuationToken', nextToken);
    
    // Estimate remaining files
    const estimatedTotal = estimateRemainingFiles(stats);
    console.log(`Progress: ${stats.totalFiles}/${estimatedTotal} files (estimated)`);
  } else {
    // Final run - generate complete reports
    generateFinalReports(spreadsheet, stats);
    
    // Update status
    updateInventoryStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('continuationToken');
    scriptProperties.deleteProperty('inventoryStats');
    scriptProperties.deleteProperty('autoMode');
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Initialize statistics object
 */
function initializeStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    filesByType: stats.filesByType || {},
    filesByFolder: stats.filesByFolder || {},
    filesByOwner: stats.filesByOwner || {},
    filesByYear: stats.filesByYear || {},
    largeFiles: stats.largeFiles || [],
    oldFiles: stats.oldFiles || [],
    sharedFiles: stats.sharedFiles || [],
    duplicateCandidates: stats.duplicateCandidates || {},
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Get files to process in this batch
 */
function getFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    // Start fresh - get all files
    if (CONFIG.INCLUDE_TRASHED) {
      files = DriveApp.getFiles();
    } else {
      files = DriveApp.searchFiles('trashed = false');
    }
  }
  
  const filesToProcess = [];
  let count = 0;
  
  while (files.hasNext() && count < batchSize) {
    const file = files.next();
    
    // Check if file should be included
    if (shouldIncludeFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  // Attach continuation capability to the array
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if a file should be included in inventory
 */
function shouldIncludeFile(file) {
  try {
    // Check if we should include Google files
    if (!CONFIG.INCLUDE_GOOGLE_FILES) {
      const mimeType = file.getMimeType();
      if (mimeType.startsWith('application/vnd.google-apps.')) {
        return false;
      }
    }
    
    // Check if we should include trashed files
    if (!CONFIG.INCLUDE_TRASHED && file.isTrashed()) {
      return false;
    }
    
    return true;
  } catch (error) {
    console.error(`Error checking file: ${error}`);
    return false;
  }
}

/**
 * Process a single file for inventory
 */
function processFileForInventory(file, spreadsheet, stats) {
  try {
    const fileData = extractFileData(file);
    
    // Update statistics
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Track by type
    const fileType = fileData.type;
    stats.filesByType[fileType] = (stats.filesByType[fileType] || 0) + 1;
    
    // Track by year
    const year = new Date(fileData.lastModified).getFullYear();
    stats.filesByYear[year] = (stats.filesByYear[year] || 0) + 1;
    
    // Track by folder
    if (fileData.folderPath) {
      stats.filesByFolder[fileData.folderPath] = (stats.filesByFolder[fileData.folderPath] || 0) + 1;
    }
    
    // Track by owner
    stats.filesByOwner[fileData.owner] = (stats.filesByOwner[fileData.owner] || 0) + 1;
    
    // Check for large files
    if (fileData.size > CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      stats.largeFiles.push({
        name: fileData.name,
        size: fileData.size,
        path: fileData.folderPath,
        url: fileData.url
      });
      
      // Keep only top 100 largest files
      stats.largeFiles.sort((a, b) => b.size - a.size);
      stats.largeFiles = stats.largeFiles.slice(0, 100);
    }
    
    // Check for old files
    const ageInDays = (new Date() - new Date(fileData.lastModified)) / (1000 * 60 * 60 * 24);
    if (ageInDays > CONFIG.OLD_FILE_THRESHOLD_DAYS) {
      if (stats.oldFiles.length < 100) {
        stats.oldFiles.push({
          name: fileData.name,
          lastModified: fileData.lastModified,
          ageInDays: Math.floor(ageInDays),
          path: fileData.folderPath,
          url: fileData.url
        });
      }
    }
    
    // Track shared files
    if (fileData.sharingAccess !== 'Private') {
      if (stats.sharedFiles.length < 100) {
        stats.sharedFiles.push({
          name: fileData.name,
          sharingAccess: fileData.sharingAccess,
          sharingPermission: fileData.sharingPermission,
          path: fileData.folderPath,
          url: fileData.url
        });
      }
    }
    
    // Track potential duplicates (by name and size)
    const duplicateKey = `${fileData.name}_${fileData.size}`;
    if (!stats.duplicateCandidates[duplicateKey]) {
      stats.duplicateCandidates[duplicateKey] = [];
    }
    stats.duplicateCandidates[duplicateKey].push({
      name: fileData.name,
      path: fileData.folderPath,
      size: fileData.size,
      lastModified: fileData.lastModified,
      url: fileData.url
    });
    
    // Add to file list sheet
    addToFileListSheet(spreadsheet, fileData);
    
  } catch (error) {
    console.error(`Error processing file for inventory: ${error}`);
    throw error;
  }
}

/**
 * Extract comprehensive data from a file
 */
function extractFileData(file) {
  const data = {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    type: getFileType(file.getName(), file.getMimeType()),
    size: file.getSize(),
    created: file.getDateCreated().toISOString(),
    lastModified: file.getLastUpdated().toISOString(),
    owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
    url: file.getUrl(),
    description: file.getDescription() || '',
    folderPath: '',
    sharingAccess: 'Private',
    sharingPermission: 'None',
    viewers: [],
    editors: []
  };
  
  // Get folder path
  try {
    const parents = file.getParents();
    const pathParts = [];
    
    if (parents.hasNext()) {
      const parent = parents.next();
      pathParts.push(parent.getName());
      
      // Get full path (limited depth to avoid timeout)
      let currentParent = parent;
      let depth = 0;
      while (depth < 5) {
        const grandParents = currentParent.getParents();
        if (grandParents.hasNext()) {
          currentParent = grandParents.next();
          pathParts.unshift(currentParent.getName());
          depth++;
        } else {
          break;
        }
      }
    }
    
    data.folderPath = pathParts.length > 0 ? pathParts.join('/') : 'Root';
  } catch (error) {
    data.folderPath = 'Unknown';
  }
  
  // Get sharing information
  if (CONFIG.TRACK_PERMISSIONS) {
    try {
      const access = file.getSharingAccess();
      const permission = file.getSharingPermission();
      
      data.sharingAccess = access.toString();
      data.sharingPermission = permission.toString();
      
      // Get viewers and editors (limited to avoid timeout)
      const viewers = file.getViewers();
      const editors = file.getEditors();
      
      data.viewers = viewers.slice(0, 5).map(user => user.getEmail());
      data.editors = editors.slice(0, 5).map(user => user.getEmail());
      
    } catch (error) {
      // Some files may not have accessible permissions
    }
  }
  
  return data;
}

/**
 * Determine file type from name and MIME type
 */
function getFileType(fileName, mimeType) {
  // Check for Google file types
  if (mimeType.startsWith('application/vnd.google-apps.')) {
    const googleType = mimeType.replace('application/vnd.google-apps.', '');
    return `Google ${googleType.charAt(0).toUpperCase() + googleType.slice(1)}`;
  }
  
  // Check by extension
  const extension = fileName.split('.').pop().toLowerCase();
  
  const typeMap = {
    // Documents
    'doc': 'Word Document', 'docx': 'Word Document',
    'pdf': 'PDF', 'txt': 'Text File', 'rtf': 'Rich Text',
    
    // Spreadsheets
    'xls': 'Excel', 'xlsx': 'Excel', 'csv': 'CSV',
    
    // Presentations
    'ppt': 'PowerPoint', 'pptx': 'PowerPoint',
    
    // Images
    'jpg': 'Image', 'jpeg': 'Image', 'png': 'Image', 
    'gif': 'Image', 'svg': 'Image', 'bmp': 'Image',
    
    // Videos
    'mp4': 'Video', 'avi': 'Video', 'mov': 'Video',
    'wmv': 'Video', 'flv': 'Video', 'mkv': 'Video',
    
    // Audio
    'mp3': 'Audio', 'wav': 'Audio', 'flac': 'Audio',
    
    // Archives
    'zip': 'Archive', 'rar': 'Archive', '7z': 'Archive',
    
    // Code
    'js': 'Code', 'html': 'Code', 'css': 'Code',
    'py': 'Code', 'java': 'Code', 'cpp': 'Code'
  };
  
  return typeMap[extension] || extension.toUpperCase() || 'Unknown';
}

/**
 * Get or create inventory spreadsheet
 */
function getOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    const spreadsheet = SpreadsheetApp.create(name);
    console.log(`Created new inventory spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Initialize sheets in the spreadsheet
 */
function initializeSheets(spreadsheet) {
  // Ensure all required sheets exist
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      // Set up headers based on sheet type
      switch (sheetName) {
        case CONFIG.SHEETS.FILE_LIST:
          sheet.getRange(1, 1, 1, 12).setValues([[
            'Name', 'Type', 'Size (MB)', 'Created', 'Last Modified',
            'Owner', 'Folder Path', 'Sharing', 'Permission', 
            'Viewers', 'Editors', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.LARGE_FILES:
          sheet.getRange(1, 1, 1, 4).setValues([[
            'Name', 'Size (MB)', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.OLD_FILES:
          sheet.getRange(1, 1, 1, 5).setValues([[
            'Name', 'Last Modified', 'Age (Days)', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.DUPLICATES:
          sheet.getRange(1, 1, 1, 5).setValues([[
            'Name', 'Size (MB)', 'Count', 'Locations', 'URLs'
          ]]);
          sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.SHARED_FILES:
          sheet.getRange(1, 1, 1, 5).setValues([[
            'Name', 'Sharing Access', 'Permission', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
          break;
      }
    }
  }
  
  // Remove default "Sheet1" if it exists
  const sheet1 = spreadsheet.getSheetByName('Sheet1');
  if (sheet1 && spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(sheet1);
  }
}

/**
 * Add file data to the file list sheet
 */
function addToFileListSheet(spreadsheet, fileData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.FILE_LIST);
  
  sheet.appendRow([
    fileData.name,
    fileData.type,
    (fileData.size / 1024 / 1024).toFixed(2),
    fileData.created,
    fileData.lastModified,
    fileData.owner,
    fileData.folderPath,
    fileData.sharingAccess,
    fileData.sharingPermission,
    fileData.viewers.join(', '),
    fileData.editors.join(', '),
    fileData.url
  ]);
}

/**
 * Update the overview sheet with current statistics
 */
function updateOverviewSheet(spreadsheet, stats, isFinal) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  // Title
  sheet.getRange(1, 1).setValue('Google Drive Inventory Report')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(isFinal ? 'Final Report' : 'Progress Report')
    .setFontSize(12);
  
  sheet.getRange(3, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary Statistics
  sheet.getRange(5, 1).setValue('SUMMARY STATISTICS').setFontWeight('bold');
  
  const summaryData = [
    ['Total Files:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Average File Size:', formatBytes(stats.totalSize / Math.max(stats.totalFiles, 1))],
    ['Large Files (>' + CONFIG.LARGE_FILE_THRESHOLD_MB + 'MB):', stats.largeFiles.length],
    ['Old Files (>' + CONFIG.OLD_FILE_THRESHOLD_DAYS + ' days):', stats.oldFiles.length],
    ['Shared Files:', stats.sharedFiles.length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(6, 1, summaryData.length, 2).setValues(summaryData);
  
  // File Types Distribution
  const typeRow = 14;
  sheet.getRange(typeRow, 1).setValue('FILE TYPES DISTRIBUTION').setFontWeight('bold');
  
  const typeEntries = Object.entries(stats.filesByType)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);
  
  if (typeEntries.length > 0) {
    sheet.getRange(typeRow + 1, 1, typeEntries.length, 2).setValues(typeEntries);
  }
  
  // Files by Year
  const yearRow = typeRow + typeEntries.length + 3;
  sheet.getRange(yearRow, 1).setValue('FILES BY YEAR').setFontWeight('bold');
  
  const yearEntries = Object.entries(stats.filesByYear)
    .sort((a, b) => b[0] - a[0])
    .slice(0, 10);
  
  if (yearEntries.length > 0) {
    sheet.getRange(yearRow + 1, 1, yearEntries.length, 2).setValues(yearEntries);
  }
  
  // Top Folders
  const folderRow = yearRow + yearEntries.length + 3;
  sheet.getRange(folderRow, 1).setValue('TOP FOLDERS BY FILE COUNT').setFontWeight('bold');
  
  const folderEntries = Object.entries(stats.filesByFolder)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (folderEntries.length > 0) {
    sheet.getRange(folderRow + 1, 1, folderEntries.length, 2).setValues(folderEntries);
  }
  
  // Format columns
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final reports after all files are processed
 */
function generateFinalReports(spreadsheet, stats) {
  console.log("Generating final reports...");
  
  // Update overview as final
  updateOverviewSheet(spreadsheet, stats, true);
  
  // Generate large files report
  generateLargeFilesReport(spreadsheet, stats);
  
  // Generate old files report
  generateOldFilesReport(spreadsheet, stats);
  
  // Generate duplicates report
  generateDuplicatesReport(spreadsheet, stats);
  
  // Generate shared files report
  generateSharedFilesReport(spreadsheet, stats);
  
  // Generate file types analysis
  generateFileTypesAnalysis(spreadsheet, stats);
  
  console.log(`Reports generated! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate large files report
 */
function generateLargeFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LARGE_FILES);
  
  // Clear existing data (except header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Add large files data
  if (stats.largeFiles.length > 0) {
    const data = stats.largeFiles.map(file => [
      file.name,
      (file.size / 1024 / 1024).toFixed(2),
      file.path,
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 4).setValues(data);
  }
}

/**
 * Generate old files report
 */
function generateOldFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OLD_FILES);
  
  // Clear existing data (except header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Add old files data
  if (stats.oldFiles.length > 0) {
    const data = stats.oldFiles.map(file => [
      file.name,
      file.lastModified,
      file.ageInDays,
      file.path,
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 5).setValues(data);
  }
}

/**
 * Generate duplicates report
 */
function generateDuplicatesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DUPLICATES);
  
  // Clear existing data (except header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Find actual duplicates (more than 1 file with same name and size)
  const duplicates = Object.entries(stats.duplicateCandidates)
    .filter(([key, files]) => files.length > 1)
    .sort((a, b) => b[1].length - a[1].length)
    .slice(0, CONFIG.MAX_DUPLICATE_GROUPS);
  
  if (duplicates.length > 0) {
    const data = duplicates.map(([key, files]) => [
      files[0].name,
      (files[0].size / 1024 / 1024).toFixed(2),
      files.length,
      files.map(f => f.path).join('\n'),
      files.map(f => f.url).join('\n')
    ]);
    
    sheet.getRange(2, 1, data.length, 5).setValues(data);
  }
}

/**
 * Generate shared files report
 */
function generateSharedFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SHARED_FILES);
  
  // Clear existing data (except header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Add shared files data
  if (stats.sharedFiles.length > 0) {
    const data = stats.sharedFiles.map(file => [
      file.name,
      file.sharingAccess,
      file.sharingPermission,
      file.path,
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 5).setValues(data);
  }
}

/**
 * Generate file types analysis
 */
function generateFileTypesAnalysis(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.FILE_TYPES);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('FILE TYPES ANALYSIS')
    .setFontSize(14).setFontWeight('bold');
  
  // Headers
  sheet.getRange(3, 1, 1, 3).setValues([['File Type', 'Count', 'Percentage']])
    .setFontWeight('bold');
  
  // Calculate percentages and sort
  const total = stats.totalFiles;
  const typeData = Object.entries(stats.filesByType)
    .sort((a, b) => b[1] - a[1])
    .map(([type, count]) => [
      type,
      count,
      ((count / total) * 100).toFixed(2) + '%'
    ]);
  
  if (typeData.length > 0) {
    sheet.getRange(4, 1, typeData.length, 3).setValues(typeData);
  }
  
  // Format
  sheet.autoResizeColumns(1, 3);
}

/**
 * Format bytes to human readable
 */
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Start automatic inventory - will run continuously until complete
 */
function startAutomaticInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Reset if starting fresh
  const continuationToken = scriptProperties.getProperty('continuationToken');
  if (!continuationToken) {
    resetInventory();
  }
  
  // Enable auto mode
  scriptProperties.setProperty('autoMode', 'true');
  
  console.log("Starting automatic inventory...");
  
  // Run immediately
  inventoryDrive();
}

/**
 * Stop automatic inventory
 */
function stopAutomaticInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('autoMode', 'false');
  
  // Cancel scheduled runs
  cancelScheduledRuns();
  
  // Update status
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  const stats = JSON.parse(scriptProperties.getProperty('inventoryStats') || '{}');
  updateInventoryStatus(spreadsheet, 'PAUSED', stats);
  
  console.log("Automatic inventory stopped. Progress has been saved.");
}

/**
 * Schedule the next automatic run
 */
function scheduleNextRun() {
  // Cancel existing triggers first
  cancelScheduledRuns();
  
  // Create a new trigger for 1 minute from now
  ScriptApp.newTrigger('continueAutomaticInventory')
    .timeBased()
    .after(1 * 60 * 1000) // 1 minute
    .create();
  
  console.log("Next run scheduled for 1 minute from now");
}

/**
 * Continue automatic inventory (called by trigger)
 */
function continueAutomaticInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Check if auto mode is still enabled
  if (scriptProperties.getProperty('autoMode') !== 'true') {
    console.log("Auto mode disabled, stopping.");
    cancelScheduledRuns();
    return;
  }
  
  // Continue inventory
  inventoryDrive();
}

/**
 * Cancel all scheduled runs
 */
function cancelScheduledRuns() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueAutomaticInventory' ||
        trigger.getHandlerFunction() === 'inventoryDrive' ||
        trigger.getHandlerFunction() === 'runScheduledInventory') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log("Cancelled all scheduled inventory runs");
}

/**
 * Update inventory status in the spreadsheet
 */
function updateInventoryStatus(spreadsheet, status, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  
  // Add status indicator
  const statusCell = sheet.getRange(4, 1);
  const statusColor = {
    'RUNNING': '#4CAF50',
    'PAUSED': '#FFC107',
    'COMPLETE': '#2196F3',
    'ERROR': '#F44336'
  };
  
  statusCell.setValue(`Status: ${status}`)
    .setFontWeight('bold')
    .setBackground(statusColor[status] || '#FFFFFF');
  
  // Add progress bar if running
  if (status === 'RUNNING' && stats.totalFiles > 0) {
    const estimated = estimateRemainingFiles(stats);
    const progress = Math.min((stats.totalFiles / estimated) * 100, 100);
    
    sheet.getRange(4, 2).setValue(`Progress: ${progress.toFixed(1)}%`);
  }
}

/**
 * Estimate total files based on current processing rate
 */
function estimateRemainingFiles(stats) {
  // Simple estimation based on folder distribution
  // Assumes similar file density across folders
  const avgFilesPerFolder = stats.totalFiles / Math.max(Object.keys(stats.filesByFolder).length, 1);
  const estimatedFolders = Object.keys(stats.filesByFolder).length * 1.5; // Rough estimate
  
  return Math.max(stats.totalFiles, Math.round(avgFilesPerFolder * estimatedFolders));
}

/**
 * Get inventory status and progress
 */
function getInventoryStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const stats = JSON.parse(scriptProperties.getProperty('inventoryStats') || '{}');
  const continuationToken = scriptProperties.getProperty('continuationToken');
  const autoMode = scriptProperties.getProperty('autoMode') === 'true';
  
  const status = {
    isRunning: !!continuationToken,
    autoMode: autoMode,
    filesProcessed: stats.totalFiles || 0,
    errors: stats.errors || 0,
    startTime: stats.startTime || null,
    currentSize: formatBytes(stats.totalSize || 0)
  };
  
  // Check for active triggers
  const triggers = ScriptApp.getProjectTriggers();
  status.hasScheduledRun = triggers.some(t => 
    t.getHandlerFunction() === 'continueAutomaticInventory' ||
    t.getHandlerFunction() === 'runScheduledInventory'
  );
  
  console.log("Inventory Status:");
  console.log(`- Running: ${status.isRunning}`);
  console.log(`- Auto Mode: ${status.autoMode}`);
  console.log(`- Files Processed: ${status.filesProcessed}`);
  console.log(`- Total Size: ${status.currentSize}`);
  console.log(`- Errors: ${status.errors}`);
  console.log(`- Scheduled Run: ${status.hasScheduledRun}`);
  
  if (status.startTime) {
    const elapsed = new Date() - new Date(status.startTime);
    console.log(`- Running for: ${(elapsed / 1000 / 60).toFixed(1)} minutes`);
  }
  
  return status;
}

/**
 * Run inventory with hourly triggers until complete
 */
function setupHourlyInventory() {
  // Cancel existing triggers
  cancelScheduledRuns();
  
  // Create hourly trigger
  ScriptApp.newTrigger('runScheduledInventory')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log("Hourly inventory scheduled. Will run every hour until complete.");
  
  // Run first batch immediately
  runScheduledInventory();
}

/**
 * Run scheduled inventory (called by hourly trigger)
 */
function runScheduledInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const continuationToken = scriptProperties.getProperty('continuationToken');
  
  if (!continuationToken) {
    // Check if we've already completed
    const stats = JSON.parse(scriptProperties.getProperty('inventoryStats') || '{}');
    if (stats.totalFiles > 0) {
      console.log("Inventory already complete. Cancelling scheduled runs.");
      cancelScheduledRuns();
      return;
    }
  }
  
  console.log("Running scheduled inventory batch...");
  
  // Set auto mode for continuous processing
  scriptProperties.setProperty('autoMode', 'true');
  
  // Run inventory
  inventoryDrive();
}

/**
 * Reset the inventory process
 */
function resetInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('continuationToken');
  scriptProperties.deleteProperty('inventoryStats');
  scriptProperties.deleteProperty('autoMode');
  console.log("Inventory reset. Next run will start from the beginning.");
}

/**
 * Quick stats without full inventory
 */
function getQuickStats() {
  console.log("Gathering quick statistics...");
  
  const stats = {
    totalFiles: 0,
    totalSize: 0,
    largestFile: null,
    oldestFile: null,
    newestFile: null
  };
  
  const files = DriveApp.searchFiles('trashed = false');
  let count = 0;
  const maxCheck = 1000; // Check up to 1000 files for quick stats
  
  while (files.hasNext() && count < maxCheck) {
    const file = files.next();
    count++;
    
    stats.totalFiles++;
    const size = file.getSize();
    stats.totalSize += size;
    
    // Track largest
    if (!stats.largestFile || size > stats.largestFile.size) {
      stats.largestFile = {
        name: file.getName(),
        size: size,
        sizeFormatted: formatBytes(size)
      };
    }
    
    // Track dates
    const lastModified = file.getLastUpdated();
    if (!stats.oldestFile || lastModified < stats.oldestFile.date) {
      stats.oldestFile = {
        name: file.getName(),
        date: lastModified
      };
    }
    
    if (!stats.newestFile || lastModified > stats.newestFile.date) {
      stats.newestFile = {
        name: file.getName(),
        date: lastModified
      };
    }
  }
  
  console.log("Quick Stats (based on first " + count + " files):");
  console.log("Total files checked: " + stats.totalFiles);
  console.log("Total size: " + formatBytes(stats.totalSize));
  console.log("Average size: " + formatBytes(stats.totalSize / stats.totalFiles));
  
  if (stats.largestFile) {
    console.log("Largest file: " + stats.largestFile.name + " (" + stats.largestFile.sizeFormatted + ")");
  }
  
  if (stats.oldestFile) {
    console.log("Oldest file: " + stats.oldestFile.name + " (" + stats.oldestFile.date.toDateString() + ")");
  }
  
  if (stats.newestFile) {
    console.log("Newest file: " + stats.newestFile.name + " (" + stats.newestFile.date.toDateString() + ")");
  }
  
  return stats;
}

/**
 * Find specific types of files
 */
function findFilesByType(fileType) {
  console.log(`Searching for ${fileType} files...`);
  
  const results = [];
  let query = '';
  
  // Build query based on file type
  const typeQueries = {
    'images': "mimeType contains 'image/'",
    'videos': "mimeType contains 'video/'",
    'documents': "mimeType contains 'document' or mimeType contains 'text'",
    'spreadsheets': "mimeType contains 'spreadsheet'",
    'pdfs': "mimeType = 'application/pdf'",
    'large': "", // Will filter by size
    'old': "", // Will filter by date
    'shared': "visibility != 'limited'"
  };
  
  query = typeQueries[fileType.toLowerCase()] || "";
  
  if (query) {
    query += " and trashed = false";
  } else {
    query = "trashed = false";
  }
  
  const files = DriveApp.searchFiles(query);
  let count = 0;
  const maxResults = 50;
  
  while (files.hasNext() && count < maxResults) {
    const file = files.next();
    
    // Additional filters
    if (fileType.toLowerCase() === 'large' && file.getSize() < CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      continue;
    }
    
    if (fileType.toLowerCase() === 'old') {
      const ageInDays = (new Date() - file.getLastUpdated()) / (1000 * 60 * 60 * 24);
      if (ageInDays < CONFIG.OLD_FILE_THRESHOLD_DAYS) {
        continue;
      }
    }
    
    results.push({
      name: file.getName(),
      size: formatBytes(file.getSize()),
      lastModified: file.getLastUpdated().toDateString(),
      url: file.getUrl()
    });
    
    count++;
  }
  
  console.log(`Found ${results.length} ${fileType} files`);
  results.forEach(file => {
    console.log(`- ${file.name} (${file.size}) - ${file.lastModified}`);
  });
  
  return results;
}

/**
 * MENU FUNCTIONS - Add these to Google Sheets menu
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📊 Drive Inventory')
      .addItem('▶️ Start Automatic Scan', 'startAutomaticInventory')
      .addItem('⏸️ Pause Scan', 'stopAutomaticInventory')
      .addItem('📈 Check Status', 'showInventoryStatus')
      .addSeparator()
      .addItem('🔄 Reset Inventory', 'confirmReset')
      .addItem('⏰ Setup Hourly Scans', 'setupHourlyInventory')
      .addItem('🚫 Cancel All Schedules', 'cancelScheduledRuns')
      .addToUi();
  } catch (e) {
    // Not in Sheets context
  }
}

/**
 * Show inventory status in a dialog
 */
function showInventoryStatus() {
  const status = getInventoryStatus();
  
  let message = `
    <b>Inventory Status</b><br><br>
    <b>Status:</b> ${status.isRunning ? '🟢 Running' : '🔴 Stopped'}<br>
    <b>Auto Mode:</b> ${status.autoMode ? 'Enabled' : 'Disabled'}<br>
    <b>Files Processed:</b> ${status.filesProcessed.toLocaleString()}<br>
    <b>Total Size:</b> ${status.currentSize}<br>
    <b>Errors:</b> ${status.errors}<br>
    <b>Scheduled Run:</b> ${status.hasScheduledRun ? 'Yes' : 'No'}
  `;
  
  if (status.startTime) {
    const elapsed = new Date() - new Date(status.startTime);
    message += `<br><b>Running for:</b> ${(elapsed / 1000 / 60).toFixed(1)} minutes`;
  }
  
  // For standalone script execution (not in Sheets)
  console.log("=== Drive Inventory Status ===");
  console.log(status.isRunning ? 'Status: 🟢 Running' : 'Status: 🔴 Stopped');
  console.log(`Files Processed: ${status.filesProcessed}`);
  console.log(`Total Size: ${status.currentSize}`);
  
  // Try to show UI dialog if in Sheets context
  try {
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(message)
      .setWidth(400)
      .setHeight(300);
    ui.showModalDialog(htmlOutput, 'Drive Inventory Status');
  } catch (e) {
    // Not in Sheets context, already logged to console
  }
}

/**
 * Confirm before resetting inventory
 */
function confirmReset() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Reset Inventory?',
      'This will clear all progress and start over. Are you sure?',
      ui.ButtonSet.YES_NO
    );
    
    if (response == ui.Button.YES) {
      resetInventory();
      ui.alert('Inventory has been reset. Click "Start Automatic Scan" to begin again.');
    }
  } catch (e) {
    // Not in Sheets context, just reset with console confirmation
    console.log("Warning: This will reset all inventory progress!");
    console.log("Run resetInventory() to confirm reset.");
  }
}

/**
 * QUICK START FUNCTIONS
 */

/**
 * One-click full automatic inventory
 */
function runCompleteInventory() {
  console.log("Starting complete automatic inventory...");
  console.log("This will run continuously until all files are processed.");
  console.log("You can close this window - the script will continue running.");
  
  startAutomaticInventory();
}

/**
 * Test run with small batch
 */
function testInventory() {
  // Temporarily change config for testing
  const originalBatchSize = CONFIG.BATCH_SIZE;
  CONFIG.BATCH_SIZE = 10;
  
  console.log("Running test inventory (10 files)...");
  
  // Reset to start fresh
  resetInventory();
  
  // Run single batch - inline version for testing
  try {
    // Get or create the inventory spreadsheet
    const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
    
    // Initialize sheets if needed
    initializeSheets(spreadsheet);
    
    // Get the continuation token from previous run (should be null after reset)
    const scriptProperties = PropertiesService.getScriptProperties();
    let continuationToken = scriptProperties.getProperty('continuationToken');
    
    // Get current stats
    let stats = JSON.parse(scriptProperties.getProperty('inventoryStats') || '{}');
    stats = initializeStats(stats);
    
    // Get files to process
    const files = getFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
    
    if (files.length === 0) {
      console.log("No files found to process!");
      CONFIG.BATCH_SIZE = originalBatchSize;
      return;
    }
    
    console.log(`Processing ${files.length} test files...`);
    
    // Process files and update stats
    for (const file of files) {
      try {
        processFileForInventory(file, spreadsheet, stats);
      } catch (error) {
        console.error(`Error processing file: ${error}`);
        stats.errors++;
      }
    }
    
    // Save stats
    scriptProperties.setProperty('inventoryStats', JSON.stringify(stats));
    
    // Update overview sheet with current progress
    updateOverviewSheet(spreadsheet, stats, false);
    
    console.log(`Test complete. Processed ${files.length} files.`);
    console.log(`View results at: ${spreadsheet.getUrl()}`);
    
    // Check if there are more files
    if (files.hasNext && files.hasNext()) {
      const nextToken = files.getContinuationToken();
      scriptProperties.setProperty('continuationToken', nextToken);
      console.log("More files available. Run inventoryDrive() to continue.");
    }
    
  } catch (error) {
    console.error(`Test failed: ${error}`);
    console.error(`Error stack: ${error.stack}`);
  }
  
  // Restore original config
  CONFIG.BATCH_SIZE = originalBatchSize;
}

/**
 * Monitor progress in real-time
 */
function monitorProgress() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const stats = JSON.parse(scriptProperties.getProperty('inventoryStats') || '{}');
  
  if (!stats.totalFiles) {
    console.log("No inventory in progress.");
    return;
  }
  
  console.log("=== INVENTORY PROGRESS ===");
  console.log(`Files Processed: ${stats.totalFiles}`);
  console.log(`Total Size: ${formatBytes(stats.totalSize)}`);
  console.log(`Errors: ${stats.errors}`);
  
  // Show top file types
  if (stats.filesByType) {
    console.log("\nTop File Types:");
    const types = Object.entries(stats.filesByType)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
    
    types.forEach(([type, count]) => {
      console.log(`  ${type}: ${count}`);
    });
  }
  
  // Estimate completion
  const estimatedTotal = estimateRemainingFiles(stats);
  const percentComplete = (stats.totalFiles / estimatedTotal * 100).toFixed(1);
  console.log(`\nEstimated Progress: ${percentComplete}%`);
  
  // Check if still running
  const continuationToken = scriptProperties.getProperty('continuationToken');
  if (continuationToken) {
    console.log("Status: 🟢 RUNNING");
  } else {
    console.log("Status: ✅ COMPLETE");
  }
}

/**
 * Check functions availability (for debugging)
 */
function checkFunctions() {
  console.log("Checking functions...");
  console.log("runInventoryBatch exists: " + (typeof runInventoryBatch !== 'undefined'));
  console.log("inventoryDrive exists: " + (typeof inventoryDrive !== 'undefined'));
  console.log("initializeStats exists: " + (typeof initializeStats !== 'undefined'));
  console.log("getFilesToProcess exists: " + (typeof getFilesToProcess !== 'undefined'));
  console.log("processFileForInventory exists: " + (typeof processFileForInventory !== 'undefined'));
  console.log("All functions defined: Script ready!");
}