/**
 * Google Drive Large Files Finder Script
 * Specialized script for finding and analyzing large files in Google Drive
 * Focuses on storage optimization and cleanup opportunities
 */

// Configuration for large files analysis
const CONFIG = {
  BATCH_SIZE: 50, // Smaller batch size for large file processing
  INVENTORY_SPREADSHEET_NAME: "📦 Large Files Analysis Report",
  
  SHEETS: {
    OVERVIEW: "Overview",
    LARGE_FILES: "Large Files",
    BY_SIZE_CATEGORY: "By Size Category",
    BY_FILE_TYPE: "By File Type",
    BY_FOLDER: "By Folder",
    CLEANUP_CANDIDATES: "Cleanup Candidates",
    RECOMMENDATIONS: "Recommendations"
  },
  
  // Size thresholds in MB
  SIZE_THRESHOLDS: {
    HUGE: 500,      // 500MB+
    VERY_LARGE: 100, // 100-500MB
    LARGE: 50,      // 50-100MB
    MEDIUM: 10      // 10-50MB (minimum for inclusion)
  },
  
  OLD_FILE_THRESHOLD_DAYS: 365,
  INCLUDE_TRASHED: false,
  TRACK_PERMISSIONS: true,
  
  // Analysis settings
  TRACK_ACCESS_PATTERNS: true,
  IDENTIFY_DUPLICATES: true,
  SUGGEST_CLEANUP: true,
  MAX_FILES_PER_CATEGORY: 200
};

/**
 * Main function to find large files
 */
function findLargeFiles() {
  console.log("Starting large files analysis...");
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const isAutoMode = scriptProperties.getProperty('largeFilesAutoMode') === 'true';
  
  if (!isAutoMode) {
    runLargeFilesAnalysisBatch();
  } else {
    runLargeFilesAnalysisContinuously();
  }
}

/**
 * Run large files analysis continuously
 */
function runLargeFilesAnalysisContinuously() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 5 * 60 * 1000;
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let batchCount = 0;
  let totalProcessed = 0;
  let hasMoreFiles = true;
  
  console.log("Running large files analysis in continuous mode...");
  
  while (hasMoreFiles) {
    const currentTime = new Date().getTime();
    const elapsedTime = currentTime - startTime;
    
    if (elapsedTime > MAX_RUNTIME_MS) {
      console.log(`Approaching time limit after ${(elapsedTime/1000).toFixed(0)} seconds`);
      scheduleNextLargeFilesRun();
      console.log(`Processed ${totalProcessed} large files in ${batchCount} batches. Scheduled next run.`);
      return;
    }
    
    const result = runLargeFilesAnalysisBatch();
    batchCount++;
    totalProcessed += result.processedCount;
    
    if (!result.hasMore) {
      hasMoreFiles = false;
      console.log(`Large files analysis complete! Processed ${totalProcessed} files in ${batchCount} batches.`);
      cancelLargeFilesScheduledRuns();
    } else {
      Utilities.sleep(200); // Longer pause for large file processing
    }
  }
}

/**
 * Run a single batch of large files analysis
 */
function runLargeFilesAnalysisBatch() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeLargeFilesSheets(spreadsheet);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('largeFilesContinuationToken');
  
  let stats = JSON.parse(scriptProperties.getProperty('largeFilesStats') || '{}');
  stats = initializeLargeFilesStats(stats);
  
  updateLargeFilesStatus(spreadsheet, 'RUNNING', stats);
  
  const files = getLargeFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more large files to process!");
    generateLargeFilesFinalReports(spreadsheet, stats);
    updateLargeFilesStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('largeFilesContinuationToken');
    scriptProperties.deleteProperty('largeFilesStats');
    scriptProperties.deleteProperty('largeFilesAutoMode');
    
    cancelLargeFilesScheduledRuns();
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing batch of ${files.length} large files...`);
  
  let processedCount = 0;
  for (const file of files) {
    try {
      processLargeFile(file, spreadsheet, stats);
      processedCount++;
    } catch (error) {
      console.error(`Error processing large file: ${error}`);
      stats.errors++;
    }
  }
  
  scriptProperties.setProperty('largeFilesStats', JSON.stringify(stats));
  updateLargeFilesOverviewSheet(spreadsheet, stats, false);
  
  console.log(`Processed ${processedCount} large files. Total so far: ${stats.totalFiles}`);
  
  const hasMore = files.hasNext && files.hasNext();
  
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('largeFilesContinuationToken', nextToken);
  } else {
    generateLargeFilesFinalReports(spreadsheet, stats);
    updateLargeFilesStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('largeFilesContinuationToken');
    scriptProperties.deleteProperty('largeFilesStats');
    scriptProperties.deleteProperty('largeFilesAutoMode');
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Get large files to process
 */
function getLargeFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    // Get all files, we'll filter by size during processing
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  let checkedCount = 0;
  const maxCheck = batchSize * 20; // Check more files to find large ones
  
  while (files.hasNext() && count < batchSize && checkedCount < maxCheck) {
    const file = files.next();
    checkedCount++;
    
    if (isLargeFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if file is considered large
 */
function isLargeFile(file) {
  try {
    const sizeInMB = file.getSize() / (1024 * 1024);
    return sizeInMB >= CONFIG.SIZE_THRESHOLDS.MEDIUM;
  } catch (error) {
    console.error(`Error checking file size: ${error}`);
    return false;
  }
}

/**
 * Process a single large file
 */
function processLargeFile(file, spreadsheet, stats) {
  try {
    const fileData = extractLargeFileData(file);
    
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Categorize by size
    const sizeCategory = getSizeCategory(fileData.size);
    stats.bySizeCategory[sizeCategory] = (stats.bySizeCategory[sizeCategory] || 0) + 1;
    
    // Track by file type
    const fileType = fileData.type;
    stats.byFileType[fileType] = (stats.byFileType[fileType] || 0) + 1;
    
    // Track by folder
    if (fileData.folderPath) {
      if (!stats.byFolder[fileData.folderPath]) {
        stats.byFolder[fileData.folderPath] = { count: 0, totalSize: 0 };
      }
      stats.byFolder[fileData.folderPath].count++;
      stats.byFolder[fileData.folderPath].totalSize += fileData.size;
    }
    
    // Track by owner
    stats.byOwner[fileData.owner] = (stats.byOwner[fileData.owner] || 0) + 1;
    
    // Add to master list (keep sorted by size)
    stats.largeFiles.push({
      name: fileData.name,
      size: fileData.size,
      type: fileData.type,
      sizeCategory: sizeCategory,
      created: fileData.created,
      lastModified: fileData.lastModified,
      owner: fileData.owner,
      path: fileData.folderPath,
      sharingAccess: fileData.sharingAccess,
      url: fileData.url,
      ageInDays: fileData.ageInDays,
      cleanupScore: calculateCleanupScore(fileData)
    });
    
    // Keep only top files by size
    stats.largeFiles.sort((a, b) => b.size - a.size);
    stats.largeFiles = stats.largeFiles.slice(0, CONFIG.MAX_FILES_PER_CATEGORY);
    
    // Track cleanup candidates
    const cleanupScore = calculateCleanupScore(fileData);
    if (cleanupScore >= 70) { // High cleanup score
      stats.cleanupCandidates.push({
        name: fileData.name,
        size: fileData.size,
        type: fileData.type,
        path: fileData.folderPath,
        url: fileData.url,
        cleanupScore: cleanupScore,
        reasons: getCleanupReasons(fileData)
      });
      
      // Keep only top cleanup candidates
      stats.cleanupCandidates.sort((a, b) => b.cleanupScore - a.cleanupScore);
      stats.cleanupCandidates = stats.cleanupCandidates.slice(0, 100);
    }
    
    // Track potential duplicates
    const duplicateKey = `${fileData.name}_${fileData.size}`;
    if (!stats.duplicateCandidates[duplicateKey]) {
      stats.duplicateCandidates[duplicateKey] = [];
    }
    stats.duplicateCandidates[duplicateKey].push({
      name: fileData.name,
      path: fileData.folderPath,
      size: fileData.size,
      type: fileData.type,
      lastModified: fileData.lastModified,
      url: fileData.url
    });
    
    addToLargeFilesSheet(spreadsheet, fileData);
    
  } catch (error) {
    console.error(`Error processing large file: ${error}`);
    throw error;
  }
}

/**
 * Extract large file specific data
 */
function extractLargeFileData(file) {
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
    ageInDays: 0
  };
  
  // Calculate age
  data.ageInDays = Math.floor((new Date() - new Date(data.lastModified)) / (1000 * 60 * 60 * 24));
  
  // Get folder path
  try {
    const parents = file.getParents();
    const pathParts = [];
    
    if (parents.hasNext()) {
      const parent = parents.next();
      pathParts.push(parent.getName());
      
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
    } catch (error) {
      // Some files may not have accessible permissions
    }
  }
  
  return data;
}

/**
 * Get size category for a file
 */
function getSizeCategory(sizeInBytes) {
  const sizeInMB = sizeInBytes / (1024 * 1024);
  
  if (sizeInMB >= CONFIG.SIZE_THRESHOLDS.HUGE) return 'Huge (500MB+)';
  if (sizeInMB >= CONFIG.SIZE_THRESHOLDS.VERY_LARGE) return 'Very Large (100-500MB)';
  if (sizeInMB >= CONFIG.SIZE_THRESHOLDS.LARGE) return 'Large (50-100MB)';
  if (sizeInMB >= CONFIG.SIZE_THRESHOLDS.MEDIUM) return 'Medium (10-50MB)';
  return 'Other';
}

/**
 * Calculate cleanup score (0-100) for a file
 */
function calculateCleanupScore(fileData) {
  let score = 0;
  
  // Age factor (0-40 points)
  if (fileData.ageInDays > 730) score += 40; // 2+ years
  else if (fileData.ageInDays > 365) score += 30; // 1+ year
  else if (fileData.ageInDays > 180) score += 20; // 6+ months
  else if (fileData.ageInDays > 90) score += 10; // 3+ months
  
  // Size factor (0-30 points)
  const sizeInMB = fileData.size / (1024 * 1024);
  if (sizeInMB > 1000) score += 30; // 1GB+
  else if (sizeInMB > 500) score += 25; // 500MB+
  else if (sizeInMB > 100) score += 20; // 100MB+
  else if (sizeInMB > 50) score += 15; // 50MB+
  else score += 10; // 10MB+
  
  // File type factor (0-20 points)
  const suspiciousTypes = ['zip', 'rar', 'tar', 'backup', 'tmp', 'log'];
  const fileType = fileData.type.toLowerCase();
  if (suspiciousTypes.some(type => fileType.includes(type))) {
    score += 20;
  } else if (fileType.includes('video') || fileType.includes('audio')) {
    score += 10; // Media files are often large but may be needed
  }
  
  // Sharing factor (0-10 points)
  if (fileData.sharingAccess === 'Private') {
    score += 10; // Private files are safer to clean up
  }
  
  return Math.min(score, 100);
}

/**
 * Get cleanup reasons for a file
 */
function getCleanupReasons(fileData) {
  const reasons = [];
  
  if (fileData.ageInDays > 365) {
    reasons.push(`Not modified in ${Math.floor(fileData.ageInDays / 365)} year(s)`);
  }
  
  const sizeInMB = fileData.size / (1024 * 1024);
  if (sizeInMB > 500) {
    reasons.push(`Very large file (${sizeInMB.toFixed(0)}MB)`);
  }
  
  const suspiciousTypes = ['zip', 'rar', 'backup', 'tmp', 'log'];
  if (suspiciousTypes.some(type => fileData.type.toLowerCase().includes(type))) {
    reasons.push('Potentially temporary or archive file');
  }
  
  if (fileData.sharingAccess === 'Private') {
    reasons.push('Private file (safer to remove)');
  }
  
  return reasons;
}

/**
 * Initialize large files statistics
 */
function initializeLargeFilesStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    bySizeCategory: stats.bySizeCategory || {},
    byFileType: stats.byFileType || {},
    byFolder: stats.byFolder || {},
    byOwner: stats.byOwner || {},
    largeFiles: stats.largeFiles || [],
    cleanupCandidates: stats.cleanupCandidates || [],
    duplicateCandidates: stats.duplicateCandidates || {},
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Initialize large files specific sheets
 */
function initializeLargeFilesSheets(spreadsheet) {
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      switch (sheetName) {
        case CONFIG.SHEETS.LARGE_FILES:
          sheet.getRange(1, 1, 1, 10).setValues([[
            'Name', 'Size (MB)', 'Type', 'Category', 'Age (Days)', 
            'Owner', 'Folder Path', 'Sharing', 'Cleanup Score', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.BY_SIZE_CATEGORY:
          sheet.getRange(1, 1, 1, 3).setValues([[
            'Size Category', 'File Count', 'Total Size (GB)'
          ]]);
          sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.CLEANUP_CANDIDATES:
          sheet.getRange(1, 1, 1, 7).setValues([[
            'Name', 'Size (MB)', 'Type', 'Cleanup Score', 'Reasons', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.BY_FOLDER:
          sheet.getRange(1, 1, 1, 4).setValues([[
            'Folder Path', 'File Count', 'Total Size (GB)', 'Avg File Size (MB)'
          ]]);
          sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
          break;
      }
    }
  }
  
  const sheet1 = spreadsheet.getSheetByName('Sheet1');
  if (sheet1 && spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(sheet1);
  }
}

/**
 * Add large file data to the sheet
 */
function addToLargeFilesSheet(spreadsheet, fileData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LARGE_FILES);
  
  const cleanupScore = calculateCleanupScore(fileData);
  
  sheet.appendRow([
    fileData.name,
    (fileData.size / 1024 / 1024).toFixed(2),
    fileData.type,
    getSizeCategory(fileData.size),
    fileData.ageInDays,
    fileData.owner,
    fileData.folderPath,
    fileData.sharingAccess,
    cleanupScore,
    fileData.url
  ]);
}

/**
 * Update the large files overview sheet
 */
function updateLargeFilesOverviewSheet(spreadsheet, stats, isFinal) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('Google Drive Large Files Analysis')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(isFinal ? 'Final Report' : 'Progress Report')
    .setFontSize(12);
  
  sheet.getRange(3, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary Statistics
  sheet.getRange(5, 1).setValue('LARGE FILES SUMMARY').setFontWeight('bold');
  
  const summaryData = [
    ['Total Large Files:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Average File Size:', formatBytes(stats.totalSize / Math.max(stats.totalFiles, 1))],
    ['Cleanup Candidates:', stats.cleanupCandidates.length],
    ['Potential Space Savings:', formatBytes(stats.cleanupCandidates.slice(0, 50).reduce((sum, file) => sum + file.size, 0))],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(6, 1, summaryData.length, 2).setValues(summaryData);
  
  // Size Categories Distribution
  const categoryRow = 14;
  sheet.getRange(categoryRow, 1).setValue('SIZE CATEGORIES DISTRIBUTION').setFontWeight('bold');
  
  const categoryEntries = Object.entries(stats.bySizeCategory)
    .sort((a, b) => b[1] - a[1]);
  
  if (categoryEntries.length > 0) {
    sheet.getRange(categoryRow + 1, 1, categoryEntries.length, 2).setValues(categoryEntries);
  }
  
  // File Types Distribution  
  const typeRow = categoryRow + categoryEntries.length + 3;
  sheet.getRange(typeRow, 1).setValue('FILE TYPES DISTRIBUTION').setFontWeight('bold');
  
  const typeEntries = Object.entries(stats.byFileType)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (typeEntries.length > 0) {
    sheet.getRange(typeRow + 1, 1, typeEntries.length, 2).setValues(typeEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final large files reports
 */
function generateLargeFilesFinalReports(spreadsheet, stats) {
  console.log("Generating final large files reports...");
  
  updateLargeFilesOverviewSheet(spreadsheet, stats, true);
  generateSizeCategoryReport(spreadsheet, stats);
  generateFileTypeReport(spreadsheet, stats);
  generateFolderAnalysisReport(spreadsheet, stats);
  generateCleanupCandidatesReport(spreadsheet, stats);
  generateRecommendationsReport(spreadsheet, stats);
  
  console.log(`Large files reports generated! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate size category report
 */
function generateSizeCategoryReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_SIZE_CATEGORY);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  const categoryData = Object.entries(stats.bySizeCategory)
    .map(([category, count]) => {
      const totalSizeGB = (stats.totalSize / count / 1024 / 1024 / 1024).toFixed(2);
      return [category, count, totalSizeGB];
    })
    .sort((a, b) => b[1] - a[1]);
  
  if (categoryData.length > 0) {
    sheet.getRange(2, 1, categoryData.length, 3).setValues(categoryData);
  }
}

/**
 * Generate file type report
 */
function generateFileTypeReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_FILE_TYPE);
  
  sheet.clear();
  sheet.getRange(1, 1, 1, 4).setValues([[
    'File Type', 'Count', 'Total Size (GB)', 'Avg Size (MB)'
  ]]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  const typeData = Object.entries(stats.byFileType)
    .map(([type, count]) => {
      const avgSizeBytes = stats.totalSize / count;
      const totalSizeGB = (avgSizeBytes * count / 1024 / 1024 / 1024).toFixed(2);
      const avgSizeMB = (avgSizeBytes / 1024 / 1024).toFixed(2);
      return [type, count, totalSizeGB, avgSizeMB];
    })
    .sort((a, b) => b[1] - a[1]);
  
  if (typeData.length > 0) {
    sheet.getRange(2, 1, typeData.length, 4).setValues(typeData);
  }
}

/**
 * Generate folder analysis report
 */
function generateFolderAnalysisReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_FOLDER);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  const folderData = Object.entries(stats.byFolder)
    .map(([folder, data]) => [
      folder,
      data.count,
      (data.totalSize / 1024 / 1024 / 1024).toFixed(2),
      (data.totalSize / data.count / 1024 / 1024).toFixed(2)
    ])
    .sort((a, b) => b[2] - a[2]) // Sort by total size
    .slice(0, 50);
  
  if (folderData.length > 0) {
    sheet.getRange(2, 1, folderData.length, 4).setValues(folderData);
  }
}

/**
 * Generate cleanup candidates report
 */
function generateCleanupCandidatesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CLEANUP_CANDIDATES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.cleanupCandidates.length > 0) {
    const data = stats.cleanupCandidates.map(file => [
      file.name,
      (file.size / 1024 / 1024).toFixed(2),
      file.type,
      file.cleanupScore,
      file.reasons.join('; '),
      file.path,
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 7).setValues(data);
  }
}

/**
 * Generate recommendations report
 */
function generateRecommendationsReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.RECOMMENDATIONS);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('STORAGE OPTIMIZATION RECOMMENDATIONS')
    .setFontSize(14).setFontWeight('bold');
  
  let currentRow = 3;
  
  // Top recommendations
  const recommendations = generateStorageRecommendations(stats);
  
  recommendations.forEach(rec => {
    sheet.getRange(currentRow, 1).setValue(rec.title).setFontWeight('bold');
    sheet.getRange(currentRow + 1, 1).setValue(rec.description);
    
    if (rec.potentialSavings) {
      sheet.getRange(currentRow + 2, 1).setValue(`Potential Savings: ${rec.potentialSavings}`);
    }
    
    currentRow += 4;
  });
  
  sheet.autoResizeColumns(1, 1);
}

/**
 * Generate storage recommendations
 */
function generateStorageRecommendations(stats) {
  const recommendations = [];
  
  // Cleanup candidates recommendation
  if (stats.cleanupCandidates.length > 0) {
    const totalSavings = stats.cleanupCandidates.slice(0, 20).reduce((sum, file) => sum + file.size, 0);
    recommendations.push({
      title: '🧹 Clean Up Old Large Files',
      description: `You have ${stats.cleanupCandidates.length} files that are good candidates for cleanup. Focus on files that haven't been accessed recently and are taking up significant space.`,
      potentialSavings: formatBytes(totalSavings)
    });
  }
  
  // Duplicate files recommendation
  const duplicateGroups = Object.entries(stats.duplicateCandidates)
    .filter(([key, files]) => files.length > 1);
  
  if (duplicateGroups.length > 0) {
    const duplicateSavings = duplicateGroups.reduce((sum, [key, files]) => {
      return sum + (files.slice(1).reduce((fileSum, file) => fileSum + file.size, 0));
    }, 0);
    
    recommendations.push({
      title: '🔄 Remove Duplicate Files',
      description: `Found ${duplicateGroups.length} groups of potential duplicate files. Removing duplicates can free up space while keeping the original files.`,
      potentialSavings: formatBytes(duplicateSavings)
    });
  }
  
  // Large folders recommendation
  const largeFolders = Object.entries(stats.byFolder)
    .sort((a, b) => b[1].totalSize - a[1].totalSize)
    .slice(0, 5);
  
  if (largeFolders.length > 0) {
    recommendations.push({
      title: '📁 Review Large Folders',
      description: `Your largest folders contain significant amounts of data. Consider archiving or organizing files in: ${largeFolders.map(([folder]) => folder).join(', ')}`,
      potentialSavings: null
    });
  }
  
  // File type recommendations
  const mediaFiles = Object.entries(stats.byFileType)
    .filter(([type]) => type.toLowerCase().includes('video') || type.toLowerCase().includes('audio'))
    .reduce((sum, [type, count]) => sum + count, 0);
  
  if (mediaFiles > 10) {
    recommendations.push({
      title: '🎥 Optimize Media Files',
      description: `You have ${mediaFiles} media files. Consider compressing videos or moving them to specialized storage solutions like Google Photos.`,
      potentialSavings: null
    });
  }
  
  return recommendations;
}

// Utility functions

function scheduleNextLargeFilesRun() {
  cancelLargeFilesScheduledRuns();
  
  ScriptApp.newTrigger('continueLargeFilesAnalysis')
    .timeBased()
    .after(1 * 60 * 1000)
    .create();
  
  console.log("Next large files analysis scheduled for 1 minute from now");
}

function continueLargeFilesAnalysis() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (scriptProperties.getProperty('largeFilesAutoMode') !== 'true') {
    console.log("Large files auto mode disabled, stopping.");
    cancelLargeFilesScheduledRuns();
    return;
  }
  
  findLargeFiles();
}

function cancelLargeFilesScheduledRuns() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueLargeFilesAnalysis' ||
        trigger.getHandlerFunction() === 'findLargeFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log("Cancelled all scheduled large files analysis runs");
}

function updateLargeFilesStatus(spreadsheet, status, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  
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
}

function getFileType(fileName, mimeType) {
  // Similar to the main script's getFileType function
  if (mimeType.startsWith('application/vnd.google-apps.')) {
    const googleType = mimeType.replace('application/vnd.google-apps.', '');
    return `Google ${googleType.charAt(0).toUpperCase() + googleType.slice(1)}`;
  }
  
  const extension = fileName.split('.').pop().toLowerCase();
  
  const typeMap = {
    'doc': 'Word Document', 'docx': 'Word Document',
    'pdf': 'PDF', 'txt': 'Text File',
    'xls': 'Excel', 'xlsx': 'Excel', 'csv': 'CSV',
    'ppt': 'PowerPoint', 'pptx': 'PowerPoint',
    'jpg': 'Image', 'jpeg': 'Image', 'png': 'Image', 'gif': 'Image',
    'mp4': 'Video', 'avi': 'Video', 'mov': 'Video', 'wmv': 'Video',
    'mp3': 'Audio', 'wav': 'Audio', 'flac': 'Audio',
    'zip': 'Archive', 'rar': 'Archive', '7z': 'Archive',
    'js': 'Code', 'html': 'Code', 'css': 'Code', 'py': 'Code'
  };
  
  return typeMap[extension] || extension.toUpperCase() || 'Unknown';
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
    console.log(`Created new large files analysis spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Start automatic large files analysis
 */
function startAutomaticLargeFilesAnalysis() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const continuationToken = scriptProperties.getProperty('largeFilesContinuationToken');
  if (!continuationToken) {
    resetLargeFilesAnalysis();
  }
  
  scriptProperties.setProperty('largeFilesAutoMode', 'true');
  
  console.log("Starting automatic large files analysis...");
  findLargeFiles();
}

/**
 * Reset large files analysis
 */
function resetLargeFilesAnalysis() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('largeFilesContinuationToken');
  scriptProperties.deleteProperty('largeFilesStats');
  scriptProperties.deleteProperty('largeFilesAutoMode');
  console.log("Large files analysis reset. Next run will start from the beginning.");
}

/**
 * Quick large files stats
 */
function getQuickLargeFilesStats() {
  console.log("Gathering quick large files statistics...");
  
  const stats = {
    totalLargeFiles: 0,
    totalSize: 0,
    largestFile: null,
    sizeCategories: {}
  };
  
  const files = DriveApp.searchFiles('trashed = false');
  let count = 0;
  const maxCheck = 1000;
  
  while (files.hasNext() && count < maxCheck) {
    const file = files.next();
    
    if (isLargeFile(file)) {
      count++;
      stats.totalLargeFiles++;
      
      const size = file.getSize();
      stats.totalSize += size;
      
      const category = getSizeCategory(size);
      stats.sizeCategories[category] = (stats.sizeCategories[category] || 0) + 1;
      
      if (!stats.largestFile || size > stats.largestFile.size) {
        stats.largestFile = {
          name: file.getName(),
          size: size,
          sizeFormatted: formatBytes(size)
        };
      }
    }
  }
  
  console.log(`Quick Large Files Stats (based on first ${count} large files):`);
  console.log(`Total large files: ${stats.totalLargeFiles}`);
  console.log(`Total size: ${formatBytes(stats.totalSize)}`);
  console.log(`Average size: ${formatBytes(stats.totalSize / stats.totalLargeFiles)}`);
  
  if (stats.largestFile) {
    console.log(`Largest file: ${stats.largestFile.name} (${stats.largestFile.sizeFormatted})`);
  }
  
  console.log('\nSize category distribution:');
  Object.entries(stats.sizeCategories)
    .sort((a, b) => b[1] - a[1])
    .forEach(([category, count]) => {
      console.log(`  ${category}: ${count}`);
    });
  
  return stats;
}