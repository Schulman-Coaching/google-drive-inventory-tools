/**
 * Google Drive Image Files Inventory Script
 * Specialized script for analyzing image files in Google Drive
 * Focuses on photos, graphics, and visual content
 */

// Configuration for image analysis
const CONFIG = {
  BATCH_SIZE: 100,
  INVENTORY_SPREADSHEET_NAME: "📸 Image Files Inventory Report",
  
  SHEETS: {
    OVERVIEW: "Overview",
    IMAGE_LIST: "Image Files",
    LARGE_IMAGES: "Large Images",
    OLD_IMAGES: "Old Images", 
    DUPLICATES: "Potential Duplicates",
    BY_FORMAT: "By Format",
    DIMENSIONS: "Image Dimensions"
  },
  
  LARGE_FILE_THRESHOLD_MB: 10, // Lower threshold for images
  OLD_FILE_THRESHOLD_DAYS: 730, // 2 years for images
  INCLUDE_TRASHED: false,
  TRACK_PERMISSIONS: true,
  
  // Image-specific settings
  ANALYZE_DIMENSIONS: true,
  COMMON_IMAGE_FORMATS: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp', 'tiff', 'raw', 'heic'],
  TRACK_PHOTO_METADATA: true
};

/**
 * Main function to inventory image files
 */
function inventoryImages() {
  console.log("Starting image files inventory...");
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const isAutoMode = scriptProperties.getProperty('imageAutoMode') === 'true';
  
  if (!isAutoMode) {
    runImageInventoryBatch();
  } else {
    runImageInventoryContinuously();
  }
}

/**
 * Run image inventory continuously
 */
function runImageInventoryContinuously() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 5 * 60 * 1000;
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let batchCount = 0;
  let totalProcessed = 0;
  let hasMoreFiles = true;
  
  console.log("Running image inventory in continuous mode...");
  
  while (hasMoreFiles) {
    const currentTime = new Date().getTime();
    const elapsedTime = currentTime - startTime;
    
    if (elapsedTime > MAX_RUNTIME_MS) {
      console.log(`Approaching time limit after ${(elapsedTime/1000).toFixed(0)} seconds`);
      scheduleNextImageRun();
      console.log(`Processed ${totalProcessed} images in ${batchCount} batches. Scheduled next run.`);
      return;
    }
    
    const result = runImageInventoryBatch();
    batchCount++;
    totalProcessed += result.processedCount;
    
    if (!result.hasMore) {
      hasMoreFiles = false;
      console.log(`Image inventory complete! Processed ${totalProcessed} images in ${batchCount} batches.`);
      cancelImageScheduledRuns();
    } else {
      Utilities.sleep(100);
    }
  }
}

/**
 * Run a single batch of image inventory processing
 */
function runImageInventoryBatch() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeImageSheets(spreadsheet);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('imageContinuationToken');
  
  let stats = JSON.parse(scriptProperties.getProperty('imageInventoryStats') || '{}');
  stats = initializeImageStats(stats);
  
  updateImageInventoryStatus(spreadsheet, 'RUNNING', stats);
  
  const files = getImageFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more image files to process!");
    generateImageFinalReports(spreadsheet, stats);
    updateImageInventoryStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('imageContinuationToken');
    scriptProperties.deleteProperty('imageInventoryStats');
    scriptProperties.deleteProperty('imageAutoMode');
    
    cancelImageScheduledRuns();
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing batch of ${files.length} image files...`);
  
  let processedCount = 0;
  for (const file of files) {
    try {
      processImageFile(file, spreadsheet, stats);
      processedCount++;
    } catch (error) {
      console.error(`Error processing image file: ${error}`);
      stats.errors++;
    }
  }
  
  scriptProperties.setProperty('imageInventoryStats', JSON.stringify(stats));
  updateImageOverviewSheet(spreadsheet, stats, false);
  
  console.log(`Processed ${processedCount} image files. Total so far: ${stats.totalFiles}`);
  
  const hasMore = files.hasNext && files.hasNext();
  
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('imageContinuationToken', nextToken);
  } else {
    generateImageFinalReports(spreadsheet, stats);
    updateImageInventoryStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('imageContinuationToken');
    scriptProperties.deleteProperty('imageInventoryStats');
    scriptProperties.deleteProperty('imageAutoMode');
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Get image files to process
 */
function getImageFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    // Search for all non-trashed files, we'll filter images during processing
    // This is more reliable than complex file name queries
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  
  while (files.hasNext() && count < batchSize) {
    const file = files.next();
    
    if (isImageFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if file is an image
 */
function isImageFile(file) {
  try {
    const mimeType = file.getMimeType();
    const fileName = file.getName().toLowerCase();
    
    // Check by MIME type
    if (mimeType.startsWith('image/')) {
      return true;
    }
    
    // Check by extension
    const extension = fileName.split('.').pop();
    return CONFIG.COMMON_IMAGE_FORMATS.includes(extension);
    
  } catch (error) {
    console.error(`Error checking if file is image: ${error}`);
    return false;
  }
}

/**
 * Process a single image file
 */
function processImageFile(file, spreadsheet, stats) {
  try {
    const fileData = extractImageFileData(file);
    
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Track by format
    const format = fileData.format;
    stats.filesByFormat[format] = (stats.filesByFormat[format] || 0) + 1;
    
    // Track by year
    const year = new Date(fileData.lastModified).getFullYear();
    stats.filesByYear[year] = (stats.filesByYear[year] || 0) + 1;
    
    // Track by folder
    if (fileData.folderPath) {
      stats.filesByFolder[fileData.folderPath] = (stats.filesByFolder[fileData.folderPath] || 0) + 1;
    }
    
    // Check for large images
    if (fileData.size > CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      stats.largeImages.push({
        name: fileData.name,
        size: fileData.size,
        format: fileData.format,
        dimensions: fileData.dimensions,
        path: fileData.folderPath,
        url: fileData.url
      });
      
      stats.largeImages.sort((a, b) => b.size - a.size);
      stats.largeImages = stats.largeImages.slice(0, 100);
    }
    
    // Check for old images
    const ageInDays = (new Date() - new Date(fileData.lastModified)) / (1000 * 60 * 60 * 24);
    if (ageInDays > CONFIG.OLD_FILE_THRESHOLD_DAYS) {
      if (stats.oldImages.length < 100) {
        stats.oldImages.push({
          name: fileData.name,
          lastModified: fileData.lastModified,
          ageInDays: Math.floor(ageInDays),
          format: fileData.format,
          path: fileData.folderPath,
          url: fileData.url
        });
      }
    }
    
    // Track potential duplicates by name and size
    const duplicateKey = `${fileData.name}_${fileData.size}`;
    if (!stats.duplicateCandidates[duplicateKey]) {
      stats.duplicateCandidates[duplicateKey] = [];
    }
    stats.duplicateCandidates[duplicateKey].push({
      name: fileData.name,
      path: fileData.folderPath,
      size: fileData.size,
      format: fileData.format,
      lastModified: fileData.lastModified,
      url: fileData.url
    });
    
    // Track dimensions
    if (fileData.dimensions && fileData.dimensions !== 'Unknown') {
      stats.dimensionStats[fileData.dimensions] = (stats.dimensionStats[fileData.dimensions] || 0) + 1;
    }
    
    addToImageListSheet(spreadsheet, fileData);
    
  } catch (error) {
    console.error(`Error processing image file: ${error}`);
    throw error;
  }
}

/**
 * Extract image-specific file data
 */
function extractImageFileData(file) {
  const data = {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    format: getImageFormat(file.getName(), file.getMimeType()),
    size: file.getSize(),
    created: file.getDateCreated().toISOString(),
    lastModified: file.getLastUpdated().toISOString(),
    owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
    url: file.getUrl(),
    description: file.getDescription() || '',
    folderPath: '',
    dimensions: 'Unknown',
    sharingAccess: 'Private',
    sharingPermission: 'None'
  };
  
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
 * Determine image format
 */
function getImageFormat(fileName, mimeType) {
  if (mimeType.startsWith('image/')) {
    const format = mimeType.replace('image/', '');
    return format.toUpperCase();
  }
  
  const extension = fileName.split('.').pop().toLowerCase();
  return extension.toUpperCase() || 'Unknown';
}

/**
 * Initialize image-specific statistics
 */
function initializeImageStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    filesByFormat: stats.filesByFormat || {},
    filesByFolder: stats.filesByFolder || {},
    filesByYear: stats.filesByYear || {},
    largeImages: stats.largeImages || [],
    oldImages: stats.oldImages || [],
    duplicateCandidates: stats.duplicateCandidates || {},
    dimensionStats: stats.dimensionStats || {},
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Initialize image-specific sheets
 */
function initializeImageSheets(spreadsheet) {
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      switch (sheetName) {
        case CONFIG.SHEETS.IMAGE_LIST:
          sheet.getRange(1, 1, 1, 10).setValues([[
            'Name', 'Format', 'Size (MB)', 'Dimensions', 'Created', 'Last Modified',
            'Owner', 'Folder Path', 'Sharing', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.LARGE_IMAGES:
          sheet.getRange(1, 1, 1, 6).setValues([[
            'Name', 'Size (MB)', 'Format', 'Dimensions', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.OLD_IMAGES:
          sheet.getRange(1, 1, 1, 6).setValues([[
            'Name', 'Last Modified', 'Age (Days)', 'Format', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.BY_FORMAT:
          sheet.getRange(1, 1, 1, 3).setValues([[
            'Format', 'Count', 'Total Size (MB)'
          ]]);
          sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
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
 * Add image data to the image list sheet
 */
function addToImageListSheet(spreadsheet, fileData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.IMAGE_LIST);
  
  sheet.appendRow([
    fileData.name,
    fileData.format,
    (fileData.size / 1024 / 1024).toFixed(2),
    fileData.dimensions,
    fileData.created,
    fileData.lastModified,
    fileData.owner,
    fileData.folderPath,
    fileData.sharingAccess,
    fileData.url
  ]);
}

/**
 * Update the image overview sheet
 */
function updateImageOverviewSheet(spreadsheet, stats, isFinal) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('Google Drive Image Files Inventory')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(isFinal ? 'Final Report' : 'Progress Report')
    .setFontSize(12);
  
  sheet.getRange(3, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary Statistics
  sheet.getRange(5, 1).setValue('IMAGE SUMMARY STATISTICS').setFontWeight('bold');
  
  const summaryData = [
    ['Total Image Files:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Average File Size:', formatBytes(stats.totalSize / Math.max(stats.totalFiles, 1))],
    ['Large Images (>' + CONFIG.LARGE_FILE_THRESHOLD_MB + 'MB):', stats.largeImages.length],
    ['Old Images (>' + CONFIG.OLD_FILE_THRESHOLD_DAYS + ' days):', stats.oldImages.length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(6, 1, summaryData.length, 2).setValues(summaryData);
  
  // Image Formats Distribution
  const formatRow = 14;
  sheet.getRange(formatRow, 1).setValue('IMAGE FORMATS DISTRIBUTION').setFontWeight('bold');
  
  const formatEntries = Object.entries(stats.filesByFormat)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);
  
  if (formatEntries.length > 0) {
    sheet.getRange(formatRow + 1, 1, formatEntries.length, 2).setValues(formatEntries);
  }
  
  // Images by Year
  const yearRow = formatRow + formatEntries.length + 3;
  sheet.getRange(yearRow, 1).setValue('IMAGES BY YEAR').setFontWeight('bold');
  
  const yearEntries = Object.entries(stats.filesByYear)
    .sort((a, b) => b[0] - a[0])
    .slice(0, 10);
  
  if (yearEntries.length > 0) {
    sheet.getRange(yearRow + 1, 1, yearEntries.length, 2).setValues(yearEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final image reports
 */
function generateImageFinalReports(spreadsheet, stats) {
  console.log("Generating final image reports...");
  
  updateImageOverviewSheet(spreadsheet, stats, true);
  generateImageFormatReport(spreadsheet, stats);
  generateLargeImagesReport(spreadsheet, stats);
  generateOldImagesReport(spreadsheet, stats);
  generateImageDuplicatesReport(spreadsheet, stats);
  
  console.log(`Image reports generated! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate image format analysis report
 */
function generateImageFormatReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_FORMAT);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  const formatData = Object.entries(stats.filesByFormat)
    .sort((a, b) => b[1] - a[1])
    .map(([format, count]) => {
      // Calculate total size for this format (simplified)
      const avgSize = stats.totalSize / stats.totalFiles;
      const totalSizeMB = (count * avgSize / 1024 / 1024).toFixed(2);
      
      return [format, count, totalSizeMB];
    });
  
  if (formatData.length > 0) {
    sheet.getRange(2, 1, formatData.length, 3).setValues(formatData);
  }
}

/**
 * Generate large images report
 */
function generateLargeImagesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LARGE_IMAGES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.largeImages.length > 0) {
    const data = stats.largeImages.map(image => [
      image.name,
      (image.size / 1024 / 1024).toFixed(2),
      image.format,
      image.dimensions || 'Unknown',
      image.path,
      image.url
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

/**
 * Generate old images report
 */
function generateOldImagesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OLD_IMAGES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.oldImages.length > 0) {
    const data = stats.oldImages.map(image => [
      image.name,
      image.lastModified,
      image.ageInDays,
      image.format,
      image.path,
      image.url
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

/**
 * Generate image duplicates report
 */
function generateImageDuplicatesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DUPLICATES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Set up headers
  sheet.getRange(1, 1, 1, 6).setValues([[
    'Name', 'Format', 'Size (MB)', 'Count', 'Locations', 'URLs'
  ]]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  
  const duplicates = Object.entries(stats.duplicateCandidates)
    .filter(([key, files]) => files.length > 1)
    .sort((a, b) => b[1].length - a[1].length)
    .slice(0, 100);
  
  if (duplicates.length > 0) {
    const data = duplicates.map(([key, files]) => [
      files[0].name,
      files[0].format,
      (files[0].size / 1024 / 1024).toFixed(2),
      files.length,
      files.map(f => f.path).join('\n'),
      files.map(f => f.url).join('\n')
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

// Utility functions

function scheduleNextImageRun() {
  cancelImageScheduledRuns();
  
  ScriptApp.newTrigger('continueImageInventory')
    .timeBased()
    .after(1 * 60 * 1000)
    .create();
  
  console.log("Next image inventory run scheduled for 1 minute from now");
}

function continueImageInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (scriptProperties.getProperty('imageAutoMode') !== 'true') {
    console.log("Image auto mode disabled, stopping.");
    cancelImageScheduledRuns();
    return;
  }
  
  inventoryImages();
}

function cancelImageScheduledRuns() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueImageInventory' ||
        trigger.getHandlerFunction() === 'inventoryImages') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log("Cancelled all scheduled image inventory runs");
}

function updateImageInventoryStatus(spreadsheet, status, stats) {
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
    console.log(`Created new image inventory spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Start automatic image inventory
 */
function startAutomaticImageInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const continuationToken = scriptProperties.getProperty('imageContinuationToken');
  if (!continuationToken) {
    resetImageInventory();
  }
  
  scriptProperties.setProperty('imageAutoMode', 'true');
  
  console.log("Starting automatic image inventory...");
  inventoryImages();
}

/**
 * Reset image inventory
 */
function resetImageInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('imageContinuationToken');
  scriptProperties.deleteProperty('imageInventoryStats');
  scriptProperties.deleteProperty('imageAutoMode');
  console.log("Image inventory reset. Next run will start from the beginning.");
}

/**
 * Quick image stats
 */
function getQuickImageStats() {
  console.log("Gathering quick image statistics...");
  
  const stats = {
    totalImages: 0,
    totalSize: 0,
    formatCounts: {},
    largestImage: null
  };
  
  const imageQuery = CONFIG.COMMON_IMAGE_FORMATS
    .map(format => `name contains '.${format}'`)
    .join(' or ');
  
  const files = DriveApp.searchFiles(`(${imageQuery}) and trashed = false`);
  let count = 0;
  const maxCheck = 500;
  
  while (files.hasNext() && count < maxCheck) {
    const file = files.next();
    
    if (isImageFile(file)) {
      count++;
      stats.totalImages++;
      
      const size = file.getSize();
      stats.totalSize += size;
      
      const format = getImageFormat(file.getName(), file.getMimeType());
      stats.formatCounts[format] = (stats.formatCounts[format] || 0) + 1;
      
      if (!stats.largestImage || size > stats.largestImage.size) {
        stats.largestImage = {
          name: file.getName(),
          size: size,
          sizeFormatted: formatBytes(size),
          format: format
        };
      }
    }
  }
  
  console.log(`Quick Image Stats (based on first ${count} images):`);
  console.log(`Total images: ${stats.totalImages}`);
  console.log(`Total size: ${formatBytes(stats.totalSize)}`);
  console.log(`Average size: ${formatBytes(stats.totalSize / stats.totalImages)}`);
  
  if (stats.largestImage) {
    console.log(`Largest image: ${stats.largestImage.name} (${stats.largestImage.sizeFormatted}) - ${stats.largestImage.format}`);
  }
  
  console.log('\nFormat distribution:');
  Object.entries(stats.formatCounts)
    .sort((a, b) => b[1] - a[1])
    .forEach(([format, count]) => {
      console.log(`  ${format}: ${count}`);
    });
  
  return stats;
}