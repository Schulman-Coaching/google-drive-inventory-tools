/**
 * Google Drive Document Files Inventory Script - BUGFIX VERSION
 * This version fixes the search query issue and should work reliably
 * 
 * QUICK START:
 * 1. Copy this entire script to Google Apps Script
 * 2. Run: inventoryDocuments()
 * 3. Check the generated spreadsheet for results
 */

// Configuration for document analysis
const CONFIG = {
  BATCH_SIZE: 50, // Smaller batch size for better reliability
  INVENTORY_SPREADSHEET_NAME: "📄 Document Files Inventory Report",
  
  SHEETS: {
    OVERVIEW: "Overview",
    DOCUMENT_LIST: "Document Files",
    LARGE_DOCS: "Large Documents",
    OLD_DOCS: "Old Documents",
    BY_TYPE: "By Document Type",
    GOOGLE_DOCS: "Google Workspace Files"
  },
  
  LARGE_FILE_THRESHOLD_MB: 25,
  OLD_FILE_THRESHOLD_DAYS: 365,
  INCLUDE_TRASHED: false,
  TRACK_PERMISSIONS: true,
  
  // Document formats we're looking for
  DOCUMENT_FORMATS: {
    'google-document': 'Google Docs',
    'google-spreadsheet': 'Google Sheets', 
    'google-presentation': 'Google Slides',
    'google-form': 'Google Forms',
    'doc': 'Word Document',
    'docx': 'Word Document',
    'xls': 'Excel Spreadsheet',
    'xlsx': 'Excel Spreadsheet',
    'ppt': 'PowerPoint',
    'pptx': 'PowerPoint',
    'pdf': 'PDF Document',
    'txt': 'Text File',
    'csv': 'CSV File'
  }
};

/**
 * Main function - run this to start the document inventory
 */
function inventoryDocuments() {
  console.log("Starting document inventory...");
  
  try {
    const result = runDocumentInventoryBatch();
    console.log(`Processed ${result.processedCount} documents`);
    
    if (result.hasMore) {
      console.log("More documents to process. Run inventoryDocuments() again to continue.");
    } else {
      console.log("Document inventory complete!");
    }
  } catch (error) {
    console.error(`Error in document inventory: ${error}`);
    console.error(`Error details: ${error.stack}`);
  }
}

/**
 * Process a batch of documents
 */
function runDocumentInventoryBatch() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeSheets(spreadsheet);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('documentContinuationToken');
  
  let stats = JSON.parse(scriptProperties.getProperty('documentInventoryStats') || '{}');
  stats = initializeStats(stats);
  
  // Get files to process
  const files = getDocumentFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more document files found!");
    generateFinalReports(spreadsheet, stats);
    
    // Clean up
    scriptProperties.deleteProperty('documentContinuationToken');
    scriptProperties.deleteProperty('documentInventoryStats');
    
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing ${files.length} documents...`);
  
  // Process each file
  let processedCount = 0;
  for (const file of files) {
    try {
      if (isDocumentFile(file)) {
        processDocumentFile(file, spreadsheet, stats);
        processedCount++;
      }
    } catch (error) {
      console.error(`Error processing file ${file.getName()}: ${error}`);
      stats.errors = (stats.errors || 0) + 1;
    }
  }
  
  // Save progress
  scriptProperties.setProperty('documentInventoryStats', JSON.stringify(stats));
  updateOverviewSheet(spreadsheet, stats);
  
  // Check if more files to process
  const hasMore = files.hasNext && files.hasNext();
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('documentContinuationToken', nextToken);
    console.log(`Processed ${processedCount} documents. More files available.`);
  } else {
    generateFinalReports(spreadsheet, stats);
    scriptProperties.deleteProperty('documentContinuationToken');
    scriptProperties.deleteProperty('documentInventoryStats');
    console.log(`Final batch: processed ${processedCount} documents.`);
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Get document files to process (FIXED VERSION)
 */
function getDocumentFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    // Simple, reliable query - get all non-trashed files
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  let checkedCount = 0;
  const maxCheck = batchSize * 10; // Check more files to find documents
  
  while (files.hasNext() && count < batchSize && checkedCount < maxCheck) {
    const file = files.next();
    checkedCount++;
    
    if (isDocumentFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  console.log(`Found ${count} documents after checking ${checkedCount} files`);
  
  // Attach continuation methods
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if a file is a document (IMPROVED VERSION)
 */
function isDocumentFile(file) {
  try {
    const mimeType = file.getMimeType();
    const fileName = file.getName().toLowerCase();
    
    // Check Google Workspace files first
    if (mimeType.startsWith('application/vnd.google-apps.')) {
      const googleType = mimeType.replace('application/vnd.google-apps.', '');
      return ['document', 'spreadsheet', 'presentation', 'form'].includes(googleType);
    }
    
    // Check common document MIME types
    const documentMimeTypes = [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-powerpoint',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'text/plain',
      'text/csv'
    ];
    
    if (documentMimeTypes.includes(mimeType)) {
      return true;
    }
    
    // Check by file extension as fallback
    const extension = fileName.split('.').pop();
    const documentExtensions = ['doc', 'docx', 'pdf', 'txt', 'xls', 'xlsx', 'ppt', 'pptx', 'csv'];
    return documentExtensions.includes(extension);
    
  } catch (error) {
    console.error(`Error checking file type for ${file.getName()}: ${error}`);
    return false;
  }
}

/**
 * Process a single document file
 */
function processDocumentFile(file, spreadsheet, stats) {
  try {
    const fileData = extractFileData(file);
    
    stats.totalFiles++;
    stats.totalSize = (stats.totalSize || 0) + fileData.size;
    
    // Track by document type
    const docType = fileData.documentType;
    stats.filesByType = stats.filesByType || {};
    stats.filesByType[docType] = (stats.filesByType[docType] || 0) + 1;
    
    // Track Google vs other files
    if (fileData.isGoogleFile) {
      stats.googleWorkspaceFiles = (stats.googleWorkspaceFiles || 0) + 1;
    } else {
      stats.otherDocumentFiles = (stats.otherDocumentFiles || 0) + 1;
    }
    
    // Track by owner
    stats.filesByOwner = stats.filesByOwner || {};
    stats.filesByOwner[fileData.owner] = (stats.filesByOwner[fileData.owner] || 0) + 1;
    
    // Check for large documents
    if (fileData.size > CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      stats.largeDocuments = stats.largeDocuments || [];
      stats.largeDocuments.push({
        name: fileData.name,
        size: fileData.size,
        documentType: fileData.documentType,
        path: fileData.folderPath,
        url: fileData.url
      });
      
      // Keep only top 50 largest
      stats.largeDocuments.sort((a, b) => b.size - a.size);
      stats.largeDocuments = stats.largeDocuments.slice(0, 50);
    }
    
    // Check for old documents
    const ageInDays = (new Date() - new Date(fileData.lastModified)) / (1000 * 60 * 60 * 24);
    if (ageInDays > CONFIG.OLD_FILE_THRESHOLD_DAYS) {
      stats.oldDocuments = stats.oldDocuments || [];
      if (stats.oldDocuments.length < 50) {
        stats.oldDocuments.push({
          name: fileData.name,
          lastModified: fileData.lastModified,
          ageInDays: Math.floor(ageInDays),
          documentType: fileData.documentType,
          path: fileData.folderPath,
          url: fileData.url
        });
      }
    }
    
    // Add to main document list
    addToDocumentListSheet(spreadsheet, fileData);
    
  } catch (error) {
    console.error(`Error processing document file: ${error}`);
    throw error;
  }
}

/**
 * Extract file data
 */
function extractFileData(file) {
  const data = {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    documentType: getDocumentType(file.getName(), file.getMimeType()),
    isGoogleFile: file.getMimeType().startsWith('application/vnd.google-apps.'),
    size: file.getSize(),
    created: file.getDateCreated().toISOString(),
    lastModified: file.getLastUpdated().toISOString(),
    owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
    url: file.getUrl(),
    folderPath: 'Root',
    sharingAccess: 'Private'
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
  try {
    const access = file.getSharingAccess();
    data.sharingAccess = access.toString();
  } catch (error) {
    // Some files may not have accessible permissions
  }
  
  return data;
}

/**
 * Determine document type
 */
function getDocumentType(fileName, mimeType) {
  // Check for Google Workspace files
  if (mimeType.startsWith('application/vnd.google-apps.')) {
    const googleType = mimeType.replace('application/vnd.google-apps.', '');
    return CONFIG.DOCUMENT_FORMATS[`google-${googleType}`] || `Google ${googleType}`;
  }
  
  // Check by extension
  const extension = fileName.split('.').pop().toLowerCase();
  return CONFIG.DOCUMENT_FORMATS[extension] || extension.toUpperCase() || 'Unknown Document';
}

/**
 * Initialize statistics
 */
function initializeStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    filesByType: stats.filesByType || {},
    filesByOwner: stats.filesByOwner || {},
    largeDocuments: stats.largeDocuments || [],
    oldDocuments: stats.oldDocuments || [],
    googleWorkspaceFiles: stats.googleWorkspaceFiles || 0,
    otherDocumentFiles: stats.otherDocumentFiles || 0,
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Initialize spreadsheet sheets
 */
function initializeSheets(spreadsheet) {
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      if (sheetName === CONFIG.SHEETS.DOCUMENT_LIST) {
        sheet.getRange(1, 1, 1, 9).setValues([[
          'Name', 'Document Type', 'Google File', 'Size (MB)', 'Created', 
          'Last Modified', 'Owner', 'Folder Path', 'URL'
        ]]);
        sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
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
 * Add document to the main list sheet
 */
function addToDocumentListSheet(spreadsheet, fileData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DOCUMENT_LIST);
  
  sheet.appendRow([
    fileData.name,
    fileData.documentType,
    fileData.isGoogleFile ? 'Yes' : 'No',
    (fileData.size / 1024 / 1024).toFixed(2),
    fileData.created,
    fileData.lastModified,
    fileData.owner,
    fileData.folderPath,
    fileData.url
  ]);
}

/**
 * Update overview sheet
 */
function updateOverviewSheet(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('Document Files Inventory')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary statistics
  const summaryData = [
    ['Total Documents:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Google Workspace Files:', stats.googleWorkspaceFiles],
    ['Other Document Files:', stats.otherDocumentFiles],
    ['Large Documents:', stats.largeDocuments.length],
    ['Old Documents:', stats.oldDocuments.length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(4, 1, summaryData.length, 2).setValues(summaryData);
  
  // Document types
  sheet.getRange(12, 1).setValue('Document Types').setFontWeight('bold');
  
  const typeEntries = Object.entries(stats.filesByType)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (typeEntries.length > 0) {
    sheet.getRange(13, 1, typeEntries.length, 2).setValues(typeEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final reports
 */
function generateFinalReports(spreadsheet, stats) {
  console.log("Generating final reports...");
  
  // Update overview
  updateOverviewSheet(spreadsheet, stats);
  
  // Generate document type report
  generateDocumentTypeReport(spreadsheet, stats);
  
  // Generate large documents report
  generateLargeDocumentsReport(spreadsheet, stats);
  
  console.log(`Document inventory complete! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate document type breakdown
 */
function generateDocumentTypeReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_TYPE);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 3).setValues([['Document Type', 'Count', 'Percentage']]);
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

/**
 * Generate large documents report
 */
function generateLargeDocumentsReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LARGE_DOCS);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Size (MB)', 'Type', 'Folder', 'URL']]);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  
  if (stats.largeDocuments.length > 0) {
    const data = stats.largeDocuments.map(doc => [
      doc.name,
      (doc.size / 1024 / 1024).toFixed(2),
      doc.documentType,
      doc.path,
      doc.url
    ]);
    
    sheet.getRange(2, 1, data.length, 5).setValues(data);
  }
}

// Utility functions

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
    console.log(`Created new spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Quick document statistics (for testing)
 */
function getQuickDocumentStats() {
  console.log("Getting quick document stats...");
  
  const stats = {
    totalDocuments: 0,
    googleWorkspace: 0,
    otherDocs: 0,
    typeBreakdown: {}
  };
  
  const files = DriveApp.searchFiles('trashed = false');
  let count = 0;
  const maxCheck = 200;
  
  while (files.hasNext() && count < maxCheck) {
    const file = files.next();
    
    if (isDocumentFile(file)) {
      stats.totalDocuments++;
      
      if (file.getMimeType().startsWith('application/vnd.google-apps.')) {
        stats.googleWorkspace++;
      } else {
        stats.otherDocs++;
      }
      
      const docType = getDocumentType(file.getName(), file.getMimeType());
      stats.typeBreakdown[docType] = (stats.typeBreakdown[docType] || 0) + 1;
    }
    count++;
  }
  
  console.log(`Found ${stats.totalDocuments} documents in first ${count} files`);
  console.log(`Google Workspace: ${stats.googleWorkspace}, Other: ${stats.otherDocs}`);
  console.log('Type breakdown:', stats.typeBreakdown);
  
  return stats;
}

/**
 * USAGE INSTRUCTIONS:
 * 
 * 1. Copy this entire script to Google Apps Script
 * 2. Run: inventoryDocuments() - this will process documents in batches
 * 3. If it says "More documents to process", run inventoryDocuments() again
 * 4. Check the generated spreadsheet for your results
 * 5. For testing: run getQuickDocumentStats() first to see what you have
 */