/**
 * Google Drive Document Files Inventory Script
 * Specialized script for analyzing document files in Google Drive
 * Focuses on text documents, PDFs, presentations, and spreadsheets
 */

// Configuration for document analysis
const CONFIG = {
  BATCH_SIZE: 100,
  INVENTORY_SPREADSHEET_NAME: "📄 Document Files Inventory Report",
  
  SHEETS: {
    OVERVIEW: "Overview",
    DOCUMENT_LIST: "Document Files",
    LARGE_DOCS: "Large Documents",
    OLD_DOCS: "Old Documents",
    DUPLICATES: "Potential Duplicates",
    BY_TYPE: "By Document Type",
    GOOGLE_DOCS: "Google Workspace Files",
    SHARING_ANALYSIS: "Sharing Analysis"
  },
  
  LARGE_FILE_THRESHOLD_MB: 25, // Threshold for documents
  OLD_FILE_THRESHOLD_DAYS: 365, // 1 year for documents
  INCLUDE_TRASHED: false,
  TRACK_PERMISSIONS: true,
  
  // Document-specific settings
  DOCUMENT_FORMATS: {
    // Google Workspace files
    'google-document': 'Google Docs',
    'google-spreadsheet': 'Google Sheets', 
    'google-presentation': 'Google Slides',
    'google-form': 'Google Forms',
    'google-drawing': 'Google Drawings',
    
    // Microsoft Office
    'doc': 'Word Document',
    'docx': 'Word Document',
    'xls': 'Excel Spreadsheet',
    'xlsx': 'Excel Spreadsheet',
    'ppt': 'PowerPoint',
    'pptx': 'PowerPoint',
    
    // PDFs and text
    'pdf': 'PDF Document',
    'txt': 'Text File',
    'rtf': 'Rich Text Format',
    'csv': 'CSV File',
    
    // Other formats
    'odt': 'OpenDocument Text',
    'ods': 'OpenDocument Spreadsheet',
    'odp': 'OpenDocument Presentation'
  }
};

/**
 * Main function to inventory document files
 */
function inventoryDocuments() {
  console.log("Starting document files inventory...");
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const isAutoMode = scriptProperties.getProperty('documentAutoMode') === 'true';
  
  if (!isAutoMode) {
    runDocumentInventoryBatch();
  } else {
    runDocumentInventoryContinuously();
  }
}

/**
 * Run document inventory continuously
 */
function runDocumentInventoryContinuously() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 5 * 60 * 1000;
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let batchCount = 0;
  let totalProcessed = 0;
  let hasMoreFiles = true;
  
  console.log("Running document inventory in continuous mode...");
  
  while (hasMoreFiles) {
    const currentTime = new Date().getTime();
    const elapsedTime = currentTime - startTime;
    
    if (elapsedTime > MAX_RUNTIME_MS) {
      console.log(`Approaching time limit after ${(elapsedTime/1000).toFixed(0)} seconds`);
      scheduleNextDocumentRun();
      console.log(`Processed ${totalProcessed} documents in ${batchCount} batches. Scheduled next run.`);
      return;
    }
    
    const result = runDocumentInventoryBatch();
    batchCount++;
    totalProcessed += result.processedCount;
    
    if (!result.hasMore) {
      hasMoreFiles = false;
      console.log(`Document inventory complete! Processed ${totalProcessed} documents in ${batchCount} batches.`);
      cancelDocumentScheduledRuns();
    } else {
      Utilities.sleep(100);
    }
  }
}

/**
 * Run a single batch of document inventory processing
 */
function runDocumentInventoryBatch() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeDocumentSheets(spreadsheet);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('documentContinuationToken');
  
  let stats = JSON.parse(scriptProperties.getProperty('documentInventoryStats') || '{}');
  stats = initializeDocumentStats(stats);
  
  updateDocumentInventoryStatus(spreadsheet, 'RUNNING', stats);
  
  const files = getDocumentFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more document files to process!");
    generateDocumentFinalReports(spreadsheet, stats);
    updateDocumentInventoryStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('documentContinuationToken');
    scriptProperties.deleteProperty('documentInventoryStats');
    scriptProperties.deleteProperty('documentAutoMode');
    
    cancelDocumentScheduledRuns();
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing batch of ${files.length} document files...`);
  
  let processedCount = 0;
  for (const file of files) {
    try {
      processDocumentFile(file, spreadsheet, stats);
      processedCount++;
    } catch (error) {
      console.error(`Error processing document file: ${error}`);
      stats.errors++;
    }
  }
  
  scriptProperties.setProperty('documentInventoryStats', JSON.stringify(stats));
  updateDocumentOverviewSheet(spreadsheet, stats, false);
  
  console.log(`Processed ${processedCount} document files. Total so far: ${stats.totalFiles}`);
  
  const hasMore = files.hasNext && files.hasNext();
  
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('documentContinuationToken', nextToken);
  } else {
    generateDocumentFinalReports(spreadsheet, stats);
    updateDocumentInventoryStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('documentContinuationToken');
    scriptProperties.deleteProperty('documentInventoryStats');
    scriptProperties.deleteProperty('documentAutoMode');
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Get document files to process
 */
function getDocumentFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    // Search for all non-trashed files, we'll filter documents during processing
    // This is more reliable than complex MIME type queries
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  
  while (files.hasNext() && count < batchSize) {
    const file = files.next();
    
    if (isDocumentFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if file is a document
 */
function isDocumentFile(file) {
  try {
    const mimeType = file.getMimeType();
    const fileName = file.getName().toLowerCase();
    
    // Check Google Workspace files
    if (mimeType.startsWith('application/vnd.google-apps.')) {
      const googleType = mimeType.replace('application/vnd.google-apps.', '');
      return ['document', 'spreadsheet', 'presentation', 'form', 'drawing'].includes(googleType);
    }
    
    // Check other document types
    if (mimeType.includes('document') || 
        mimeType.includes('text') || 
        mimeType.includes('spreadsheet') || 
        mimeType.includes('presentation') ||
        mimeType === 'application/pdf') {
      return true;
    }
    
    // Check by extension
    const extension = fileName.split('.').pop();
    return Object.keys(CONFIG.DOCUMENT_FORMATS).includes(extension);
    
  } catch (error) {
    console.error(`Error checking if file is document: ${error}`);
    return false;
  }
}

/**
 * Process a single document file
 */
function processDocumentFile(file, spreadsheet, stats) {
  try {
    const fileData = extractDocumentFileData(file);
    
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Track by document type
    const docType = fileData.documentType;
    stats.filesByType[docType] = (stats.filesByType[docType] || 0) + 1;
    
    // Track Google Workspace vs other files
    if (fileData.isGoogleFile) {
      stats.googleWorkspaceFiles++;
    } else {
      stats.otherDocumentFiles++;
    }
    
    // Track by year
    const year = new Date(fileData.lastModified).getFullYear();
    stats.filesByYear[year] = (stats.filesByYear[year] || 0) + 1;
    
    // Track by folder
    if (fileData.folderPath) {
      stats.filesByFolder[fileData.folderPath] = (stats.filesByFolder[fileData.folderPath] || 0) + 1;
    }
    
    // Track by owner
    stats.filesByOwner[fileData.owner] = (stats.filesByOwner[fileData.owner] || 0) + 1;
    
    // Check for large documents
    if (fileData.size > CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      stats.largeDocuments.push({
        name: fileData.name,
        size: fileData.size,
        documentType: fileData.documentType,
        path: fileData.folderPath,
        url: fileData.url,
        isGoogleFile: fileData.isGoogleFile
      });
      
      stats.largeDocuments.sort((a, b) => b.size - a.size);
      stats.largeDocuments = stats.largeDocuments.slice(0, 100);
    }
    
    // Check for old documents
    const ageInDays = (new Date() - new Date(fileData.lastModified)) / (1000 * 60 * 60 * 24);
    if (ageInDays > CONFIG.OLD_FILE_THRESHOLD_DAYS) {
      if (stats.oldDocuments.length < 100) {
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
    
    // Track shared documents
    if (fileData.sharingAccess !== 'Private') {
      stats.sharedDocuments.push({
        name: fileData.name,
        sharingAccess: fileData.sharingAccess,
        sharingPermission: fileData.sharingPermission,
        documentType: fileData.documentType,
        viewers: fileData.viewers,
        editors: fileData.editors,
        path: fileData.folderPath,
        url: fileData.url
      });
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
      documentType: fileData.documentType,
      lastModified: fileData.lastModified,
      url: fileData.url
    });
    
    addToDocumentListSheet(spreadsheet, fileData);
    
  } catch (error) {
    console.error(`Error processing document file: ${error}`);
    throw error;
  }
}

/**
 * Extract document-specific file data
 */
function extractDocumentFileData(file) {
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
 * Initialize document-specific statistics
 */
function initializeDocumentStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    filesByType: stats.filesByType || {},
    filesByFolder: stats.filesByFolder || {},
    filesByOwner: stats.filesByOwner || {},
    filesByYear: stats.filesByYear || {},
    largeDocuments: stats.largeDocuments || [],
    oldDocuments: stats.oldDocuments || [],
    sharedDocuments: stats.sharedDocuments || [],
    duplicateCandidates: stats.duplicateCandidates || {},
    googleWorkspaceFiles: stats.googleWorkspaceFiles || 0,
    otherDocumentFiles: stats.otherDocumentFiles || 0,
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Initialize document-specific sheets
 */
function initializeDocumentSheets(spreadsheet) {
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      switch (sheetName) {
        case CONFIG.SHEETS.DOCUMENT_LIST:
          sheet.getRange(1, 1, 1, 11).setValues([[
            'Name', 'Document Type', 'Google File', 'Size (MB)', 'Created', 'Last Modified',
            'Owner', 'Folder Path', 'Sharing', 'Collaborators', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.LARGE_DOCS:
          sheet.getRange(1, 1, 1, 6).setValues([[
            'Name', 'Size (MB)', 'Document Type', 'Google File', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.OLD_DOCS:
          sheet.getRange(1, 1, 1, 6).setValues([[
            'Name', 'Last Modified', 'Age (Days)', 'Document Type', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.BY_TYPE:
          sheet.getRange(1, 1, 1, 4).setValues([[
            'Document Type', 'Count', 'Total Size (MB)', 'Percentage'
          ]]);
          sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.SHARING_ANALYSIS:
          sheet.getRange(1, 1, 1, 7).setValues([[
            'Name', 'Document Type', 'Sharing Level', 'Viewers', 'Editors', 'Folder Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
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
 * Add document data to the document list sheet
 */
function addToDocumentListSheet(spreadsheet, fileData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DOCUMENT_LIST);
  
  const collaborators = [...fileData.viewers, ...fileData.editors].join(', ');
  
  sheet.appendRow([
    fileData.name,
    fileData.documentType,
    fileData.isGoogleFile ? 'Yes' : 'No',
    (fileData.size / 1024 / 1024).toFixed(2),
    fileData.created,
    fileData.lastModified,
    fileData.owner,
    fileData.folderPath,
    fileData.sharingAccess,
    collaborators,
    fileData.url
  ]);
}

/**
 * Update the document overview sheet
 */
function updateDocumentOverviewSheet(spreadsheet, stats, isFinal) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('Google Drive Document Files Inventory')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(isFinal ? 'Final Report' : 'Progress Report')
    .setFontSize(12);
  
  sheet.getRange(3, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary Statistics
  sheet.getRange(5, 1).setValue('DOCUMENT SUMMARY STATISTICS').setFontWeight('bold');
  
  const summaryData = [
    ['Total Document Files:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Average File Size:', formatBytes(stats.totalSize / Math.max(stats.totalFiles, 1))],
    ['Google Workspace Files:', stats.googleWorkspaceFiles],
    ['Other Document Files:', stats.otherDocumentFiles],
    ['Large Documents (>' + CONFIG.LARGE_FILE_THRESHOLD_MB + 'MB):', stats.largeDocuments.length],
    ['Old Documents (>' + CONFIG.OLD_FILE_THRESHOLD_DAYS + ' days):', stats.oldDocuments.length],
    ['Shared Documents:', stats.sharedDocuments.length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(6, 1, summaryData.length, 2).setValues(summaryData);
  
  // Document Types Distribution
  const typeRow = 16;
  sheet.getRange(typeRow, 1).setValue('DOCUMENT TYPES DISTRIBUTION').setFontWeight('bold');
  
  const typeEntries = Object.entries(stats.filesByType)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);
  
  if (typeEntries.length > 0) {
    sheet.getRange(typeRow + 1, 1, typeEntries.length, 2).setValues(typeEntries);
  }
  
  // Documents by Year
  const yearRow = typeRow + typeEntries.length + 3;
  sheet.getRange(yearRow, 1).setValue('DOCUMENTS BY YEAR').setFontWeight('bold');
  
  const yearEntries = Object.entries(stats.filesByYear)
    .sort((a, b) => b[0] - a[0])
    .slice(0, 10);
  
  if (yearEntries.length > 0) {
    sheet.getRange(yearRow + 1, 1, yearEntries.length, 2).setValues(yearEntries);
  }
  
  // Top Document Owners
  const ownerRow = yearRow + yearEntries.length + 3;
  sheet.getRange(ownerRow, 1).setValue('TOP DOCUMENT OWNERS').setFontWeight('bold');
  
  const ownerEntries = Object.entries(stats.filesByOwner)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (ownerEntries.length > 0) {
    sheet.getRange(ownerRow + 1, 1, ownerEntries.length, 2).setValues(ownerEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final document reports
 */
function generateDocumentFinalReports(spreadsheet, stats) {
  console.log("Generating final document reports...");
  
  updateDocumentOverviewSheet(spreadsheet, stats, true);
  generateDocumentTypeReport(spreadsheet, stats);
  generateLargeDocumentsReport(spreadsheet, stats);
  generateOldDocumentsReport(spreadsheet, stats);
  generateDocumentSharingReport(spreadsheet, stats);
  generateDocumentDuplicatesReport(spreadsheet, stats);
  
  console.log(`Document reports generated! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate document type analysis report
 */
function generateDocumentTypeReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_TYPE);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  const typeData = Object.entries(stats.filesByType)
    .sort((a, b) => b[1] - a[1])
    .map(([type, count]) => {
      const percentage = ((count / stats.totalFiles) * 100).toFixed(2);
      const avgSize = stats.totalSize / stats.totalFiles;
      const totalSizeMB = (count * avgSize / 1024 / 1024).toFixed(2);
      
      return [type, count, totalSizeMB, percentage + '%'];
    });
  
  if (typeData.length > 0) {
    sheet.getRange(2, 1, typeData.length, 4).setValues(typeData);
  }
}

/**
 * Generate large documents report
 */
function generateLargeDocumentsReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LARGE_DOCS);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.largeDocuments.length > 0) {
    const data = stats.largeDocuments.map(doc => [
      doc.name,
      (doc.size / 1024 / 1024).toFixed(2),
      doc.documentType,
      doc.isGoogleFile ? 'Yes' : 'No',
      doc.path,
      doc.url
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

/**
 * Generate old documents report
 */
function generateOldDocumentsReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OLD_DOCS);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.oldDocuments.length > 0) {
    const data = stats.oldDocuments.map(doc => [
      doc.name,
      doc.lastModified,
      doc.ageInDays,
      doc.documentType,
      doc.path,
      doc.url
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

/**
 * Generate document sharing analysis report
 */
function generateDocumentSharingReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SHARING_ANALYSIS);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.sharedDocuments.length > 0) {
    const data = stats.sharedDocuments.slice(0, 200).map(doc => [
      doc.name,
      doc.documentType,
      doc.sharingAccess,
      doc.viewers.join(', '),
      doc.editors.join(', '),
      doc.path,
      doc.url
    ]);
    
    sheet.getRange(2, 1, data.length, 7).setValues(data);
  }
}

/**
 * Generate document duplicates report
 */
function generateDocumentDuplicatesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DUPLICATES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Set up headers
  sheet.getRange(1, 1, 1, 6).setValues([[
    'Name', 'Document Type', 'Size (MB)', 'Count', 'Locations', 'URLs'
  ]]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  
  const duplicates = Object.entries(stats.duplicateCandidates)
    .filter(([key, files]) => files.length > 1)
    .sort((a, b) => b[1].length - a[1].length)
    .slice(0, 100);
  
  if (duplicates.length > 0) {
    const data = duplicates.map(([key, files]) => [
      files[0].name,
      files[0].documentType,
      (files[0].size / 1024 / 1024).toFixed(2),
      files.length,
      files.map(f => f.path).join('\n'),
      files.map(f => f.url).join('\n')
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

// Utility functions for document inventory

function scheduleNextDocumentRun() {
  cancelDocumentScheduledRuns();
  
  ScriptApp.newTrigger('continueDocumentInventory')
    .timeBased()
    .after(1 * 60 * 1000)
    .create();
  
  console.log("Next document inventory run scheduled for 1 minute from now");
}

function continueDocumentInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (scriptProperties.getProperty('documentAutoMode') !== 'true') {
    console.log("Document auto mode disabled, stopping.");
    cancelDocumentScheduledRuns();
    return;
  }
  
  inventoryDocuments();
}

function cancelDocumentScheduledRuns() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueDocumentInventory' ||
        trigger.getHandlerFunction() === 'inventoryDocuments') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log("Cancelled all scheduled document inventory runs");
}

function updateDocumentInventoryStatus(spreadsheet, status, stats) {
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
    console.log(`Created new document inventory spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Start automatic document inventory
 */
function startAutomaticDocumentInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const continuationToken = scriptProperties.getProperty('documentContinuationToken');
  if (!continuationToken) {
    resetDocumentInventory();
  }
  
  scriptProperties.setProperty('documentAutoMode', 'true');
  
  console.log("Starting automatic document inventory...");
  inventoryDocuments();
}

/**
 * Reset document inventory
 */
function resetDocumentInventory() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('documentContinuationToken');
  scriptProperties.deleteProperty('documentInventoryStats');
  scriptProperties.deleteProperty('documentAutoMode');
  console.log("Document inventory reset. Next run will start from the beginning.");
}

/**
 * Quick document stats
 */
function getQuickDocumentStats() {
  console.log("Gathering quick document statistics...");
  
  const stats = {
    totalDocuments: 0,
    totalSize: 0,
    typeCounts: {},
    googleWorkspaceCount: 0,
    otherDocCount: 0
  };
  
  const docQuery = [
    "mimeType contains 'document'",
    "mimeType contains 'text'", 
    "mimeType contains 'spreadsheet'",
    "mimeType contains 'presentation'",
    "mimeType = 'application/pdf'"
  ].join(' or ');
  
  const files = DriveApp.searchFiles(`(${docQuery}) and trashed = false`);
  let count = 0;
  const maxCheck = 500;
  
  while (files.hasNext() && count < maxCheck) {
    const file = files.next();
    
    if (isDocumentFile(file)) {
      count++;
      stats.totalDocuments++;
      
      const size = file.getSize();
      stats.totalSize += size;
      
      const docType = getDocumentType(file.getName(), file.getMimeType());
      stats.typeCounts[docType] = (stats.typeCounts[docType] || 0) + 1;
      
      if (file.getMimeType().startsWith('application/vnd.google-apps.')) {
        stats.googleWorkspaceCount++;
      } else {
        stats.otherDocCount++;
      }
    }
  }
  
  console.log(`Quick Document Stats (based on first ${count} documents):`);
  console.log(`Total documents: ${stats.totalDocuments}`);
  console.log(`Total size: ${formatBytes(stats.totalSize)}`);
  console.log(`Average size: ${formatBytes(stats.totalSize / stats.totalDocuments)}`);
  console.log(`Google Workspace files: ${stats.googleWorkspaceCount}`);
  console.log(`Other document files: ${stats.otherDocCount}`);
  
  console.log('\nDocument type distribution:');
  Object.entries(stats.typeCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .forEach(([type, count]) => {
      console.log(`  ${type}: ${count}`);
    });
  
  return stats;
}