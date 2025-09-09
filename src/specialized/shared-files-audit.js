/**
 * Google Drive Shared Files Security Audit Script
 * Specialized script for analyzing file sharing and permissions in Google Drive
 * Focuses on security, compliance, and access management
 */

// Configuration for shared files analysis
const CONFIG = {
  BATCH_SIZE: 100,
  INVENTORY_SPREADSHEET_NAME: "🔒 Shared Files Security Audit Report",
  
  SHEETS: {
    OVERVIEW: "Overview",
    SHARED_FILES: "All Shared Files",
    PUBLIC_FILES: "Public Files",
    DOMAIN_SHARED: "Domain Shared",
    EXTERNALLY_SHARED: "External Sharing",
    BY_PERMISSION: "By Permission Level",
    BY_OWNER: "By Owner",
    HIGH_RISK: "High Risk Files",
    RECOMMENDATIONS: "Security Recommendations"
  },
  
  INCLUDE_TRASHED: false,
  TRACK_DETAILED_PERMISSIONS: true,
  
  // Security analysis settings
  ANALYZE_EXTERNAL_DOMAINS: true,
  FLAG_PUBLIC_FILES: true,
  FLAG_SENSITIVE_FILES: true,
  MAX_EXTERNAL_DOMAINS: 100,
  
  // Risk assessment
  HIGH_RISK_KEYWORDS: [
    'confidential', 'secret', 'private', 'internal', 'restricted',
    'salary', 'budget', 'financial', 'contract', 'agreement',
    'password', 'credential', 'api', 'key', 'token',
    'personal', 'ssn', 'social security', 'bank', 'account'
  ]
};

/**
 * Main function to audit shared files
 */
function auditSharedFiles() {
  console.log("Starting shared files security audit...");
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const isAutoMode = scriptProperties.getProperty('sharedFilesAutoMode') === 'true';
  
  if (!isAutoMode) {
    runSharedFilesAuditBatch();
  } else {
    runSharedFilesAuditContinuously();
  }
}

/**
 * Run shared files audit continuously
 */
function runSharedFilesAuditContinuously() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 5 * 60 * 1000;
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let batchCount = 0;
  let totalProcessed = 0;
  let hasMoreFiles = true;
  
  console.log("Running shared files audit in continuous mode...");
  
  while (hasMoreFiles) {
    const currentTime = new Date().getTime();
    const elapsedTime = currentTime - startTime;
    
    if (elapsedTime > MAX_RUNTIME_MS) {
      console.log(`Approaching time limit after ${(elapsedTime/1000).toFixed(0)} seconds`);
      scheduleNextSharedFilesRun();
      console.log(`Processed ${totalProcessed} shared files in ${batchCount} batches. Scheduled next run.`);
      return;
    }
    
    const result = runSharedFilesAuditBatch();
    batchCount++;
    totalProcessed += result.processedCount;
    
    if (!result.hasMore) {
      hasMoreFiles = false;
      console.log(`Shared files audit complete! Processed ${totalProcessed} files in ${batchCount} batches.`);
      cancelSharedFilesScheduledRuns();
    } else {
      Utilities.sleep(150); // Slightly longer pause for permission checking
    }
  }
}

/**
 * Run a single batch of shared files audit
 */
function runSharedFilesAuditBatch() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeSharedFilesSheets(spreadsheet);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('sharedFilesContinuationToken');
  
  let stats = JSON.parse(scriptProperties.getProperty('sharedFilesStats') || '{}');
  stats = initializeSharedFilesStats(stats);
  
  updateSharedFilesStatus(spreadsheet, 'RUNNING', stats);
  
  const files = getSharedFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more shared files to process!");
    generateSharedFilesFinalReports(spreadsheet, stats);
    updateSharedFilesStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('sharedFilesContinuationToken');
    scriptProperties.deleteProperty('sharedFilesStats');
    scriptProperties.deleteProperty('sharedFilesAutoMode');
    
    cancelSharedFilesScheduledRuns();
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing batch of ${files.length} shared files...`);
  
  let processedCount = 0;
  for (const file of files) {
    try {
      processSharedFile(file, spreadsheet, stats);
      processedCount++;
    } catch (error) {
      console.error(`Error processing shared file: ${error}`);
      stats.errors++;
    }
  }
  
  scriptProperties.setProperty('sharedFilesStats', JSON.stringify(stats));
  updateSharedFilesOverviewSheet(spreadsheet, stats, false);
  
  console.log(`Processed ${processedCount} shared files. Total so far: ${stats.totalFiles}`);
  
  const hasMore = files.hasNext && files.hasNext();
  
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('sharedFilesContinuationToken', nextToken);
  } else {
    generateSharedFilesFinalReports(spreadsheet, stats);
    updateSharedFilesStatus(spreadsheet, 'COMPLETE', stats);
    
    scriptProperties.deleteProperty('sharedFilesContinuationToken');
    scriptProperties.deleteProperty('sharedFilesStats');
    scriptProperties.deleteProperty('sharedFilesAutoMode');
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Get shared files to process
 */
function getSharedFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    // Search for all non-trashed files, we'll filter shared files during processing
    // This is more reliable than complex visibility queries
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  let checkedCount = 0;
  const maxCheck = batchSize * 10; // Check more files to find shared ones
  
  while (files.hasNext() && count < batchSize && checkedCount < maxCheck) {
    const file = files.next();
    checkedCount++;
    
    if (isSharedFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if file is shared
 */
function isSharedFile(file) {
  try {
    const access = file.getSharingAccess();
    return access.toString() !== 'PRIVATE';
  } catch (error) {
    console.error(`Error checking file sharing: ${error}`);
    return false;
  }
}

/**
 * Process a single shared file
 */
function processSharedFile(file, spreadsheet, stats) {
  try {
    const fileData = extractSharedFileData(file);
    
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Track by sharing access level
    const accessLevel = fileData.sharingAccess;
    stats.byAccessLevel[accessLevel] = (stats.byAccessLevel[accessLevel] || 0) + 1;
    
    // Track by permission level
    const permissionLevel = fileData.sharingPermission;
    stats.byPermission[permissionLevel] = (stats.byPermission[permissionLevel] || 0) + 1;
    
    // Track by owner
    stats.byOwner[fileData.owner] = (stats.byOwner[fileData.owner] || 0) + 1;
    
    // Track by file type
    const fileType = fileData.type;
    stats.byFileType[fileType] = (stats.byFileType[fileType] || 0) + 1;
    
    // Categorize sharing level
    categorizeSharing(fileData, stats);
    
    // Security risk assessment
    const riskLevel = assessSecurityRisk(fileData);
    if (riskLevel >= 70) { // High risk threshold
      stats.highRiskFiles.push({
        name: fileData.name,
        type: fileData.type,
        size: fileData.size,
        owner: fileData.owner,
        sharingAccess: fileData.sharingAccess,
        sharingPermission: fileData.sharingPermission,
        path: fileData.folderPath,
        url: fileData.url,
        riskScore: riskLevel,
        riskFactors: getRiskFactors(fileData),
        externalDomains: fileData.externalDomains,
        viewerCount: fileData.viewers.length,
        editorCount: fileData.editors.length
      });
      
      // Sort by risk score and keep top 100
      stats.highRiskFiles.sort((a, b) => b.riskScore - a.riskScore);
      stats.highRiskFiles = stats.highRiskFiles.slice(0, 100);
    }
    
    // Track external domains
    fileData.externalDomains.forEach(domain => {
      stats.externalDomains[domain] = (stats.externalDomains[domain] || 0) + 1;
    });
    
    addToSharedFilesSheet(spreadsheet, fileData, riskLevel);
    
  } catch (error) {
    console.error(`Error processing shared file: ${error}`);
    throw error;
  }
}

/**
 * Extract shared file specific data
 */
function extractSharedFileData(file) {
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
    editors: [],
    externalDomains: [],
    isPublic: false,
    isDomainShared: false
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
  
  // Get detailed sharing information
  try {
    const access = file.getSharingAccess();
    const permission = file.getSharingPermission();
    
    data.sharingAccess = access.toString();
    data.sharingPermission = permission.toString();
    
    data.isPublic = (access.toString() === 'ANYONE' || access.toString() === 'ANYONE_WITH_LINK');
    data.isDomainShared = (access.toString() === 'DOMAIN' || access.toString() === 'DOMAIN_WITH_LINK');
    
    // Get viewers and editors
    const viewers = file.getViewers();
    const editors = file.getEditors();
    
    data.viewers = viewers.slice(0, 20).map(user => user.getEmail()); // Limit to avoid timeout
    data.editors = editors.slice(0, 20).map(user => user.getEmail());
    
    // Identify external domains
    const ownerDomain = data.owner.split('@')[1];
    const allUsers = [...data.viewers, ...data.editors];
    
    data.externalDomains = [...new Set(allUsers
      .map(email => email.split('@')[1])
      .filter(domain => domain && domain !== ownerDomain)
    )];
    
  } catch (error) {
    console.error(`Error getting sharing details: ${error}`);
  }
  
  return data;
}

/**
 * Categorize sharing level
 */
function categorizeSharing(fileData, stats) {
  if (fileData.isPublic) {
    stats.publicFiles++;
  } else if (fileData.isDomainShared) {
    stats.domainSharedFiles++;
  } else if (fileData.externalDomains.length > 0) {
    stats.externallySharedFiles++;
  } else {
    stats.internallySharedFiles++;
  }
}

/**
 * Assess security risk level (0-100)
 */
function assessSecurityRisk(fileData) {
  let riskScore = 0;
  
  // Public access (highest risk)
  if (fileData.isPublic) {
    riskScore += 40;
  } else if (fileData.isDomainShared) {
    riskScore += 20;
  } else if (fileData.externalDomains.length > 0) {
    riskScore += 25;
  }
  
  // Permission level risk
  if (fileData.sharingPermission === 'EDIT' && fileData.isPublic) {
    riskScore += 20;
  } else if (fileData.sharingPermission === 'EDIT') {
    riskScore += 10;
  }
  
  // File content sensitivity
  const fileName = fileData.name.toLowerCase();
  const description = fileData.description.toLowerCase();
  
  CONFIG.HIGH_RISK_KEYWORDS.forEach(keyword => {
    if (fileName.includes(keyword) || description.includes(keyword)) {
      riskScore += 5; // Max 25 points for sensitive content
    }
  });
  
  riskScore = Math.min(riskScore, 25); // Cap sensitive content score
  
  // File type risk
  const sensitiveTypes = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'csv'];
  if (sensitiveTypes.includes(fileData.type.toLowerCase())) {
    riskScore += 10;
  }
  
  // Number of external users
  if (fileData.externalDomains.length > 5) {
    riskScore += 15;
  } else if (fileData.externalDomains.length > 2) {
    riskScore += 10;
  } else if (fileData.externalDomains.length > 0) {
    riskScore += 5;
  }
  
  return Math.min(riskScore, 100);
}

/**
 * Get risk factors for a file
 */
function getRiskFactors(fileData) {
  const factors = [];
  
  if (fileData.isPublic) {
    factors.push('Publicly accessible');
  }
  
  if (fileData.sharingPermission === 'EDIT' && fileData.isPublic) {
    factors.push('Public edit access');
  }
  
  if (fileData.externalDomains.length > 0) {
    factors.push(`Shared with ${fileData.externalDomains.length} external domain(s)`);
  }
  
  const fileName = fileData.name.toLowerCase();
  const sensitiveKeywords = CONFIG.HIGH_RISK_KEYWORDS.filter(keyword => 
    fileName.includes(keyword)
  );
  
  if (sensitiveKeywords.length > 0) {
    factors.push(`Contains sensitive keywords: ${sensitiveKeywords.join(', ')}`);
  }
  
  if (fileData.viewers.length + fileData.editors.length > 20) {
    factors.push('Shared with many users');
  }
  
  return factors;
}

/**
 * Initialize shared files statistics
 */
function initializeSharedFilesStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    byAccessLevel: stats.byAccessLevel || {},
    byPermission: stats.byPermission || {},
    byOwner: stats.byOwner || {},
    byFileType: stats.byFileType || {},
    publicFiles: stats.publicFiles || 0,
    domainSharedFiles: stats.domainSharedFiles || 0,
    externallySharedFiles: stats.externallySharedFiles || 0,
    internallySharedFiles: stats.internallySharedFiles || 0,
    highRiskFiles: stats.highRiskFiles || [],
    externalDomains: stats.externalDomains || {},
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Initialize shared files specific sheets
 */
function initializeSharedFilesSheets(spreadsheet) {
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      switch (sheetName) {
        case CONFIG.SHEETS.SHARED_FILES:
          sheet.getRange(1, 1, 1, 12).setValues([[
            'Name', 'Type', 'Owner', 'Sharing Access', 'Permission', 
            'Viewers', 'Editors', 'External Domains', 'Risk Score', 
            'Folder Path', 'Size (MB)', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.HIGH_RISK:
          sheet.getRange(1, 1, 1, 10).setValues([[
            'Name', 'Type', 'Owner', 'Sharing Level', 'Risk Score', 
            'Risk Factors', 'External Users', 'Folder Path', 'Size (MB)', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.PUBLIC_FILES:
          sheet.getRange(1, 1, 1, 8).setValues([[
            'Name', 'Type', 'Owner', 'Permission', 'Risk Score',
            'Folder Path', 'Size (MB)', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.EXTERNALLY_SHARED:
          sheet.getRange(1, 1, 1, 9).setValues([[
            'Name', 'Type', 'Owner', 'External Domains', 'Viewers',
            'Editors', 'Folder Path', 'Size (MB)', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
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
 * Add shared file data to the sheet
 */
function addToSharedFilesSheet(spreadsheet, fileData, riskScore) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SHARED_FILES);
  
  sheet.appendRow([
    fileData.name,
    fileData.type,
    fileData.owner,
    fileData.sharingAccess,
    fileData.sharingPermission,
    fileData.viewers.slice(0, 5).join(', '), // Limit display
    fileData.editors.slice(0, 5).join(', '),
    fileData.externalDomains.join(', '),
    riskScore,
    fileData.folderPath,
    (fileData.size / 1024 / 1024).toFixed(2),
    fileData.url
  ]);
}

/**
 * Update the shared files overview sheet
 */
function updateSharedFilesOverviewSheet(spreadsheet, stats, isFinal) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('Google Drive Shared Files Security Audit')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(isFinal ? 'Final Report' : 'Progress Report')
    .setFontSize(12);
  
  sheet.getRange(3, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary Statistics
  sheet.getRange(5, 1).setValue('SHARING SECURITY SUMMARY').setFontWeight('bold');
  
  const summaryData = [
    ['Total Shared Files:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Public Files:', stats.publicFiles],
    ['Domain Shared Files:', stats.domainSharedFiles],
    ['Externally Shared Files:', stats.externallySharedFiles],
    ['Internally Shared Files:', stats.internallySharedFiles],
    ['High Risk Files:', stats.highRiskFiles.length],
    ['External Domains:', Object.keys(stats.externalDomains).length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(6, 1, summaryData.length, 2).setValues(summaryData);
  
  // Risk Assessment
  const riskRow = 16;
  sheet.getRange(riskRow, 1).setValue('SECURITY RISK ASSESSMENT').setFontWeight('bold');
  
  let riskLevel = 'LOW';
  let riskColor = '#4CAF50';
  
  if (stats.publicFiles > 10 || stats.highRiskFiles.length > 5) {
    riskLevel = 'HIGH';
    riskColor = '#F44336';
  } else if (stats.publicFiles > 0 || stats.externallySharedFiles > 10) {
    riskLevel = 'MEDIUM';
    riskColor = '#FF9800';
  }
  
  sheet.getRange(riskRow + 1, 1).setValue(`Overall Risk Level: ${riskLevel}`)
    .setBackground(riskColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  
  // Access Levels Distribution
  const accessRow = riskRow + 4;
  sheet.getRange(accessRow, 1).setValue('SHARING ACCESS DISTRIBUTION').setFontWeight('bold');
  
  const accessEntries = Object.entries(stats.byAccessLevel)
    .sort((a, b) => b[1] - a[1]);
  
  if (accessEntries.length > 0) {
    sheet.getRange(accessRow + 1, 1, accessEntries.length, 2).setValues(accessEntries);
  }
  
  // Top External Domains
  const domainRow = accessRow + accessEntries.length + 3;
  sheet.getRange(domainRow, 1).setValue('TOP EXTERNAL DOMAINS').setFontWeight('bold');
  
  const domainEntries = Object.entries(stats.externalDomains)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (domainEntries.length > 0) {
    sheet.getRange(domainRow + 1, 1, domainEntries.length, 2).setValues(domainEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final shared files reports
 */
function generateSharedFilesFinalReports(spreadsheet, stats) {
  console.log("Generating final shared files reports...");
  
  updateSharedFilesOverviewSheet(spreadsheet, stats, true);
  generatePublicFilesReport(spreadsheet, stats);
  generateExternalSharingReport(spreadsheet, stats);
  generateHighRiskFilesReport(spreadsheet, stats);
  generatePermissionAnalysisReport(spreadsheet, stats);
  generateSecurityRecommendationsReport(spreadsheet, stats);
  
  console.log(`Shared files audit reports generated! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate public files report
 */
function generatePublicFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PUBLIC_FILES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  // Filter for public files from high risk files
  const publicFiles = stats.highRiskFiles.filter(file => 
    file.sharingAccess === 'ANYONE' || file.sharingAccess === 'ANYONE_WITH_LINK'
  );
  
  if (publicFiles.length > 0) {
    const data = publicFiles.map(file => [
      file.name,
      file.type,
      file.owner,
      file.sharingPermission,
      file.riskScore,
      file.path,
      (file.size / 1024 / 1024).toFixed(2),
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 8).setValues(data);
  }
}

/**
 * Generate external sharing report
 */
function generateExternalSharingReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.EXTERNALLY_SHARED);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  const externalFiles = stats.highRiskFiles.filter(file => 
    file.externalDomains && file.externalDomains.length > 0
  );
  
  if (externalFiles.length > 0) {
    const data = externalFiles.map(file => [
      file.name,
      file.type,
      file.owner,
      file.externalDomains.join(', '),
      file.viewerCount,
      file.editorCount,
      file.path,
      (file.size / 1024 / 1024).toFixed(2),
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 9).setValues(data);
  }
}

/**
 * Generate high risk files report
 */
function generateHighRiskFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.HIGH_RISK);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.highRiskFiles.length > 0) {
    const data = stats.highRiskFiles.map(file => [
      file.name,
      file.type,
      file.owner,
      file.sharingAccess,
      file.riskScore,
      file.riskFactors.join('; '),
      file.viewerCount + file.editorCount,
      file.path,
      (file.size / 1024 / 1024).toFixed(2),
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 10).setValues(data);
  }
}

/**
 * Generate permission analysis report
 */
function generatePermissionAnalysisReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_PERMISSION);
  
  sheet.clear();
  sheet.getRange(1, 1, 1, 4).setValues([[
    'Permission Level', 'File Count', 'Percentage', 'Risk Level'
  ]]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  const total = stats.totalFiles;
  const permissionData = Object.entries(stats.byPermission)
    .sort((a, b) => b[1] - a[1])
    .map(([permission, count]) => {
      const percentage = ((count / total) * 100).toFixed(1) + '%';
      const riskLevel = permission === 'EDIT' ? 'High' : 'Medium';
      return [permission, count, percentage, riskLevel];
    });
  
  if (permissionData.length > 0) {
    sheet.getRange(2, 1, permissionData.length, 4).setValues(permissionData);
  }
}

/**
 * Generate security recommendations report
 */
function generateSecurityRecommendationsReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.RECOMMENDATIONS);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('SECURITY RECOMMENDATIONS')
    .setFontSize(14).setFontWeight('bold');
  
  let currentRow = 3;
  
  const recommendations = generateSecurityRecommendations(stats);
  
  recommendations.forEach(rec => {
    sheet.getRange(currentRow, 1).setValue(rec.priority).setFontWeight('bold');
    sheet.getRange(currentRow, 2).setValue(rec.title).setFontWeight('bold');
    sheet.getRange(currentRow + 1, 2).setValue(rec.description);
    
    if (rec.action) {
      sheet.getRange(currentRow + 2, 2).setValue(`Action: ${rec.action}`);
    }
    
    currentRow += 4;
  });
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate security recommendations
 */
function generateSecurityRecommendations(stats) {
  const recommendations = [];
  
  // Public files recommendation
  if (stats.publicFiles > 0) {
    recommendations.push({
      priority: '🔴 URGENT',
      title: 'Review Public Files',
      description: `You have ${stats.publicFiles} files that are publicly accessible. This poses significant security risks.`,
      action: 'Review each public file and change sharing to private or restricted access.'
    });
  }
  
  // High risk files recommendation
  if (stats.highRiskFiles.length > 0) {
    recommendations.push({
      priority: '🟠 HIGH',
      title: 'Address High Risk Files',
      description: `${stats.highRiskFiles.length} files have been flagged as high risk due to their sharing settings and content.`,
      action: 'Review the High Risk Files sheet and adjust permissions accordingly.'
    });
  }
  
  // External sharing recommendation
  if (stats.externallySharedFiles > 10) {
    recommendations.push({
      priority: '🟡 MEDIUM',
      title: 'Review External Sharing',
      description: `${stats.externallySharedFiles} files are shared with external users. Ensure this is necessary and appropriate.`,
      action: 'Audit external sharing and implement regular access reviews.'
    });
  }
  
  // Domain policy recommendation
  if (stats.domainSharedFiles > stats.internallySharedFiles) {
    recommendations.push({
      priority: '🟡 MEDIUM',
      title: 'Review Domain Sharing Policy',
      description: 'Many files are shared at the domain level. Consider implementing more restrictive default sharing.',
      action: 'Update organizational sharing policies to be more restrictive by default.'
    });
  }
  
  // Regular audit recommendation
  recommendations.push({
    priority: '🔵 INFO',
    title: 'Implement Regular Security Audits',
    description: 'Schedule regular sharing audits to maintain security hygiene.',
    action: 'Set up monthly or quarterly sharing reviews using this script.'
  });
  
  return recommendations;
}

// Utility functions

function scheduleNextSharedFilesRun() {
  cancelSharedFilesScheduledRuns();
  
  ScriptApp.newTrigger('continueSharedFilesAudit')
    .timeBased()
    .after(1 * 60 * 1000)
    .create();
  
  console.log("Next shared files audit scheduled for 1 minute from now");
}

function continueSharedFilesAudit() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (scriptProperties.getProperty('sharedFilesAutoMode') !== 'true') {
    console.log("Shared files auto mode disabled, stopping.");
    cancelSharedFilesScheduledRuns();
    return;
  }
  
  auditSharedFiles();
}

function cancelSharedFilesScheduledRuns() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueSharedFilesAudit' ||
        trigger.getHandlerFunction() === 'auditSharedFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  console.log("Cancelled all scheduled shared files audit runs");
}

function updateSharedFilesStatus(spreadsheet, status, stats) {
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
    'jpg': 'Image', 'jpeg': 'Image', 'png': 'Image',
    'mp4': 'Video', 'avi': 'Video', 'mov': 'Video',
    'mp3': 'Audio', 'wav': 'Audio',
    'zip': 'Archive', 'rar': 'Archive'
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
    console.log(`Created new shared files audit spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Start automatic shared files audit
 */
function startAutomaticSharedFilesAudit() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const continuationToken = scriptProperties.getProperty('sharedFilesContinuationToken');
  if (!continuationToken) {
    resetSharedFilesAudit();
  }
  
  scriptProperties.setProperty('sharedFilesAutoMode', 'true');
  
  console.log("Starting automatic shared files audit...");
  auditSharedFiles();
}

/**
 * Reset shared files audit
 */
function resetSharedFilesAudit() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('sharedFilesContinuationToken');
  scriptProperties.deleteProperty('sharedFilesStats');
  scriptProperties.deleteProperty('sharedFilesAutoMode');
  console.log("Shared files audit reset. Next run will start from the beginning.");
}

/**
 * Quick shared files security check
 */
function getQuickSecurityCheck() {
  console.log("Running quick security check...");
  
  const stats = {
    totalSharedFiles: 0,
    publicFiles: 0,
    externallyShared: 0,
    highRiskFound: 0,
    externalDomains: new Set()
  };
  
  // Quick search for public files
  let publicFiles = DriveApp.searchFiles('visibility = "anyoneCanFind" and trashed = false');
  while (publicFiles.hasNext()) {
    publicFiles.next();
    stats.publicFiles++;
  }
  
  // Quick search for domain shared files
  let domainFiles = DriveApp.searchFiles('visibility = "domainCanFind" and trashed = false');
  let count = 0;
  const maxCheck = 100; // Limit for quick check
  
  while (domainFiles.hasNext() && count < maxCheck) {
    const file = domainFiles.next();
    count++;
    stats.totalSharedFiles++;
    
    // Check for external sharing
    try {
      const viewers = file.getViewers();
      const editors = file.getEditors();
      const owner = file.getOwner();
      const ownerDomain = owner ? owner.getEmail().split('@')[1] : '';
      
      [...viewers, ...editors].forEach(user => {
        const userDomain = user.getEmail().split('@')[1];
        if (userDomain !== ownerDomain) {
          stats.externallyShared++;
          stats.externalDomains.add(userDomain);
        }
      });
      
      // Quick risk check
      const fileName = file.getName().toLowerCase();
      if (CONFIG.HIGH_RISK_KEYWORDS.some(keyword => fileName.includes(keyword))) {
        stats.highRiskFound++;
      }
      
    } catch (error) {
      // Skip files with permission errors
    }
  }
  
  console.log('Quick Security Check Results:');
  console.log(`Total shared files checked: ${stats.totalSharedFiles}`);
  console.log(`Public files found: ${stats.publicFiles}`);
  console.log(`Externally shared files: ${stats.externallyShared}`);
  console.log(`High-risk files found: ${stats.highRiskFound}`);
  console.log(`External domains: ${stats.externalDomains.size}`);
  
  // Risk assessment
  let overallRisk = 'LOW';
  if (stats.publicFiles > 5 || stats.highRiskFound > 2) {
    overallRisk = 'HIGH';
  } else if (stats.publicFiles > 0 || stats.externallyShared > 10) {
    overallRisk = 'MEDIUM';
  }
  
  console.log(`Overall Risk Level: ${overallRisk}`);
  
  return stats;
}