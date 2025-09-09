/**
 * Google Drive Markdown Files Inventory Script
 * Specialized script for analyzing Markdown (.md) files in Google Drive
 * Perfect for documentation, README files, and technical writing analysis
 */

// Configuration for markdown analysis
const CONFIG = {
  BATCH_SIZE: 100,
  INVENTORY_SPREADSHEET_NAME: "📝 Markdown Files Inventory Report",
  
  SHEETS: {
    OVERVIEW: "Overview",
    MARKDOWN_LIST: "All Markdown Files",
    LARGE_FILES: "Large Markdown Files",
    OLD_FILES: "Old Markdown Files",
    BY_PROJECT: "By Project/Folder",
    README_FILES: "README Files",
    DOCUMENTATION: "Documentation Files",
    ORPHANED: "Orphaned Files"
  },
  
  LARGE_FILE_THRESHOLD_MB: 1, // Smaller threshold for markdown
  OLD_FILE_THRESHOLD_DAYS: 180, // 6 months for markdown
  INCLUDE_TRASHED: false,
  TRACK_PERMISSIONS: true,
  
  // Markdown-specific settings
  MARKDOWN_EXTENSIONS: ['md', 'markdown', 'mdown', 'mkd', 'mkdn'],
  DOCUMENTATION_KEYWORDS: [
    'readme', 'documentation', 'docs', 'guide', 'manual', 'tutorial',
    'howto', 'faq', 'api', 'changelog', 'license', 'contributing',
    'install', 'setup', 'getting-started', 'quickstart'
  ],
  
  // Project identification patterns
  PROJECT_INDICATORS: [
    'readme.md', 'package.json', '.git', 'dockerfile', 'requirements.txt',
    'composer.json', 'cargo.toml', 'pom.xml', 'build.gradle'
  ]
};

/**
 * Main function to inventory markdown files
 */
function inventoryMarkdownFiles() {
  console.log("Starting Markdown files inventory...");
  
  try {
    const result = runMarkdownInventoryBatch();
    console.log(`Processed ${result.processedCount} Markdown files`);
    
    if (result.hasMore) {
      console.log("More Markdown files to process. Run inventoryMarkdownFiles() again to continue.");
    } else {
      console.log("Markdown inventory complete!");
    }
  } catch (error) {
    console.error(`Error in Markdown inventory: ${error}`);
    console.error(`Error details: ${error.stack}`);
  }
}

/**
 * Process a batch of markdown files
 */
function runMarkdownInventoryBatch() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeMarkdownSheets(spreadsheet);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('markdownContinuationToken');
  
  let stats = JSON.parse(scriptProperties.getProperty('markdownInventoryStats') || '{}');
  stats = initializeMarkdownStats(stats);
  
  const files = getMarkdownFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more Markdown files found!");
    generateMarkdownFinalReports(spreadsheet, stats);
    
    // Clean up
    scriptProperties.deleteProperty('markdownContinuationToken');
    scriptProperties.deleteProperty('markdownInventoryStats');
    
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing ${files.length} Markdown files...`);
  
  let processedCount = 0;
  for (const file of files) {
    try {
      if (isMarkdownFile(file)) {
        processMarkdownFile(file, spreadsheet, stats);
        processedCount++;
      }
    } catch (error) {
      console.error(`Error processing file ${file.getName()}: ${error}`);
      stats.errors = (stats.errors || 0) + 1;
    }
  }
  
  // Save progress
  scriptProperties.setProperty('markdownInventoryStats', JSON.stringify(stats));
  updateMarkdownOverviewSheet(spreadsheet, stats);
  
  const hasMore = files.hasNext && files.hasNext();
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('markdownContinuationToken', nextToken);
  } else {
    generateMarkdownFinalReports(spreadsheet, stats);
    scriptProperties.deleteProperty('markdownContinuationToken');
    scriptProperties.deleteProperty('markdownInventoryStats');
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Get markdown files to process
 */
function getMarkdownFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  let checkedCount = 0;
  const maxCheck = batchSize * 20; // Check more files to find markdown files
  
  while (files.hasNext() && count < batchSize && checkedCount < maxCheck) {
    const file = files.next();
    checkedCount++;
    
    if (isMarkdownFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  console.log(`Found ${count} Markdown files after checking ${checkedCount} files`);
  
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if file is a markdown file
 */
function isMarkdownFile(file) {
  try {
    const fileName = file.getName().toLowerCase();
    const extension = fileName.split('.').pop();
    
    // Check by extension
    if (CONFIG.MARKDOWN_EXTENSIONS.includes(extension)) {
      return true;
    }
    
    // Check by MIME type (some markdown files may have text/plain)
    const mimeType = file.getMimeType();
    if ((mimeType === 'text/plain' || mimeType === 'text/markdown') && 
        fileName.includes('.md')) {
      return true;
    }
    
    return false;
  } catch (error) {
    console.error(`Error checking if file is markdown: ${error}`);
    return false;
  }
}

/**
 * Process a single markdown file
 */
function processMarkdownFile(file, spreadsheet, stats) {
  try {
    const fileData = extractMarkdownFileData(file);
    
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Categorize markdown file
    const category = categorizeMarkdownFile(fileData);
    stats.byCategory[category] = (stats.byCategory[category] || 0) + 1;
    
    // Track by project/folder
    if (fileData.folderPath) {
      const project = extractProjectName(fileData.folderPath);
      stats.byProject[project] = (stats.byProject[project] || 0) + 1;
    }
    
    // Track by owner
    stats.byOwner[fileData.owner] = (stats.byOwner[fileData.owner] || 0) + 1;
    
    // Track by year
    const year = new Date(fileData.lastModified).getFullYear();
    stats.byYear[year] = (stats.byYear[year] || 0) + 1;
    
    // Check for large files
    if (fileData.size > CONFIG.LARGE_FILE_THRESHOLD_MB * 1024 * 1024) {
      stats.largeFiles.push({
        name: fileData.name,
        size: fileData.size,
        category: category,
        path: fileData.folderPath,
        url: fileData.url,
        lastModified: fileData.lastModified
      });
      
      stats.largeFiles.sort((a, b) => b.size - a.size);
      stats.largeFiles = stats.largeFiles.slice(0, 50);
    }
    
    // Check for old files
    const ageInDays = (new Date() - new Date(fileData.lastModified)) / (1000 * 60 * 60 * 24);
    if (ageInDays > CONFIG.OLD_FILE_THRESHOLD_DAYS) {
      if (stats.oldFiles.length < 50) {
        stats.oldFiles.push({
          name: fileData.name,
          lastModified: fileData.lastModified,
          ageInDays: Math.floor(ageInDays),
          category: category,
          path: fileData.folderPath,
          url: fileData.url
        });
      }
    }
    
    // Special handling for README files
    if (fileData.name.toLowerCase().includes('readme')) {
      stats.readmeFiles.push({
        name: fileData.name,
        path: fileData.folderPath,
        project: extractProjectName(fileData.folderPath),
        size: fileData.size,
        lastModified: fileData.lastModified,
        url: fileData.url
      });
    }
    
    // Check if orphaned (not in a clear project structure)
    if (isOrphanedFile(fileData)) {
      stats.orphanedFiles.push({
        name: fileData.name,
        path: fileData.folderPath,
        lastModified: fileData.lastModified,
        url: fileData.url
      });
    }
    
    // Add to main list
    addToMarkdownListSheet(spreadsheet, fileData, category);
    
  } catch (error) {
    console.error(`Error processing markdown file: ${error}`);
    throw error;
  }
}

/**
 * Extract markdown-specific file data
 */
function extractMarkdownFileData(file) {
  const data = {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    extension: file.getName().split('.').pop().toLowerCase(),
    size: file.getSize(),
    created: file.getDateCreated().toISOString(),
    lastModified: file.getLastUpdated().toISOString(),
    owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
    url: file.getUrl(),
    description: file.getDescription() || '',
    folderPath: 'Root',
    sharingAccess: 'Private'
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
      while (depth < 3) { // Limit depth for markdown analysis
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
  
  // Get sharing info
  try {
    const access = file.getSharingAccess();
    data.sharingAccess = access.toString();
  } catch (error) {
    // Some files may not have accessible permissions
  }
  
  return data;
}

/**
 * Categorize markdown file by purpose
 */
function categorizeMarkdownFile(fileData) {
  const fileName = fileData.name.toLowerCase();
  
  if (fileName.includes('readme')) {
    return 'README';
  }
  
  for (const keyword of CONFIG.DOCUMENTATION_KEYWORDS) {
    if (fileName.includes(keyword)) {
      return 'Documentation';
    }
  }
  
  if (fileName.includes('changelog') || fileName.includes('history')) {
    return 'Changelog';
  }
  
  if (fileName.includes('license') || fileName.includes('copyright')) {
    return 'License';
  }
  
  if (fileName.includes('todo') || fileName.includes('task')) {
    return 'TODO/Tasks';
  }
  
  if (fileName.includes('note') || fileName.includes('memo')) {
    return 'Notes';
  }
  
  return 'General';
}

/**
 * Extract project name from folder path
 */
function extractProjectName(folderPath) {
  if (!folderPath || folderPath === 'Root' || folderPath === 'Unknown') {
    return 'Uncategorized';
  }
  
  // Take the first folder as project name, or last folder if it looks like a project
  const parts = folderPath.split('/');
  
  // Look for common project indicators
  for (const part of parts) {
    const lowerPart = part.toLowerCase();
    if (lowerPart.includes('project') || 
        lowerPart.includes('repo') || 
        lowerPart.includes('code') ||
        lowerPart.includes('dev')) {
      return part;
    }
  }
  
  // Return the deepest folder (most specific)
  return parts[parts.length - 1] || 'Uncategorized';
}

/**
 * Check if file appears to be orphaned (not in project structure)
 */
function isOrphanedFile(fileData) {
  const path = fileData.folderPath.toLowerCase();
  const name = fileData.name.toLowerCase();
  
  // If it's just sitting in root or a generic folder
  if (path === 'root' || path === 'unknown') {
    return true;
  }
  
  // If it's not a README and doesn't have clear project context
  if (!name.includes('readme') && 
      !path.includes('project') && 
      !path.includes('repo') && 
      !path.includes('code') &&
      !path.includes('dev') &&
      !path.includes('docs')) {
    return true;
  }
  
  return false;
}

/**
 * Initialize markdown statistics
 */
function initializeMarkdownStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    byCategory: stats.byCategory || {},
    byProject: stats.byProject || {},
    byOwner: stats.byOwner || {},
    byYear: stats.byYear || {},
    largeFiles: stats.largeFiles || [],
    oldFiles: stats.oldFiles || [],
    readmeFiles: stats.readmeFiles || [],
    orphanedFiles: stats.orphanedFiles || [],
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Initialize markdown-specific sheets
 */
function initializeMarkdownSheets(spreadsheet) {
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      switch (sheetName) {
        case CONFIG.SHEETS.MARKDOWN_LIST:
          sheet.getRange(1, 1, 1, 10).setValues([[
            'Name', 'Extension', 'Category', 'Size (KB)', 'Project/Folder', 
            'Last Modified', 'Owner', 'Folder Path', 'Sharing', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.README_FILES:
          sheet.getRange(1, 1, 1, 6).setValues([[
            'README File', 'Project', 'Size (KB)', 'Last Modified', 'Path', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.ORPHANED:
          sheet.getRange(1, 1, 1, 4).setValues([[
            'File Name', 'Path', 'Last Modified', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
          break;
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
 * Add markdown file to main list sheet
 */
function addToMarkdownListSheet(spreadsheet, fileData, category) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.MARKDOWN_LIST);
  
  sheet.appendRow([
    fileData.name,
    fileData.extension,
    category,
    (fileData.size / 1024).toFixed(2), // KB for markdown files
    extractProjectName(fileData.folderPath),
    fileData.lastModified,
    fileData.owner,
    fileData.folderPath,
    fileData.sharingAccess,
    fileData.url
  ]);
}

/**
 * Update markdown overview sheet
 */
function updateMarkdownOverviewSheet(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('Markdown Files Inventory')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary statistics
  const summaryData = [
    ['Total Markdown Files:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Average Size:', formatBytes(stats.totalSize / Math.max(stats.totalFiles, 1))],
    ['README Files:', stats.readmeFiles.length],
    ['Large Files (>' + CONFIG.LARGE_FILE_THRESHOLD_MB + 'MB):', stats.largeFiles.length],
    ['Old Files (>' + CONFIG.OLD_FILE_THRESHOLD_DAYS + ' days):', stats.oldFiles.length],
    ['Orphaned Files:', stats.orphanedFiles.length],
    ['Unique Projects:', Object.keys(stats.byProject).length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(4, 1, summaryData.length, 2).setValues(summaryData);
  
  // Categories breakdown
  sheet.getRange(15, 1).setValue('BY CATEGORY').setFontWeight('bold');
  
  const categoryEntries = Object.entries(stats.byCategory)
    .sort((a, b) => b[1] - a[1]);
  
  if (categoryEntries.length > 0) {
    sheet.getRange(16, 1, categoryEntries.length, 2).setValues(categoryEntries);
  }
  
  // Projects breakdown
  const projectRow = 16 + categoryEntries.length + 2;
  sheet.getRange(projectRow, 1).setValue('TOP PROJECTS').setFontWeight('bold');
  
  const projectEntries = Object.entries(stats.byProject)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (projectEntries.length > 0) {
    sheet.getRange(projectRow + 1, 1, projectEntries.length, 2).setValues(projectEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final markdown reports
 */
function generateMarkdownFinalReports(spreadsheet, stats) {
  console.log("Generating final Markdown reports...");
  
  updateMarkdownOverviewSheet(spreadsheet, stats);
  generateReadmeFilesReport(spreadsheet, stats);
  generateOrphanedFilesReport(spreadsheet, stats);
  generateProjectAnalysisReport(spreadsheet, stats);
  
  console.log(`Markdown inventory complete! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate README files report
 */
function generateReadmeFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.README_FILES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.readmeFiles.length > 0) {
    const data = stats.readmeFiles.map(file => [
      file.name,
      file.project,
      (file.size / 1024).toFixed(2),
      file.lastModified,
      file.path,
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
  }
}

/**
 * Generate orphaned files report
 */
function generateOrphanedFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ORPHANED);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  if (stats.orphanedFiles.length > 0) {
    const data = stats.orphanedFiles.map(file => [
      file.name,
      file.path,
      file.lastModified,
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 4).setValues(data);
  }
}

/**
 * Generate project analysis report
 */
function generateProjectAnalysisReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_PROJECT);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 3).setValues([['Project/Folder', 'File Count', 'Percentage']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  
  const total = stats.totalFiles;
  const projectData = Object.entries(stats.byProject)
    .sort((a, b) => b[1] - a[1])
    .map(([project, count]) => [
      project,
      count,
      ((count / total) * 100).toFixed(1) + '%'
    ]);
  
  if (projectData.length > 0) {
    sheet.getRange(2, 1, projectData.length, 3).setValues(projectData);
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
    console.log(`Created new Markdown inventory spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Quick markdown stats for testing
 */
function getQuickMarkdownStats() {
  console.log("Getting quick Markdown stats...");
  
  const stats = {
    totalMarkdown: 0,
    readmeCount: 0,
    categories: {},
    projects: {}
  };
  
  const files = DriveApp.searchFiles('trashed = false');
  let count = 0;
  const maxCheck = 500; // Check more files since markdown files are less common
  
  while (files.hasNext() && count < maxCheck) {
    const file = files.next();
    
    if (isMarkdownFile(file)) {
      stats.totalMarkdown++;
      
      const fileName = file.getName().toLowerCase();
      if (fileName.includes('readme')) {
        stats.readmeCount++;
      }
      
      // Try to get folder info
      try {
        const parents = file.getParents();
        if (parents.hasNext()) {
          const parent = parents.next();
          const project = parent.getName();
          stats.projects[project] = (stats.projects[project] || 0) + 1;
        }
      } catch (error) {
        // Skip folder analysis if error
      }
      
      const category = categorizeMarkdownFile({ name: file.getName() });
      stats.categories[category] = (stats.categories[category] || 0) + 1;
    }
    count++;
  }
  
  console.log(`Found ${stats.totalMarkdown} Markdown files in first ${count} files`);
  console.log(`README files: ${stats.readmeCount}`);
  console.log('Categories:', stats.categories);
  console.log('Top projects:', Object.entries(stats.projects).sort((a,b) => b[1] - a[1]).slice(0, 5));
  
  return stats;
}

/**
 * USAGE INSTRUCTIONS:
 * 
 * 1. Copy this entire script to Google Apps Script
 * 2. For testing: run getQuickMarkdownStats() first
 * 3. Run: inventoryMarkdownFiles() for full analysis
 * 4. Check the generated spreadsheet for detailed results
 * 
 * Perfect for analyzing:
 * - Documentation files across projects
 * - README file organization
 * - Orphaned markdown files that need organizing
 * - Project documentation completeness
 */