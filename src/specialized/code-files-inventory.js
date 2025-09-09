/**
 * Google Drive Code Files Inventory Script
 * Specialized script for analyzing programming and code files in Google Drive
 * Perfect for developers tracking code repositories, scripts, and project files
 */

// Configuration for code files analysis
const CONFIG = {
  BATCH_SIZE: 100,
  INVENTORY_SPREADSHEET_NAME: "💻 Code Files Inventory Report",
  
  SHEETS: {
    OVERVIEW: "Overview",
    CODE_FILES: "All Code Files",
    BY_LANGUAGE: "By Programming Language",
    BY_PROJECT: "By Project/Repository",
    LARGE_FILES: "Large Code Files",
    OLD_FILES: "Old Code Files",
    CONFIG_FILES: "Configuration Files",
    SCRIPTS: "Scripts & Executables",
    REPOSITORIES: "Repository Analysis"
  },
  
  LARGE_FILE_THRESHOLD_MB: 5, // Smaller threshold for code files
  OLD_FILE_THRESHOLD_DAYS: 365, // 1 year for code files
  INCLUDE_TRASHED: false,
  TRACK_PERMISSIONS: true,
  
  // Programming languages and extensions
  PROGRAMMING_LANGUAGES: {
    // Web Development
    'js': 'JavaScript', 'jsx': 'React/JSX', 'ts': 'TypeScript', 'tsx': 'TypeScript React',
    'html': 'HTML', 'htm': 'HTML', 'css': 'CSS', 'scss': 'SASS/SCSS', 'sass': 'SASS/SCSS',
    'less': 'LESS', 'vue': 'Vue.js', 'svelte': 'Svelte',
    
    // Backend Languages
    'py': 'Python', 'java': 'Java', 'c': 'C', 'cpp': 'C++', 'cc': 'C++', 'cxx': 'C++',
    'cs': 'C#', 'php': 'PHP', 'rb': 'Ruby', 'go': 'Go', 'rs': 'Rust',
    'kt': 'Kotlin', 'scala': 'Scala', 'clj': 'Clojure', 'hs': 'Haskell',
    'swift': 'Swift', 'dart': 'Dart', 'lua': 'Lua', 'r': 'R',
    
    // Mobile Development
    'm': 'Objective-C', 'mm': 'Objective-C++',
    
    // Data & Config
    'sql': 'SQL', 'json': 'JSON', 'xml': 'XML', 'yaml': 'YAML', 'yml': 'YAML',
    'toml': 'TOML', 'ini': 'INI', 'cfg': 'Config', 'conf': 'Config',
    
    // Scripting
    'sh': 'Shell Script', 'bash': 'Bash Script', 'ps1': 'PowerShell', 
    'bat': 'Batch File', 'cmd': 'Command File',
    
    // Markup & Documentation
    'md': 'Markdown', 'rst': 'reStructuredText', 'tex': 'LaTeX',
    
    // Other
    'dockerfile': 'Docker', 'makefile': 'Makefile', 'cmake': 'CMake'
  },
  
  // Configuration file patterns
  CONFIG_FILES: [
    'package.json', 'composer.json', 'requirements.txt', 'gemfile', 'cargo.toml',
    'pom.xml', 'build.gradle', 'tsconfig.json', 'webpack.config.js', 'gulpfile.js',
    'dockerfile', 'docker-compose.yml', '.gitignore', '.env', '.env.example',
    'makefile', 'cmake', '.editorconfig', '.prettierrc', '.eslintrc'
  ],
  
  // Repository indicators
  REPO_INDICATORS: [
    '.git', 'package.json', 'composer.json', 'requirements.txt', 'gemfile',
    'cargo.toml', 'pom.xml', 'build.gradle', 'dockerfile', 'readme.md'
  ],
  
  // Script file patterns
  SCRIPT_PATTERNS: ['script', 'bin', 'tool', 'util', 'helper', 'run', 'build', 'deploy']
};

/**
 * Main function to inventory code files
 */
function inventoryCodeFiles() {
  console.log("Starting code files inventory...");
  
  try {
    const result = runCodeInventoryBatch();
    console.log(`Processed ${result.processedCount} code files`);
    
    if (result.hasMore) {
      console.log("More code files to process. Run inventoryCodeFiles() again to continue.");
    } else {
      console.log("Code files inventory complete!");
    }
  } catch (error) {
    console.error(`Error in code inventory: ${error}`);
    console.error(`Error details: ${error.stack}`);
  }
}

/**
 * Process a batch of code files
 */
function runCodeInventoryBatch() {
  const spreadsheet = getOrCreateSpreadsheet(CONFIG.INVENTORY_SPREADSHEET_NAME);
  initializeCodeSheets(spreadsheet);
  
  const scriptProperties = PropertiesService.getScriptProperties();
  let continuationToken = scriptProperties.getProperty('codeContinuationToken');
  
  let stats = JSON.parse(scriptProperties.getProperty('codeInventoryStats') || '{}');
  stats = initializeCodeStats(stats);
  
  const files = getCodeFilesToProcess(continuationToken, CONFIG.BATCH_SIZE);
  
  if (files.length === 0) {
    console.log("No more code files found!");
    generateCodeFinalReports(spreadsheet, stats);
    
    // Clean up
    scriptProperties.deleteProperty('codeContinuationToken');
    scriptProperties.deleteProperty('codeInventoryStats');
    
    return { processedCount: 0, hasMore: false };
  }
  
  console.log(`Processing ${files.length} code files...`);
  
  let processedCount = 0;
  for (const file of files) {
    try {
      if (isCodeFile(file)) {
        processCodeFile(file, spreadsheet, stats);
        processedCount++;
      }
    } catch (error) {
      console.error(`Error processing file ${file.getName()}: ${error}`);
      stats.errors = (stats.errors || 0) + 1;
    }
  }
  
  // Save progress
  scriptProperties.setProperty('codeInventoryStats', JSON.stringify(stats));
  updateCodeOverviewSheet(spreadsheet, stats);
  
  const hasMore = files.hasNext && files.hasNext();
  if (hasMore) {
    const nextToken = files.getContinuationToken();
    scriptProperties.setProperty('codeContinuationToken', nextToken);
  } else {
    generateCodeFinalReports(spreadsheet, stats);
    scriptProperties.deleteProperty('codeContinuationToken');
    scriptProperties.deleteProperty('codeInventoryStats');
  }
  
  return { processedCount: processedCount, hasMore: hasMore };
}

/**
 * Get code files to process
 */
function getCodeFilesToProcess(continuationToken, batchSize) {
  let files;
  
  if (continuationToken) {
    files = DriveApp.continueFileIterator(continuationToken);
  } else {
    files = DriveApp.searchFiles('trashed = false');
  }
  
  const filesToProcess = [];
  let count = 0;
  let checkedCount = 0;
  const maxCheck = batchSize * 15; // Check more files to find code files
  
  while (files.hasNext() && count < batchSize && checkedCount < maxCheck) {
    const file = files.next();
    checkedCount++;
    
    if (isCodeFile(file)) {
      filesToProcess.push(file);
      count++;
    }
  }
  
  console.log(`Found ${count} code files after checking ${checkedCount} files`);
  
  filesToProcess.hasNext = () => files.hasNext();
  filesToProcess.getContinuationToken = () => files.getContinuationToken();
  
  return filesToProcess;
}

/**
 * Check if file is a code file
 */
function isCodeFile(file) {
  try {
    const fileName = file.getName().toLowerCase();
    const extension = fileName.split('.').pop();
    const mimeType = file.getMimeType();
    
    // Check by extension first
    if (CONFIG.PROGRAMMING_LANGUAGES[extension]) {
      return true;
    }
    
    // Check for specific config files
    if (CONFIG.CONFIG_FILES.includes(fileName)) {
      return true;
    }
    
    // Check for files without extensions but with code-like names
    if (!fileName.includes('.') && (
      CONFIG.SCRIPT_PATTERNS.some(pattern => fileName.includes(pattern)) ||
      fileName === 'makefile' || fileName === 'dockerfile'
    )) {
      return true;
    }
    
    // Check MIME type for text files that might be code
    if (mimeType === 'text/plain' && (
      CONFIG.PROGRAMMING_LANGUAGES[extension] ||
      fileName.includes('script') ||
      fileName.includes('config')
    )) {
      return true;
    }
    
    return false;
  } catch (error) {
    console.error(`Error checking if file is code: ${error}`);
    return false;
  }
}

/**
 * Process a single code file
 */
function processCodeFile(file, spreadsheet, stats) {
  try {
    const fileData = extractCodeFileData(file);
    
    stats.totalFiles++;
    stats.totalSize += fileData.size;
    
    // Track by programming language
    const language = fileData.language;
    stats.byLanguage[language] = (stats.byLanguage[language] || 0) + 1;
    
    // Track by file category
    const category = fileData.category;
    stats.byCategory[category] = (stats.byCategory[category] || 0) + 1;
    
    // Track by project/repository
    if (fileData.project) {
      stats.byProject[fileData.project] = (stats.byProject[fileData.project] || 0) + 1;
      
      // Track repository info
      if (fileData.isInRepository) {
        if (!stats.repositories[fileData.project]) {
          stats.repositories[fileData.project] = {
            fileCount: 0,
            languages: new Set(),
            hasConfig: false,
            lastActivity: fileData.lastModified
          };
        }
        stats.repositories[fileData.project].fileCount++;
        stats.repositories[fileData.project].languages.add(language);
        if (fileData.category === 'Configuration') {
          stats.repositories[fileData.project].hasConfig = true;
        }
        if (fileData.lastModified > stats.repositories[fileData.project].lastActivity) {
          stats.repositories[fileData.project].lastActivity = fileData.lastModified;
        }
      }
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
        language: language,
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
          language: language,
          lastModified: fileData.lastModified,
          ageInDays: Math.floor(ageInDays),
          path: fileData.folderPath,
          url: fileData.url
        });
      }
    }
    
    // Add to main list
    addToCodeListSheet(spreadsheet, fileData);
    
  } catch (error) {
    console.error(`Error processing code file: ${error}`);
    throw error;
  }
}

/**
 * Extract code-specific file data
 */
function extractCodeFileData(file) {
  const fileName = file.getName();
  const data = {
    id: file.getId(),
    name: fileName,
    mimeType: file.getMimeType(),
    extension: fileName.split('.').pop().toLowerCase(),
    language: getLanguage(fileName),
    category: getFileCategory(fileName),
    size: file.getSize(),
    created: file.getDateCreated().toISOString(),
    lastModified: file.getLastUpdated().toISOString(),
    owner: file.getOwner() ? file.getOwner().getEmail() : 'Unknown',
    url: file.getUrl(),
    description: file.getDescription() || '',
    folderPath: 'Root',
    project: 'Uncategorized',
    isInRepository: false,
    sharingAccess: 'Private'
  };
  
  // Get folder path and project info
  try {
    const parents = file.getParents();
    const pathParts = [];
    
    if (parents.hasNext()) {
      const parent = parents.next();
      pathParts.push(parent.getName());
      
      let currentParent = parent;
      let depth = 0;
      while (depth < 4) { // Check deeper for code projects
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
    data.project = extractProjectName(data.folderPath);
    data.isInRepository = checkIfInRepository(data.folderPath, fileName);
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
 * Get programming language from filename
 */
function getLanguage(fileName) {
  const lowerName = fileName.toLowerCase();
  const extension = lowerName.split('.').pop();
  
  // Check extension first
  if (CONFIG.PROGRAMMING_LANGUAGES[extension]) {
    return CONFIG.PROGRAMMING_LANGUAGES[extension];
  }
  
  // Check for special cases
  if (lowerName === 'makefile' || lowerName.startsWith('makefile.')) {
    return 'Makefile';
  }
  
  if (lowerName === 'dockerfile' || lowerName.startsWith('dockerfile.')) {
    return 'Docker';
  }
  
  if (CONFIG.CONFIG_FILES.includes(lowerName)) {
    return 'Configuration';
  }
  
  return 'Unknown';
}

/**
 * Categorize file by purpose
 */
function getFileCategory(fileName) {
  const lowerName = fileName.toLowerCase();
  
  // Configuration files
  if (CONFIG.CONFIG_FILES.includes(lowerName) ||
      lowerName.includes('config') ||
      lowerName.includes('.env') ||
      lowerName.startsWith('.')) {
    return 'Configuration';
  }
  
  // Scripts
  if (CONFIG.SCRIPT_PATTERNS.some(pattern => lowerName.includes(pattern)) ||
      ['sh', 'bash', 'ps1', 'bat', 'cmd'].includes(lowerName.split('.').pop())) {
    return 'Script';
  }
  
  // Tests
  if (lowerName.includes('test') || lowerName.includes('spec')) {
    return 'Test';
  }
  
  // Documentation
  if (lowerName.includes('readme') || lowerName.includes('doc')) {
    return 'Documentation';
  }
  
  // Build files
  if (lowerName.includes('build') || lowerName.includes('make') ||
      ['pom.xml', 'build.gradle', 'cargo.toml'].includes(lowerName)) {
    return 'Build';
  }
  
  return 'Source Code';
}

/**
 * Extract project name from folder path
 */
function extractProjectName(folderPath) {
  if (!folderPath || folderPath === 'Root' || folderPath === 'Unknown') {
    return 'Uncategorized';
  }
  
  const parts = folderPath.split('/');
  
  // Look for repository-like folders
  for (const part of parts) {
    const lowerPart = part.toLowerCase();
    if (lowerPart.includes('project') || 
        lowerPart.includes('repo') || 
        lowerPart.includes('code') ||
        lowerPart.includes('dev') ||
        lowerPart.includes('src') ||
        lowerPart.includes('app')) {
      return part;
    }
  }
  
  // Return the first folder (likely the main project folder)
  return parts[0] || 'Uncategorized';
}

/**
 * Check if file appears to be in a repository
 */
function checkIfInRepository(folderPath, fileName) {
  const path = folderPath.toLowerCase();
  const name = fileName.toLowerCase();
  
  // Look for repository indicators in path
  return CONFIG.REPO_INDICATORS.some(indicator => 
    path.includes(indicator) || name === indicator
  ) || path.includes('src') || path.includes('lib') || path.includes('app');
}

/**
 * Initialize code statistics
 */
function initializeCodeStats(stats) {
  return {
    totalFiles: stats.totalFiles || 0,
    totalSize: stats.totalSize || 0,
    byLanguage: stats.byLanguage || {},
    byCategory: stats.byCategory || {},
    byProject: stats.byProject || {},
    byOwner: stats.byOwner || {},
    byYear: stats.byYear || {},
    largeFiles: stats.largeFiles || [],
    oldFiles: stats.oldFiles || [],
    repositories: stats.repositories || {},
    errors: stats.errors || 0,
    startTime: stats.startTime || new Date().toISOString()
  };
}

/**
 * Initialize code-specific sheets
 */
function initializeCodeSheets(spreadsheet) {
  for (const sheetName of Object.values(CONFIG.SHEETS)) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      
      switch (sheetName) {
        case CONFIG.SHEETS.CODE_FILES:
          sheet.getRange(1, 1, 1, 11).setValues([[
            'Name', 'Language', 'Category', 'Extension', 'Size (KB)', 'Project',
            'Last Modified', 'Owner', 'Folder Path', 'Sharing', 'URL'
          ]]);
          sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
          break;
          
        case CONFIG.SHEETS.REPOSITORIES:
          sheet.getRange(1, 1, 1, 5).setValues([[
            'Repository/Project', 'File Count', 'Languages', 'Last Activity', 'Has Config'
          ]]);
          sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
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
 * Add code file to main list sheet
 */
function addToCodeListSheet(spreadsheet, fileData) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CODE_FILES);
  
  sheet.appendRow([
    fileData.name,
    fileData.language,
    fileData.category,
    fileData.extension,
    (fileData.size / 1024).toFixed(2), // KB for code files
    fileData.project,
    fileData.lastModified,
    fileData.owner,
    fileData.folderPath,
    fileData.sharingAccess,
    fileData.url
  ]);
}

/**
 * Update code overview sheet
 */
function updateCodeOverviewSheet(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.OVERVIEW);
  sheet.clear();
  
  sheet.getRange(1, 1).setValue('Code Files Inventory')
    .setFontSize(16).setFontWeight('bold');
  
  sheet.getRange(2, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10);
  
  // Summary statistics
  const summaryData = [
    ['Total Code Files:', stats.totalFiles],
    ['Total Size:', formatBytes(stats.totalSize)],
    ['Programming Languages:', Object.keys(stats.byLanguage).length],
    ['Projects/Repositories:', Object.keys(stats.byProject).length],
    ['Large Files (>' + CONFIG.LARGE_FILE_THRESHOLD_MB + 'MB):', stats.largeFiles.length],
    ['Old Files (>' + CONFIG.OLD_FILE_THRESHOLD_DAYS + ' days):', stats.oldFiles.length],
    ['Active Repositories:', Object.keys(stats.repositories).length],
    ['Processing Errors:', stats.errors]
  ];
  
  sheet.getRange(4, 1, summaryData.length, 2).setValues(summaryData);
  
  // Top languages
  sheet.getRange(14, 1).setValue('TOP PROGRAMMING LANGUAGES').setFontWeight('bold');
  
  const languageEntries = Object.entries(stats.byLanguage)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  if (languageEntries.length > 0) {
    sheet.getRange(15, 1, languageEntries.length, 2).setValues(languageEntries);
  }
  
  // File categories
  const categoryRow = 15 + languageEntries.length + 2;
  sheet.getRange(categoryRow, 1).setValue('FILE CATEGORIES').setFontWeight('bold');
  
  const categoryEntries = Object.entries(stats.byCategory)
    .sort((a, b) => b[1] - a[1]);
  
  if (categoryEntries.length > 0) {
    sheet.getRange(categoryRow + 1, 1, categoryEntries.length, 2).setValues(categoryEntries);
  }
  
  sheet.autoResizeColumns(1, 2);
}

/**
 * Generate final code reports
 */
function generateCodeFinalReports(spreadsheet, stats) {
  console.log("Generating final code reports...");
  
  updateCodeOverviewSheet(spreadsheet, stats);
  generateLanguageAnalysisReport(spreadsheet, stats);
  generateRepositoryAnalysisReport(spreadsheet, stats);
  generateLargeFilesReport(spreadsheet, stats);
  
  console.log(`Code inventory complete! View at: ${spreadsheet.getUrl()}`);
}

/**
 * Generate programming language analysis
 */
function generateLanguageAnalysisReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.BY_LANGUAGE);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 3).setValues([['Programming Language', 'File Count', 'Percentage']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  
  const total = stats.totalFiles;
  const languageData = Object.entries(stats.byLanguage)
    .sort((a, b) => b[1] - a[1])
    .map(([language, count]) => [
      language,
      count,
      ((count / total) * 100).toFixed(1) + '%'
    ]);
  
  if (languageData.length > 0) {
    sheet.getRange(2, 1, languageData.length, 3).setValues(languageData);
  }
}

/**
 * Generate repository analysis report
 */
function generateRepositoryAnalysisReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.REPOSITORIES);
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  
  const repoData = Object.entries(stats.repositories)
    .sort((a, b) => b[1].fileCount - a[1].fileCount)
    .map(([repo, data]) => [
      repo,
      data.fileCount,
      Array.from(data.languages).join(', '),
      data.lastActivity,
      data.hasConfig ? 'Yes' : 'No'
    ]);
  
  if (repoData.length > 0) {
    sheet.getRange(2, 1, repoData.length, 5).setValues(repoData);
  }
}

/**
 * Generate large code files report
 */
function generateLargeFilesReport(spreadsheet, stats) {
  const sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.LARGE_FILES);
  sheet.clear();
  
  sheet.getRange(1, 1, 1, 6).setValues([['File Name', 'Size (MB)', 'Language', 'Category', 'Project', 'URL']]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  
  if (stats.largeFiles.length > 0) {
    const data = stats.largeFiles.map(file => [
      file.name,
      (file.size / 1024 / 1024).toFixed(2),
      file.language,
      file.category,
      extractProjectName(file.path),
      file.url
    ]);
    
    sheet.getRange(2, 1, data.length, 6).setValues(data);
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
    console.log(`Created new code inventory spreadsheet: ${spreadsheet.getUrl()}`);
    return spreadsheet;
  }
}

/**
 * Quick code stats for testing
 */
function getQuickCodeStats() {
  console.log("Getting quick code stats...");
  
  const stats = {
    totalCode: 0,
    languages: {},
    categories: {},
    projects: {}
  };
  
  const files = DriveApp.searchFiles('trashed = false');
  let count = 0;
  const maxCheck = 500; // Check more files since code files might be less common
  
  while (files.hasNext() && count < maxCheck) {
    const file = files.next();
    
    if (isCodeFile(file)) {
      stats.totalCode++;
      
      const language = getLanguage(file.getName());
      stats.languages[language] = (stats.languages[language] || 0) + 1;
      
      const category = getFileCategory(file.getName());
      stats.categories[category] = (stats.categories[category] || 0) + 1;
      
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
    }
    count++;
  }
  
  console.log(`Found ${stats.totalCode} code files in first ${count} files`);
  console.log('Top languages:', Object.entries(stats.languages).sort((a,b) => b[1] - a[1]).slice(0, 5));
  console.log('Categories:', stats.categories);
  console.log('Top projects:', Object.entries(stats.projects).sort((a,b) => b[1] - a[1]).slice(0, 5));
  
  return stats;
}

/**
 * USAGE INSTRUCTIONS:
 * 
 * 1. Copy this entire script to Google Apps Script
 * 2. For testing: run getQuickCodeStats() first
 * 3. Run: inventoryCodeFiles() for full analysis
 * 4. Check the generated spreadsheet for detailed results
 * 
 * Perfect for analyzing:
 * - Code repositories and project organization
 * - Programming language usage across projects
 * - Large code files that might need optimization
 * - Configuration files and project structure
 * - Repository activity and maintenance status
 */