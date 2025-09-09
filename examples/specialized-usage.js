/**
 * Specialized Scripts Usage Examples
 * 
 * This file shows how to use each specialized inventory script
 * Copy the relevant functions to your Google Apps Script project
 */

// ========================================
// MARKDOWN FILES INVENTORY
// ========================================

/**
 * Example: Analyze all Markdown files
 * Perfect for documentation analysis and README file organization
 */
function runMarkdownAnalysis() {
  console.log("=== MARKDOWN FILES ANALYSIS ===");
  
  // Quick overview first
  console.log("Getting quick Markdown overview...");
  const quickStats = getQuickMarkdownStats();
  
  if (quickStats.totalMarkdown === 0) {
    console.log("No Markdown files found in your Drive");
    return;
  }
  
  console.log(`Found ${quickStats.totalMarkdown} Markdown files`);
  console.log(`README files: ${quickStats.readmeCount}`);
  
  // Run full inventory
  console.log("Starting full Markdown inventory...");
  inventoryMarkdownFiles();
}

/**
 * Example: Find all README files across projects
 */
function findAllReadmeFiles() {
  // This would be part of the markdown inventory results
  // Check the "README Files" sheet in the generated spreadsheet
  console.log("Run inventoryMarkdownFiles() first, then check the 'README Files' sheet");
}

// ========================================
// CODE FILES INVENTORY  
// ========================================

/**
 * Example: Analyze all programming/code files
 * Perfect for developers tracking code repositories and projects
 */
function runCodeAnalysis() {
  console.log("=== CODE FILES ANALYSIS ===");
  
  // Quick overview first
  console.log("Getting quick code overview...");
  const quickStats = getQuickCodeStats();
  
  if (quickStats.totalCode === 0) {
    console.log("No code files found in your Drive");
    return;
  }
  
  console.log(`Found ${quickStats.totalCode} code files`);
  console.log("Top languages:", Object.entries(quickStats.languages)
    .sort((a,b) => b[1] - a[1])
    .slice(0, 3)
    .map(([lang, count]) => `${lang}: ${count}`)
    .join(', '));
  
  // Run full inventory
  console.log("Starting full code inventory...");
  inventoryCodeFiles();
}

/**
 * Example: Repository analysis
 */
function analyzeRepositories() {
  console.log("Run inventoryCodeFiles() first, then check the 'Repository Analysis' sheet");
  console.log("This will show you:");
  console.log("- Files per repository/project");
  console.log("- Programming languages used");
  console.log("- Last activity dates");
  console.log("- Whether projects have configuration files");
}

// ========================================
// DOCUMENT ANALYSIS (FIXED VERSION)
// ========================================

/**
 * Example: Document analysis with error handling
 * Use this if you encountered the search query error
 */
function runDocumentAnalysisSafe() {
  console.log("=== DOCUMENT ANALYSIS (SAFE VERSION) ===");
  
  try {
    // Quick test first
    const quickStats = getQuickDocumentStats();
    console.log(`Found ${quickStats.totalDocuments} documents`);
    console.log(`Google Workspace: ${quickStats.googleWorkspace}, Other: ${quickStats.otherDocs}`);
    
    // Run full inventory
    inventoryDocuments();
    
  } catch (error) {
    console.error("Error running document analysis:", error);
    console.log("Try using the document-inventory-fixed.js script instead");
  }
}

// ========================================
// IMAGE FILES ANALYSIS
// ========================================

/**
 * Example: Comprehensive image analysis
 * Perfect for photo organization and storage optimization
 */
function runImageAnalysis() {
  console.log("=== IMAGE FILES ANALYSIS ===");
  
  // Quick overview
  const quickStats = getQuickImageStats();
  console.log(`Found ${quickStats.totalImages} images`);
  
  if (quickStats.formatCounts) {
    console.log("Top formats:");
    Object.entries(quickStats.formatCounts)
      .sort((a,b) => b[1] - a[1])
      .slice(0, 5)
      .forEach(([format, count]) => {
        console.log(`  ${format}: ${count}`);
      });
  }
  
  // Run full analysis
  startAutomaticImageInventory();
}

// ========================================
// LARGE FILES ANALYSIS
// ========================================

/**
 * Example: Storage optimization analysis
 * Perfect for cleaning up Drive space
 */
function runStorageOptimization() {
  console.log("=== STORAGE OPTIMIZATION ANALYSIS ===");
  
  // Quick overview
  const quickStats = getQuickLargeFilesStats();
  console.log(`Found ${quickStats.totalLargeFiles} large files`);
  console.log(`Total size: ${quickStats.totalSize}`);
  
  if (quickStats.largestFile) {
    console.log(`Largest file: ${quickStats.largestFile.name} (${quickStats.largestFile.sizeFormatted})`);
  }
  
  // Run full analysis with cleanup recommendations
  startAutomaticLargeFilesAnalysis();
}

// ========================================
// SECURITY AUDIT
// ========================================

/**
 * Example: Comprehensive sharing security audit
 * Essential for compliance and data protection
 */
function runSecurityAudit() {
  console.log("=== SECURITY AUDIT ===");
  
  // Quick security check first
  const securityCheck = getQuickSecurityCheck();
  
  console.log(`Public files found: ${securityCheck.publicFiles}`);
  console.log(`Externally shared files: ${securityCheck.externallyShared}`);
  console.log(`High-risk files: ${securityCheck.highRiskFound}`);
  console.log(`External domains: ${securityCheck.externalDomains.size}`);
  
  if (securityCheck.publicFiles > 0) {
    console.log("⚠️  WARNING: Public files detected! Review immediately.");
  }
  
  // Run full audit
  startAutomaticSharedFilesAudit();
}

// ========================================
// COMBINED ANALYSIS WORKFLOWS
// ========================================

/**
 * Example: Complete developer workflow
 * Analyzes code, documentation, and project organization
 */
function runDeveloperWorkflow() {
  console.log("=== DEVELOPER COMPLETE ANALYSIS ===");
  
  console.log("1. Analyzing code files...");
  // Copy code from getQuickCodeStats() and inventoryCodeFiles()
  
  console.log("2. Analyzing documentation (Markdown)...");  
  // Copy code from getQuickMarkdownStats() and inventoryMarkdownFiles()
  
  console.log("3. Analyzing project documents...");
  // Copy code from getQuickDocumentStats() and inventoryDocuments()
  
  console.log("Run each analysis separately, then compare results across spreadsheets");
}

/**
 * Example: Content creator workflow
 * Focuses on media, images, and large files
 */
function runContentCreatorWorkflow() {
  console.log("=== CONTENT CREATOR ANALYSIS ===");
  
  console.log("1. Analyzing images and graphics...");
  // Copy from image inventory functions
  
  console.log("2. Finding large files for storage optimization...");
  // Copy from large files analysis
  
  console.log("3. Checking sharing permissions for published content...");
  // Copy from security audit functions
}

/**
 * Example: Business/organizational workflow
 * Focuses on documents, sharing, and compliance
 */
function runBusinessWorkflow() {
  console.log("=== BUSINESS ORGANIZATION ANALYSIS ===");
  
  console.log("1. Analyzing all documents...");
  // Document inventory
  
  console.log("2. Security and sharing audit...");
  // Shared files audit
  
  console.log("3. Storage cost optimization...");
  // Large files analysis
}

// ========================================
// AUTOMATION EXAMPLES
// ========================================

/**
 * Example: Set up weekly automated analysis
 */
function setupWeeklyAnalysis() {
  console.log("Setting up weekly automation...");
  
  // You would set up triggers in Google Apps Script:
  // 1. Go to Triggers in the Apps Script editor
  // 2. Add trigger for runWeeklyAnalysis()
  // 3. Set to weekly schedule
  
  console.log("Trigger setup required in Google Apps Script interface");
}

function runWeeklyAnalysis() {
  console.log("Running weekly automated analysis...");
  
  // Run security audit (most important)
  getQuickSecurityCheck();
  
  // Check for large files (storage management)
  getQuickLargeFilesStats();
  
  // Send summary email (you'd implement this)
  // sendWeeklySummaryEmail();
}

/**
 * Example: Monthly comprehensive review
 */
function runMonthlyReview() {
  console.log("Running monthly comprehensive review...");
  
  // Run all quick stats
  console.log("=== MONTHLY DRIVE REVIEW ===");
  
  try { getQuickDocumentStats(); } catch(e) { console.log("Documents: Error"); }
  try { getQuickImageStats(); } catch(e) { console.log("Images: Error"); }  
  try { getQuickMarkdownStats(); } catch(e) { console.log("Markdown: Error"); }
  try { getQuickCodeStats(); } catch(e) { console.log("Code: Error"); }
  try { getQuickLargeFilesStats(); } catch(e) { console.log("Large files: Error"); }
  try { getQuickSecurityCheck(); } catch(e) { console.log("Security: Error"); }
  
  console.log("Review complete. Run full inventories for areas of concern.");
}

// ========================================
// USAGE NOTES
// ========================================

/**
 * IMPORTANT USAGE NOTES:
 * 
 * 1. ALWAYS run the quick stats functions first to see what you have
 * 2. Copy only the specific specialized script you need
 * 3. The functions referenced here need to be copied from the appropriate specialized scripts
 * 4. Each specialized script is standalone - you don't need all of them
 * 5. Start with small batches if you have a large Drive
 * 
 * RECOMMENDED ORDER FOR NEW USERS:
 * 1. getQuickDocumentStats() - see your document overview
 * 2. getQuickSecurityCheck() - check for security issues
 * 3. getQuickLargeFilesStats() - see storage usage
 * 4. Run full inventories based on your needs
 * 
 * TROUBLESHOOTING:
 * - If you get "Invalid argument: q" errors, use the -fixed versions
 * - If scripts timeout, reduce BATCH_SIZE in the CONFIG
 * - If you see permission errors, the script may not have access to some files
 */