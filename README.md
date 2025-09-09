# Google Drive Inventory Tools

A comprehensive collection of Google Apps Script tools for analyzing and managing your Google Drive files. This project provides both a complete inventory solution and specialized scripts for different file types and use cases.

## 🚀 Quick Start

1. **Complete Inventory**: Use `src/core/drive-inventory-complete.js` for comprehensive Drive analysis
2. **Specialized Tools**: Choose from `src/specialized/` for targeted file type analysis
3. **Examples**: Check `examples/` for implementation templates

## 📋 Features

### Core Inventory Script
- **Comprehensive Analysis**: Scans all files in your Google Drive
- **Multiple Reports**: Overview, large files, old files, duplicates, shared files, and file type analysis
- **Batch Processing**: Handles large drives with automatic continuation
- **Progress Tracking**: Real-time status updates and progress monitoring
- **Automatic Scheduling**: Set up hourly or continuous scanning

### Specialized Scripts
- **Image Files**: Focus on photos, graphics, and visual content analysis
- **Documents**: Word docs, PDFs, and text files with collaboration details  
- **Large Files**: Identify storage space consumers with cleanup recommendations
- **Shared Files**: Security audit with permission analysis and risk assessment
- **Markdown Files**: Documentation, README files, and technical writing analysis
- **Code Files**: Programming files with language detection and repository analysis

## 🔧 Installation

1. **Create a Google Apps Script Project**:
   - Go to [script.google.com](https://script.google.com)
   - Create a new project
   - Copy the desired script from this repository

2. **Set Up Permissions**:
   - The script will request Google Drive access
   - Authorize the necessary permissions when prompted

3. **Configure Settings**:
   - Modify the `CONFIG` object in each script for your needs
   - Adjust batch sizes, thresholds, and report preferences

## 📊 Reports Generated

### Overview Report
- Total files and storage usage
- File type distribution
- Files by creation year
- Top folders by file count

### Detailed Analysis
- **Large Files**: Files exceeding size threshold (configurable)
- **Old Files**: Files not modified recently (configurable threshold)
- **Duplicates**: Potential duplicate files by name and size
- **Shared Files**: Files with sharing permissions
- **File Types**: Comprehensive breakdown by file format

## 🎯 Use Cases

- **Storage Optimization**: Identify large files consuming space
- **Security Audit**: Find publicly shared or over-shared files
- **Cleanup Planning**: Locate old, unused files for deletion
- **Organization**: Understand your file structure and patterns
- **Compliance**: Generate reports for data governance

## 📁 Project Structure

```
src/
├── core/
│   └── drive-inventory-complete.js    # Main comprehensive script
├── specialized/
│   ├── image-inventory.js            # Image files analysis
│   ├── document-inventory.js         # Document files analysis
│   ├── large-files-finder.js         # Storage optimization
│   ├── shared-files-audit.js         # Security audit
│   ├── markdown-inventory.js         # Markdown/documentation files
│   └── code-files-inventory.js       # Programming files analysis
├── utils/
│   └── common-functions.js           # Shared utilities
docs/
├── setup-guide.md                    # Detailed installation
├── configuration.md                  # Config options
└── troubleshooting.md                # Common issues
examples/
├── basic-setup.js                    # Simple implementation
├── custom-config.js                  # Configuration examples
└── scheduled-runs.js                 # Automation setup
```

## ⚙️ Configuration Options

### Basic Settings
```javascript
const CONFIG = {
  BATCH_SIZE: 100,                     // Files per processing batch
  INVENTORY_SPREADSHEET_NAME: "Drive Inventory Report",
  LARGE_FILE_THRESHOLD_MB: 50,         // Large file threshold
  OLD_FILE_THRESHOLD_DAYS: 365,        // Old file threshold
  INCLUDE_GOOGLE_FILES: true,          // Include Docs, Sheets, etc.
  INCLUDE_TRASHED: false,              // Include deleted files
  TRACK_PERMISSIONS: true              // Analyze sharing settings
};
```

### Advanced Options
- Custom file type detection
- Folder exclusion patterns
- Report formatting preferences
- Automation triggers and schedules

## 🔄 Automation

### Continuous Mode
```javascript
function startAutomaticInventory() {
  // Runs continuously until complete
  // Handles Google Apps Script time limits
  // Automatically schedules continuation
}
```

### Scheduled Runs
```javascript
function setupHourlyInventory() {
  // Runs every hour until complete
  // Perfect for very large drives
  // Email notifications on completion
}
```

## 🛠️ Troubleshooting

### Common Issues
1. **Script Timeout**: Use batch processing or scheduled runs for large drives
2. **Permission Errors**: Ensure proper Drive API access
3. **Quota Limits**: Implement delays between API calls

### Performance Tips
- Adjust `BATCH_SIZE` based on your drive size
- Use specialized scripts for focused analysis
- Enable automatic continuation for large inventories

## 📈 Performance

- **Small Drives** (< 1,000 files): Complete in single run
- **Medium Drives** (1,000-10,000 files): Automatic batching
- **Large Drives** (10,000+ files): Scheduled hourly runs
- **Enterprise Drives**: Custom batch sizing and optimization

## 🔒 Security & Privacy

- **No Data Transmission**: All processing happens in your Google Apps Script environment
- **Local Results**: Reports stored in your own Google Drive
- **Permission Control**: You control all access permissions
- **Privacy First**: No external services or data sharing

## 📝 License

MIT License - feel free to modify and distribute

## 🤝 Contributing

Contributions welcome! Please feel free to submit a Pull Request.

## 📞 Support

- Check the `docs/` folder for detailed guides
- Review `examples/` for implementation help
- Open an issue for bugs or feature requests

---

**Note**: These scripts require Google Apps Script and Google Drive API access. All processing is done within Google's infrastructure for maximum security and performance.