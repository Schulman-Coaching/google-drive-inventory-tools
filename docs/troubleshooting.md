# Troubleshooting Guide - Google Drive Inventory Tools

This guide helps resolve common issues when running the inventory scripts.

## 🚨 Property Storage Quota Error

**Error Message:**
```
Exception: You have exceeded the property storage quota. Please remove some properties and try again.
```

### What This Means
Google Apps Script limits how much data can be stored in script properties. Large drives (5,000+ files) often hit this limit.

### ✅ Solution: Use the Optimized Version

**Instead of:** `src/core/drive-inventory-complete.js`  
**Use:** `src/core/drive-inventory-optimized.js`

**Quick Fix:**
1. Copy the contents of `drive-inventory-optimized.js`
2. Replace your current script with this optimized version
3. Run: `inventoryDriveLarge()` (instead of `inventoryDrive()`)

**Benefits of Optimized Version:**
- ✅ No property storage limits
- ✅ Handles drives with 10,000+ files
- ✅ Real-time progress tracking
- ✅ Automatic recovery from timeouts
- ✅ Memory efficient processing

---

## 🔍 Search Query Error

**Error Message:**
```
Exception: Invalid argument: q
```

### Solution
This happens with complex search queries in specialized scripts.

**Fixed Versions Available:**
- `document-inventory-fixed.js` - For document analysis
- All other specialized scripts have been updated with reliable search queries

---

## ⏱️ Script Timeout Errors

**Error Message:**
```
Exception: Exceeded maximum execution time
```

### Solutions

**Option 1: Use Automatic Mode**
```javascript
// Instead of single run
inventoryDrive()

// Use automatic continuation
startAutomaticInventory()
```

**Option 2: Reduce Batch Size**
```javascript
const CONFIG = {
  BATCH_SIZE: 50, // Reduce from 100-200
  // ... other settings
};
```

**Option 3: Use Scheduled Runs**
```javascript
setupHourlyInventory() // Runs every hour until complete
```

---

## 🔐 Permission Errors

**Error Message:**
```
Exception: You don't have permission to access this file
```

### Solutions

1. **Re-authorize the script:**
   - Delete the current script authorization
   - Run the script again to re-authorize

2. **Skip problematic files:**
   - The scripts automatically skip files they can't access
   - Check the error count in the final report

3. **Enterprise accounts:**
   - Some enterprise Google accounts restrict Apps Script access
   - Contact your IT administrator if needed

---

## 💾 Memory and Performance Issues

### Large Drive Optimization

**For drives with 10,000+ files:**
```javascript
// Use the memory-optimized version
inventoryDriveLarge()
```

**For drives with 50,000+ files:**
```javascript
// Use smaller batches and scheduled runs
const CONFIG = {
  BATCH_SIZE: 25,
  // ... other settings
};
setupHourlyInventory();
```

### Performance Tips

1. **Start with quick stats functions:**
   ```javascript
   getQuickStats()        // Test with ~1000 files first
   getQuickDocumentStats()
   getQuickImageStats()
   ```

2. **Use specialized scripts instead of complete inventory:**
   - Faster processing for specific file types
   - Lower memory usage
   - More focused results

3. **Run during off-peak hours:**
   - Google Apps Script has better performance during off-peak times
   - Set up scheduled runs for overnight processing

---

## 🔧 Common Configuration Issues

### Batch Size Guidelines

| Drive Size | Recommended BATCH_SIZE | Script Version |
|------------|----------------------|----------------|
| < 1,000 files | 200 | Any |
| 1,000-5,000 files | 100 | Standard |
| 5,000-10,000 files | 50 | Standard |
| 10,000+ files | 200 | **Optimized version required** |

### Memory Settings

```javascript
const CONFIG = {
  // Reduce these limits for very large drives
  MAX_LARGE_FILES: 50,        // Default: 100
  MAX_OLD_FILES: 50,          // Default: 100
  MAX_SHARED_FILES: 50,       // Default: 100
  MAX_DUPLICATE_GROUPS: 25,   // Default: 100
};
```

---

## 📊 Monitoring Progress

### Check Progress During Long Runs

**Optimized Version:** Check the "Progress Tracking" sheet in your generated spreadsheet

**Standard Version:** Check Google Apps Script logs

### Estimated Run Times

| Files | Batch Size | Estimated Time | Recommended Approach |
|-------|------------|----------------|---------------------|
| 1,000 | 100 | 2-5 minutes | Single run |
| 5,000 | 100 | 10-20 minutes | Automatic mode |
| 10,000 | 100 | 30+ minutes | Optimized version |
| 25,000+ | 50 | Hours | Scheduled runs + optimized |

---

## 🆘 Emergency Recovery

### If Script Gets Stuck

1. **Stop all triggers:**
   ```javascript
   cancelScheduledRuns()
   ```

2. **Reset progress:**
   ```javascript
   resetInventory()
   ```

3. **Check partial results:**
   - Even failed runs often produce partial spreadsheets
   - Check your Google Drive for inventory report files

### Partial Results Are Still Useful

Even if a script doesn't complete, you'll typically get:
- ✅ Overview statistics for processed files
- ✅ Complete file list for processed files  
- ✅ Large files found so far
- ✅ File type distribution

---

## 📞 Getting Help

### Before Asking for Help

1. **Check the error message** against this troubleshooting guide
2. **Try the optimized version** for property quota errors
3. **Test with quick stats functions** first
4. **Check Google Apps Script quotas:** [script.google.com/quotas](https://script.google.com/quotas)

### Useful Information to Include

When reporting issues, include:
- ✅ Exact error message
- ✅ Approximate number of files in your Drive
- ✅ Which script version you're using
- ✅ Your batch size setting
- ✅ How long the script ran before failing

### Script Quotas and Limits

Google Apps Script has daily quotas:
- **Execution time:** 6 minutes per run
- **Triggers:** 20 time-based triggers per script
- **Property storage:** 9KB for script properties
- **Daily runtime:** 6 hours total per day

The optimized versions are designed to work within these limits.

---

## ✅ Success Tips

1. **Always start with quick stats functions**
2. **Use the right tool for your drive size**
3. **Monitor progress regularly**
4. **Be patient with large drives**
5. **Keep Google Apps Script quotas in mind**

Most issues can be resolved by using the appropriate script version for your drive size and following the batch size recommendations above.