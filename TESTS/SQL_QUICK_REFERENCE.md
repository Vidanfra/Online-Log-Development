# SQLite Synchronization - Quick Reference Card

## 🚀 Quick Start (First Time Setup)

1. **Enable SQLite Mirroring:**
   - Open Settings (button on right side)
   - Go to "Database Sync" tab
   - Check ✓ "Enable SQLite Mirroring"
   - Click "Save Settings"

2. **Wait for Initial Sync:**
   - Watch status bar at bottom
   - Wait for "Initial SQLite sync complete" message
   - A `.db` file is created next to your Excel file

3. **Verify It's Working:**
   - Status indicator shows "SQLite Mirror: Enabled" ✓
   - After logging events, status shows "Excel: OK. SQLite: OK."

---

## 📊 What's New in This Version

### excel_row Column
- **First column** in the database
- Tracks the **exact Excel row position** for each record
- Helps you find which SQL row corresponds to which Excel row
- Example: `excel_row = 248` means this data is in row 248 of Excel

### Automatic Validation
- **Fixes empty excel_row values** automatically
- **Fixes duplicate excel_row values** automatically
- **Fixes invalid excel_row formats** automatically
- Runs during every full synchronization

---

## ✅ How to Test (30 seconds)

### Test 1: Basic Functionality
1. Press any event button
2. Check status bar → Should show "Excel: OK. SQLite: OK."
3. ✓ Working!

### Test 2: Concurrent Operations
1. Press 5 event buttons rapidly
2. All should log successfully
3. No "database is locked" errors
4. ✓ Working!

### Test 3: Auto-Sync
1. Wait 15 minutes (or set shorter interval in Settings)
2. Watch console for "Starting Full Synchronization"
3. Press event button **during sync**
4. Both should complete without errors
5. ✓ Working!

---

## 🐛 Troubleshooting

### "Database is locked" Error
**Quick Fix:**
- Close Excel file if open
- Wait 10 seconds
- Try again

### SQLite Not Updating
**Quick Fix:**
1. Go to Settings
2. **Uncheck** "Enable SQLite Mirroring"
3. Save Settings
4. **Check** "Enable SQLite Mirroring" again
5. Save Settings (this triggers fresh sync)

### Missing excel_row Column
**Quick Fix:**
1. Close application
2. Delete the `.db` file
3. Reopen application
4. Enable SQLite mirroring
5. Fresh database will be created with excel_row

---

## 📈 Performance Expectations

| Operation | Expected Time | Red Flag |
|-----------|--------------|----------|
| Single event log | < 2 seconds | > 5 seconds |
| Auto-sync (500 rows) | < 30 seconds | > 60 seconds |
| Initial sync | < 30 seconds | > 60 seconds |

---

## 🔍 How to Inspect the Database

### Option 1: DB Browser for SQLite (Recommended)
1. Download: https://sqlitebrowser.org/
2. Open your `.db` file
3. Go to "Browse Data" tab
4. Select your table
5. See all columns including `excel_row`

### Option 2: Command Line
```powershell
# Open SQLite command line
sqlite3 "path\to\your\logfile.db"

# View table structure
.schema

# View first 10 rows
SELECT * FROM TableName LIMIT 10;

# Check excel_row column
SELECT excel_row FROM TableName LIMIT 10;

# Exit
.quit
```

---

## 📞 Need Help?

### Check Console Output
- Console window shows detailed operations
- Look for ERROR or WARNING messages
- Copy and send error messages if issues occur

### Run Automated Test
```powershell
python "Python Script\test_sqlite_sync.py" "path\to\your\logfile.db"
```
- This runs 15 automated checks
- Shows PASS/FAIL for each test
- Helps identify specific issues

---

## ⚙️ Settings Reference

### Database Sync Tab
- **Enable SQLite Mirroring:** Turns on database synchronization
- **Auto-Sync Enabled:** Periodic full sync (recommended: ON)
- **Auto-Sync Interval:** How often to sync (default: 15 minutes)

### What Happens During Auto-Sync?
1. Reads entire Excel sheet
2. Validates excel_row values
3. Validates UUID values (if column exists)
4. Clears database
5. Inserts fresh data
6. Takes ~30 seconds for 500 rows

---

## 💡 Tips for Boat Operations

### Daily Routine
1. **Start of shift:** Enable SQLite mirroring
2. **During shift:** Log events normally (no changes to workflow)
3. **End of shift:** Data is automatically synced
4. **Next shift:** Database ready with all previous data

### If Excel Crashes
1. Reopen application
2. Database still has all data up to last auto-sync
3. Recent events (since last sync) might need manual recovery from backup

### Backup Strategy
- Excel file = Primary data (always)
- SQLite database = Mirror copy
- Keep Excel file in version control or regular backups
- Database can always be regenerated from Excel

---

## 🎯 Key Points to Remember

✅ **SQLite is a MIRROR** - Excel is still the primary data source
✅ **excel_row** tracks Excel row positions
✅ **Auto-sync** runs every 15 minutes by default
✅ **Concurrent operations** are safe (logging during sync works)
✅ **Validation** automatically fixes data issues
✅ **WAL mode** prevents "database is locked" errors

---

## 🔒 Data Safety

### What's Safe:
✓ Logging events while auto-sync is running
✓ Multiple rapid button presses
✓ Editing Excel manually (will sync on next auto-sync)
✓ Restarting application
✓ Deleting `.db` file (it will recreate)

### What to Avoid:
⚠️ Don't manually edit the `.db` file (changes will be overwritten)
⚠️ Don't delete Excel file (database can't regenerate)
⚠️ Don't disable mirroring mid-sync (wait for completion)

---

## 📋 Pre-Shift Checklist

- [ ] Application opens without errors
- [ ] SQLite mirror shows "Enabled"
- [ ] Test one event log (should show "Excel: OK. SQLite: OK.")
- [ ] Check console for any WARNING or ERROR messages
- [ ] If any issues, try disable/enable mirroring

**If everything checks out → You're good to go! 🚢**

---

*Version 2.2 - SQLite Synchronization with excel_row tracking*
*Questions? Check SQL_TESTING_GUIDE.md for detailed testing procedures*
