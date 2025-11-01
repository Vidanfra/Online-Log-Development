# SQLite Synchronization Testing Guide

## 🎯 Overview
This guide will help you thoroughly test the SQLite mirroring feature to ensure it's production-ready for your colleagues on the boat.

---

## 📋 Pre-Testing Checklist

### 1. Backup Your Data
- [ ] Create a backup copy of your Excel log file
- [ ] Store it in a safe location
- [ ] Name it with date (e.g., `LogFile_Backup_2025-11-01.xlsb`)

### 2. Verify Settings
- [ ] Open the application
- [ ] Go to Settings → Database Sync
- [ ] Ensure SQLite mirror is enabled
- [ ] Note the database file location

---

## 🧪 Testing Scenarios

### **TEST 1: Initial Synchronization**
**Purpose:** Verify full sync creates database correctly with excel_row column

**Steps:**
1. Close the application completely
2. Delete the `.db` file if it exists
3. Open the application
4. Enable SQLite mirroring in Settings
5. Wait for "Initial SQLite sync complete" message

**Expected Results:**
- ✅ Database file created (same location as Excel file)
- ✅ Status shows "SQLite: OK"
- ✅ Console shows "excel_row column validated"
- ✅ Console shows row range (e.g., "Range: 15 to 247")

**Pass/Fail:** ________

---

### **TEST 2: Single Event Logging**
**Purpose:** Verify instant row addition with excel_row tracking

**Steps:**
1. Note current row count in Excel
2. Press "Event" button (or any custom button)
3. Check status bar for "Excel: OK. SQLite: OK."
4. Open the Excel file
5. Check the last row number (e.g., row 248)

**Expected Results:**
- ✅ New row added to Excel
- ✅ Status shows both Excel and SQLite success
- ✅ Console shows "Added 1 row to 'TableName' (Excel row 248)"
- ✅ No "WARNING - No excel_row provided" message

**Pass/Fail:** ________

---

### **TEST 3: Multiple Rapid Events**
**Purpose:** Test concurrent logging (stress test)

**Steps:**
1. Press multiple event buttons rapidly (5-10 clicks in 5 seconds)
2. Wait for all operations to complete
3. Check status bar after each
4. Open Excel and count new rows

**Expected Results:**
- ✅ All events logged successfully
- ✅ No "database is locked" errors
- ✅ All rows have unique excel_row values
- ✅ Excel row count matches button presses

**Pass/Fail:** ________

---

### **TEST 4: Logging During Auto-Sync**
**Purpose:** Verify concurrent operations don't conflict

**Steps:**
1. Set auto-sync interval to 1 minute (Settings → Database Sync)
2. Wait for auto-sync to start (check console for "Starting Full Synchronization")
3. **IMMEDIATELY** press an event button while sync is running
4. Watch console output
5. Verify both operations complete

**Expected Results:**
- ✅ Both operations complete without errors
- ✅ No "database is locked" errors
- ✅ New event appears in database
- ✅ Auto-sync completes successfully

**Pass/Fail:** ________

---

### **TEST 5: Manual Excel Modification During Sync**
**Purpose:** Test handling of Excel changes during synchronization

**Steps:**
1. Trigger auto-sync or press "Manual KP Log"
2. **While sync is running**, open Excel
3. Manually change a value in the Event column
4. Save Excel file
5. Wait for sync to complete
6. Check database for the change

**Expected Results:**
- ✅ Sync completes without crashes
- ✅ Next sync picks up the manual change
- ⚠️ Current sync might not include change (expected behavior)

**Pass/Fail:** ________

---

### **TEST 6: UUID Column Validation (if used)** ⚠️ CRITICAL
**Purpose:** Verify UUID corrections are written to BOTH Excel and SQLite

**Steps:**
1. Open Excel file
2. Find a row with a UUID
3. Delete or corrupt the UUID value (e.g., change to "CORRUPTED")
4. Save Excel and close it
5. Trigger full sync (disable/enable SQLite mirroring)
6. Check console output
7. **CRITICAL:** Reopen Excel and verify the UUID was fixed there too

**Expected Results:**
- ✅ Console shows "Generating/fixing X UUIDs"
- ✅ Console shows "Wrote X UUID values to Excel"
- ✅ Console shows "Excel file saved with corrections"
- ✅ **Excel file now has valid UUID (not empty or corrupted)**
- ✅ SQLite database has the SAME valid UUID as Excel
- ✅ excel_row column remains unchanged

**Pass/Fail:** ________

**⚠️ CRITICAL:** This test verifies UUIDs match in both Excel and SQLite!
If UUIDs differ, the databases are NOT properly synchronized.

---

### **TEST 7: Duplicate excel_row Detection**
**Purpose:** Verify validation catches and fixes duplicates

**Steps:**
1. Open Excel file
2. Manually create duplicate excel_row values:
   - Copy row 50's excel_row value to row 51
3. Save Excel
4. Trigger full sync (disable/enable SQLite mirroring)
5. Check console output

**Expected Results:**
- ✅ Console shows "Found X duplicate excel_row values. Regenerating..."
- ✅ Duplicates are fixed with correct row numbers
- ✅ Database has no duplicates

**Pass/Fail:** ________

---

### **TEST 8: Empty excel_row Detection**
**Purpose:** Verify validation catches empty values

**Steps:**
1. Open Excel file
2. Delete the excel_row value from a few rows
3. Save Excel
4. Trigger full sync
5. Check console output

**Expected Results:**
- ✅ Console shows "Fixing X excel_row values..."
- ✅ Empty values filled with correct row numbers
- ✅ No NULL values in database

**Pass/Fail:** ________

---

### **TEST 9: Large Dataset Performance**
**Purpose:** Verify performance with realistic data volume

**Steps:**
1. Ensure Excel file has at least 100 rows
2. Trigger full sync
3. Time how long it takes
4. Check console for row count

**Expected Results:**
- ✅ Sync completes in reasonable time (< 30 seconds for 500 rows)
- ✅ Console shows correct row count
- ✅ No timeout errors

**Sync Duration:** ________ seconds
**Pass/Fail:** ________

---

### **TEST 10: Application Restart**
**Purpose:** Verify database persists correctly

**Steps:**
1. Log a few events
2. Note the last excel_row number
3. Close the application completely
4. Wait 10 seconds
5. Reopen the application
6. Check SQLite status

**Expected Results:**
- ✅ Database file still exists
- ✅ SQLite mirror reconnects automatically
- ✅ Previous data intact
- ✅ New events log correctly

**Pass/Fail:** ________

---

## 🔧 Automated Testing

### Running the Test Script

1. **Open PowerShell** in the project directory
2. **Run the test script:**
   ```powershell
   python "Python Script\test_sqlite_sync.py" "path\to\your\logfile.db"
   ```
   
3. **Example:**
   ```powershell
   python "Python Script\test_sqlite_sync.py" "C:\Projects\Online-Log-Development\SQL Database\Online_Log_Test.db"
   ```

### Test Script Checks:
- ✅ Database file existence
- ✅ Table structure (excel_row first column)
- ✅ Data integrity (no NULLs, no duplicates)
- ✅ Concurrent read operations (10 simultaneous readers)
- ✅ WAL mode configuration
- ✅ Query performance
- ✅ Stress test (100 rapid queries)

### Expected Output:
```
==================================================================
  SQLite Synchronization Test Suite
==================================================================

Database: C:\...\Online_Log_Test.db

==================================================================
  TEST 1: Database File Existence
==================================================================

✓ PASS - Database file exists
       Path: C:\...\Online_Log_Test.db

...

==================================================================
  TEST SUMMARY
==================================================================

  Total Tests:  15
  Passed:       15
  Failed:       0
  Success Rate: 100.0%

✓ All tests passed successfully!
```

---

## 🐛 Common Issues and Solutions

### Issue 1: "Database is locked"
**Symptoms:** Error message when logging events
**Solutions:**
- Close Excel file (it might have exclusive lock)
- Wait for auto-sync to complete
- Check console for stuck operations
- Restart application

### Issue 2: "No excel_row provided"
**Symptoms:** Warning in console when logging
**Solutions:**
- This is a warning, not an error
- Check if `save_to_excel_and_sqlite` passes `next_row` parameter
- Verify the code update was applied correctly

### Issue 3: Missing excel_row column
**Symptoms:** Database doesn't have excel_row as first column
**Solutions:**
- Delete the `.db` file
- Disable and re-enable SQLite mirroring
- This will trigger fresh full sync with excel_row

### Issue 4: Duplicates not being fixed
**Symptoms:** Duplicate excel_row values persist
**Solutions:**
- Trigger manual full sync (disable/enable mirroring)
- Check console output for validation messages
- Verify `_validate_and_fix_excel_rows` is being called

---

## 📊 Performance Benchmarks

### Acceptable Performance:
- **Initial Sync (500 rows):** < 30 seconds
- **Single Event Logging:** < 2 seconds total (Excel + SQLite)
- **Auto-Sync (500 rows):** < 30 seconds
- **Concurrent Reads:** No errors with 10+ simultaneous connections
- **Query Response:** < 100ms for simple SELECT

### Red Flags:
- ⚠️ Sync takes > 60 seconds for 500 rows
- ⚠️ "Database is locked" errors during normal operation
- ⚠️ Event logging takes > 5 seconds
- ⚠️ Frequent timeout errors

---

## ✅ Pre-Deployment Checklist

Before sending to colleagues:

- [ ] All 10 manual tests passed
- [ ] Automated test script shows 100% pass rate
- [ ] Tested with realistic data volume (100+ rows)
- [ ] Tested concurrent operations (logging during sync)
- [ ] Tested Excel manual modifications during sync
- [ ] Verified excel_row column is always first
- [ ] Verified no NULL or duplicate excel_row values
- [ ] Verified UUID validation still works (if used)
- [ ] Performance is acceptable (see benchmarks above)
- [ ] No "database is locked" errors in normal use
- [ ] Application handles restarts correctly
- [ ] Console output is clean (no errors or warnings)

---

## 📝 Notes for Deployment

### What to Tell Your Colleagues:

1. **Enable SQLite Mirror:**
   - Go to Settings → Database Sync
   - Check "Enable SQLite Mirroring"
   - Wait for initial sync to complete

2. **What They'll See:**
   - A `.db` file created next to their Excel file
   - Status indicator shows "SQLite Mirror: Enabled"
   - Status bar shows "Excel: OK. SQLite: OK." after each event

3. **What's Different:**
   - Database now tracks exact Excel row positions
   - First column in database is `excel_row`
   - This helps identify which SQL row = which Excel row
   - Automatic validation fixes duplicates and empty values

4. **If Problems Occur:**
   - Disable and re-enable SQLite mirroring (triggers fresh sync)
   - Check console output for error messages
   - Delete the `.db` file and let it recreate

5. **Performance Notes:**
   - Auto-sync runs every 15 minutes by default
   - During sync, logging still works (concurrent operations)
   - WAL mode prevents "database is locked" errors

---

## 🎯 Success Criteria

The SQLite synchronization is ready for deployment if:
1. ✅ All manual tests pass
2. ✅ Automated test shows 100% pass rate
3. ✅ No errors during concurrent operations
4. ✅ excel_row column is always first and valid
5. ✅ Performance meets benchmarks
6. ✅ Database handles restarts correctly

---

## 📞 Support

If issues arise in production:
1. Check console output for specific error messages
2. Run the automated test script on their machine
3. Verify their Excel file structure matches expectations
4. Test with a fresh database file (delete `.db` and recreate)

Good luck with deployment! 🚢
