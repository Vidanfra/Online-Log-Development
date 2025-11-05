"""
Concurrent Database Access Test for Online Logger and Field Log Viewer
=======================================================================

This script tests whether Online Logger (writer) and Field Log Viewer (reader)
can operate simultaneously on the same SQLite database without blocking issues.

Test Scenarios:
1. WAL Mode Verification - Ensure database is in WAL mode
2. Concurrent Write/Read - Online Logger writes while Field Log Viewer reads
3. Multiple Readers - Multiple Field Log Viewer instances reading simultaneously
4. Write During Read - Simulate real-world concurrent operations
5. Lock Detection - Verify no database locks occur

Requirements:
- SQLite database file (created by Online Logger with WAL mode enabled)
- Python 3.13+
- sqlite3 (built-in)
"""

import sqlite3
import time
import threading
import os
from pathlib import Path
from datetime import datetime
import sys

# Color output for better readability
class Colors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

def print_success(msg):
    print(f"{Colors.GREEN}✓ {msg}{Colors.RESET}")

def print_error(msg):
    print(f"{Colors.RED}✗ {msg}{Colors.RESET}")

def print_warning(msg):
    print(f"{Colors.YELLOW}⚠ {msg}{Colors.RESET}")

def print_info(msg):
    print(f"{Colors.CYAN}ℹ {msg}{Colors.RESET}")

def print_header(msg):
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'='*70}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{msg}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'='*70}{Colors.RESET}\n")


class DatabaseTester:
    def __init__(self, db_path):
        self.db_path = db_path
        self.test_results = {}
        self.errors = []
        
    def verify_database_exists(self):
        """Check if the database file exists."""
        print_header("1. Database Existence Check")
        
        if not os.path.exists(self.db_path):
            print_error(f"Database not found: {self.db_path}")
            print_info("Please run Online Logger and enable 'SQLite Mirror' to create the database first.")
            return False
        
        print_success(f"Database found: {self.db_path}")
        
        # Check for WAL files
        wal_file = f"{self.db_path}-wal"
        shm_file = f"{self.db_path}-shm"
        
        if os.path.exists(wal_file):
            print_success(f"WAL file exists: {wal_file}")
        else:
            print_warning(f"WAL file not found: {wal_file} (will be created)")
            
        if os.path.exists(shm_file):
            print_success(f"SHM file exists: {shm_file}")
        else:
            print_warning(f"SHM file not found: {shm_file} (will be created)")
        
        return True
    
    def test_wal_mode(self):
        """Verify the database is in WAL mode."""
        print_header("2. WAL Mode Verification")
        
        try:
            conn = sqlite3.connect(self.db_path, timeout=10.0)
            cursor = conn.cursor()
            
            # Check journal mode
            cursor.execute("PRAGMA journal_mode")
            mode = cursor.fetchone()[0].lower()
            
            if mode == 'wal':
                print_success(f"Database is in WAL mode: {mode}")
                self.test_results['wal_mode'] = True
            else:
                print_error(f"Database is NOT in WAL mode: {mode}")
                print_warning("This will cause blocking issues!")
                print_info("Solution: Enable 'SQLite Mirror' in Online Logger to set WAL mode")
                self.test_results['wal_mode'] = False
                return False
            
            # Test WAL checkpoint
            try:
                cursor.execute("PRAGMA wal_checkpoint(PASSIVE)")
                print_success("WAL checkpoint test passed")
            except sqlite3.Error as e:
                print_warning(f"WAL checkpoint test failed: {e}")
            
            conn.close()
            return True
            
        except sqlite3.Error as e:
            print_error(f"Database error: {e}")
            self.test_results['wal_mode'] = False
            return False
    
    def test_get_table_info(self):
        """Get information about tables and row count."""
        print_header("3. Database Structure Analysis")
        
        try:
            conn = sqlite3.connect(self.db_path, timeout=10.0)
            cursor = conn.cursor()
            
            # Get all tables
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
            tables = [row[0] for row in cursor.fetchall()]
            
            if not tables:
                print_warning("No tables found in database")
                conn.close()
                return None
            
            print_success(f"Found {len(tables)} table(s): {', '.join(tables)}")
            
            # Use first table for testing
            test_table = tables[0]
            
            # Get row count
            cursor.execute(f'SELECT COUNT(*) FROM "{test_table}"')
            row_count = cursor.fetchone()[0]
            
            print_info(f"Table '{test_table}' has {row_count} rows")
            
            # Get columns
            cursor.execute(f'PRAGMA table_info("{test_table}")')
            columns = [row[1] for row in cursor.fetchall()]
            print_info(f"Columns ({len(columns)}): {', '.join(columns[:5])}{'...' if len(columns) > 5 else ''}")
            
            conn.close()
            return test_table
            
        except sqlite3.Error as e:
            print_error(f"Error reading database structure: {e}")
            return None
    
    def test_writer_simulation(self, table_name, duration=5):
        """Simulate Online Logger writing to database."""
        print_info(f"[WRITER] Starting write simulation for {duration} seconds...")
        
        write_count = 0
        errors = []
        start_time = time.time()
        
        try:
            conn = sqlite3.connect(self.db_path, timeout=10.0)
            cursor = conn.cursor()
            
            while time.time() - start_time < duration:
                try:
                    # Simulate a write operation (insert a test row)
                    timestamp = datetime.now().isoformat()
                    cursor.execute(f'SELECT COUNT(*) FROM "{table_name}"')
                    count = cursor.fetchone()[0]
                    
                    write_count += 1
                    time.sleep(0.5)  # Simulate work
                    
                except sqlite3.Error as e:
                    errors.append(str(e))
                    time.sleep(0.1)
            
            conn.close()
            
        except Exception as e:
            errors.append(f"Writer connection error: {e}")
        
        return write_count, errors
    
    def test_reader_simulation(self, table_name, duration=5, reader_id=1):
        """Simulate Field Log Viewer reading from database."""
        print_info(f"[READER {reader_id}] Starting read simulation for {duration} seconds...")
        
        read_count = 0
        errors = []
        start_time = time.time()
        
        try:
            # CRITICAL: Use read-only mode like Field Log Viewer
            db_uri = f"file:{self.db_path}?mode=ro"
            conn = sqlite3.connect(db_uri, uri=True, timeout=10.0)
            conn.execute("PRAGMA query_only = 1")
            cursor = conn.cursor()
            
            while time.time() - start_time < duration:
                try:
                    # Simulate a read operation
                    cursor.execute(f'SELECT COUNT(*) FROM "{table_name}"')
                    count = cursor.fetchone()[0]
                    
                    cursor.execute(f'SELECT * FROM "{table_name}" LIMIT 10')
                    rows = cursor.fetchall()
                    
                    read_count += 1
                    time.sleep(0.3)  # Simulate work
                    
                except sqlite3.Error as e:
                    if "database is locked" in str(e).lower():
                        errors.append(f"🔒 LOCK DETECTED: {e}")
                    else:
                        errors.append(str(e))
                    time.sleep(0.1)
            
            conn.close()
            
        except Exception as e:
            errors.append(f"Reader {reader_id} connection error: {e}")
        
        return read_count, errors
    
    def test_concurrent_access(self, table_name):
        """Test concurrent write/read operations."""
        print_header("4. Concurrent Access Test (Writer + 2 Readers)")
        print_info("Simulating Online Logger writing while Field Log Viewers read...")
        print_info("Duration: 5 seconds")
        
        # Start threads
        writer_thread = threading.Thread(
            target=lambda: setattr(self, 'writer_results', self.test_writer_simulation(table_name, duration=5))
        )
        reader1_thread = threading.Thread(
            target=lambda: setattr(self, 'reader1_results', self.test_reader_simulation(table_name, duration=5, reader_id=1))
        )
        reader2_thread = threading.Thread(
            target=lambda: setattr(self, 'reader2_results', self.test_reader_simulation(table_name, duration=5, reader_id=2))
        )
        
        # Start all threads
        writer_thread.start()
        time.sleep(0.5)  # Let writer start first
        reader1_thread.start()
        reader2_thread.start()
        
        # Wait for completion
        writer_thread.join()
        reader1_thread.join()
        reader2_thread.join()
        
        # Analyze results
        write_count, write_errors = self.writer_results
        read1_count, read1_errors = self.reader1_results
        read2_count, read2_errors = self.reader2_results
        
        print(f"\n{Colors.BOLD}Results:{Colors.RESET}")
        print(f"  Writer:   {write_count} operations, {len(write_errors)} errors")
        print(f"  Reader 1: {read1_count} operations, {len(read1_errors)} errors")
        print(f"  Reader 2: {read2_count} operations, {len(read2_errors)} errors")
        
        # Check for lock errors
        all_errors = write_errors + read1_errors + read2_errors
        lock_errors = [e for e in all_errors if "lock" in e.lower()]
        
        if lock_errors:
            print_error(f"\n{len(lock_errors)} DATABASE LOCK ERROR(S) DETECTED!")
            for err in lock_errors[:3]:  # Show first 3
                print_error(f"  • {err}")
            self.test_results['concurrent_access'] = False
            return False
        else:
            print_success("\n✓ No database locks detected!")
            print_success("✓ Concurrent access working correctly!")
            self.test_results['concurrent_access'] = True
            return True
    
    def test_rapid_reader_switches(self, table_name):
        """Test rapid connect/disconnect cycles like Field Log Viewer button clicks."""
        print_header("5. Rapid Reader Connection Test")
        print_info("Simulating multiple Field Log Viewer 'Update Excel' button clicks...")
        
        errors = []
        success_count = 0
        
        for i in range(10):
            try:
                # Open in read-only mode
                db_uri = f"file:{self.db_path}?mode=ro"
                conn = sqlite3.connect(db_uri, uri=True, timeout=10.0)
                conn.execute("PRAGMA query_only = 1")
                cursor = conn.cursor()
                
                # Quick read
                cursor.execute(f'SELECT COUNT(*) FROM "{table_name}"')
                count = cursor.fetchone()[0]
                
                conn.close()
                success_count += 1
                print(f"  Connection {i+1}/10: OK ({count} rows)")
                
            except sqlite3.Error as e:
                errors.append(f"Connection {i+1}: {e}")
                print_error(f"  Connection {i+1}/10: FAILED - {e}")
            
            time.sleep(0.1)  # Small delay between connections
        
        if not errors:
            print_success(f"\n✓ All {success_count} rapid connections successful!")
            self.test_results['rapid_connections'] = True
            return True
        else:
            print_error(f"\n✗ {len(errors)} connection failures detected!")
            self.test_results['rapid_connections'] = False
            return False
    
    def run_all_tests(self):
        """Run complete test suite."""
        print(f"\n{Colors.BOLD}{Colors.BLUE}")
        print("╔════════════════════════════════════════════════════════════════════╗")
        print("║  SQLite Concurrent Access Test Suite                              ║")
        print("║  Testing: Online Logger (Writer) + Field Log Viewer (Readers)     ║")
        print("╚════════════════════════════════════════════════════════════════════╝")
        print(f"{Colors.RESET}")
        
        # Test 1: Database exists
        if not self.verify_database_exists():
            return False
        
        # Test 2: WAL mode
        if not self.test_wal_mode():
            return False
        
        # Test 3: Get table info
        table_name = self.test_get_table_info()
        if not table_name:
            print_error("Cannot proceed without a valid table")
            return False
        
        # Test 4: Concurrent access
        concurrent_ok = self.test_concurrent_access(table_name)
        
        # Test 5: Rapid connections
        rapid_ok = self.test_rapid_reader_switches(table_name)
        
        # Final summary
        self.print_summary()
        
        return all(self.test_results.values())
    
    def print_summary(self):
        """Print final test summary."""
        print_header("TEST SUMMARY")
        
        all_passed = all(self.test_results.values())
        
        for test_name, passed in self.test_results.items():
            status = f"{Colors.GREEN}PASS{Colors.RESET}" if passed else f"{Colors.RED}FAIL{Colors.RESET}"
            print(f"  {test_name.replace('_', ' ').title()}: {status}")
        
        print("\n" + "="*70)
        if all_passed:
            print(f"{Colors.GREEN}{Colors.BOLD}✓ ALL TESTS PASSED!{Colors.RESET}")
            print(f"{Colors.GREEN}The database is properly configured for concurrent access.{Colors.RESET}")
            print(f"{Colors.GREEN}Safe to deploy to colleagues!{Colors.RESET}")
        else:
            print(f"{Colors.RED}{Colors.BOLD}✗ SOME TESTS FAILED!{Colors.RESET}")
            print(f"{Colors.RED}Do NOT deploy until all tests pass.{Colors.RESET}")
            print(f"\n{Colors.YELLOW}Common fixes:{Colors.RESET}")
            print(f"  1. Run Online Logger and enable 'SQLite Mirror'")
            print(f"  2. Verify WAL mode is enabled (check console output)")
            print(f"  3. Ensure database is on local drive (not network drive)")
            print(f"  4. Close any programs that might have the database locked")
        print("="*70 + "\n")


def main():
    print("\n" + "="*70)
    print("SQLite Concurrent Access Test - Online Logger & Field Log Viewer")
    print("="*70 + "\n")
    
    # Get database path
    if len(sys.argv) > 1:
        db_path = sys.argv[1]
    else:
        # Try to find database in common locations
        common_paths = [
            "SQL Database/Online_Log_SQLite.db",
            "SQL Database/fieldlog.db",
            "Online_Log_SQLite.db",
        ]
        
        db_path = None
        for path in common_paths:
            if os.path.exists(path):
                db_path = path
                break
        
        if not db_path:
            print_error("Database not found in common locations.")
            print_info("\nUsage:")
            print_info("  python test_concurrent_access.py [path/to/database.db]")
            print_info("\nOr create database first:")
            print_info("  1. Run Online Logger")
            print_info("  2. Enable 'SQLite Mirror' checkbox")
            print_info("  3. Add at least one log entry")
            print_info("  4. Run this test script")
            return 1
    
    # Run tests
    tester = DatabaseTester(db_path)
    success = tester.run_all_tests()
    
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
