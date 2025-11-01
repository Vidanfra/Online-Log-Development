"""
SQLite Synchronization Test Suite
==================================

This script performs comprehensive tests on the SQLite mirroring feature
to ensure it handles concurrent operations, edge cases, and stress scenarios.

Usage:
    python test_sqlite_sync.py

Requirements:
    - An Excel log file with some test data
    - SQLite mirroring enabled in the application
"""

import sqlite3
import os
import time
import threading
import uuid as uuid_module
from pathlib import Path
import pandas as pd
import random
import sys

# Color codes for terminal output
GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
BLUE = '\033[94m'
RESET = '\033[0m'

class SQLiteSyncTester:
    def __init__(self, db_path):
        self.db_path = db_path
        self.test_results = []
        self.table_name = Path(db_path).stem.replace('.db', '')
        
    def connect_db(self):
        """Create a connection to the database."""
        return sqlite3.connect(self.db_path, timeout=30.0)
    
    def print_header(self, text):
        """Print a formatted header."""
        print(f"\n{BLUE}{'='*70}")
        print(f"  {text}")
        print(f"{'='*70}{RESET}\n")
    
    def print_test(self, test_name, passed, message=""):
        """Print test result."""
        status = f"{GREEN}✓ PASS{RESET}" if passed else f"{RED}✗ FAIL{RESET}"
        print(f"{status} - {test_name}")
        if message:
            print(f"       {message}")
        self.test_results.append((test_name, passed, message))
    
    def test_1_database_exists(self):
        """Test 1: Verify database file exists."""
        self.print_header("TEST 1: Database File Existence")
        exists = os.path.exists(self.db_path)
        self.print_test("Database file exists", exists, 
                       f"Path: {self.db_path}" if exists else "Database file not found!")
        return exists
    
    def test_2_table_structure(self):
        """Test 2: Verify table structure and excel_row column."""
        self.print_header("TEST 2: Table Structure")
        try:
            conn = self.connect_db()
            cursor = conn.cursor()
            
            # Get table info
            cursor.execute(f"PRAGMA table_info('{self.table_name}')")
            columns = cursor.fetchall()
            
            if not columns:
                self.print_test("Table exists", False, "Table not found in database!")
                conn.close()
                return False
            
            column_names = [col[1] for col in columns]
            
            # Test: Table exists
            self.print_test("Table exists", True, f"Table '{self.table_name}' found")
            
            # Test: excel_row is first column
            is_first = column_names[0] == 'excel_row'
            self.print_test("excel_row is first column", is_first,
                          f"First column: {column_names[0]}")
            
            # Test: Column count
            self.print_test("Table has columns", len(columns) > 1,
                          f"Total columns: {len(columns)}")
            
            # Print all columns
            print(f"\n  {YELLOW}Table Structure:{RESET}")
            for i, col_name in enumerate(column_names, 1):
                marker = "◄ PRIMARY ID" if col_name == 'excel_row' else ""
                print(f"    {i}. {col_name} {marker}")
            
            conn.close()
            return is_first and len(columns) > 1
            
        except Exception as e:
            self.print_test("Table structure check", False, f"Error: {e}")
            return False
    
    def test_3_data_integrity(self):
        """Test 3: Verify data integrity and excel_row values."""
        self.print_header("TEST 3: Data Integrity")
        try:
            conn = self.connect_db()
            cursor = conn.cursor()
            
            # Count total rows
            cursor.execute(f"SELECT COUNT(*) FROM '{self.table_name}'")
            total_rows = cursor.fetchone()[0]
            
            self.print_test("Database has data", total_rows > 0,
                          f"Total rows: {total_rows}")
            
            if total_rows == 0:
                conn.close()
                return False
            
            # Check excel_row values
            cursor.execute(f"SELECT excel_row FROM '{self.table_name}'")
            excel_rows = [row[0] for row in cursor.fetchall()]
            
            # Test: No NULL excel_row values
            null_count = sum(1 for r in excel_rows if r is None or r == '')
            self.print_test("No NULL excel_row values", null_count == 0,
                          f"NULL values: {null_count}")
            
            # Test: No duplicate excel_row values
            unique_count = len(set(excel_rows))
            has_duplicates = unique_count != len(excel_rows)
            self.print_test("No duplicate excel_row values", not has_duplicates,
                          f"Unique: {unique_count}, Total: {len(excel_rows)}")
            
            # Test: excel_row values are positive integers
            try:
                excel_row_ints = [int(r) for r in excel_rows if r]
                all_positive = all(r > 0 for r in excel_row_ints)
                self.print_test("All excel_row values are positive integers", all_positive,
                              f"Range: {min(excel_row_ints) if excel_row_ints else 'N/A'} to {max(excel_row_ints) if excel_row_ints else 'N/A'}")
            except ValueError as e:
                self.print_test("All excel_row values are positive integers", False,
                              f"Invalid format: {e}")
            
            # Check UUID column if exists
            cursor.execute(f"PRAGMA table_info('{self.table_name}')")
            columns = [col[1].lower() for col in cursor.fetchall()]
            
            if 'uuid' in columns:
                cursor.execute(f"SELECT UUID FROM '{self.table_name}'")
                uuids = [row[0] for row in cursor.fetchall()]
                
                # Test UUID integrity
                null_uuids = sum(1 for u in uuids if not u or u.strip() == '')
                self.print_test("UUID column has no empty values", null_uuids == 0,
                              f"Empty UUIDs: {null_uuids}")
                
                # Check for valid UUID format
                valid_uuids = 0
                for u in uuids:
                    if u:
                        try:
                            uuid_module.UUID(str(u))
                            valid_uuids += 1
                        except:
                            pass
                
                self.print_test("All UUIDs are valid format", valid_uuids == len(uuids),
                              f"Valid: {valid_uuids}/{len(uuids)}")
            
            conn.close()
            return True
            
        except Exception as e:
            self.print_test("Data integrity check", False, f"Error: {e}")
            return False
    
    def test_4_concurrent_reads(self):
        """Test 4: Verify database handles concurrent read operations."""
        self.print_header("TEST 4: Concurrent Read Operations")
        
        errors = []
        results = []
        
        def read_worker(worker_id):
            try:
                conn = self.connect_db()
                cursor = conn.cursor()
                cursor.execute(f"SELECT COUNT(*) FROM '{self.table_name}'")
                count = cursor.fetchone()[0]
                results.append((worker_id, count))
                conn.close()
            except Exception as e:
                errors.append((worker_id, str(e)))
        
        # Spawn 10 concurrent readers
        threads = []
        for i in range(10):
            t = threading.Thread(target=read_worker, args=(i,))
            threads.append(t)
            t.start()
        
        # Wait for all threads
        for t in threads:
            t.join()
        
        # Check results
        self.print_test("All concurrent reads successful", len(errors) == 0,
                       f"Successful: {len(results)}/10, Errors: {len(errors)}")
        
        if errors:
            for worker_id, error in errors[:3]:  # Show first 3 errors
                print(f"       Worker {worker_id}: {error}")
        
        # Verify all readers got same count
        if results:
            counts = [r[1] for r in results]
            all_same = len(set(counts)) == 1
            self.print_test("All readers see consistent data", all_same,
                          f"Row counts: {set(counts)}")
        
        return len(errors) == 0
    
    def test_5_wal_mode(self):
        """Test 5: Verify WAL mode is enabled."""
        self.print_header("TEST 5: Database Configuration")
        try:
            conn = self.connect_db()
            cursor = conn.cursor()
            
            # Check journal mode
            cursor.execute("PRAGMA journal_mode")
            journal_mode = cursor.fetchone()[0].upper()
            
            self.print_test("WAL mode enabled", journal_mode == 'WAL',
                          f"Journal mode: {journal_mode}")
            
            # Check busy timeout
            cursor.execute("PRAGMA busy_timeout")
            timeout = cursor.fetchone()[0]
            
            self.print_test("Busy timeout configured", timeout >= 30000,
                          f"Timeout: {timeout}ms")
            
            conn.close()
            return journal_mode == 'WAL'
            
        except Exception as e:
            self.print_test("Database configuration check", False, f"Error: {e}")
            return False
    
    def test_6_query_performance(self):
        """Test 6: Measure query performance."""
        self.print_header("TEST 6: Query Performance")
        try:
            conn = self.connect_db()
            cursor = conn.cursor()
            
            # Test 1: Simple SELECT
            start = time.time()
            cursor.execute(f"SELECT * FROM '{self.table_name}' LIMIT 100")
            cursor.fetchall()
            select_time = (time.time() - start) * 1000
            
            self.print_test("SELECT query performance", select_time < 100,
                          f"Time: {select_time:.2f}ms")
            
            # Test 2: excel_row lookup
            cursor.execute(f"SELECT excel_row FROM '{self.table_name}' LIMIT 1")
            sample_row = cursor.fetchone()
            
            if sample_row:
                start = time.time()
                cursor.execute(f"SELECT * FROM '{self.table_name}' WHERE excel_row = ?", (sample_row[0],))
                cursor.fetchone()
                lookup_time = (time.time() - start) * 1000
                
                self.print_test("excel_row lookup performance", lookup_time < 50,
                              f"Time: {lookup_time:.2f}ms")
            
            # Test 3: Full table scan
            start = time.time()
            cursor.execute(f"SELECT COUNT(*) FROM '{self.table_name}'")
            cursor.fetchone()
            count_time = (time.time() - start) * 1000
            
            self.print_test("COUNT query performance", count_time < 200,
                          f"Time: {count_time:.2f}ms")
            
            conn.close()
            return True
            
        except Exception as e:
            self.print_test("Query performance test", False, f"Error: {e}")
            return False
    
    def test_7_stress_test(self):
        """Test 7: Stress test with rapid queries."""
        self.print_header("TEST 7: Stress Test (100 rapid queries)")
        
        errors = []
        success_count = 0
        
        def stress_worker(worker_id):
            nonlocal success_count
            try:
                conn = self.connect_db()
                cursor = conn.cursor()
                
                # Perform 10 random operations
                for _ in range(10):
                    operation = random.choice(['count', 'select', 'lookup'])
                    
                    if operation == 'count':
                        cursor.execute(f"SELECT COUNT(*) FROM '{self.table_name}'")
                        cursor.fetchone()
                    elif operation == 'select':
                        cursor.execute(f"SELECT * FROM '{self.table_name}' LIMIT 10")
                        cursor.fetchall()
                    else:
                        cursor.execute(f"SELECT excel_row FROM '{self.table_name}' LIMIT 1")
                        cursor.fetchone()
                    
                    success_count += 1
                
                conn.close()
            except Exception as e:
                errors.append((worker_id, str(e)))
        
        start_time = time.time()
        
        # Spawn 10 workers doing 10 operations each = 100 queries
        threads = []
        for i in range(10):
            t = threading.Thread(target=stress_worker, args=(i,))
            threads.append(t)
            t.start()
        
        for t in threads:
            t.join()
        
        duration = time.time() - start_time
        
        self.print_test("Stress test completed", len(errors) == 0,
                       f"Duration: {duration:.2f}s, Queries/sec: {success_count/duration:.1f}")
        
        if errors:
            self.print_test("All queries successful", False,
                          f"Errors: {len(errors)}")
            for worker_id, error in errors[:3]:
                print(f"       Worker {worker_id}: {error}")
        else:
            self.print_test("All queries successful", True,
                          f"{success_count} queries completed")
        
        return len(errors) == 0
    
    def run_all_tests(self):
        """Run all tests and print summary."""
        print(f"\n{BLUE}{'='*70}")
        print(f"  SQLite Synchronization Test Suite")
        print(f"{'='*70}{RESET}\n")
        print(f"Database: {self.db_path}\n")
        
        # Run tests
        self.test_1_database_exists()
        self.test_2_table_structure()
        self.test_3_data_integrity()
        self.test_4_concurrent_reads()
        self.test_5_wal_mode()
        self.test_6_query_performance()
        self.test_7_stress_test()
        
        # Summary
        self.print_header("TEST SUMMARY")
        
        total = len(self.test_results)
        passed = sum(1 for _, p, _ in self.test_results if p)
        failed = total - passed
        
        print(f"  Total Tests:  {total}")
        print(f"  {GREEN}Passed:       {passed}{RESET}")
        print(f"  {RED}Failed:       {failed}{RESET}")
        print(f"  Success Rate: {(passed/total*100):.1f}%\n")
        
        if failed > 0:
            print(f"{RED}⚠ Some tests failed. Review the output above for details.{RESET}\n")
        else:
            print(f"{GREEN}✓ All tests passed successfully!{RESET}\n")
        
        return failed == 0


def main():
    """Main test execution."""
    print("\n" + "="*70)
    print("  SQLite Synchronization Test Suite")
    print("="*70 + "\n")
    
    # Ask for database path
    print("Please provide the path to your SQLite database file.")
    print("Example: C:\\Path\\To\\Your\\LogFile.db\n")
    
    if len(sys.argv) > 1:
        db_path = sys.argv[1]
    else:
        db_path = input("Database path: ").strip().strip('"')
    
    if not db_path:
        print(f"{RED}Error: No database path provided.{RESET}")
        return
    
    if not os.path.exists(db_path):
        print(f"{RED}Error: Database file not found: {db_path}{RESET}")
        return
    
    # Run tests
    tester = SQLiteSyncTester(db_path)
    success = tester.run_all_tests()
    
    # Exit with appropriate code
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
