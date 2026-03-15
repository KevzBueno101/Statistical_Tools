"""
Database Integration Test Script
Comprehensive testing for the Statistical Analysis Suite database system

Run this script to verify all database functionality works correctly.
"""

import sys
import os
from datetime import datetime

# Add modules directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

# Import database module
try:
    from modules import database
    print("✓ Database module imported successfully")
except ImportError as e:
    print(f"✗ Failed to import database module: {e}")
    sys.exit(1)


def test_initialization():
    """Test database initialization"""
    print("\n" + "=" * 70)
    print("TEST 1: DATABASE INITIALIZATION")
    print("=" * 70)
    
    try:
        success = database.initialize_database("test_stats_app.db")
        if success:
            print("✓ Database initialized successfully")
            return True
        else:
            print("✗ Database initialization failed")
            return False
    except Exception as e:
        print(f"✗ Error during initialization: {e}")
        return False


def test_save_result():
    """Test saving analysis results"""
    print("\n" + "=" * 70)
    print("TEST 2: SAVE ANALYSIS RESULT")
    print("=" * 70)
    
    try:
        # Test data
        test_data = {
            'analysis_type': 'Test ANOVA',
            'input_data': {
                'groups': ['A', 'B', 'C'],
                'n_total': 30,
                'alpha': 0.05
            },
            'result': {
                'F_statistic': 4.567,
                'p_value': 0.023,
                'df_between': 2,
                'df_within': 27,
                'is_significant': True
            },
            'interpretation': 'Test interpretation: Significant difference found between groups.',
            'metadata': {
                'test_run': True,
                'timestamp': datetime.now().isoformat()
            }
        }
        
        record_id = database.save_result(
            analysis_type=test_data['analysis_type'],
            input_data=test_data['input_data'],
            result=test_data['result'],
            interpretation=test_data['interpretation'],
            metadata=test_data['metadata'],
            db_path="test_stats_app.db"
        )
        
        if record_id:
            print(f"✓ Result saved successfully (ID: {record_id})")
            return record_id
        else:
            print("✗ Failed to save result")
            return None
            
    except Exception as e:
        print(f"✗ Error saving result: {e}")
        return None


def test_retrieve_result(record_id):
    """Test retrieving a specific result"""
    print("\n" + "=" * 70)
    print("TEST 3: RETRIEVE RESULT BY ID")
    print("=" * 70)
    
    try:
        record = database.get_result_by_id(record_id, db_path="test_stats_app.db")
        
        if record:
            print(f"✓ Retrieved record successfully")
            print(f"   ID: {record['id']}")
            print(f"   Type: {record['analysis_type']}")
            print(f"   Date: {record['date_created']}")
            return True
        else:
            print("✗ Failed to retrieve record")
            return False
            
    except Exception as e:
        print(f"✗ Error retrieving result: {e}")
        return False


def test_retrieve_all():
    """Test retrieving all results"""
    print("\n" + "=" * 70)
    print("TEST 4: RETRIEVE ALL RESULTS")
    print("=" * 70)
    
    try:
        results = database.get_all_results(db_path="test_stats_app.db")
        
        print(f"✓ Retrieved {len(results)} record(s)")
        
        for i, record in enumerate(results, 1):
            print(f"\n   Record {i}:")
            print(f"   - ID: {record['id']}")
            print(f"   - Type: {record['analysis_type']}")
            print(f"   - Date: {record['date_created']}")
        
        return len(results) > 0
        
    except Exception as e:
        print(f"✗ Error retrieving results: {e}")
        return False


def test_search():
    """Test search functionality"""
    print("\n" + "=" * 70)
    print("TEST 5: SEARCH BY ANALYSIS TYPE")
    print("=" * 70)
    
    try:
        # Search for test data
        results = database.search_by_analysis_type("Test", db_path="test_stats_app.db")
        
        print(f"✓ Search completed - found {len(results)} matching record(s)")
        
        for record in results:
            print(f"   - {record['analysis_type']} (ID: {record['id']})")
        
        return True
        
    except Exception as e:
        print(f"✗ Error during search: {e}")
        return False


def test_database_stats():
    """Test database statistics"""
    print("\n" + "=" * 70)
    print("TEST 6: DATABASE STATISTICS")
    print("=" * 70)
    
    try:
        stats = database.get_database_stats(db_path="test_stats_app.db")
        
        print(f"✓ Statistics retrieved successfully:")
        print(f"   - Total Records: {stats.get('total_records', 0)}")
        print(f"   - Database Size: {stats.get('database_size_kb', 0)} KB")
        
        if stats.get('by_type'):
            print(f"   - Records by Type:")
            for atype, count in stats['by_type'].items():
                print(f"      • {atype}: {count}")
        
        print(f"   - First Record: {stats.get('first_record', 'N/A')}")
        print(f"   - Last Record: {stats.get('last_record', 'N/A')}")
        
        return True
        
    except Exception as e:
        print(f"✗ Error getting statistics: {e}")
        return False


def test_export_csv():
    """Test CSV export"""
    print("\n" + "=" * 70)
    print("TEST 7: EXPORT TO CSV")
    print("=" * 70)
    
    try:
        output_file = "test_export.csv"
        
        success = database.export_to_csv(
            output_file,
            db_path="test_stats_app.db"
        )
        
        if success and os.path.exists(output_file):
            print(f"✓ CSV export successful: {output_file}")
            
            # Check file content
            with open(output_file, 'r') as f:
                lines = f.readlines()
                print(f"   - File contains {len(lines)} lines")
            
            return True
        else:
            print("✗ CSV export failed")
            return False
            
    except Exception as e:
        print(f"✗ Error exporting CSV: {e}")
        return False


def test_export_json():
    """Test JSON export"""
    print("\n" + "=" * 70)
    print("TEST 8: EXPORT TO JSON")
    print("=" * 70)
    
    try:
        output_file = "test_export.json"
        
        success = database.export_full_to_json(
            output_file,
            db_path="test_stats_app.db"
        )
        
        if success and os.path.exists(output_file):
            print(f"✓ JSON export successful: {output_file}")
            
            # Check file size
            file_size = os.path.getsize(output_file)
            print(f"   - File size: {file_size} bytes")
            
            return True
        else:
            print("✗ JSON export failed")
            return False
            
    except Exception as e:
        print(f"✗ Error exporting JSON: {e}")
        return False


def test_delete_result(record_id):
    """Test deleting a result"""
    print("\n" + "=" * 70)
    print("TEST 9: DELETE RESULT")
    print("=" * 70)
    
    try:
        success = database.delete_result(record_id, db_path="test_stats_app.db")
        
        if success:
            print(f"✓ Record {record_id} deleted successfully")
            
            # Verify deletion
            record = database.get_result_by_id(record_id, db_path="test_stats_app.db")
            if record is None:
                print("✓ Deletion verified - record no longer exists")
                return True
            else:
                print("✗ Record still exists after deletion")
                return False
        else:
            print("✗ Delete operation failed")
            return False
            
    except Exception as e:
        print(f"✗ Error deleting result: {e}")
        return False


def test_multiple_saves():
    """Test saving multiple different analysis types"""
    print("\n" + "=" * 70)
    print("TEST 10: MULTIPLE ANALYSIS TYPES")
    print("=" * 70)
    
    try:
        analyses = [
            {
                'type': 'ANOVA',
                'input': {'groups': 3, 'n': 30},
                'result': {'F': 4.56, 'p': 0.023},
                'interpretation': 'Significant difference found'
            },
            {
                'type': 't-test',
                'input': {'n1': 15, 'n2': 15},
                'result': {'t': 2.34, 'p': 0.045},
                'interpretation': 'Significant difference'
            },
            {
                'type': 'Cronbach\'s Alpha',
                'input': {'items': 10, 'respondents': 100},
                'result': {'alpha': 0.85},
                'interpretation': 'Good reliability'
            }
        ]
        
        saved_ids = []
        
        for analysis in analyses:
            record_id = database.save_result(
                analysis_type=analysis['type'],
                input_data=analysis['input'],
                result=analysis['result'],
                interpretation=analysis['interpretation'],
                db_path="test_stats_app.db"
            )
            
            if record_id:
                saved_ids.append(record_id)
                print(f"✓ {analysis['type']} saved (ID: {record_id})")
            else:
                print(f"✗ Failed to save {analysis['type']}")
        
        print(f"\n✓ Saved {len(saved_ids)} of {len(analyses)} analyses")
        return len(saved_ids) == len(analyses)
        
    except Exception as e:
        print(f"✗ Error in multiple saves: {e}")
        return False


def cleanup_test_files():
    """Clean up test files"""
    print("\n" + "=" * 70)
    print("CLEANUP: REMOVING TEST FILES")
    print("=" * 70)
    
    files_to_remove = [
        "test_stats_app.db",
        "test_export.csv",
        "test_export.json"
    ]
    
    for filename in files_to_remove:
        if os.path.exists(filename):
            try:
                os.remove(filename)
                print(f"✓ Removed: {filename}")
            except Exception as e:
                print(f"✗ Failed to remove {filename}: {e}")


def run_all_tests():
    """Run all database tests"""
    print("\n")
    print("=" * 70)
    print("DATABASE INTEGRATION TEST SUITE")
    print("=" * 70)
    print(f"Start Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Track results
    results = []
    
    # Run tests
    results.append(("Initialization", test_initialization()))
    
    # Save a test record
    test_id = test_save_result()
    results.append(("Save Result", test_id is not None))
    
    if test_id:
        results.append(("Retrieve by ID", test_retrieve_result(test_id)))
    else:
        results.append(("Retrieve by ID", False))
    
    results.append(("Retrieve All", test_retrieve_all()))
    results.append(("Search", test_search()))
    results.append(("Statistics", test_database_stats()))
    results.append(("Export CSV", test_export_csv()))
    results.append(("Export JSON", test_export_json()))
    results.append(("Multiple Types", test_multiple_saves()))
    
    # Delete test should be last
    if test_id:
        results.append(("Delete Result", test_delete_result(test_id)))
    else:
        results.append(("Delete Result", False))
    
    # Summary
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✓ PASS" if result else "✗ FAIL"
        print(f"{status} - {test_name}")
    
    print("\n" + "-" * 70)
    print(f"Results: {passed}/{total} tests passed ({passed/total*100:.1f}%)")
    print("-" * 70)
    
    # Cleanup
    cleanup_test_files()
    
    print(f"\nEnd Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    
    # Return success status
    return passed == total


if __name__ == "__main__":
    success = run_all_tests()
    
    if success:
        print("\n🎉 All tests passed! Database system is working correctly.")
        sys.exit(0)
    else:
        print("\n⚠️ Some tests failed. Please review the output above.")
        sys.exit(1)