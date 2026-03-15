"""
Database Module for Statistical Analysis Application
Provides SQLite database operations for storing analysis history

Features:
- Complete CRUD operations
- Parameterized queries for SQL injection prevention
- Comprehensive error handling
- Connection management with context managers
- Export functionality to CSV

Author: Statistical Analysis Suite
Version: 1.0.0
"""

import sqlite3
import json
import csv
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any
import os


class DatabaseManager:
    """
    Manages all database operations for the statistical analysis application.
    Uses context managers for safe connection handling.
    """
    
    def __init__(self, db_path: str = "stats_app.db"):
        """
        Initialize the database manager.
        
        Parameters:
        -----------
        db_path : str
            Path to the SQLite database file (default: "stats_app.db")
        """
        self.db_path = db_path
        self.connection = None
        
    def __enter__(self):
        """Context manager entry - opens database connection"""
        self.connection = sqlite3.connect(self.db_path)
        self.connection.row_factory = sqlite3.Row  # Enable column access by name
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - closes database connection"""
        if self.connection:
            if exc_type is None:
                self.connection.commit()
            else:
                self.connection.rollback()
            self.connection.close()
        return False
    
    def get_connection(self) -> sqlite3.Connection:
        """
        Get a database connection.
        Creates new connection if one doesn't exist.
        
        Returns:
        --------
        sqlite3.Connection : Database connection object
        """
        if self.connection is None:
            self.connection = sqlite3.connect(self.db_path)
            self.connection.row_factory = sqlite3.Row
        return self.connection
    
    def close_connection(self):
        """Close the database connection if open"""
        if self.connection:
            self.connection.close()
            self.connection = None


# ============================================================================
# DATABASE INITIALIZATION
# ============================================================================

def connect_db(db_path: str = "stats_app.db") -> sqlite3.Connection:
    """
    Connect to the SQLite database.
    Creates the database file if it doesn't exist.
    
    Parameters:
    -----------
    db_path : str
        Path to the database file
        
    Returns:
    --------
    sqlite3.Connection : Database connection object
    
    Raises:
    -------
    sqlite3.Error : If connection fails
    """
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row  # Enable accessing columns by name
        print(f"✓ Database connected: {db_path}")
        return conn
    except sqlite3.Error as e:
        print(f"✗ Database connection error: {e}")
        raise


def create_tables(db_path: str = "stats_app.db") -> bool:
    """
    Create all necessary tables in the database.
    
    Creates:
    - analysis_history: Stores all statistical analysis results
    
    Parameters:
    -----------
    db_path : str
        Path to the database file
        
    Returns:
    --------
    bool : True if successful, False otherwise
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        # Create analysis_history table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS analysis_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                analysis_type TEXT NOT NULL,
                input_data TEXT NOT NULL,
                result TEXT NOT NULL,
                interpretation TEXT,
                date_created TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                metadata TEXT
            )
        """)
        
        # Create indexes for faster queries
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_analysis_type 
            ON analysis_history(analysis_type)
        """)
        
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_date_created 
            ON analysis_history(date_created DESC)
        """)
        
        conn.commit()
        conn.close()
        
        print("✓ Database tables created successfully")
        return True
        
    except sqlite3.Error as e:
        print(f"✗ Error creating tables: {e}")
        return False


# ============================================================================
# CRUD OPERATIONS
# ============================================================================

def save_result(
    analysis_type: str,
    input_data: Any,
    result: Any,
    interpretation: str = "",
    metadata: Dict = None,
    db_path: str = "stats_app.db"
) -> Optional[int]:
    """
    Save an analysis result to the database.
    
    Parameters:
    -----------
    analysis_type : str
        Type of analysis (e.g., "ANOVA", "Cronbach's Alpha", "t-test")
    input_data : Any
        Input data used for the analysis (will be JSON serialized)
    result : Any
        Analysis results (will be JSON serialized)
    interpretation : str, optional
        Human-readable interpretation of the results
    metadata : Dict, optional
        Additional metadata (e.g., user name, settings used)
    db_path : str
        Path to the database file
        
    Returns:
    --------
    Optional[int] : ID of the inserted record, or None if failed
    """
    try:
        # Serialize complex data structures to JSON
        input_data_json = json.dumps(input_data, default=str)
        result_json = json.dumps(result, default=str)
        metadata_json = json.dumps(metadata or {}, default=str)
        
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        # Use parameterized query to prevent SQL injection
        cursor.execute("""
            INSERT INTO analysis_history 
            (analysis_type, input_data, result, interpretation, metadata)
            VALUES (?, ?, ?, ?, ?)
        """, (analysis_type, input_data_json, result_json, interpretation, metadata_json))
        
        record_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        print(f"✓ Analysis saved successfully (ID: {record_id})")
        return record_id
        
    except (sqlite3.Error, json.JSONDecodeError) as e:
        print(f"✗ Error saving result: {e}")
        return None


def get_all_results(
    db_path: str = "stats_app.db",
    limit: int = None
) -> List[Dict]:
    """
    Retrieve all analysis results from the database.
    
    Parameters:
    -----------
    db_path : str
        Path to the database file
    limit : int, optional
        Maximum number of results to return
        
    Returns:
    --------
    List[Dict] : List of analysis records as dictionaries
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        query = "SELECT * FROM analysis_history ORDER BY date_created DESC"
        if limit:
            query += f" LIMIT {limit}"
        
        cursor.execute(query)
        rows = cursor.fetchall()
        conn.close()
        
        # Convert Row objects to dictionaries
        results = []
        for row in rows:
            record = dict(row)
            # Parse JSON fields
            try:
                record['input_data'] = json.loads(record['input_data'])
                record['result'] = json.loads(record['result'])
                record['metadata'] = json.loads(record.get('metadata', '{}'))
            except json.JSONDecodeError:
                pass  # Keep as string if JSON parsing fails
            results.append(record)
        
        return results
        
    except sqlite3.Error as e:
        print(f"✗ Error retrieving results: {e}")
        return []


def get_result_by_id(
    record_id: int,
    db_path: str = "stats_app.db"
) -> Optional[Dict]:
    """
    Retrieve a specific analysis result by ID.
    
    Parameters:
    -----------
    record_id : int
        ID of the record to retrieve
    db_path : str
        Path to the database file
        
    Returns:
    --------
    Optional[Dict] : Analysis record as dictionary, or None if not found
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            "SELECT * FROM analysis_history WHERE id = ?",
            (record_id,)
        )
        
        row = cursor.fetchone()
        conn.close()
        
        if row:
            record = dict(row)
            # Parse JSON fields
            try:
                record['input_data'] = json.loads(record['input_data'])
                record['result'] = json.loads(record['result'])
                record['metadata'] = json.loads(record.get('metadata', '{}'))
            except json.JSONDecodeError:
                pass
            return record
        else:
            print(f"✗ No record found with ID: {record_id}")
            return None
            
    except sqlite3.Error as e:
        print(f"✗ Error retrieving result: {e}")
        return None


def delete_result(
    record_id: int,
    db_path: str = "stats_app.db"
) -> bool:
    """
    Delete an analysis result from the database.
    
    Parameters:
    -----------
    record_id : int
        ID of the record to delete
    db_path : str
        Path to the database file
        
    Returns:
    --------
    bool : True if successful, False otherwise
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        # Check if record exists
        cursor.execute(
            "SELECT id FROM analysis_history WHERE id = ?",
            (record_id,)
        )
        
        if cursor.fetchone() is None:
            print(f"✗ No record found with ID: {record_id}")
            conn.close()
            return False
        
        # Delete the record
        cursor.execute(
            "DELETE FROM analysis_history WHERE id = ?",
            (record_id,)
        )
        
        conn.commit()
        conn.close()
        
        print(f"✓ Record {record_id} deleted successfully")
        return True
        
    except sqlite3.Error as e:
        print(f"✗ Error deleting result: {e}")
        return False


# ============================================================================
# SEARCH & FILTER OPERATIONS
# ============================================================================

def search_by_analysis_type(
    analysis_type: str,
    db_path: str = "stats_app.db"
) -> List[Dict]:
    """
    Search for results by analysis type.
    
    Parameters:
    -----------
    analysis_type : str
        Type of analysis to search for (e.g., "ANOVA", "t-test")
    db_path : str
        Path to the database file
        
    Returns:
    --------
    List[Dict] : List of matching analysis records
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT * FROM analysis_history 
            WHERE analysis_type LIKE ?
            ORDER BY date_created DESC
        """, (f"%{analysis_type}%",))
        
        rows = cursor.fetchall()
        conn.close()
        
        results = []
        for row in rows:
            record = dict(row)
            try:
                record['input_data'] = json.loads(record['input_data'])
                record['result'] = json.loads(record['result'])
                record['metadata'] = json.loads(record.get('metadata', '{}'))
            except json.JSONDecodeError:
                pass
            results.append(record)
        
        return results
        
    except sqlite3.Error as e:
        print(f"✗ Error searching results: {e}")
        return []


def search_by_date_range(
    start_date: str,
    end_date: str,
    db_path: str = "stats_app.db"
) -> List[Dict]:
    """
    Search for results within a date range.
    
    Parameters:
    -----------
    start_date : str
        Start date in format 'YYYY-MM-DD'
    end_date : str
        End date in format 'YYYY-MM-DD'
    db_path : str
        Path to the database file
        
    Returns:
    --------
    List[Dict] : List of analysis records within the date range
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT * FROM analysis_history 
            WHERE date(date_created) BETWEEN date(?) AND date(?)
            ORDER BY date_created DESC
        """, (start_date, end_date))
        
        rows = cursor.fetchall()
        conn.close()
        
        results = []
        for row in rows:
            record = dict(row)
            try:
                record['input_data'] = json.loads(record['input_data'])
                record['result'] = json.loads(record['result'])
                record['metadata'] = json.loads(record.get('metadata', '{}'))
            except json.JSONDecodeError:
                pass
            results.append(record)
        
        return results
        
    except sqlite3.Error as e:
        print(f"✗ Error searching by date: {e}")
        return []


def get_analysis_types(db_path: str = "stats_app.db") -> List[str]:
    """
    Get a list of all unique analysis types in the database.
    
    Parameters:
    -----------
    db_path : str
        Path to the database file
        
    Returns:
    --------
    List[str] : List of unique analysis types
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT DISTINCT analysis_type 
            FROM analysis_history 
            ORDER BY analysis_type
        """)
        
        types = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        return types
        
    except sqlite3.Error as e:
        print(f"✗ Error retrieving analysis types: {e}")
        return []


# ============================================================================
# EXPORT OPERATIONS
# ============================================================================

def export_to_csv(
    output_file: str,
    analysis_type: str = None,
    db_path: str = "stats_app.db"
) -> bool:
    """
    Export analysis results to a CSV file.
    
    Parameters:
    -----------
    output_file : str
        Path to the output CSV file
    analysis_type : str, optional
        Filter by analysis type (exports all if None)
    db_path : str
        Path to the database file
        
    Returns:
    --------
    bool : True if successful, False otherwise
    """
    try:
        # Get results
        if analysis_type:
            results = search_by_analysis_type(analysis_type, db_path)
        else:
            results = get_all_results(db_path)
        
        if not results:
            print("✗ No results to export")
            return False
        
        # Write to CSV
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['id', 'analysis_type', 'date_created', 'interpretation']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
            
            writer.writeheader()
            for result in results:
                # Simplify for CSV export
                simplified = {
                    'id': result['id'],
                    'analysis_type': result['analysis_type'],
                    'date_created': result['date_created'],
                    'interpretation': result['interpretation']
                }
                writer.writerow(simplified)
        
        print(f"✓ Exported {len(results)} results to {output_file}")
        return True
        
    except (sqlite3.Error, IOError) as e:
        print(f"✗ Error exporting to CSV: {e}")
        return False


def export_full_to_json(
    output_file: str,
    db_path: str = "stats_app.db"
) -> bool:
    """
    Export all analysis results to a JSON file with complete data.
    
    Parameters:
    -----------
    output_file : str
        Path to the output JSON file
    db_path : str
        Path to the database file
        
    Returns:
    --------
    bool : True if successful, False otherwise
    """
    try:
        results = get_all_results(db_path)
        
        if not results:
            print("✗ No results to export")
            return False
        
        with open(output_file, 'w', encoding='utf-8') as jsonfile:
            json.dump(results, jsonfile, indent=2, default=str)
        
        print(f"✓ Exported {len(results)} results to {output_file}")
        return True
        
    except (sqlite3.Error, IOError, json.JSONDecodeError) as e:
        print(f"✗ Error exporting to JSON: {e}")
        return False


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def get_database_stats(db_path: str = "stats_app.db") -> Dict:
    """
    Get statistics about the database.
    
    Parameters:
    -----------
    db_path : str
        Path to the database file
        
    Returns:
    --------
    Dict : Dictionary containing database statistics
    """
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        # Total records
        cursor.execute("SELECT COUNT(*) FROM analysis_history")
        total_records = cursor.fetchone()[0]
        
        # Records by type
        cursor.execute("""
            SELECT analysis_type, COUNT(*) as count 
            FROM analysis_history 
            GROUP BY analysis_type
            ORDER BY count DESC
        """)
        by_type = dict(cursor.fetchall())
        
        # Date range
        cursor.execute("""
            SELECT MIN(date_created), MAX(date_created) 
            FROM analysis_history
        """)
        date_range = cursor.fetchone()
        
        # Database file size
        db_size = os.path.getsize(db_path) if os.path.exists(db_path) else 0
        
        conn.close()
        
        return {
            'total_records': total_records,
            'by_type': by_type,
            'first_record': date_range[0] if date_range[0] else 'N/A',
            'last_record': date_range[1] if date_range[1] else 'N/A',
            'database_size_kb': round(db_size / 1024, 2)
        }
        
    except sqlite3.Error as e:
        print(f"✗ Error getting database stats: {e}")
        return {}


def clear_all_data(db_path: str = "stats_app.db", confirm: bool = False) -> bool:
    """
    Clear all data from the database (USE WITH CAUTION).
    
    Parameters:
    -----------
    db_path : str
        Path to the database file
    confirm : bool
        Must be True to execute (safety check)
        
    Returns:
    --------
    bool : True if successful, False otherwise
    """
    if not confirm:
        print("✗ Confirmation required to clear all data")
        return False
    
    try:
        conn = connect_db(db_path)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM analysis_history")
        cursor.execute("DELETE FROM sqlite_sequence WHERE name='analysis_history'")
        
        conn.commit()
        conn.close()
        
        print("✓ All data cleared from database")
        return True
        
    except sqlite3.Error as e:
        print(f"✗ Error clearing data: {e}")
        return False


# ============================================================================
# INITIALIZATION FUNCTION
# ============================================================================

def initialize_database(db_path: str = "stats_app.db") -> bool:
    """
    Initialize the database system.
    Creates database file and tables if they don't exist.
    
    Parameters:
    -----------
    db_path : str
        Path to the database file
        
    Returns:
    --------
    bool : True if successful, False otherwise
    """
    try:
        print("=" * 70)
        print("DATABASE INITIALIZATION")
        print("=" * 70)
        
        # Check if database exists
        db_exists = os.path.exists(db_path)
        
        if db_exists:
            print(f"✓ Database file found: {db_path}")
        else:
            print(f"→ Creating new database: {db_path}")
        
        # Create tables
        success = create_tables(db_path)
        
        if success:
            # Get database stats
            stats = get_database_stats(db_path)
            print(f"\n📊 Database Statistics:")
            print(f"   Total Records: {stats.get('total_records', 0)}")
            print(f"   Database Size: {stats.get('database_size_kb', 0)} KB")
            
            if stats.get('by_type'):
                print(f"\n   Records by Analysis Type:")
                for atype, count in stats['by_type'].items():
                    print(f"      • {atype}: {count}")
            
            print("\n" + "=" * 70)
            print("✓ Database system ready")
            print("=" * 70 + "\n")
            return True
        else:
            return False
            
    except Exception as e:
        print(f"✗ Database initialization failed: {e}")
        return False


# ============================================================================
# MAIN (FOR TESTING)
# ============================================================================

if __name__ == "__main__":
    """Test the database module"""
    
    print("\n" + "=" * 70)
    print("DATABASE MODULE TEST")
    print("=" * 70 + "\n")
    
    # Initialize
    initialize_database()
    
    # Test save
    test_result = save_result(
        analysis_type="Test Analysis",
        input_data={"values": [1, 2, 3, 4, 5]},
        result={"mean": 3.0, "std": 1.41},
        interpretation="Test interpretation",
        metadata={"user": "test_user"}
    )
    
    if test_result:
        print(f"\n✓ Test record saved with ID: {test_result}")
        
        # Test retrieve
        record = get_result_by_id(test_result)
        if record:
            print(f"✓ Retrieved record: {record['analysis_type']}")
        
        # Test delete
        delete_result(test_result)
    
    # Show stats
    stats = get_database_stats()
    print(f"\n📊 Final Database Statistics:")
    print(f"   Total Records: {stats.get('total_records', 0)}")
    
    print("\n" + "=" * 70)
    print("✓ Database module test complete")
    print("=" * 70 + "\n")