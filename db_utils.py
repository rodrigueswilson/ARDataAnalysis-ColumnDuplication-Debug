#!/usr/bin/env python3
"""
Unified Database Connection Utility for ARDataAnalysis Project

This module provides a single source of truth for MongoDB database connections
across all scripts in the ARDataAnalysis project.

Usage:
    from db_utils import get_db_connection
    
    db = get_db_connection()
    collection = db.media_records
"""

import pymongo
from pymongo import MongoClient
import sys
from typing import Optional

# SINGLE SOURCE OF TRUTH FOR DATABASE CONFIGURATION
DEFAULT_HOST = 'localhost'
DEFAULT_PORT = 27017
DEFAULT_DATABASE_NAME = 'ARDataAnalysis'  # Correct database name with actual data
DEFAULT_COLLECTION_NAME = 'media_records'

def _get_db_client(host: str = DEFAULT_HOST, port: int = DEFAULT_PORT, timeout_ms: int = 5000) -> MongoClient:
    """Establishes and tests a MongoDB client connection."""
    try:
        client = MongoClient(host=host, port=port, serverSelectionTimeoutMS=timeout_ms)
        client.admin.command('ping')
        return client
    except (pymongo.errors.ServerSelectionTimeoutError, pymongo.errors.ConnectionFailure) as e:
        error_msg = f"[ERROR] Failed to connect to MongoDB at {host}:{port}: {e}"
        print(error_msg, file=sys.stderr)
        raise ConnectionError(error_msg) from e

def get_db_connection(
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT, 
    db_name: str = DEFAULT_DATABASE_NAME,
    timeout_ms: int = 5000
) -> pymongo.database.Database:
    """Gets a MongoDB database object using a centralized client."""
    try:
        client = _get_db_client(host=host, port=port, timeout_ms=timeout_ms)
        database = client[db_name]
        print(f"[OK] Connected to MongoDB: {host}:{port}/{db_name}")
        return database
    except Exception as e:
        error_msg = f"[ERROR] Unexpected error getting database {db_name}: {e}"
        print(error_msg, file=sys.stderr)
        raise ConnectionError(error_msg) from e

def get_collection(
    collection_name: str = DEFAULT_COLLECTION_NAME,
    db_name: str = DEFAULT_DATABASE_NAME,
    **kwargs
) -> pymongo.collection.Collection:
    """Gets a MongoDB collection using a consistent database connection."""
    database = get_db_connection(db_name=db_name, **kwargs)
    return database[collection_name]

def verify_database_connection(db_name: str = DEFAULT_DATABASE_NAME) -> bool:
    """Verifies the database connection and checks for data."""
    try:
        db = get_db_connection(db_name=db_name)
        collection = db[DEFAULT_COLLECTION_NAME]
        document_count = collection.count_documents({})
        if document_count > 0:
            print(f"[OK] DB verification successful: {db_name}.{DEFAULT_COLLECTION_NAME} has {document_count:,} documents")
            return True
        else:
            print(f"[INFO] DB {db_name}.{DEFAULT_COLLECTION_NAME} exists but is empty.")
            return False
    except Exception as e:
        print(f"[ERROR] Database verification failed: {e}")
        return False

def list_available_databases() -> list:
    """Lists all available databases on the server."""
    client = None
    try:
        client = _get_db_client()
        databases = client.list_database_names()
        print(f"[INFO] Available databases on {DEFAULT_HOST}:{DEFAULT_PORT}:")
        for db_name in databases:
            if db_name not in ['admin', 'config', 'local']:
                try:
                    db = client[db_name]
                    collections = db.list_collection_names()
                    print(f"  - {db_name}: {len(collections)} collections")
                    for coll_name in collections:
                        count = db[coll_name].count_documents({})
                        print(f"    └── {coll_name}: {count:,} documents")
                except Exception as e:
                    print(f"    └── Error reading {db_name}: {e}")
        return databases
    except Exception as e:
        print(f"[ERROR] Failed to list databases: {e}")
        return []
    finally:
        if client:
            client.close()

if __name__ == "__main__":
    """
    Test the database connection utility.
    """
    print("=== Database Connection Utility Test ===")
    
    # Test basic connection
    print("\n1. Testing basic connection...")
    try:
        db = get_db_connection()
        print(f"   [OK] Successfully connected to {DEFAULT_DATABASE_NAME}")
    except Exception as e:
        print(f"   [ERROR] Connection failed: {e}")
    
    # Verify database
    print("\n2. Verifying database...")
    verify_database_connection()
    
    # List available databases
    print("\n3. Listing available databases...")
    list_available_databases()
    
    print("\n=== Test Complete ===")
