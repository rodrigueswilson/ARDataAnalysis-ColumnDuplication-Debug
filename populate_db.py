import os
import argparse
import datetime
import json
import time

from pathlib import Path

import mutagen
import piexif
from pymongo import MongoClient
from pymongo.errors import ConnectionFailure

# Import shared utilities
from ar_utils import (
    get_school_calendar,
    get_non_collection_days,
    precompute_collection_days,
    is_collection_day
)

print("[DEBUG] Script starting - imports completed...")

# Import unified database connection utility
from db_utils import get_db_connection

def find_media_files(root_dir):
    """Finds all JPG and MP3 files in media directories only."""
    print(f"Scanning for media files in: {root_dir}")
    
    # Define media directories to scan (only the four main directories)
    media_dirs = [
        '21_22 Photos', '22_23 Photos',  # Photo directories
        '21_22 Audio', '22_23 Audio'     # Audio directories
    ]
    
    jpg_files_set = set()
    mp3_files_set = set()
    
    for media_dir in media_dirs:
        media_path = Path(root_dir) / media_dir
        if media_path.exists() and media_path.is_dir():
            print(f"  Scanning {media_dir}...")
            
            # Search for JPG files (both cases)
            jpg_files_set.update(media_path.rglob('*.jpg'))
            jpg_files_set.update(media_path.rglob('*.JPG'))
            jpg_files_set.update(media_path.rglob('*.jpeg'))
            jpg_files_set.update(media_path.rglob('*.JPEG'))
            
            # Search for MP3 files (both cases)
            mp3_files_set.update(media_path.rglob('*.mp3'))
            mp3_files_set.update(media_path.rglob('*.MP3'))
        else:
            print(f"  Warning: {media_dir} not found, skipping...")
    
    jpg_files = sorted(list(jpg_files_set))
    mp3_files = sorted(list(mp3_files_set))
    
    print(f"Found {len(jpg_files)} JPG files and {len(mp3_files)} MP3 files.")
    return jpg_files, mp3_files

def extract_mp3_metadata(file_path, collection_day_map=None):
    """Extracts all relevant metadata from an MP3 file into a dictionary."""
    try:
        audio = mutagen.File(file_path)
        if not audio or not audio.tags:
            return None

        # Initialize a dictionary to hold our structured data
        doc = {
            "file_name": file_path.name,
            "file_path": str(file_path.resolve()),
            "file_type": "MP3",
            "_creation_timestamp": datetime.datetime.now(datetime.timezone.utc)
        }

        # Extract all TXXX (custom) tags
        for key in audio.tags:
            if key.startswith('TXXX:'):
                # TXXX tags are stored as a list, we take the first item
                tag_name = key.split(':', 1)[1]
                value = audio.tags[key].text[0]
                
                # Convert numeric-like strings to actual numbers
                try:
                    if '.' in value:
                        value = float(value)
                    else:
                        value = int(value)
                except (ValueError, TypeError):
                    pass # Keep as string if conversion fails
                
                # Convert boolean-like strings to actual booleans
                if isinstance(value, str):
                    if value.lower() == 'true':
                        value = True
                    elif value.lower() == 'false':
                        value = False

                doc[tag_name] = value

        # Extract MP3 file properties
        doc["File_Size_MB"] = round(file_path.stat().st_size / (1024 * 1024), 3)
        
        if hasattr(audio.info, 'length'):
            doc["Duration_Seconds"] = round(audio.info.length, 1)
        
        if hasattr(audio.info, 'bitrate'):
            # Convert to kbps for readability
            doc["Bitrate_kbps"] = round(audio.info.bitrate / 1000)
            
        if hasattr(audio.info, 'channels'):
            doc["Channels"] = audio.info.channels
            
        # Add is_collection_day field if we can determine it
        if collection_day_map is not None and "ISO_Date" in doc:
            try:
                # Convert ISO_Date string to date object for lookup
                date_str = doc.get("ISO_Date", "")
                if date_str:
                    date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                    doc["is_collection_day"] = is_collection_day(date_obj, collection_day_map)
            except Exception as e:
                print(f"Warning: Could not determine collection day status for {file_path.name}: {e}")
                doc["is_collection_day"] = False

        return doc
    except Exception as e:
        print(f"Error processing MP3 {file_path.name}: {e}")
        return None

def extract_jpg_metadata(file_path, collection_day_map=None):
    """Extracts all relevant metadata from a JPG file into a dictionary."""
    try:
        exif_dict = piexif.load(str(file_path))

        # The JSON data is stored in the UserComment field
        user_comment_bytes = exif_dict.get('Exif', {}).get(piexif.ExifIFD.UserComment, b'')
        
        # UserComment has an 8-byte encoding prefix we need to strip
        if user_comment_bytes.startswith(b'UNICODE\x00'):
            user_comment_bytes = user_comment_bytes[8:]
        
        # Decode from UTF-16LE and load the JSON
        try:
            json_str = user_comment_bytes.decode('utf-16-le')
            doc = json.loads(json_str)
        except (UnicodeDecodeError, json.JSONDecodeError):
            # If UserComment is not valid, we can't get the contextual data
            print(f"Warning: Could not decode UserComment for {file_path.name}. Skipping.")
            return None

        # Add file system and core EXIF info to the document
        doc["file_name"] = file_path.name
        doc["file_path"] = str(file_path.resolve())
        doc["file_type"] = "JPG"
        doc["_creation_timestamp"] = datetime.datetime.now(datetime.timezone.utc)
        
        # Add a few key technical details from standard EXIF tags
        zeroth_ifd = exif_dict.get('0th', {})
        exif_ifd = exif_dict.get('Exif', {})
        
        doc["Camera_Make"] = zeroth_ifd.get(piexif.ImageIFD.Make, b'').decode('utf-8', 'ignore').strip()
        doc["Camera_Model"] = zeroth_ifd.get(piexif.ImageIFD.Model, b'').decode('utf-8', 'ignore').strip()
        
        width = exif_ifd.get(piexif.ExifIFD.PixelXDimension)
        height = exif_ifd.get(piexif.ExifIFD.PixelYDimension)
        if width and height:
            doc["Image_Dimensions"] = f"{width}x{height}"
            
        doc["File_Size_MB"] = round(file_path.stat().st_size / (1024 * 1024), 3)
        
        # Add is_collection_day field if we can determine it
        if collection_day_map is not None and "ISO_Date" in doc:
            try:
                # Convert ISO_Date string to date object for lookup
                date_str = doc.get("ISO_Date", "")
                if date_str:
                    date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
                    doc["is_collection_day"] = is_collection_day(date_obj, collection_day_map)
            except Exception as e:
                print(f"Warning: Could not determine collection day status for {file_path.name}: {e}")
                doc["is_collection_day"] = False

        return doc
    except Exception as e:
        print(f"Error processing JPG {file_path.name}: {e}")
        return None


def main():
    """Main function to drive the database population."""
    print("[DEBUG] Starting DB population...")
    
    parser = argparse.ArgumentParser(
        description='Scans media files, extracts metadata, and populates a MongoDB database.',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        '--source_dir', 
        required=False, 
        help='The root directory to scan. Defaults to the script\'s location.'
    )
    parser.add_argument(
        '--db_path', 
        required=True, 
        help='The directory where MongoDB should store its data files.\nExample: D:\\ARDataAnalysis\\db'
    )
    args = parser.parse_args()
    print(f"[DEBUG] Arguments parsed: source_dir={args.source_dir}, db_path={args.db_path}")

    # Determine the source directory
    if args.source_dir:
        source_dir = args.source_dir
        print(f"[DEBUG] Using provided source directory: {source_dir}")
    else:
        # Default to the directory where the script is located
        source_dir = os.path.dirname(os.path.abspath(__file__))
        print(f"[DEBUG] --source_dir not provided. Defaulting to script directory: {source_dir}")
    
    # Precompute collection days for tagging
    print("Precomputing collection days from school calendar...")
    school_calendar = get_school_calendar()
    non_collection_days = get_non_collection_days()
    collection_day_map = precompute_collection_days(school_calendar, non_collection_days)
    print(f"Identified {len(collection_day_map)} valid collection days across all school periods.")

    # --- Connect to DB ---
    print("[DEBUG] Connecting to database...")
    db = get_db_connection()
    if db is None:
        print("[ERROR] Database connection failed - exiting")
        return # Exit if connection fails
    print("[DEBUG] Database connection successful")

    # --- Clear existing data ---
    collection_name = 'media_records'
    print(f"\nClearing existing data from collection: '{collection_name}'...")
    db[collection_name].delete_many({})
    print("Collection cleared.")

    # --- Process Files ---
    print("[DEBUG] Scanning for media files...")
    jpg_files, mp3_files = find_media_files(source_dir)
    all_files = jpg_files + mp3_files
    print(f"[DEBUG] Processing {len(jpg_files)} JPG and {len(mp3_files)} MP3 files (total: {len(all_files)})")
    
    if len(all_files) == 0:
        print("[WARNING] No media files found - check source directory!")
        return
    
    documents_to_insert = []
    processed_count = 0
    error_count = 0

    for file in all_files:
        print(f"[DEBUG] Reading metadata from: {file.name}")
        doc = None
        if file.suffix.lower() in ['.jpg', '.jpeg']:
            doc = extract_jpg_metadata(file, collection_day_map)
        elif file.suffix.lower() == '.mp3':
            doc = extract_mp3_metadata(file, collection_day_map)

        if doc:
            documents_to_insert.append(doc)
            processed_count += 1
        else:
            print(f"[WARNING] Skipped file: {file.name} (metadata extraction failed)")
            error_count += 1
    
    # --- Insert into MongoDB ---
    print(f"\n[DEBUG] Ready to insert {len(documents_to_insert)} documents into DB")
    if documents_to_insert:
        print(f"[DEBUG] Inserting {len(documents_to_insert)} documents into MongoDB...")
        db[collection_name].insert_many(documents_to_insert)
        print("[DEBUG] Database insertion complete!")
    else:
        print("[ERROR] No valid documents found to insert - check metadata extraction!")
        raise RuntimeError("No documents parsed -- check metadata extraction.")

    # --- Summary ---
    print("\n==================== SUMMARY ====================")
    print(f"Total files scanned: {len(all_files)}")
    print(f"Successfully processed and prepared for DB: {processed_count}")
    print(f"Files with errors (skipped): {error_count}")
    print("==============================================")




if __name__ == '__main__':
    # Before running, ensure you have the required libraries installed:
    # pip install pymongo mutagen piexif
    print("[DEBUG] About to call main() function...")
    main()
    print("[DEBUG] main() function completed!")
