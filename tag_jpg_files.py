import os
import argparse
import datetime
from pathlib import Path
import piexif
from PIL import Image
import json

# Import shared utilities
from ar_utils import (
    get_school_calendar,
    get_non_collection_days,
    get_activity_schedule,
    precompute_collection_days,
    get_contextual_info
)



# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def find_jpg_files(root_dir):
    """Recursively finds all JPG files in the given directory."""
    return sorted(list(Path(root_dir).rglob('*.jpg'))) + sorted(list(Path(root_dir).rglob('*.jpeg')))

def get_image_datetime(file_path):
    """
    Extracts the 'DateTimeOriginal' from EXIF data. Falls back to filename.
    """
    try:
        exif_dict = piexif.load(str(file_path))
        datetime_original_bytes = exif_dict.get('Exif', {}).get(piexif.ExifIFD.DateTimeOriginal)
        if datetime_original_bytes:
            datetime_str = datetime_original_bytes.decode('utf-8')
            return datetime.datetime.strptime(datetime_str, '%Y:%m:%d %H:%M:%S')
    except Exception as e:
        # This can happen if the file is not a valid JPG or has no EXIF
        print(f"Info: Could not read EXIF datetime for {file_path.name}. Reason: {e}. Falling back to filename.")

    # Fallback to filename parsing if EXIF fails
    try:
        stem = file_path.stem
        parts = stem.split('_')
        if len(parts) >= 2:
            timestamp_str = f"{parts[0]}_{parts[1]}"
            return datetime.datetime.strptime(timestamp_str, "%y%m%d_%H%M%S")
    except (ValueError, IndexError):
        print(f"Warning: Could not parse datetime from filename for {file_path.name}")
        return None
    return None




def write_tags_to_jpg(file_path, tags_to_write):
    """
    Writes contextual tags to a JPG file's EXIF data.
    - UserComment: Stores all tags as a JSON object for full data retrieval.
    - XPComment: Stores a human-readable summary.
    """
    try:
        exif_dict = piexif.load(str(file_path))

        # --- Store all data as JSON in UserComment for machine readability ---
        # The UserComment tag requires a specific 8-byte prefix for encoding
        json_data = json.dumps(tags_to_write, indent=4)
        exif_dict['Exif'][piexif.ExifIFD.UserComment] = b'UNICODE\x00' + json_data.encode('utf-16-le')
        
        # --- Store a human-readable summary in XPComment ---
        summary_lines = [f"{k}: {v}" for k, v in tags_to_write.items()]
        summary_str = "\n".join(summary_lines)
        # XPComment (and other XP tags) use UTF-16LE encoding
        exif_dict['0th'][piexif.ImageIFD.XPComment] = summary_str.encode('utf-16-le')

        # Get the EXIF bytes and insert them into the file
        exif_bytes = piexif.dump(exif_dict)
        piexif.insert(exif_bytes, str(file_path))
        
        print(f"Successfully wrote tags to: {file_path.name}")
        return True

    except Exception as e:
        print(f"Error writing tags to {file_path.name}: {e}")
        return False

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    """Main function to drive the script."""
    parser = argparse.ArgumentParser(
        description='Adds custom school calendar metadata to JPG files based on their EXIF timestamps.',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('root_dir', type=str, help='The root directory to scan for JPG files.')
    args = parser.parse_args()

    if not os.path.isdir(args.root_dir):
        print(f"Error: Directory not found at {args.root_dir}")
        return

    # --- Load Configuration ---
    school_calendar = get_school_calendar()
    non_collection_days = get_non_collection_days()
    activity_schedule = get_activity_schedule()

    # --- Pre-computation Step for Efficiency ---
    print("Pre-computing collection day numbers...")
    collection_day_map = precompute_collection_days(school_calendar, non_collection_days)
    print(f"Mapped {len(collection_day_map)} total collection days.")

    # --- Find and Process Files ---
    print(f"\nScanning for JPG files in {args.root_dir}...")
    jpg_files = find_jpg_files(args.root_dir)
    print(f"Found {len(jpg_files)} JPG files to process.\n")

    processed_count = 0
    error_count = 0
    for file in jpg_files:
        print(f"--- Processing: {file.name} ---")
        
        # 1. Get the primary timestamp from EXIF data
        file_datetime = get_image_datetime(file)
        if not file_datetime:
            error_count += 1
            print(f"Skipping file due to missing timestamp: {file.name}\n")
            continue

        # 2. Check for outlier status
        is_outlier = 'outliers' in file.parts

        # 3. Get all contextual information
        context_info = get_contextual_info(
            file_datetime, school_calendar, 
            non_collection_days, activity_schedule, collection_day_map,
            is_outlier=is_outlier
        )
        if not context_info:
            error_count += 1
            print(f"Skipping file due to missing context: {file.name}\n")
            continue
        
        # Print the derived info for verification
        for key, val in context_info.items():
            print(f"  - {key}: {val}")

        # 4. Write all information as tags into the JPG file
        if write_tags_to_jpg(file, context_info):
            processed_count += 1
        else:
            error_count += 1
        print("-" * (len(file.name) + 18) + "\n")


    print("\n==================== SUMMARY ====================")
    print(f"Total files found: {len(jpg_files)}")
    print(f"Successfully processed and tagged: {processed_count}")
    print(f"Files with errors (skipped): {error_count}")
    print("==============================================")


if __name__ == '__main__':
    # Before running, ensure you have piexif installed:
    # pip install piexif
    main()
