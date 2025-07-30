import os
import argparse
import datetime
from pathlib import Path
import mutagen
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3NoHeaderError

# Import shared utilities
from ar_utils import (
    get_school_calendar,
    get_non_collection_days,
    get_activity_schedule,
    precompute_collection_days,
    get_contextual_info,
    seconds_to_iso_duration,
    seconds_to_hms
)



# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def find_mp3_files(root_dir):
    """Recursively finds all MP3 files in the given directory."""
    return sorted(list(Path(root_dir).rglob('*.mp3')))

def parse_filename_to_datetime(file_path):
    """
    Extracts date and time from a filename like 'YYMMDD_HHMM.mp3' or 'YYMMDD_HHMM_SS.mp3'.
    Returns a datetime object or None if the format is incorrect.
    """
    try:
        stem = file_path.stem
        parts = stem.split('_')
        if len(parts) < 2:
            return None
        timestamp_str = f"{parts[0]}_{parts[1]}"
        return datetime.datetime.strptime(timestamp_str, "%y%m%d_%H%M")
    except (ValueError, IndexError):
        print(f"Warning: Could not parse datetime from filename: {file_path.name}")
        return None



def write_tags_to_mp3(file_path, tags_to_write):
    """
    Writes a dictionary of tags to an MP3 file's metadata.
    """
    try:
        audio = mutagen.File(file_path)
        if audio is None:
            print(f"Error: Could not load file {file_path.name}")
            return False

        if not audio.tags:
            audio.add_tags()

        dt_obj = parse_filename_to_datetime(file_path)
        if dt_obj:
            audio.tags.add(mutagen.id3.TDRC(encoding=3, text=dt_obj.strftime('%Y-%m-%dT%H:%M:%S')))

        duration_sec = tags_to_write.get('Duration_Seconds')
        if duration_sec is not None:
            duration_ms = str(int(float(duration_sec) * 1000))
            audio.tags.add(mutagen.id3.TLEN(encoding=3, text=duration_ms))

        for key, value in tags_to_write.items():
            audio.tags.add(mutagen.id3.TXXX(encoding=3, desc=key, text=str(value)))

        audio.save()
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
        description='Adds custom school calendar metadata to MP3 files based on their filenames.',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('root_dir', type=str, help='The root directory to scan for MP3 files.')
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
    print(f"\nScanning for MP3 files in {args.root_dir}...")
    mp3_files = find_mp3_files(args.root_dir)
    print(f"Found {len(mp3_files)} MP3 files to process.\n")

    processed_count = 0
    error_count = 0
    for file in mp3_files:
        print(f"--- Processing: {file.name} ---")
        
        # 1. Get all file system and audio info first
        audio_properties = {
            'duration': 0.0, 'file_size': 0, 'bitrate': 0, 'channels': 0, 'is_outlier': False
        }
        try:
            audio_file = mutagen.File(file)
            if audio_file and audio_file.info:
                audio_properties['duration'] = audio_file.info.length
                # Bitrate is in bps, convert to kbps
                audio_properties['bitrate'] = int(audio_file.info.bitrate / 1000)
                audio_properties['channels'] = audio_file.info.channels
            audio_properties['file_size'] = file.stat().st_size
            # Check for outlier status based on path
            audio_properties['is_outlier'] = 'outliers' in file.parts
        except Exception as e:
            print(f"Warning: Could not read file info for {file.name}: {e}")
            error_count += 1
            continue

        # 2. Parse filename to get the base datetime
        file_datetime = parse_filename_to_datetime(file)
        if not file_datetime:
            error_count += 1
            continue

        # 3. Get all contextual information based on all file properties
        context_info = get_contextual_info(
            file_datetime, school_calendar, non_collection_days, 
            activity_schedule, collection_day_map, audio_props=audio_properties
        )
        if not context_info:
            error_count += 1
            continue
        
        # Print the derived info for verification
        for key, val in context_info.items():
            print(f"  - {key}: {val}")

        # 4. Write all information as tags into the MP3 file
        if write_tags_to_mp3(file, context_info):
            processed_count += 1
        else:
            error_count += 1
        print("-" * (len(file.name) + 18) + "\n")


    print("\n==================== SUMMARY ====================")
    print(f"Total files found: {len(mp3_files)}")
    print(f"Successfully processed and tagged: {processed_count}")
    print(f"Files with errors (skipped): {error_count}")
    print("==============================================")


if __name__ == '__main__':
    main()
