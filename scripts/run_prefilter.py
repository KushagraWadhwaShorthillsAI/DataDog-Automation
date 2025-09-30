#!/usr/bin/env python3
"""
Run Pre-Filter Script
Simple wrapper script to run the column pre-filtering process.
"""

import sys
import os
from pathlib import Path

# Add scripts directory to path
sys.path.append(os.path.join(os.path.dirname(__file__)))

from pre_filter_columns import ColumnPreFilter


def print_usage():
    """Print usage instructions."""
    print("🎯 COLUMN PRE-FILTERING")
    print("=" * 50)
    print("This script removes irrelevant columns and keeps only the specific required columns")
    print("based on sheet-wise mapping configuration.")
    print()
    print("Usage:")
    print("  python run_prefilter.py <input_path> [options]")
    print()
    print("Parameters:")
    print("  input_path  : Path to input file or directory")
    print("  -o, --output : Output file or directory path (optional)")
    print("  -s, --sheet  : Sheet name (optional, auto-detected from filename)")
    print("  --strict     : Use strict column matching (exact match only)")
    print("  --fuzzy      : Use fuzzy column matching (partial match)")
    print("  --config     : Path to column mapping configuration file")
    print()
    print("Examples:")
    print("  # Filter single file (auto-detect sheet name)")
    print("  python run_prefilter.py source_data/QnA.xlsx")
    print()
    print("  # Filter single file with specific sheet name")
    print("  python run_prefilter.py source_data/QnA.xlsx -s QnA")
    print()
    print("  # Filter single file with custom output")
    print("  python run_prefilter.py source_data/QnA.xlsx -o filtered_data/QnA_filtered.xlsx")
    print()
    print("  # Filter entire directory")
    print("  python run_prefilter.py source_data/ -o filtered_data/")
    print()
    print("  # Use fuzzy matching (partial column name matching)")
    print("  python run_prefilter.py source_data/QnA.xlsx --fuzzy")
    print()
    print("  # Use custom configuration file")
    print("  python run_prefilter.py source_data/QnA.xlsx --config custom_mapping.json")
    print()
    print("Supported file formats:")
    print("  ✓ Excel files (.xlsx, .xls)")
    print("  ✓ CSV files (.csv)")
    print("  ✓ JSON files (.json)")
    print()
    print("Configuration:")
    print("  The script uses 'column_mapping_config.json' by default.")
    print("  This file contains the required columns for each sheet:")
    print("  - QnA, Search, Summary, RelevantDoc, PrepSubmission, celery logs")
    print()


def main():
    """Main function."""
    if len(sys.argv) < 2:
        print_usage()
        return 1
    
    # Check if help is requested
    if sys.argv[1] in ['-h', '--help', 'help']:
        print_usage()
        return 0
    
    # Parse command line arguments manually for simplicity
    input_path = sys.argv[1]
    output_path = None
    sheet_name = None
    strict_mode = True
    config_path = None
    
    # Parse additional arguments
    i = 2
    while i < len(sys.argv):
        arg = sys.argv[i]
        
        if arg in ['-o', '--output'] and i + 1 < len(sys.argv):
            output_path = sys.argv[i + 1]
            i += 2
        elif arg in ['-s', '--sheet'] and i + 1 < len(sys.argv):
            sheet_name = sys.argv[i + 1]
            i += 2
        elif arg == '--strict':
            strict_mode = True
            i += 1
        elif arg == '--fuzzy':
            strict_mode = False
            i += 1
        elif arg == '--config' and i + 1 < len(sys.argv):
            config_path = sys.argv[i + 1]
            i += 2
        else:
            print(f"❌ Unknown argument: {arg}")
            print_usage()
            return 1
    
    # Validate input path
    if not os.path.exists(input_path):
        print(f"❌ Input path not found: {input_path}")
        return 1
    
    # Initialize pre-filter
    try:
        prefilter = ColumnPreFilter(config_path)
    except Exception as e:
        print(f"❌ Error initializing pre-filter: {e}")
        return 1
    
    # Process based on input type
    input_path_obj = Path(input_path)
    
    if input_path_obj.is_file():
        # Process single file
        print(f"🎯 PRE-FILTERING SINGLE FILE")
        print(f"=" * 50)
        print(f"Input: {input_path}")
        if output_path:
            print(f"Output: {output_path}")
        if sheet_name:
            print(f"Sheet: {sheet_name}")
        print(f"Mode: {'Strict' if strict_mode else 'Fuzzy'}")
        print()
        
        success = prefilter.process_file(
            input_path, 
            output_path, 
            sheet_name, 
            strict_mode
        )
        
        if success:
            print(f"\n✅ File processed successfully!")
            return 0
        else:
            print(f"\n❌ File processing failed!")
            return 1
    
    elif input_path_obj.is_dir():
        # Process directory
        print(f"🎯 PRE-FILTERING DIRECTORY")
        print(f"=" * 50)
        print(f"Input Directory: {input_path}")
        if output_path:
            print(f"Output Directory: {output_path}")
        print(f"Mode: {'Strict' if strict_mode else 'Fuzzy'}")
        print()
        
        results = prefilter.process_directory(
            input_path, 
            output_path, 
            strict_mode
        )
        
        successful = sum(1 for success in results.values() if success)
        total = len(results)
        
        print(f"\n📊 PROCESSING SUMMARY")
        print(f"=" * 30)
        print(f"Total files: {total}")
        print(f"Successful: {successful}")
        print(f"Failed: {total - successful}")
        
        if total - successful > 0:
            print(f"\n❌ Failed files:")
            for filename, success in results.items():
                if not success:
                    print(f"  ✗ {filename}")
        
        return 0 if successful == total else 1
    
    else:
        print(f"❌ Invalid input path: {input_path}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
