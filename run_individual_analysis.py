#!/usr/bin/env python3
"""
Run Individual Analysis
Simple script to analyze individual files with charts and metrics only
No comprehensive reports - focused on individual analysis
"""

import sys
import os
from datetime import datetime

# Add scripts directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from simple_individual_analyzer import analyze_all_source_files, analyze_file


def print_usage():
    """Print usage instructions"""
    print("🎯 INDIVIDUAL FILE ANALYSIS")
    print("=" * 50)
    print("Usage:")
    print("  python run_individual_analysis.py                    # Analyze all files")
    print("  python run_individual_analysis.py <file_name>        # Analyze specific file")
    print("  python run_individual_analysis.py <file_name> <start_date> <end_date>  # Analyze with date filtering")
    print()
    print("Parameters:")
    print("  file_name  : Name of the file in source_data directory (e.g., 'QnA.xlsx')")
    print("  start_date : Start date for analysis (format: YYYY-MM-DD) - OPTIONAL")
    print("  end_date   : End date for analysis (format: YYYY-MM-DD) - OPTIONAL")
    print()
    print("Examples:")
    print("  python run_individual_analysis.py                    # Analyze all files")
    print("  python run_individual_analysis.py QnA.xlsx           # Analyze QnA.xlsx")
    print("  python run_individual_analysis.py data.csv            # Analyze data.csv (will convert to XLSX)")
    print("  python run_individual_analysis.py QnA.xlsx 2024-01-01 2024-01-31  # Analyze with date filtering")
    print()
    print("This script will:")
    print("  ✓ Analyze each file individually")
    print("  ✓ Convert CSV files to XLSX format automatically")
    print("  ✓ Generate DAU and DAUU charts")
    print("  ✓ Calculate response time & LLM cost metrics")
    print("  ✓ Show accurate error rates and breakdown")
    print("  ✓ Save individual analysis to TXT files")
    print("  ✓ Generate combined Excel and PDF reports")
    print()


def validate_date(date_string):
    """Validate date format"""
    try:
        datetime.strptime(date_string, '%Y-%m-%d')
        return True
    except ValueError:
        return False


def main():
    print("🎯 INDIVIDUAL FILE ANALYSIS & COMBINED REPORTS")
    print("=" * 50)
    print("This script will:")
    print("  ✓ Analyze each file individually")
    print("  ✓ Convert CSV files to XLSX format automatically")
    print("  ✓ Generate DAU and DAUU charts")
    print("  ✓ Calculate response time & LLM cost metrics")
    print("  ✓ Show accurate error rates and breakdown")
    print("  ✓ Save individual analysis to TXT files")
    print("  ✓ Generate combined Excel and PDF reports")
    print()
    
    if len(sys.argv) == 1:
        # Analyze all files (normal analysis)
        print("🔄 Analyzing all files in source_data directory...")
        analyze_all_source_files()
        
        # Generate combined reports
        print(f"\n🔄 Generating combined reports...")
        from final_combined_report import FinalPolishedCombinedReport
        generator = FinalPolishedCombinedReport()
        generator.generate_reports()
        
        return 0
        
    elif len(sys.argv) == 2:
        # Analyze specific file (normal analysis)
        file_path = sys.argv[1]
        if not os.path.isabs(file_path):
            # Assume it's in source_data directory
            file_path = f"/Users/shtlpmac027/Documents/DataDog/source_data/{file_path}"
        
        if os.path.exists(file_path):
            print(f"🔄 Analyzing specific file: {os.path.basename(file_path)}")
            success = analyze_file(file_path)
            
            if success:
                # Generate combined reports
                print(f"\n🔄 Generating combined reports...")
                from final_combined_report import FinalPolishedCombinedReport
                generator = FinalPolishedCombinedReport()
                generator.generate_reports()
            
            return 0 if success else 1
        else:
            print(f"❌ File not found: {file_path}")
            return 1
            
    elif len(sys.argv) == 4:
        # Analyze specific file with date filtering
        file_name = sys.argv[1]
        start_date = sys.argv[2]
        end_date = sys.argv[3]
        
        # Validate date formats
        if not validate_date(start_date):
            print(f"❌ Error: Invalid start date format: {start_date}")
            print("   Expected format: YYYY-MM-DD")
            return 1
        
        if not validate_date(end_date):
            print(f"❌ Error: Invalid end date format: {end_date}")
            print("   Expected format: YYYY-MM-DD")
            return 1
        
        # Construct file path
        if not os.path.isabs(file_name):
            file_path = f"/Users/shtlpmac027/Documents/DataDog/source_data/{file_name}"
        else:
            file_path = file_name
        
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"❌ Error: File not found: {file_path}")
            print("   Please check the file name and ensure it exists in source_data directory")
            return 1
        
        # Validate date range
        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
        end_dt = datetime.strptime(end_date, '%Y-%m-%d')
        
        if start_dt > end_dt:
            print(f"❌ Error: Start date ({start_date}) cannot be after end date ({end_date})")
            return 1
        
        print("🎯 INDIVIDUAL FILE ANALYSIS WITH DATE FILTERING")
        print("=" * 60)
        print(f"📁 File: {os.path.basename(file_path)}")
        print(f"📅 Date Range: {start_date} to {end_date}")
        print()
        
        # Analyze file with date parameters
        print(f"🔄 Analyzing file with date filtering...")
        success = analyze_file(file_path, compare=(start_date, end_date))
        
        if success:
            # Generate combined reports
            print(f"\n🔄 Generating combined reports...")
            from final_combined_report import FinalPolishedCombinedReport
            generator = FinalPolishedCombinedReport()
            generator.generate_reports()
            print("✅ Analysis completed successfully!")
        else:
            print("❌ Analysis failed!")
            return 1
        
        return 0
        
    else:
        # Invalid number of arguments
        print_usage()
        print("❌ Error: Incorrect number of arguments provided.")
        print(f"   Expected: 0, 1, or 3 arguments, Got: {len(sys.argv) - 1}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
