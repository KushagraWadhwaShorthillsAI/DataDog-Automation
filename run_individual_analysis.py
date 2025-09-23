#!/usr/bin/env python3
"""
Run Individual Analysis
Simple script to analyze individual files with charts and metrics only
No comprehensive reports - focused on individual analysis
"""

import sys
import os

# Add scripts directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from simple_individual_analyzer import analyze_all_source_files, analyze_file


def main():
    print("🎯 INDIVIDUAL FILE ANALYSIS & COMBINED REPORTS")
    print("=" * 50)
    print("This script will:")
    print("  ✓ Analyze each file individually")
    print("  ✓ Generate DAU and DAUU charts")
    print("  ✓ Calculate response time & LLM cost metrics")
    print("  ✓ Show accurate error rates and breakdown")
    print("  ✓ Save individual analysis to TXT files")
    print("  ✓ Generate combined Excel and PDF reports")
    print()
    
    if len(sys.argv) > 1:
        # Analyze specific file
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
    else:
        # Analyze all files
        print("🔄 Analyzing all files in source_data directory...")
        analyze_all_source_files()
        
        # Generate combined reports
        print(f"\n🔄 Generating combined reports...")
        from final_combined_report import FinalPolishedCombinedReport
        generator = FinalPolishedCombinedReport()
        generator.generate_reports()
        
        return 0


if __name__ == "__main__":
    sys.exit(main())
