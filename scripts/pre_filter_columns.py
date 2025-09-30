#!/usr/bin/env python3
"""
Pre-Filter Columns Script
Removes irrelevant columns and keeps only the specific required columns based on sheet-wise mapping.
"""

import os
import sys
import json
import pandas as pd
from typing import Dict, List, Optional, Tuple
import argparse
from pathlib import Path

# Add current directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from data_loaders import load_data_from_file


class ColumnPreFilter:
    """
    Pre-filter script to remove irrelevant columns and keep only required ones
    based on sheet-wise column mapping configuration.
    """
    
    def __init__(self, config_path: str = None):
        """
        Initialize the pre-filter with column mapping configuration.
        
        Args:
            config_path: Path to the column mapping configuration JSON file
        """
        if config_path is None:
            config_path = os.path.join(os.path.dirname(__file__), 'column_mapping_config.json')
        
        self.config_path = config_path
        self.column_mappings = self._load_config()
        
    def _load_config(self) -> Dict:
        """Load column mapping configuration from JSON file."""
        try:
            with open(self.config_path, 'r') as f:
                config = json.load(f)
            return config.get('sheet_column_mappings', {})
        except FileNotFoundError:
            print(f"❌ Configuration file not found: {self.config_path}")
            return {}
        except json.JSONDecodeError as e:
            print(f"❌ Error parsing configuration file: {e}")
            return {}
    
    def detect_sheet_type_from_data(self, df: pd.DataFrame) -> str:
        """
        Detect sheet type from the service/source column in the data.
        
        Args:
            df: DataFrame to analyze
            
        Returns:
            Detected sheet type name
        """
        # Look for service or source column
        service_col = None
        for col in df.columns:
            if col.lower() in ['service', 'source', '@service', '@source']:
                service_col = col
                break
        
        if service_col is None:
            print("⚠️  No service/source column found, using default mapping")
            return "QnA"  # Default fallback
        
        # Get unique values from service/source column
        unique_services = df[service_col].dropna().unique()
        print(f"🔍 Found services: {list(unique_services)}")
        
        # Map service values to sheet types
        service_mapping = {
            'qna': 'QnA',
            'search': 'Search', 
            'summary': 'Summary',
            'relevantdoc': 'RelevantDoc',
            'prepsubmission': 'PrepSubmission',
            'prep submission': 'PrepSubmission',
            'prepare submission': 'PrepSubmission'
        }
        
        # Check if any service matches our known types
        for service in unique_services:
            service_lower = str(service).lower().strip()
            for key, sheet_type in service_mapping.items():
                if key in service_lower:
                    print(f"✅ Detected sheet type: {sheet_type} (from service: {service})")
                    return sheet_type
        
        # If no match found, try to infer from column names
        columns_lower = [col.lower() for col in df.columns]
        
        # Check for celery-specific columns
        celery_columns = ['@processname', '@processname', '@processcreatedon', '@processstartedon', 
                          '@processcompletedon', '@requestid', '@totaltimetaken']
        
        if any(col in columns_lower for col in celery_columns):
            print("✅ Detected celery logs pattern - using Summary mapping")
            return "Summary"
        
        # Check for QnA-specific columns
        qna_columns = ['@requestpayload.mode', '@requestpayload.question', '@session_id', 
                       '@isautomode', '@redirectedmode', '@requestpayload.selectedguids']
        
        if any(col in columns_lower for col in qna_columns):
            print("✅ Detected QnA pattern")
            return "QnA"
        
        # Check for Search-specific columns
        search_columns = ['@websocket.url_details.path']
        
        if any(col in columns_lower for col in search_columns):
            print("✅ Detected Search pattern")
            return "Search"
        
        # Check for HTTP-specific columns (RelevantDoc)
        http_columns = ['@http.url_details.path', '@http.method', '@http.status_code']
        
        if any(col in columns_lower for col in http_columns):
            print("✅ Detected HTTP pattern - using RelevantDoc mapping")
            return "RelevantDoc"
        
        print("⚠️  Could not detect sheet type, using QnA as default")
        return "QnA"

    def get_required_columns(self, sheet_name: str) -> List[str]:
        """
        Get the list of required columns for a specific sheet.
        
        Args:
            sheet_name: Name of the sheet (e.g., 'QnA', 'Search', 'Summary')
            
        Returns:
            List of required column names
        """
        return self.column_mappings.get(sheet_name, [])
    
    def filter_columns(self, df: pd.DataFrame, sheet_name: str, 
                      strict_mode: bool = True) -> Tuple[pd.DataFrame, Dict]:
        """
        Filter DataFrame to keep only required columns for the specified sheet.
        
        Args:
            df: Input DataFrame
            sheet_name: Name of the sheet to determine required columns
            strict_mode: If True, only keep columns that exactly match required columns.
                        If False, also keep columns that contain required column names.
            
        Returns:
            Tuple of (filtered_dataframe, filtering_report)
        """
        required_columns = self.get_required_columns(sheet_name)
        
        if not required_columns:
            print(f"⚠️  No column mapping found for sheet: {sheet_name}")
            return df, {"status": "no_mapping", "kept_columns": list(df.columns)}
        
        available_columns = list(df.columns)
        filtering_report = {
            "sheet_name": sheet_name,
            "total_columns": len(available_columns),
            "required_columns": required_columns,
            "kept_columns": [],
            "removed_columns": [],
            "missing_columns": [],
            "status": "success"
        }
        
        if strict_mode:
            # Exact match mode - only keep columns that exactly match required columns
            kept_columns = []
            for col in available_columns:
                if col in required_columns:
                    kept_columns.append(col)
                else:
                    filtering_report["removed_columns"].append(col)
            
            # Check for missing required columns
            for req_col in required_columns:
                if req_col not in available_columns:
                    filtering_report["missing_columns"].append(req_col)
            
            filtering_report["kept_columns"] = kept_columns
            
        else:
            # Fuzzy match mode - keep columns that contain required column names
            kept_columns = []
            for col in available_columns:
                should_keep = False
                for req_col in required_columns:
                    if req_col.lower() in col.lower() or col.lower() in req_col.lower():
                        should_keep = True
                        break
                
                if should_keep:
                    kept_columns.append(col)
                else:
                    filtering_report["removed_columns"].append(col)
            
            filtering_report["kept_columns"] = kept_columns
        
        # Create filtered DataFrame
        if filtering_report["kept_columns"]:
            filtered_df = df[filtering_report["kept_columns"]].copy()
        else:
            print(f"⚠️  No columns matched for sheet: {sheet_name}")
            filtered_df = df.copy()
        
        return filtered_df, filtering_report
    
    def process_file(self, input_path: str, output_path: str = None, 
                    sheet_name: str = None, strict_mode: bool = True) -> bool:
        """
        Process a single file to filter columns.
        
        Args:
            input_path: Path to input file
            output_path: Path to output file (if None, creates filtered version in same directory)
            sheet_name: Name of the sheet (if None, tries to detect from filename)
            strict_mode: Whether to use strict column matching
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Load data
            print(f"📁 Loading data from: {input_path}")
            loader = load_data_from_file(input_path)
            df = loader.load_data()
            
            if df is None or df.empty:
                print(f"❌ No data found in file: {input_path}")
                return False
            
            # Determine sheet name if not provided
            if sheet_name is None:
                # Try to detect from data first
                sheet_name = self.detect_sheet_type_from_data(df)
            else:
                print(f"🔍 Using provided sheet name: {sheet_name}")
            
            print(f"🔍 Processing sheet: {sheet_name}")
            print(f"📊 Original columns: {len(df.columns)}")
            print(f"📋 Available columns: {list(df.columns)}")
            
            # Filter columns
            filtered_df, report = self.filter_columns(df, sheet_name, strict_mode)
            
            # Generate output path if not provided
            if output_path is None:
                input_path_obj = Path(input_path)
                output_path = input_path_obj.parent / f"{input_path_obj.stem}_filtered{input_path_obj.suffix}"
            
            # Save filtered data
            print(f"💾 Saving filtered data to: {output_path}")
            if input_path.endswith('.xlsx') or input_path.endswith('.xls'):
                filtered_df.to_excel(output_path, index=False)
            elif input_path.endswith('.csv'):
                filtered_df.to_csv(output_path, index=False)
            elif input_path.endswith('.json'):
                filtered_df.to_json(output_path, orient='records', indent=2)
            else:
                # Default to CSV
                output_path = str(output_path).replace('.xlsx', '.csv').replace('.xls', '.csv')
                filtered_df.to_csv(output_path, index=False)
            
            # Print filtering report
            self._print_filtering_report(report)
            
            return True
            
        except Exception as e:
            print(f"❌ Error processing file {input_path}: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _print_filtering_report(self, report: Dict):
        """Print a detailed filtering report."""
        print(f"\n📊 FILTERING REPORT")
        print(f"=" * 50)
        print(f"Sheet: {report['sheet_name']}")
        print(f"Total columns: {report['total_columns']}")
        print(f"Kept columns: {len(report['kept_columns'])}")
        print(f"Removed columns: {len(report['removed_columns'])}")
        print(f"Missing required columns: {len(report['missing_columns'])}")
        
        if report['kept_columns']:
            print(f"\n✅ Kept columns:")
            for col in report['kept_columns']:
                print(f"  ✓ {col}")
        
        if report['removed_columns']:
            print(f"\n❌ Removed columns:")
            for col in report['removed_columns']:
                print(f"  ✗ {col}")
        
        if report['missing_columns']:
            print(f"\n⚠️  Missing required columns:")
            for col in report['missing_columns']:
                print(f"  ? {col}")
        
        print(f"\n📈 Filtering efficiency: {len(report['kept_columns'])}/{report['total_columns']} columns kept ({len(report['kept_columns'])/report['total_columns']*100:.1f}%)")
    
    def process_directory(self, input_dir: str, output_dir: str = None, 
                         strict_mode: bool = True) -> Dict[str, bool]:
        """
        Process all files in a directory.
        
        Args:
            input_dir: Input directory path
            output_dir: Output directory path (if None, creates filtered files in same directory)
            strict_mode: Whether to use strict column matching
            
        Returns:
            Dictionary mapping file names to success status
        """
        input_path = Path(input_dir)
        if not input_path.exists():
            print(f"❌ Input directory not found: {input_dir}")
            return {}
        
        if output_dir:
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)
        
        results = {}
        supported_extensions = ['.xlsx', '.xls', '.csv', '.json']
        
        for file_path in input_path.iterdir():
            if file_path.is_file() and file_path.suffix.lower() in supported_extensions:
                print(f"\n🔄 Processing: {file_path.name}")
                
                if output_dir:
                    output_file_path = Path(output_dir) / f"{file_path.stem}_filtered{file_path.suffix}"
                else:
                    output_file_path = None
                
                success = self.process_file(
                    str(file_path), 
                    str(output_file_path) if output_file_path else None,
                    strict_mode=strict_mode
                )
                
                results[file_path.name] = success
        
        return results


def main():
    """Main function for command-line usage."""
    parser = argparse.ArgumentParser(description='Pre-filter columns based on sheet-wise mapping')
    parser.add_argument('input', help='Input file or directory path')
    parser.add_argument('-o', '--output', help='Output file or directory path')
    parser.add_argument('-s', '--sheet', help='Sheet name (if not provided, will be detected from filename)')
    parser.add_argument('--strict', action='store_true', help='Use strict column matching (exact match only)')
    parser.add_argument('--fuzzy', action='store_true', help='Use fuzzy column matching (partial match)')
    parser.add_argument('--config', help='Path to column mapping configuration file')
    
    args = parser.parse_args()
    
    # Initialize pre-filter
    prefilter = ColumnPreFilter(args.config)
    
    # Determine matching mode
    strict_mode = args.strict or not args.fuzzy
    
    input_path = Path(args.input)
    
    if input_path.is_file():
        # Process single file
        print(f"🎯 PRE-FILTERING SINGLE FILE")
        print(f"=" * 50)
        success = prefilter.process_file(
            args.input, 
            args.output, 
            args.sheet, 
            strict_mode
        )
        
        if success:
            print(f"\n✅ File processed successfully!")
        else:
            print(f"\n❌ File processing failed!")
            sys.exit(1)
    
    elif input_path.is_dir():
        # Process directory
        print(f"🎯 PRE-FILTERING DIRECTORY")
        print(f"=" * 50)
        results = prefilter.process_directory(
            args.input, 
            args.output, 
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
    
    else:
        print(f"❌ Input path not found: {args.input}")
        sys.exit(1)


if __name__ == "__main__":
    main()
