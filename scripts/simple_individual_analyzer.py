#!/usr/bin/env python3
"""
Simple Individual File Analyzer
Focus on individual analysis with two charts and metrics table only
No Excel/PDF reports - just charts and TXT metrics
"""

import os
import shutil
import sys
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import matplotlib.dates as mdates
from typing import Dict, List, Optional, Tuple
import traceback

# Add current directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from data_loaders import load_data_from_file


class SimpleIndividualAnalyzer:
    """Simple analyzer for individual files - charts and metrics only"""
    
    def __init__(self, file_path: str, compare_dates: Optional[Tuple[str, str]] = None):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path).split('.')[0]
        self.file_extension = os.path.splitext(file_path)[1].lower()
        self.df = None
        self.original_df = None  # Keep original for accurate error counts
        self.compare_dates: Optional[Tuple[str, str]] = compare_dates
        
        # Set up directory paths
        self.base_dir = "/Users/shtlpmac027/Documents/DataDog"
        self.output_dir = f"{self.base_dir}/individual_analysis/{self.file_name}"
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Column mappings
        self.column_mappings = {
            'date': None,
            'status': None,
            'response_time': None,
            'uuid': None,
            'llm_cost': None,
            'message': None,
            # Newly mapped columns
            'service': None,
            'process_name': None,
            'request_payload_mode': None,
            'redirected_mode': None
        }
        
        # Set up plotting style
        plt.style.use('default')
        sns.set_palette("husl")
        
    def _normalize_service_name(self, raw: str) -> str:
        """Normalize a service value for safe folder naming."""
        try:
            name = str(raw).strip().lower()
            # Replace separators with underscores
            for ch in [' ', '/', '\\', ':', '|', '*', '?', '"', '<', '>', '.']:
                name = name.replace(ch, '_')
            # Collapse multiple underscores
            while '__' in name:
                name = name.replace('__', '_')
            return name.strip('_') or self.file_name
        except Exception:
            return self.file_name

    def _maybe_update_output_dir_with_service(self):
        """Switch output directory to be based on Service column if present."""
        try:
            service_col = self.column_mappings.get('service')
            if not service_col or service_col not in self.df.columns:
                return
            # Get most frequent non-empty service value
            series = self.df[service_col].astype(str).str.strip()
            series = series[series != '']
            if series.empty:
                return
            top_service = series.value_counts().idxmax()
            normalized = self._normalize_service_name(top_service)
            new_dir = f"{self.base_dir}/individual_analysis/{normalized}"
            current_dir = self.output_dir
            # If only case differs (macOS case-insensitive), do a two-step rename to force case change
            if current_dir.lower() == new_dir.lower() and current_dir != new_dir:
                tmp_dir = f"{new_dir}__tmp_casefix__"
                try:
                    os.rename(current_dir, tmp_dir)
                    os.rename(tmp_dir, new_dir)
                    self.output_dir = new_dir
                    print(f"✓ Output folder case-aligned to: {self.output_dir}")
                    return
                except Exception:
                    pass
            # If folder path differs, move/rename existing output into the service-based one
            if new_dir != current_dir:
                try:
                    # If destination does not exist or is empty, move whole dir; otherwise merge files
                    if not os.path.exists(new_dir) or not os.listdir(new_dir):
                        shutil.move(current_dir, new_dir)
                    else:
                        for name in os.listdir(current_dir):
                            src = os.path.join(current_dir, name)
                            dst = os.path.join(new_dir, name)
                            if os.path.exists(dst):
                                continue
                            shutil.move(src, dst)
                        # remove old if empty
                        try:
                            os.rmdir(current_dir)
                        except Exception:
                            pass
                    self.output_dir = new_dir
                    print(f"✓ Output directory set to service-based folder: {self.output_dir}")
                except Exception as e:
                    # Fallback: ensure new dir exists and switch for subsequent writes
                    os.makedirs(new_dir, exist_ok=True)
                    self.output_dir = new_dir
                    print(f"⚠️  Switched to service-based folder (partial move): {self.output_dir} — {e}")
        except Exception as e:
            print(f"⚠️  Could not set service-based output directory: {e}")

    def load_and_detect_columns(self) -> bool:
        """Load data and detect column mappings"""
        try:
            print(f"Loading {self.file_extension} file: {self.file_name}")
            self.df = load_data_from_file(self.file_path)
            
            if self.df is None or self.df.empty:
                print("❌ No data found in file")
                return False
                
            # Store original data for accurate error counting
            self.original_df = self.df.copy()
            
            print(f"✅ Data loaded successfully! Shape: {self.df.shape}")
            
            # Detect columns
            self._detect_columns()
            return True
            
        except Exception as e:
            print(f"❌ Failed to load data: {e}")
            traceback.print_exc()
            return False
    
    def _detect_columns(self):
        """Detect column mappings using rule-based approach"""
        columns = list(self.df.columns)
        columns_lower = [col.lower() for col in columns]
        
        print(f"Available columns: {columns}")
        
        # Detection patterns
        detection_patterns = {
            'date': ['date', 'timestamp', '@timestamp', 'time', 'datetime'],
            'status': ['status', '@status', 'response_status', 'result'],
            'response_time': ['responsetime', 'response_time', 'totaltimetaken', 'total_time_taken', 
                            'duration', 'elapsed', 'time_taken', 'timetaken'],
            'uuid': ['useruuid', 'user_uuid', 'uuid', 'userid', 'user_id', 'clientid', 'client_id'],
            'llm_cost': ['meta.totalllmcost', 'totalllmcost', 'llmcost', 'totalcost', 'meta_totalllmcost',
                        'meta.total_llm_cost', 'total_llm_cost'],
            'message': ['message', 'requestpayload.message', 'requestpayloadmessage', 'error_message', '@message'],
            # Additional columns
            'service': ['service', 'service_name', '@service', 'servicename', 'source', 'source_name', '@source', 'sourcename'],
            'process_name': ['processname', 'process_name'],
            'request_payload_mode': ['requestpayloadmode', 'request_payload_mode', 'requestpayload.mode', 'resquestpayloadmode'],
            'redirected_mode': ['redirectedmode', 'redirect_mode', 'redirectionmode']
        }
        
        for mapping_key, patterns in detection_patterns.items():
            for pattern in patterns:
                for i, col in enumerate(columns_lower):
                    if pattern in col.replace(' ', '').replace('_', '').replace('.', ''):
                        self.column_mappings[mapping_key] = columns[i]
                        break
                if self.column_mappings[mapping_key]:
                    break
        
        # Special handling for message column - prefer 'Message' over '@Message' if both exist
        if 'Message' in columns and '@Message' in columns:
            self.column_mappings['message'] = 'Message'
            print(f"⚠️  Found both 'Message' and '@Message' columns, using 'Message'")
        
        print(f"Column mappings: {self.column_mappings}")
    
    def _categorize_error_messages_properly(self, error_messages: List[str]) -> Dict[str, int]:
        """Properly categorize error messages and count occurrences"""
        categories = {}
        
        # Count how many times each category appears
        for error_msg in error_messages:
            error_str = str(error_msg).lower()
            
            # Categorize based on message content
            if 'timeout' in error_str or 'timed out' in error_str or 'time out' in error_str:
                category = 'Timeout Errors'
            elif 'connection' in error_str or 'connect' in error_str or 'network' in error_str or 'socket' in error_str:
                category = 'Network/Connection Errors'
            elif 'auth' in error_str or 'permission' in error_str or 'unauthorized' in error_str or 'forbidden' in error_str:
                category = 'Authentication/Authorization Errors'
            elif 'not found' in error_str or '404' in error_str or 'missing' in error_str or 'no results' in error_str or 'contains no results' in error_str:
                category = 'Resource Not Found Errors'
            elif 'invalid data payload' in error_str or 'validation' in error_str or 'invalid' in error_str or 'bad request' in error_str or 'payload' in error_str:
                category = 'Data Validation/Payload Errors'
            elif 'internal server error' in error_str or 'server error' in error_str or '500' in error_str:
                category = 'Internal Server Errors'
            elif 'litellm' in error_str or 'llm' in error_str or 'summarize_document' in error_str:
                category = 'LLM Service Errors'
            elif 'query' in error_str or 'params' in error_str or 'parameter' in error_str or 'filtertype' in error_str:
                category = 'Query/Parameter Errors'
            elif 'exception' in error_str or 'baseexception' in error_str:
                category = 'Application Exception Errors'
            elif 'model mapping' in error_str or 'fetch' in error_str:
                category = 'Service Configuration Errors'
            elif 'json' in error_str or 'parse' in error_str or 'format' in error_str:
                category = 'Data Format Errors'
            else:
                category = 'Other/Uncategorized Errors'
            
            # Count occurrences
            if category in categories:
                categories[category] += 1
            else:
                categories[category] = 1
        
        return categories
    
    def _categorize_error_messages(self, error_messages: List[str]) -> Dict[str, int]:
        """Categorize error messages into types and count occurrences"""
        categories = {}
        
        for error_msg in error_messages:
            error_str = str(error_msg).lower()
            
            # Enhanced categorization with better patterns
            if 'timeout' in error_str or 'timed out' in error_str or 'time out' in error_str:
                category = 'Timeout Errors'
            elif 'connection' in error_str or 'connect' in error_str or 'network' in error_str or 'socket' in error_str:
                category = 'Network/Connection Errors'
            elif 'auth' in error_str or 'permission' in error_str or 'unauthorized' in error_str or 'forbidden' in error_str:
                category = 'Authentication/Authorization Errors'
            elif 'not found' in error_str or '404' in error_str or 'missing' in error_str or 'no results' in error_str or 'contains no results' in error_str:
                category = 'Resource Not Found Errors'
            elif 'invalid data payload' in error_str or 'validation' in error_str or 'invalid' in error_str or 'bad request' in error_str or 'payload' in error_str:
                category = 'Data Validation/Payload Errors'
            elif 'internal server error' in error_str or 'server error' in error_str or '500' in error_str:
                category = 'Internal Server Errors'
            elif 'litellm' in error_str or 'llm' in error_str or 'summarize_document' in error_str:
                category = 'LLM Service Errors'
            elif 'query' in error_str or 'params' in error_str or 'parameter' in error_str or 'filtertype' in error_str:
                category = 'Query/Parameter Errors'
            elif 'exception' in error_str or 'baseexception' in error_str:
                category = 'Application Exception Errors'
            elif 'model mapping' in error_str or 'fetch' in error_str:
                category = 'Service Configuration Errors'
            elif 'json' in error_str or 'parse' in error_str or 'format' in error_str:
                category = 'Data Format Errors'
            else:
                category = 'Other/Uncategorized Errors'
            
            # Count occurrences in the error breakdown
            if hasattr(self, 'df') and self.column_mappings.get('message'):
                message_col = self.column_mappings['message']
                status_col = self.column_mappings.get('status')
                if status_col:
                    error_df = self.df[self.df[status_col].str.lower() == 'error']
                    count = len(error_df[error_df[message_col] == error_msg])
                    
                    if category in categories:
                        categories[category] += count
                    else:
                        categories[category] = count
        
        return categories
    
    def preprocess_data(self) -> bool:
        """Basic preprocessing while keeping original for accurate error counts"""
        try:
            print("\n🔧 PREPROCESSING DATA")
            print("=" * 30)
            
            original_count = len(self.df)
            original_columns = len(self.df.columns)
            
            # Remove columns that are completely NaN
            columns_before = self.df.columns.tolist()
            self.df = self.df.dropna(axis=1, how='all')
            columns_after = self.df.columns.tolist()
            
            removed_columns = set(columns_before) - set(columns_after)
            if removed_columns:
                print(f"✓ Removed {len(removed_columns)} completely NaN columns: {list(removed_columns)}")
            else:
                print(f"✓ No completely NaN columns found")

            # Drop rows with blank/NaN service (service identifies the type of service)
            service_col = self.column_mappings.get('service')
            if service_col and service_col in self.df.columns:
                before_service = len(self.df)
                # Normalize service strings and drop empties
                self.df[service_col] = self.df[service_col].astype(str).str.strip()
                self.df = self.df[self.df[service_col].notna() & (self.df[service_col] != '')].copy()
                after_service = len(self.df)
                print(f"✓ Dropped rows with blank service: {before_service - after_service}")
            
            # Add formatted date column
            if self.column_mappings['date']:
                date_col = self.column_mappings['date']
                self.df[date_col] = pd.to_datetime(self.df[date_col], errors='coerce')
                self.df = self.df.dropna(subset=[date_col])  # Remove invalid dates
                self.df['formatted_date'] = self.df[date_col].dt.strftime('%Y-%m-%d')
                print(f"✓ Added formatted_date column")
            
            # Remove weekend data
            if self.column_mappings['date']:
                date_col = self.column_mappings['date']
                weekdays_mask = self.df[date_col].dt.weekday < 5
                self.df = self.df[weekdays_mask].copy()
                print(f"✓ Removed weekend data")
            
            # Filter status to only 'info' and 'error'
            status_col = self.column_mappings.get('status')
            if status_col and status_col in self.df.columns:
                before_status = len(self.df)
                status_series = self.df[status_col].astype(str).str.strip().str.lower()
                keep_mask = status_series.isin(['info', 'error'])
                self.df = self.df[keep_mask].copy()
                after_status = len(self.df)
                print(f"✓ Filtered status to ['info','error']: removed {before_status - after_status}")
            
            # Remove response time outliers (0-2000ms range)
            if self.column_mappings['response_time']:
                rt_col = self.column_mappings['response_time']
                self.df[rt_col] = pd.to_numeric(self.df[rt_col], errors='coerce')
                valid_rt_mask = (
                    (self.df[rt_col] >= 0) & 
                    (self.df[rt_col] <= 2000) & 
                    (self.df[rt_col].notna())
                )
                self.df = self.df[valid_rt_mask].copy()
                print(f"✓ Removed response time outliers")

            # Compute effective mode if mode columns exist (for QnA sheet)
            req_mode_col = self.column_mappings.get('request_payload_mode')
            redir_mode_col = self.column_mappings.get('redirected_mode')
            if req_mode_col and req_mode_col in self.df.columns:
                # Coerce to numeric where possible
                req_mode_series = pd.to_numeric(self.df[req_mode_col], errors='coerce')
                redir_mode_series = pd.to_numeric(self.df[redir_mode_col], errors='coerce') if redir_mode_col and redir_mode_col in self.df.columns else pd.Series([np.nan] * len(self.df), index=self.df.index)

                def _compute_effective(rm, rd):
                    if pd.isna(rm):
                        return np.nan
                    if int(rm) == 11:
                        # Redirected to 2 or 7 if present; else 0
                        if not pd.isna(rd) and int(rd) in (2, 7):
                            return int(rd)
                        return 0
                    return int(rm)

                self.df['effective_mode'] = [
                    _compute_effective(rm, rd) for rm, rd in zip(req_mode_series, redir_mode_series)
                ]
                print("✓ Computed effective_mode column")
            
            final_count = len(self.df)
            removed = original_count - final_count
            
            print(f"Original records: {original_count:,}")
            print(f"Final records: {final_count:,}")
            print(f"Records removed: {removed:,} ({removed/original_count*100:.1f}%)")
            
            return True
            
        except Exception as e:
            print(f"Error in preprocessing: {e}")
            return False
    
    def calculate_metrics(self) -> Dict:
        """Calculate all key metrics from preprocessed data"""
        metrics = {}
        
        try:
            # Basic counts from PREPROCESSED data (as requested)
            original_total = len(self.original_df)
            processed_total = len(self.df)
            
            print(f"\n📊 CALCULATING METRICS")
            print("=" * 30)
            
            # Status analysis from PREPROCESSED data
            status_col = self.column_mappings.get('status')
            if status_col and status_col in self.df.columns:
                # Count from preprocessed data
                processed_status_counts = self.df[status_col].str.lower().value_counts()
                processed_success = processed_status_counts.get('info', 0)
                processed_errors = processed_status_counts.get('error', 0)
                
                # Calculate rates based on PREPROCESSED data
                success_rate = (processed_success / processed_total * 100) if processed_total > 0 else 0
                error_rate = (processed_errors / processed_total * 100) if processed_total > 0 else 0
                
                metrics['status_analysis'] = {
                    'original_total': original_total,
                    'processed_total': processed_total,
                    'processed_success': processed_success,
                    'processed_errors': processed_errors,
                    'success_rate': success_rate,
                    'error_rate': error_rate
                }
                
                print(f"✓ Status Analysis (from preprocessed data):")
                print(f"  Processed Total: {processed_total:,}")
                print(f"  Processed Errors: {processed_errors:,} ({error_rate:.2f}%)")
                print(f"  Processed Success: {processed_success:,} ({success_rate:.2f}%)")
                
                # Error breakdown and categorization from preprocessed data
                if processed_errors > 0 and self.column_mappings.get('message'):
                    message_col = self.column_mappings['message']
                    
                    # STEP 1: Filter rows where status = 'error'
                    error_rows = self.df[self.df[status_col].str.lower() == 'error']
                    print(f"  Found {len(error_rows)} rows with status='error'")
                    
                    if not error_rows.empty and message_col in error_rows.columns:
                        # STEP 2: Get messages from error rows only
                        error_messages = error_rows[message_col].dropna()
                        print(f"  Found {len(error_messages)} error messages")
                        
                        if len(error_messages) > 0:
                            # STEP 3: Count each unique error message
                            error_counts = error_messages.value_counts()
                            metrics['error_breakdown'] = error_counts.to_dict()
                            
                            # STEP 4: Categorize based on actual error messages
                            error_categories = self._categorize_error_messages_properly(error_messages.tolist())
                            metrics['error_categories'] = error_categories
                            
                            print(f"  Error message types: {len(error_counts)}")
                            print(f"  Error categories: {len(error_categories)}")
                            
                            # Debug: Show actual error messages found
                            print(f"  Sample error messages:")
                            for i, (msg, count) in enumerate(error_counts.head(3).items()):
                                print(f"    {i+1}. '{msg}' (count: {count})")
                        else:
                            print(f"  No error messages found in message column")
                    else:
                        print(f"  Error rows found but no message column or empty")
            
            # Response time analysis from PROCESSED data
            rt_col = self.column_mappings.get('response_time')
            if rt_col and rt_col in self.df.columns:
                rt_data = pd.to_numeric(self.df[rt_col], errors='coerce').dropna()
                if len(rt_data) > 0:
                    metrics['response_time'] = {
                        'mean': rt_data.mean(),
                        'median': rt_data.median(),
                        'min': rt_data.min(),
                        'max': rt_data.max(),
                        'std': rt_data.std(),
                        'count': len(rt_data)
                    }
                    print(f"✓ Response Time Analysis: {len(rt_data):,} records")
            
            # LLM cost analysis from PROCESSED data
            cost_col = self.column_mappings.get('llm_cost')
            if cost_col and cost_col in self.df.columns:
                cost_data = pd.to_numeric(self.df[cost_col], errors='coerce').dropna()
                if len(cost_data) > 0:
                    metrics['llm_cost'] = {
                        'mean': cost_data.mean(),
                        'median': cost_data.median(),
                        'min': cost_data.min(),
                        'max': cost_data.max(),
                        'total': cost_data.sum(),
                        'std': cost_data.std(),
                        'count': len(cost_data)
                    }
                    print(f"✓ LLM Cost Analysis: {len(cost_data):,} records")

            # Process-wise metrics (for PrepareSubmission-like sheets)
            process_col = self.column_mappings.get('process_name')
            if process_col and process_col in self.df.columns:
                rt_col = self.column_mappings.get('response_time')
                if rt_col and rt_col in self.df.columns:
                    df_proc_rt = self.df[[process_col, rt_col]].copy()
                    df_proc_rt[rt_col] = pd.to_numeric(df_proc_rt[rt_col], errors='coerce')
                    proc_rt = df_proc_rt.dropna().groupby(process_col)[rt_col].agg(['mean','median','min','max','std','count']).sort_values('mean')
                    metrics['response_time_by_process'] = proc_rt.reset_index().to_dict(orient='records')
                    print(f"✓ Computed response time by process: {len(proc_rt)} rows")
                if cost_col and cost_col in self.df.columns:
                    df_proc_cost = self.df[[process_col, cost_col]].copy()
                    df_proc_cost[cost_col] = pd.to_numeric(df_proc_cost[cost_col], errors='coerce')
                    proc_cost = df_proc_cost.dropna().groupby(process_col)[cost_col].agg(['mean','median','min','max','sum','count']).rename(columns={'sum':'total'}).sort_values('total', ascending=False)
                    metrics['llm_cost_by_process'] = proc_cost.reset_index().to_dict(orient='records')
                    print(f"✓ Computed LLM cost by process: {len(proc_cost)} rows")
                # Failure table by process
                status_col = self.column_mappings.get('status')
                if status_col and status_col in self.df.columns:
                    df_proc_status = self.df[[process_col, status_col]].copy()
                    df_proc_status[status_col] = df_proc_status[status_col].astype(str).str.strip().str.lower()
                    pvt = df_proc_status.pivot_table(index=process_col, columns=status_col, aggfunc='size', fill_value=0)
                    if 'error' not in pvt.columns: pvt['error'] = 0
                    if 'info' not in pvt.columns: pvt['info'] = 0
                    pvt = pvt[['error','info']].reset_index()
                    pvt['total'] = pvt['error'] + pvt['info']
                    pvt['failure_pct'] = (pvt['error'] / pvt['total'] * 100).fillna(0)
                    metrics['failure_by_process'] = pvt.to_dict(orient='records')
                    print(f"✓ Computed failure rates by process: {len(pvt)} rows")

            # Effective mode-wise metrics (for QnA-like sheets)
            if 'effective_mode' in self.df.columns:
                # Mode name mapping
                mode_map = {
                    1: 'isDocument', 2: 'isInternet', 3: 'isDatabase', 4: 'isDirectTaxCode', 5: 'isGlobal',
                    6: 'isHarvey', 7: 'isDatabaseGeneric', 8: 'isNLP', 9: 'isDeepResearch', 10: 'isDraft',
                    11: 'isAutoMode', 12: 'isMultipleDbGeneric', 13: 'isDatabaseGenericVersion2', 14: 'isDatabaseGenericLite',
                    15: 'isDeepResearchWebSearch', 0: 'UnresolvedRedirect'
                }
                # Response time by effective mode
                rt_col = self.column_mappings.get('response_time')
                if rt_col and rt_col in self.df.columns:
                    df_mode_rt = self.df[['effective_mode', rt_col]].copy()
                    df_mode_rt[rt_col] = pd.to_numeric(df_mode_rt[rt_col], errors='coerce')
                    mode_rt = df_mode_rt.dropna().groupby('effective_mode')[rt_col].agg(['mean','median','min','max','std','count']).sort_values('mean')
                    mode_rt = mode_rt.reset_index()
                    mode_rt['mode_name'] = mode_rt['effective_mode'].apply(lambda m: mode_map.get(int(m), str(int(m)) if not pd.isna(m) else 'Unknown'))
                    metrics['response_time_by_effective_mode'] = mode_rt.to_dict(orient='records')
                    print(f"✓ Computed response time by effective mode: {len(mode_rt)} rows")
                # LLM cost by effective mode
                if cost_col and cost_col in self.df.columns:
                    df_mode_cost = self.df[['effective_mode', cost_col]].copy()
                    df_mode_cost[cost_col] = pd.to_numeric(df_mode_cost[cost_col], errors='coerce')
                    mode_cost = df_mode_cost.dropna().groupby('effective_mode')[cost_col].agg(['mean','median','min','max','sum','count']).rename(columns={'sum':'total'}).sort_values('total', ascending=False)
                    mode_cost = mode_cost.reset_index()
                    mode_cost['mode_name'] = mode_cost['effective_mode'].apply(lambda m: mode_map.get(int(m), str(int(m)) if not pd.isna(m) else 'Unknown'))
                    metrics['llm_cost_by_effective_mode'] = mode_cost.to_dict(orient='records')
                    print(f"✓ Computed LLM cost by effective mode: {len(mode_cost)} rows")
                # Failure table by effective mode
                status_col = self.column_mappings.get('status')
                if status_col and status_col in self.df.columns:
                    df_mode_status = self.df[['effective_mode', status_col]].copy()
                    df_mode_status[status_col] = df_mode_status[status_col].astype(str).str.strip().str.lower()
                    pivot = df_mode_status.pivot_table(index='effective_mode', columns=status_col, aggfunc='size', fill_value=0)
                    pivot = pivot.rename(columns={'error':'error', 'info':'info'})
                    if 'error' not in pivot.columns: pivot['error'] = 0
                    if 'info' not in pivot.columns: pivot['info'] = 0
                    pivot = pivot[['error', 'info']]
                    pivot = pivot.reset_index()
                    pivot['total'] = pivot['error'] + pivot['info']
                    pivot['failure_pct'] = (pivot['error'] / pivot['total'] * 100).fillna(0)
                    pivot['mode_name'] = pivot['effective_mode'].apply(lambda m: mode_map.get(int(m), str(int(m)) if not pd.isna(m) else 'Unknown'))
                    metrics['failure_by_effective_mode'] = pivot.to_dict(orient='records')
                    print(f"✓ Computed failure rates by effective mode: {len(pivot)} rows")

            # Process x Mode combined metrics (when both exist)
            process_col = self.column_mappings.get('process_name')
            if process_col and process_col in self.df.columns and 'effective_mode' in self.df.columns:
                # Response time by process x mode
                rt_col = self.column_mappings.get('response_time')
                if rt_col and rt_col in self.df.columns:
                    df_pm_rt = self.df[[process_col, 'effective_mode', rt_col]].copy()
                    df_pm_rt[rt_col] = pd.to_numeric(df_pm_rt[rt_col], errors='coerce')
                    pm_rt = df_pm_rt.dropna().groupby([process_col, 'effective_mode'])[rt_col].agg(['mean','median','min','max','std','count']).reset_index()
                    metrics['response_time_by_process_mode'] = pm_rt.to_dict(orient='records')
                    print(f"✓ Computed response time by process x mode: {len(pm_rt)} rows")
                # LLM cost by process x mode
                if cost_col and cost_col in self.df.columns:
                    df_pm_cost = self.df[[process_col, 'effective_mode', cost_col]].copy()
                    df_pm_cost[cost_col] = pd.to_numeric(df_pm_cost[cost_col], errors='coerce')
                    pm_cost = df_pm_cost.dropna().groupby([process_col, 'effective_mode'])[cost_col].agg(['mean','median','min','max','sum','count']).rename(columns={'sum':'total'}).reset_index()
                    metrics['llm_cost_by_process_mode'] = pm_cost.to_dict(orient='records')
                    print(f"✓ Computed LLM cost by process x mode: {len(pm_cost)} rows")
                # Failure table by process x mode
                status_col = self.column_mappings.get('status')
                if status_col and status_col in self.df.columns:
                    df_pm_status = self.df[[process_col, 'effective_mode', status_col]].copy()
                    df_pm_status[status_col] = df_pm_status[status_col].astype(str).str.strip().str.lower()
                    pm_pivot = df_pm_status.pivot_table(index=[process_col, 'effective_mode'], columns=status_col, aggfunc='size', fill_value=0)
                    if 'error' not in pm_pivot.columns: pm_pivot['error'] = 0
                    if 'info' not in pm_pivot.columns: pm_pivot['info'] = 0
                    pm_pivot = pm_pivot[['error', 'info']].reset_index()
                    pm_pivot['total'] = pm_pivot['error'] + pm_pivot['info']
                    pm_pivot['failure_pct'] = (pm_pivot['error'] / pm_pivot['total'] * 100).fillna(0)
                    metrics['failure_by_process_mode'] = pm_pivot.to_dict(orient='records')
                    print(f"✓ Computed failure rates by process x mode: {len(pm_pivot)} rows")
            
            return metrics
            
        except Exception as e:
            print(f"Error calculating metrics: {e}")
            traceback.print_exc()
            return {}
    
    def create_dau_dauu_charts(self) -> bool:
        """Create DAU and DAUU charts from PREPROCESSED data (weekdays only)"""
        if 'formatted_date' not in self.df.columns:
            print("❌ No formatted_date column found")
            return False
        
        uuid_col = self.column_mappings.get('uuid')
        if not uuid_col:
            print("❌ No UUID column found for DAUU calculation")
            return False
        
        try:
            print(f"\n📈 CREATING CHARTS (FROM PREPROCESSED DATA - WEEKDAYS ONLY)")
            print("=" * 60)
            
            # Use PREPROCESSED data (self.df) which already excludes weekends
            df_chart = self.df.copy()
            df_chart['formatted_date'] = pd.to_datetime(df_chart['formatted_date'])
            
            # Verify we only have weekdays
            weekdays_in_data = df_chart['formatted_date'].dt.weekday.unique()
            print(f"✓ Chart data contains only weekdays: {sorted(weekdays_in_data)} (0=Monday, 6=Sunday)")
            if any(day >= 5 for day in weekdays_in_data):
                print("⚠️  Warning: Weekend data detected in preprocessed data!")
            
            # Calculate DAUU (Daily Active Unique Users)
            dauu_data = df_chart.groupby('formatted_date')[uuid_col].nunique().reset_index()
            dauu_data.rename(columns={uuid_col: 'daily_active_unique_users'}, inplace=True)
            
            # Calculate DAU (Daily Active Users - total activities)
            dau_data = df_chart.groupby('formatted_date').size().reset_index()
            dau_data.rename(columns={0: 'daily_active_users'}, inplace=True)
            
            # Create DAUU Chart with continuous x-axis (no weekend gaps)
            plt.figure(figsize=(14, 8))
            
            # Use sequential plotting to avoid weekend gaps
            x_positions = range(len(dauu_data))
            date_labels = [pd.to_datetime(date).strftime('%m-%d') for date in dauu_data['formatted_date']]
            
            plt.plot(x_positions, dauu_data['daily_active_unique_users'], 
                    marker='o', linewidth=3, markersize=8, color='#2E86AB')
            
            plt.title(f'{self.file_name} - Daily Active Unique Users (DAUU)\nWeekdays Only (Continuous Timeline)', 
                     fontsize=16, fontweight='bold', pad=20)
            plt.xlabel('Date (Weekdays Only)', fontsize=14, fontweight='bold')
            plt.ylabel('Daily Active Unique Users', fontsize=14, fontweight='bold')
            
            # Set x-axis labels to show dates without gaps
            plt.xticks(x_positions, date_labels, rotation=45, fontsize=10, ha='right')
            plt.yticks(fontsize=12)
            plt.grid(True, alpha=0.3, linestyle='--')
            
            # Add value annotations
            for i, (_, row) in enumerate(dauu_data.iterrows()):
                plt.annotate(f'{int(row["daily_active_unique_users"])}', 
                           (i, row['daily_active_unique_users']),
                           textcoords="offset points", xytext=(0,10), ha='center', 
                           fontsize=10, fontweight='bold')
            
            plt.tight_layout()
            
            # Save DAUU chart
            dauu_chart_path = f"{self.output_dir}/dauu_chart.png"
            plt.savefig(dauu_chart_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            print(f"✓ DAUU chart saved: {dauu_chart_path}")
            
            # Create DAU Chart with continuous x-axis (no weekend gaps)
            plt.figure(figsize=(14, 8))
            
            # Use sequential plotting to avoid weekend gaps
            x_positions = range(len(dau_data))
            date_labels = [pd.to_datetime(date).strftime('%m-%d') for date in dau_data['formatted_date']]
            
            plt.plot(x_positions, dau_data['daily_active_users'], 
                    marker='s', linewidth=3, markersize=8, color='#FF6B6B')
            
            plt.title(f'{self.file_name} - Daily Active Users (DAU)\nWeekdays Only (Continuous Timeline)', 
                     fontsize=16, fontweight='bold', pad=20)
            plt.xlabel('Date (Weekdays Only)', fontsize=14, fontweight='bold')
            plt.ylabel('Daily Active Users', fontsize=14, fontweight='bold')
            
            # Set x-axis labels to show dates without gaps
            plt.xticks(x_positions, date_labels, rotation=45, fontsize=10, ha='right')
            plt.yticks(fontsize=12)
            plt.grid(True, alpha=0.3, linestyle='--')
            
            # Add value annotations
            for i, (_, row) in enumerate(dau_data.iterrows()):
                plt.annotate(f'{int(row["daily_active_users"])}', 
                           (i, row['daily_active_users']),
                           textcoords="offset points", xytext=(0,10), ha='center', 
                           fontsize=10, fontweight='bold')
            
            plt.tight_layout()
            
            # Save DAU chart
            dau_chart_path = f"{self.output_dir}/dau_chart.png"
            plt.savefig(dau_chart_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            print(f"✓ DAU chart saved: {dau_chart_path}")
            
            # Print statistics
            print(f"\nChart Statistics:")
            print(f"DAUU - Avg: {dauu_data['daily_active_unique_users'].mean():.0f}, " +
                  f"Max: {dauu_data['daily_active_unique_users'].max()}, " +
                  f"Min: {dauu_data['daily_active_unique_users'].min()}")
            print(f"DAU - Avg: {dau_data['daily_active_users'].mean():.0f}, " +
                  f"Max: {dau_data['daily_active_users'].max()}, " +
                  f"Min: {dau_data['daily_active_users'].min()}")
            
            return True
            
        except Exception as e:
            print(f"Error creating charts: {e}")
            traceback.print_exc()
            return False
    
    def create_response_time_charts(self) -> bool:
        """Create comprehensive response time analysis charts"""
        rt_col = self.column_mappings.get('response_time')
        if not rt_col or rt_col not in self.df.columns:
            print("⚠️  No response time column found, skipping response time charts")
            return True
        
        try:
            print(f"\n📈 CREATING RESPONSE TIME ANALYSIS CHARTS")
            print("=" * 50)
            
            # Get response time data
            rt_data = pd.to_numeric(self.df[rt_col], errors='coerce').dropna()
            
            if len(rt_data) == 0:
                print("❌ No valid response time data found")
                return False
            
            # Calculate percentiles
            percentiles = [50, 75, 90, 95, 99]
            percentile_values = [rt_data.quantile(p/100) for p in percentiles]
            
            print(f"✓ Response time percentiles:")
            for p, val in zip(percentiles, percentile_values):
                print(f"  {p}th percentile: {val:.2f}s")
            
            # Create comprehensive response time visualization
            fig = plt.figure(figsize=(16, 12))
            
            # Create a 2x2 grid
            gs = plt.GridSpec(2, 2, figure=fig, hspace=0.3, wspace=0.3)
            
            # 1. Histogram with percentile lines
            ax1 = fig.add_subplot(gs[0, 0])
            ax1.hist(rt_data, bins=50, alpha=0.7, color='skyblue', edgecolor='black')
            ax1.axvline(rt_data.mean(), color='red', linestyle='--', linewidth=2, label=f'Mean: {rt_data.mean():.2f}s')
            ax1.axvline(rt_data.median(), color='green', linestyle='--', linewidth=2, label=f'Median: {rt_data.median():.2f}s')
            ax1.axvline(percentile_values[3], color='orange', linestyle='--', linewidth=2, label=f'95th: {percentile_values[3]:.2f}s')
            ax1.axvline(percentile_values[4], color='purple', linestyle='--', linewidth=2, label=f'99th: {percentile_values[4]:.2f}s')
            ax1.set_xlabel('Response Time (seconds)')
            ax1.set_ylabel('Frequency')
            ax1.set_title(f'{self.file_name} - Response Time Distribution\n(Weekdays Only)')
            ax1.legend()
            ax1.grid(True, alpha=0.3)
            
            # 2. Box plot
            ax2 = fig.add_subplot(gs[0, 1])
            box_plot = ax2.boxplot(rt_data, patch_artist=True)
            box_plot['boxes'][0].set_facecolor('lightblue')
            ax2.set_ylabel('Response Time (seconds)')
            ax2.set_title('Response Time Box Plot')
            ax2.grid(True, alpha=0.3)
            
            # Add percentile annotations
            for i, (p, val) in enumerate(zip([25, 50, 75, 95, 99], [rt_data.quantile(0.25), rt_data.median(), rt_data.quantile(0.75), percentile_values[3], percentile_values[4]])):
                if p in [95, 99]:
                    ax2.annotate(f'{p}th: {val:.2f}s', xy=(1, val), xytext=(1.2, val),
                               arrowprops=dict(arrowstyle='->', color='red' if p == 95 else 'purple'),
                               fontsize=9, color='red' if p == 95 else 'purple')
            
            # 3. Percentile chart
            ax3 = fig.add_subplot(gs[1, 0])
            percentile_range = list(range(1, 101))
            percentile_vals = [rt_data.quantile(p/100) for p in percentile_range]
            ax3.plot(percentile_range, percentile_vals, color='blue', linewidth=2)
            ax3.axhline(percentile_values[3], color='orange', linestyle='--', label=f'95th: {percentile_values[3]:.2f}s')
            ax3.axhline(percentile_values[4], color='purple', linestyle='--', label=f'99th: {percentile_values[4]:.2f}s')
            ax3.set_xlabel('Percentile')
            ax3.set_ylabel('Response Time (seconds)')
            ax3.set_title('Response Time Percentiles')
            ax3.legend()
            ax3.grid(True, alpha=0.3)
            
            # Highlight critical percentiles
            critical_percentiles = [90, 95, 99]
            for cp in critical_percentiles:
                val = rt_data.quantile(cp/100)
                ax3.scatter([cp], [val], color='red', s=50, zorder=5)
                ax3.annotate(f'{cp}th\n{val:.2f}s', xy=(cp, val), xytext=(cp, val + max(percentile_vals) * 0.1),
                           ha='center', fontsize=8, fontweight='bold')
            
            # 4. Daily response time trends
            ax4 = fig.add_subplot(gs[1, 1])
            
            # Group by date and calculate daily statistics
            self.df['rt_numeric'] = pd.to_numeric(self.df[rt_col], errors='coerce')
            daily_stats = self.df.groupby('formatted_date')['rt_numeric'].agg([
                'mean', 'median', 
                lambda x: x.quantile(0.95), 
                lambda x: x.quantile(0.99)
            ]).rename(columns={'<lambda_0>': 'p95', '<lambda_1>': 'p99'})
            
            daily_stats.index = pd.to_datetime(daily_stats.index)
            
            ax4.plot(daily_stats.index, daily_stats['mean'], marker='o', label='Mean', linewidth=2, markersize=4)
            ax4.plot(daily_stats.index, daily_stats['median'], marker='s', label='Median', linewidth=2, markersize=4)
            ax4.plot(daily_stats.index, daily_stats['p95'], marker='^', label='95th percentile', linewidth=2, markersize=4, color='orange')
            ax4.plot(daily_stats.index, daily_stats['p99'], marker='v', label='99th percentile', linewidth=2, markersize=4, color='red')
            
            ax4.set_xlabel('Date')
            ax4.set_ylabel('Response Time (seconds)')
            ax4.set_title('Daily Response Time Trends\n(Weekdays Only)')
            ax4.legend()
            ax4.grid(True, alpha=0.3)
            
            # Format x-axis dates - show only actual dates (weekdays only)
            import matplotlib.dates as mdates
            actual_dates = sorted(daily_stats.index.unique())
            ax4.set_xticks(actual_dates)
            ax4.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d'))
            plt.setp(ax4.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            plt.suptitle(f'{self.file_name} - Comprehensive Response Time Analysis', fontsize=16, fontweight='bold')
            
            # Save the comprehensive chart
            rt_chart_path = f"{self.output_dir}/response_time_analysis.png"
            plt.savefig(rt_chart_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
            plt.close()
            print(f"✓ Response time analysis chart saved: {rt_chart_path}")
            
            # Create a separate simple percentile summary chart
            self._create_simple_percentile_chart(rt_data, percentiles, percentile_values)
            
            # Create daily min/max/average chart
            self._create_daily_minmax_chart(daily_stats)
            
            return True
            
        except Exception as e:
            print(f"Error creating response time charts: {e}")
            traceback.print_exc()
            return False
    
    def _create_simple_percentile_chart(self, rt_data, percentiles, percentile_values):
        """Create a simple, clean percentile chart for presentations"""
        try:
            fig, ax = plt.subplots(figsize=(12, 8))
            
            # Create bar chart of percentiles
            colors = ['#4CAF50', '#FFC107', '#FF9800', '#FF5722', '#9C27B0']
            bars = ax.bar([f'{p}th' for p in percentiles], percentile_values, 
                         color=colors, alpha=0.8, edgecolor='black', linewidth=1)
            
            # Add value labels on bars
            for bar, val in zip(bars, percentile_values):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height + max(percentile_values) * 0.01,
                       f'{val:.2f}s', ha='center', va='bottom', fontweight='bold', fontsize=12)
            
            # Add mean and max lines
            ax.axhline(y=rt_data.mean(), color='red', linestyle='--', linewidth=2, 
                      label=f'Mean: {rt_data.mean():.2f}s', alpha=0.8)
            ax.axhline(y=rt_data.max(), color='purple', linestyle=':', linewidth=2, 
                      label=f'Max: {rt_data.max():.2f}s', alpha=0.8)
            
            ax.set_ylabel('Response Time (seconds)', fontsize=14, fontweight='bold')
            ax.set_xlabel('Percentiles', fontsize=14, fontweight='bold')
            ax.set_title(f'{self.file_name} - Response Time Percentiles Summary\n'
                        f'Total Requests: {len(rt_data):,} (Weekdays Only)', 
                        fontsize=16, fontweight='bold', pad=20)
            
            ax.legend(loc='upper left', fontsize=12)
            ax.grid(True, alpha=0.3, axis='y')
            
            # Improve styling
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            plt.xticks(fontsize=12)
            plt.yticks(fontsize=12)
            
            plt.tight_layout()
            
            # Save the simple chart
            simple_chart_path = f"{self.output_dir}/response_time_percentiles.png"
            plt.savefig(simple_chart_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
            plt.close()
            print(f"✓ Simple percentile chart saved: {simple_chart_path}")
            
        except Exception as e:
            print(f"Error creating simple percentile chart: {e}")
    
    def _create_daily_minmax_chart(self, daily_stats):
        """Create daily min/max points with average line chart"""
        try:
            # Get additional daily statistics
            rt_col = self.column_mappings.get('response_time')
            self.df['rt_numeric'] = pd.to_numeric(self.df[rt_col], errors='coerce')
            
            # Calculate daily min, max, and mean
            daily_detailed = self.df.groupby('formatted_date')['rt_numeric'].agg([
                'min', 'max', 'mean', 'count'
            ])
            daily_detailed.index = pd.to_datetime(daily_detailed.index)
            
            # Create the chart with continuous x-axis (no weekend gaps)
            fig, ax = plt.subplots(figsize=(14, 8))
            
            dates = daily_detailed.index
            date_labels = [date.strftime('%m-%d') for date in dates]
            x_positions = range(len(dates))
            
            # Plot min and max as scatter points
            ax.scatter(x_positions, daily_detailed['min'], color='green', s=60, alpha=0.7, 
                      label='Daily Minimum', marker='v', zorder=3)
            ax.scatter(x_positions, daily_detailed['max'], color='red', s=60, alpha=0.7, 
                      label='Daily Maximum', marker='^', zorder=3)
            
            # Plot average as a line
            ax.plot(x_positions, daily_detailed['mean'], color='blue', linewidth=3, 
                   marker='o', markersize=6, label='Daily Average', zorder=2)
            
            # Fill area between min and max to show daily range
            ax.fill_between(x_positions, daily_detailed['min'], daily_detailed['max'], 
                           alpha=0.2, color='gray', label='Daily Range')
            
            # Customize the chart
            ax.set_xlabel('Date (Weekdays Only)', fontsize=14, fontweight='bold')
            ax.set_ylabel('Response Time (seconds)', fontsize=14, fontweight='bold')
            ax.set_title(f'{self.file_name} - Daily Response Time Range & Average\n'
                        f'Weekdays Only (Continuous Timeline)', 
                        fontsize=16, fontweight='bold', pad=20)
            
            # Set x-axis labels to show dates without gaps
            ax.set_xticks(x_positions)
            ax.set_xticklabels(date_labels, rotation=45, ha='right')
            
            # Add legend
            ax.legend(loc='upper left', fontsize=12)
            ax.grid(True, alpha=0.3)
            
            # Add annotations for key insights
            max_range_idx = (daily_detailed['max'] - daily_detailed['min']).idxmax()
            max_range_value = daily_detailed.loc[max_range_idx, 'max'] - daily_detailed.loc[max_range_idx, 'min']
            max_range_pos = list(dates).index(max_range_idx)
            
            # Annotate the day with highest variability
            ax.annotate(f'Highest variability\n{max_range_value:.1f}s range', 
                       xy=(max_range_pos, daily_detailed.loc[max_range_idx, 'max']),
                       xytext=(10, 20), textcoords='offset points',
                       bbox=dict(boxstyle='round,pad=0.5', fc='yellow', alpha=0.7),
                       arrowprops=dict(arrowstyle='->', connectionstyle='arc3,rad=0'),
                       fontsize=10)
            
            # Add summary statistics as text
            overall_stats = f"""Daily Statistics:
Avg Min: {daily_detailed['min'].mean():.1f}s
Avg Max: {daily_detailed['max'].mean():.1f}s  
Avg Range: {(daily_detailed['max'] - daily_detailed['min']).mean():.1f}s
Most Stable Day: {(daily_detailed['max'] - daily_detailed['min']).idxmin().strftime('%m-%d')}
Most Variable Day: {max_range_idx.strftime('%m-%d')}"""
            
            ax.text(0.02, 0.98, overall_stats, transform=ax.transAxes, fontsize=10,
                   verticalalignment='top', bbox=dict(boxstyle='round', facecolor='lightblue', alpha=0.8))
            
            # Improve styling
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            plt.xticks(fontsize=12)
            plt.yticks(fontsize=12)
            
            plt.tight_layout()
            
            # Save the chart
            minmax_chart_path = f"{self.output_dir}/daily_response_time_range.png"
            plt.savefig(minmax_chart_path, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
            plt.close()
            print(f"✓ Daily min/max range chart saved: {minmax_chart_path}")
            
            # Histogram trend chart removed as requested
            
        except Exception as e:
            print(f"Error creating daily min/max chart: {e}")
            traceback.print_exc()
    
    
    def generate_daily_analysis(self) -> bool:
        """Generate daily analysis with 5 key metrics comparison"""
        try:
            print(f"\n📅 GENERATING DAILY ANALYSIS")
            print("=" * 40)
            
            # Calculate daily metrics
            daily_metrics = self._calculate_daily_metrics()
            
            if not daily_metrics or len(daily_metrics) < 2:
                print("❌ Need at least 2 days of data for daily analysis")
                return False
            
            # Determine comparison dates
            dates = sorted(daily_metrics.keys())
            
            print(f"✓ Found {len(dates)} days of data")
            
            # Resolve user-specified compare dates if provided
            if self.compare_dates:
                print(f"✓ Using requested compare dates: {self.compare_dates}")
                requested = self._resolve_compare_dates(self.compare_dates, dates)
                if requested is None:
                    print("❌ Requested compare dates not found in data")
                    return False
                yesterday, today = requested
            elif len(dates) >= 2:
                # Fallback: last two days
                yesterday, today = dates[-2], dates[-1]
            else:
                print("❌ Need at least 2 consecutive days for comparison")
                return False

            # At this point we have yesterday and today resolved; run comparison and write file
            print(f"✓ Comparing: {yesterday} → {today}")
            
            analysis = self._compare_daily_metrics(
                daily_metrics[yesterday], 
                daily_metrics[today], 
                yesterday, 
                today
            )
            
            # Save single daily analysis to file
            self._save_single_daily_analysis(analysis)
            
            # Show the analysis
            print(f"\n📊 DAILY ANALYSIS (Latest Consecutive Days):")
            print("=" * 50)
            self._print_daily_analysis(analysis)
            return True
            
        except Exception as e:
            print(f"Error generating daily analysis: {e}")
            traceback.print_exc()
            return False

    def _resolve_compare_dates(self, compare_dates: Tuple[str, str], available_dates: List[str]) -> Optional[Tuple[str, str]]:
        """Resolve input date tokens to actual formatted_date strings present in data.
        Accepts tokens in 'dd/mm' or 'yyyy-mm-dd'. Chooses the latest matching year for dd/mm.
        """
        try:
            # Precompute set and map for quick lookup
            avail = sorted(available_dates)
            avail_dt = [pd.to_datetime(d) for d in avail]
            def normalize(token: str) -> Optional[str]:
                token = token.strip()
                if '/' in token and '-' not in token:
                    # dd/mm
                    d, m = token.split('/')
                    day = int(d)
                    month = int(m)
                    # Find latest year match
                    candidates = [dt for dt in avail_dt if dt.day == day and dt.month == month]
                    if not candidates:
                        return None
                    chosen = max(candidates)
                    return chosen.strftime('%Y-%m-%d')
                else:
                    # Expect full ISO date
                    try:
                        dt = pd.to_datetime(token)
                        iso = dt.strftime('%Y-%m-%d')
                        return iso if iso in available_dates else None
                    except Exception:
                        return None
            y = normalize(compare_dates[0])
            t = normalize(compare_dates[1])
            if y and t:
                return (y, t)
            return None
        except Exception:
            return None
    
    def _calculate_daily_metrics(self) -> Dict:
        """Calculate key metrics for each day"""
        daily_metrics = {}
        
        if 'formatted_date' not in self.df.columns:
            return daily_metrics
        
        # Prepare data
        rt_col = self.column_mappings.get('response_time')
        uuid_col = self.column_mappings.get('uuid')
        status_col = self.column_mappings.get('status')
        cost_col = self.column_mappings.get('llm_cost')
        
        for date in sorted(self.df['formatted_date'].unique()):
            day_data = self.df[self.df['formatted_date'] == date]
            
            metrics = {
                'date': date,
                'total_requests': len(day_data),
                'unique_users': 0,
                'avg_response_time': 0,
                'success_rate': 0,
                'total_llm_cost': 0
            }
            
            # 1. Latency Metric (Average Response Time)
            if rt_col and rt_col in day_data.columns:
                rt_data = pd.to_numeric(day_data[rt_col], errors='coerce').dropna()
                if len(rt_data) > 0:
                    metrics['avg_response_time'] = rt_data.mean()
            
            # 2. Throughput Metric (Total Requests)
            metrics['total_requests'] = len(day_data)
            
            # 3. LLM Cost Metric
            if cost_col and cost_col in day_data.columns:
                cost_data = pd.to_numeric(day_data[cost_col], errors='coerce').dropna()
                if len(cost_data) > 0:
                    metrics['total_llm_cost'] = cost_data.sum()
            
            # 4. Reliability Metric (Success Rate)
            if status_col and status_col in day_data.columns:
                total_records = len(day_data)
                success_records = len(day_data[day_data[status_col].str.lower() == 'info'])
                metrics['success_rate'] = (success_records / total_records * 100) if total_records > 0 else 0
            
            # 5. User Activity Metric (Unique Users)
            if uuid_col and uuid_col in day_data.columns:
                metrics['unique_users'] = day_data[uuid_col].nunique()
            
            daily_metrics[date] = metrics
        
        return daily_metrics
    
    def _compare_daily_metrics(self, yesterday_metrics: Dict, today_metrics: Dict, yesterday_date: str, today_date: str) -> Dict:
        """Compare metrics between two consecutive days"""
        comparison = {
            'yesterday_date': yesterday_date,
            'today_date': today_date,
            'metrics': {}
        }
        
        # 1. Latency Metric
        today_latency = today_metrics['avg_response_time']
        yesterday_latency = yesterday_metrics['avg_response_time']
        latency_change = today_latency - yesterday_latency
        latency_pct = (latency_change / yesterday_latency * 100) if yesterday_latency > 0 else 0
        
        comparison['metrics']['latency'] = {
            'today': today_latency,
            'yesterday': yesterday_latency,
            'change_absolute': latency_change,
            'change_percent': latency_pct,
            'status': self._get_latency_status(latency_pct)
        }
        
        # 2. Throughput Metric
        today_requests = today_metrics['total_requests']
        yesterday_requests = yesterday_metrics['total_requests']
        requests_change = today_requests - yesterday_requests
        requests_pct = (requests_change / yesterday_requests * 100) if yesterday_requests > 0 else 0
        
        comparison['metrics']['throughput'] = {
            'today': today_requests,
            'yesterday': yesterday_requests,
            'change_absolute': requests_change,
            'change_percent': requests_pct,
            'status': self._get_throughput_status(requests_pct)
        }
        
        # 3. LLM Cost Metric
        today_cost = today_metrics['total_llm_cost']
        yesterday_cost = yesterday_metrics['total_llm_cost']
        cost_change = today_cost - yesterday_cost
        cost_pct = (cost_change / yesterday_cost * 100) if yesterday_cost > 0 else 0
        
        comparison['metrics']['llm_cost'] = {
            'today': today_cost,
            'yesterday': yesterday_cost,
            'change_absolute': cost_change,
            'change_percent': cost_pct,
            'status': self._get_cost_status(cost_pct)
        }
        
        # 4. Reliability Metric
        today_success = today_metrics['success_rate']
        yesterday_success = yesterday_metrics['success_rate']
        success_change = today_success - yesterday_success
        success_pct = (success_change / yesterday_success * 100) if yesterday_success > 0 else 0
        
        comparison['metrics']['reliability'] = {
            'today': today_success,
            'yesterday': yesterday_success,
            'change_absolute': success_change,
            'change_percent': success_pct,
            'status': self._get_reliability_status(success_change)
        }
        
        # 5. User Activity Metric
        today_users = today_metrics['unique_users']
        yesterday_users = yesterday_metrics['unique_users']
        users_change = today_users - yesterday_users
        users_pct = (users_change / yesterday_users * 100) if yesterday_users > 0 else 0
        
        comparison['metrics']['user_activity'] = {
            'today': today_users,
            'yesterday': yesterday_users,
            'change_absolute': users_change,
            'change_percent': users_pct,
            'status': self._get_user_activity_status(users_pct)
        }
        
        return comparison
    
    def _get_latency_status(self, pct_change: float) -> str:
        """Determine latency status based on percentage change"""
        if pct_change < -5:
            return "IMPROVING"
        elif pct_change > 5:
            return "DEGRADING"
        else:
            return "STABLE"
    
    def _get_throughput_status(self, pct_change: float) -> str:
        """Determine throughput status based on percentage change"""
        if pct_change > 5:
            return "GROWING"
        elif pct_change < -5:
            return "DECLINING"
        else:
            return "STABLE"
    
    def _get_cost_status(self, pct_change: float) -> str:
        """Determine cost status based on percentage change"""
        if pct_change < -5:
            return "EFFICIENT"
        elif pct_change > 10:
            return "EXPENSIVE"
        else:
            return "STABLE"
    
    def _get_reliability_status(self, absolute_change: float) -> str:
        """Determine reliability status based on absolute change"""
        if absolute_change > 1:
            return "IMPROVING"
        elif absolute_change < -1:
            return "DEGRADING"
        else:
            return "STABLE"
    
    def _get_user_activity_status(self, pct_change: float) -> str:
        """Determine user activity status based on percentage change"""
        if pct_change > 5:
            return "GROWING"
        elif pct_change < -5:
            return "DECLINING"
        else:
            return "STABLE"
    
    def _print_daily_analysis(self, analysis: Dict):
        """Print daily analysis in the requested format"""
        metrics = analysis['metrics']
        today_date = analysis['today_date']
        yesterday_date = analysis['yesterday_date']
        
        print(f"Comparison: {yesterday_date} → {today_date}")
        print()
        
        # 1. Latency Metric (percent from actual values; more precision shown)
        latency = metrics['latency']
        change_symbol = '→' if latency['change_absolute'] == 0 else ('↓' if latency['change_absolute'] < 0 else '↑')
        print(f"1. Latency Metric")
        print(f"Today's Avg Response Time: {latency['today']:.3f}ms")
        print(f"Yesterday's Avg Response Time: {latency['yesterday']:.3f}ms")
        arrow = change_symbol
        word = 'no change' if latency['change_absolute'] == 0 else ('improvement' if latency['change_absolute'] < 0 else 'increase')
        print(f"Change: {latency['change_absolute']:+.3f}ms ({arrow}{abs(latency['change_percent']):.2f}% {word})")
        print(f"Status: {latency['status']}")
        print()
        
        # 2. Throughput Metric
        throughput = metrics['throughput']
        change_symbol = "↓" if throughput['change_absolute'] < 0 else "↑"
        print(f"2. Throughput Metric")
        print(f"Today's Total Requests: {throughput['today']:,}")
        print(f"Yesterday's Total Requests: {throughput['yesterday']:,}")
        if throughput['change_absolute'] == 0:
            arrow = '→'; word = 'no change'
        else:
            arrow = '↓' if throughput['change_absolute'] < 0 else '↑'
            word = 'decrease' if throughput['change_absolute'] < 0 else 'increase'
        print(f"Change: {throughput['change_absolute']:+,} requests ({arrow}{abs(throughput['change_percent']):.1f}% {word})")
        print(f"Status: {throughput['status']}")
        print()
        
        # 3. LLM Cost Metric (percent from actual values)
        cost = metrics['llm_cost']
        change_symbol = '→' if cost['change_absolute'] == 0 else ('↓' if cost['change_absolute'] < 0 else '↑')
        print(f"3. LLM Cost Metric")
        print(f"Today's Total Cost: ${cost['today']:.4f}")
        print(f"Yesterday's Total Cost: ${cost['yesterday']:.4f}")
        arrow = change_symbol
        word = 'no change' if cost['change_absolute'] == 0 else ('decrease' if cost['change_absolute'] < 0 else 'increase')
        print(f"Change: ${cost['change_absolute']:+.4f} ({arrow}{abs(cost['change_percent']):.2f}% {word})")
        print(f"Status: {cost['status']}")
        print()
        
        # 4. Reliability Metric (percent from actual values)
        reliability = metrics['reliability']
        change_symbol = '→' if reliability['change_absolute'] == 0 else ('↓' if reliability['change_absolute'] < 0 else '↑')
        print(f"4. Reliability Metric")
        print(f"Today's Success Rate: {reliability['today']:.2f}%")
        print(f"Yesterday's Success Rate: {reliability['yesterday']:.2f}%")
        arrow = change_symbol
        word = 'no change' if reliability['change_absolute'] == 0 else ('improvement' if reliability['change_absolute'] > 0 else 'degradation')
        print(f"Change: {reliability['change_absolute']:+.2f}% ({arrow}{abs(reliability['change_percent']):.2f}% {word})")
        print(f"Status: {reliability['status']}")
        print()
        
        # 5. User Activity Metric
        activity = metrics['user_activity']
        change_symbol = "↓" if activity['change_absolute'] < 0 else "↑"
        print(f"5. User Activity Metric")
        print(f"Today's Unique Users: {activity['today']:,}")
        print(f"Yesterday's Unique Users: {activity['yesterday']:,}")
        if activity['change_absolute'] == 0:
            arrow = '→'; word = 'no change'
        else:
            arrow = '↓' if activity['change_absolute'] < 0 else '↑'
            word = 'decline' if activity['change_absolute'] < 0 else 'growth'
        print(f"Change: {activity['change_absolute']:+,} users ({arrow}{abs(activity['change_percent']):.1f}% {word})")
        print(f"Status: {activity['status']}")
    
    def _save_single_daily_analysis(self, analysis: Dict):
        """Save single daily analysis result to file"""
        try:
            daily_analysis_path = f"{self.output_dir}/daily_analysis.txt"
            
            with open(daily_analysis_path, 'w', encoding='utf-8') as f:
                metrics = analysis['metrics']
                today_date = analysis['today_date']
                yesterday_date = analysis['yesterday_date']
                
                f.write(f"DAILY ANALYSIS REPORT - {self.file_name}\n")
                f.write(f"=" * 60 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"File: {self.file_name}{self.file_extension}\n")
                f.write(f"Comparison: {yesterday_date} → {today_date}\n\n")
                
                # 1. Latency Metric (raw percent, more precision)
                latency = metrics['latency']
                change_symbol = '→' if latency['change_absolute'] == 0 else ('↓' if latency['change_absolute'] < 0 else '↑')
                f.write(f"1. Latency Metric\n")
                f.write(f"Today's Avg Response Time: {latency['today']:.3f}ms\n")
                f.write(f"Yesterday's Avg Response Time: {latency['yesterday']:.3f}ms\n")
                arrow = change_symbol
                word = 'no change' if latency['change_absolute'] == 0 else ('improvement' if latency['change_absolute'] < 0 else 'increase')
                f.write(f"Change: {latency['change_absolute']:+.3f}ms ({arrow}{abs(latency['change_percent']):.2f}% {word})\n")
                f.write(f"Status: {latency['status']}\n\n")
                
                # 2. Throughput Metric
                throughput = metrics['throughput']
                change_symbol = "↓" if throughput['change_absolute'] < 0 else "↑"
                f.write(f"2. Throughput Metric\n")
                f.write(f"Today's Total Requests: {throughput['today']:,}\n")
                f.write(f"Yesterday's Total Requests: {throughput['yesterday']:,}\n")
                if throughput['change_absolute'] == 0:
                    arrow = '→'; word = 'no change'
                else:
                    arrow = '↓' if throughput['change_absolute'] < 0 else '↑'
                    word = 'decrease' if throughput['change_absolute'] < 0 else 'increase'
                f.write(f"Change: {throughput['change_absolute']:+,} requests ({arrow}{abs(throughput['change_percent']):.1f}% {word})\n")
                f.write(f"Status: {throughput['status']}\n\n")
                
                # 3. LLM Cost Metric
                cost = metrics['llm_cost']
                change_symbol = "↓" if cost['change_absolute'] < 0 else "↑"
                f.write(f"3. LLM Cost Metric\n")
                f.write(f"Today's Total Cost: ${cost['today']:.2f}\n")
                f.write(f"Yesterday's Total Cost: ${cost['yesterday']:.2f}\n")
                if cost['change_absolute'] == 0:
                    arrow = '→'; word = 'no change'
                else:
                    arrow = '↓' if cost['change_absolute'] < 0 else '↑'
                    word = 'decrease' if cost['change_absolute'] < 0 else 'increase'
                f.write(f"Change: ${cost['change_absolute']:+.2f} ({arrow}{abs(cost['change_percent']):.1f}% {word})\n")
                f.write(f"Status: {cost['status']}\n\n")
                
                # 4. Reliability Metric
                reliability = metrics['reliability']
                change_symbol = "↓" if reliability['change_absolute'] < 0 else "↑"
                f.write(f"4. Reliability Metric\n")
                f.write(f"Today's Success Rate: {reliability['today']:.1f}%\n")
                f.write(f"Yesterday's Success Rate: {reliability['yesterday']:.1f}%\n")
                if reliability['change_absolute'] == 0:
                    arrow = '→'; word = 'no change'
                else:
                    arrow = '↑' if reliability['change_absolute'] > 0 else '↓'
                    word = 'improvement' if reliability['change_absolute'] > 0 else 'degradation'
                f.write(f"Change: {reliability['change_absolute']:+.1f}% ({arrow}{abs(reliability['change_percent']):.1f}% {word})\n")
                f.write(f"Status: {reliability['status']}\n\n")
                
                # 5. User Activity Metric
                activity = metrics['user_activity']
                change_symbol = "↓" if activity['change_absolute'] < 0 else "↑"
                f.write(f"5. User Activity Metric\n")
                f.write(f"Today's Unique Users: {activity['today']:,}\n")
                f.write(f"Yesterday's Unique Users: {activity['yesterday']:,}\n")
                if activity['change_absolute'] == 0:
                    arrow = '→'; word = 'no change'
                else:
                    arrow = '↓' if activity['change_absolute'] < 0 else '↑'
                    word = 'decline' if activity['change_absolute'] < 0 else 'growth'
                f.write(f"Change: {activity['change_absolute']:+,} users ({arrow}{abs(activity['change_percent']):.1f}% {word})\n")
                f.write(f"Status: {activity['status']}\n\n")
            
            print(f"✓ Daily analysis saved: {daily_analysis_path}")
            return True
            
        except Exception as e:
            print(f"Error saving daily analysis: {e}")
            return False
    
    def _save_daily_analysis(self, analysis_results: List[Dict]):
        """Save daily analysis results to file"""
        try:
            daily_analysis_path = f"{self.output_dir}/daily_analysis.txt"
            
            with open(daily_analysis_path, 'w', encoding='utf-8') as f:
                f.write(f"DAILY ANALYSIS REPORT - {self.file_name}\n")
                f.write(f"=" * 60 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"File: {self.file_name}{self.file_extension}\n")
                f.write(f"Total Daily Comparisons: {len(analysis_results)}\n\n")
                
                for i, analysis in enumerate(analysis_results, 1):
                    metrics = analysis['metrics']
                    today_date = analysis['today_date']
                    yesterday_date = analysis['yesterday_date']
                    
                    f.write(f"COMPARISON #{i}: {yesterday_date} → {today_date}\n")
                    f.write("=" * 50 + "\n\n")
                    
                    # 1. Latency Metric
                    latency = metrics['latency']
                    change_symbol = "↓" if latency['change_absolute'] < 0 else "↑"
                    f.write(f"1. Latency Metric\n")
                    f.write(f"Today's Avg Response Time: {latency['today']:.0f}ms\n")
                    f.write(f"Yesterday's Avg Response Time: {latency['yesterday']:.0f}ms\n")
                    f.write(f"Change: {latency['change_absolute']:+.0f}ms ({change_symbol}{abs(latency['change_percent']):.1f}% {'improvement' if latency['change_absolute'] < 0 else 'increase'})\n")
                    f.write(f"Status: {latency['status']}\n\n")
                    
                    # 2. Throughput Metric
                    throughput = metrics['throughput']
                    change_symbol = "↓" if throughput['change_absolute'] < 0 else "↑"
                    f.write(f"2. Throughput Metric\n")
                    f.write(f"Today's Total Requests: {throughput['today']:,}\n")
                    f.write(f"Yesterday's Total Requests: {throughput['yesterday']:,}\n")
                    f.write(f"Change: {throughput['change_absolute']:+,} requests ({change_symbol}{abs(throughput['change_percent']):.1f}% {'decrease' if throughput['change_absolute'] < 0 else 'increase'})\n")
                    f.write(f"Status: {throughput['status']}\n\n")
                    
                    # 3. LLM Cost Metric (raw percent)
                    cost = metrics['llm_cost']
                    change_symbol = '→' if cost['change_absolute'] == 0 else ('↓' if cost['change_absolute'] < 0 else '↑')
                    f.write(f"3. LLM Cost Metric\n")
                    f.write(f"Today's Total Cost: ${cost['today']:.4f}\n")
                    f.write(f"Yesterday's Total Cost: ${cost['yesterday']:.4f}\n")
                    arrow = change_symbol
                    word = 'no change' if cost['change_absolute'] == 0 else ('decrease' if cost['change_absolute'] < 0 else 'increase')
                    f.write(f"Change: ${cost['change_absolute']:+.4f} ({arrow}{abs(cost['change_percent']):.2f}% {word})\n")
                    f.write(f"Status: {cost['status']}\n\n")
                    
                    # 4. Reliability Metric (raw percent)
                    reliability = metrics['reliability']
                    change_symbol = '→' if reliability['change_absolute'] == 0 else ('↓' if reliability['change_absolute'] < 0 else '↑')
                    f.write(f"4. Reliability Metric\n")
                    f.write(f"Today's Success Rate: {reliability['today']:.2f}%\n")
                    f.write(f"Yesterday's Success Rate: {reliability['yesterday']:.2f}%\n")
                    arrow = change_symbol
                    word = 'no change' if reliability['change_absolute'] == 0 else ('improvement' if reliability['change_absolute'] > 0 else 'degradation')
                    f.write(f"Change: {reliability['change_absolute']:+.2f}% ({arrow}{abs(reliability['change_percent']):.2f}% {word})\n")
                    f.write(f"Status: {reliability['status']}\n\n")
                    
                    # 5. User Activity Metric
                    activity = metrics['user_activity']
                    change_symbol = "↓" if activity['change_absolute'] < 0 else "↑"
                    f.write(f"5. User Activity Metric\n")
                    f.write(f"Today's Unique Users: {activity['today']:,}\n")
                    f.write(f"Yesterday's Unique Users: {activity['yesterday']:,}\n")
                    f.write(f"Change: {activity['change_absolute']:+,} users ({change_symbol}{abs(activity['change_percent']):.1f}% {'decline' if activity['change_absolute'] < 0 else 'growth'})\n")
                    f.write(f"Status: {activity['status']}\n\n")
                    
                    f.write("=" * 60 + "\n\n")
                
                # Summary section
                f.write("SUMMARY TRENDS\n")
                f.write("=" * 20 + "\n")
                
                # Calculate overall trends
                latency_trends = [a['metrics']['latency']['status'] for a in analysis_results]
                throughput_trends = [a['metrics']['throughput']['status'] for a in analysis_results]
                cost_trends = [a['metrics']['llm_cost']['status'] for a in analysis_results]
                reliability_trends = [a['metrics']['reliability']['status'] for a in analysis_results]
                activity_trends = [a['metrics']['user_activity']['status'] for a in analysis_results]
                
                f.write(f"Latency Trend: {self._get_dominant_trend(latency_trends)}\n")
                f.write(f"Throughput Trend: {self._get_dominant_trend(throughput_trends)}\n")
                f.write(f"Cost Trend: {self._get_dominant_trend(cost_trends)}\n")
                f.write(f"Reliability Trend: {self._get_dominant_trend(reliability_trends)}\n")
                f.write(f"User Activity Trend: {self._get_dominant_trend(activity_trends)}\n")
            
            print(f"✓ Daily analysis saved: {daily_analysis_path}")
            return True
            
        except Exception as e:
            print(f"Error saving daily analysis: {e}")
            return False
    
    def _get_dominant_trend(self, trends: List[str]) -> str:
        """Get the dominant trend from a list of status values"""
        from collections import Counter
        trend_counts = Counter(trends)
        return trend_counts.most_common(1)[0][0] if trend_counts else "STABLE"
    
    def save_metrics_to_txt(self, metrics: Dict) -> bool:
        """Save all metrics to a comprehensive TXT file"""
        try:
            txt_path = f"{self.output_dir}/metrics_analysis.txt"
            
            with open(txt_path, 'w', encoding='utf-8') as f:
                # Header: Service/Source display name for downstream report naming
                service_display = None
                service_col = self.column_mappings.get('service')
                if service_col and service_col in self.df.columns:
                    s = self.df[service_col].astype(str).str.strip()
                    s = s[s.astype(bool)]
                    if not s.empty:
                        service_display = s.value_counts().idxmax()
                if not service_display:
                    service_display = self.file_name
                f.write(f"SERVICE NAME: {service_display}\n\n")
                f.write(f"INDIVIDUAL ANALYSIS REPORT\n")
                f.write(f"=" * 50 + "\n")
                f.write(f"File: {self.file_name}{self.file_extension}\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Output Directory: {self.output_dir}\n\n")
                
                # Response Time and LLM Cost Metrics
                f.write(f"RESPONSE TIME AND LLM COST METRICS\n")
                f.write(f"=" * 40 + "\n")
                
                # Response Time Table
                if 'response_time' in metrics:
                    rt = metrics['response_time']
                    f.write(f"Response Time Metrics:\n")
                    f.write(f"{'Metric':<25} {'Value':<15}\n")
                    f.write(f"{'-'*40}\n")
                    f.write(f"{'Avg Time Taken':<25} {rt['mean']:.2f} s\n")
                    f.write(f"{'Min Time Taken':<25} {rt['min']:.2f} s\n")
                    f.write(f"{'Max Time Taken':<25} {rt['max']:.2f} s\n")
                    f.write(f"{'Median Time':<25} {rt['median']:.2f} s\n")
                    f.write(f"{'Std Deviation':<25} {rt['std']:.2f} s\n")
                    f.write(f"{'Records Analyzed':<25} {rt['count']:,}\n")
                    f.write(f"\n")
                else:
                    f.write(f"Response Time Metrics: Not Available\n\n")
                
                # LLM Cost Table
                if 'llm_cost' in metrics:
                    cost = metrics['llm_cost']
                    f.write(f"LLM Cost Metrics:\n")
                    f.write(f"{'Metric':<25} {'Value':<15}\n")
                    f.write(f"{'-'*40}\n")
                    f.write(f"{'Avg LLM Cost':<25} ${cost['mean']:.4f}\n")
                    f.write(f"{'Min LLM Cost':<25} ${cost['min']:.4f}\n")
                    f.write(f"{'Max LLM Cost':<25} ${cost['max']:.4f}\n")
                    f.write(f"{'Total LLM Cost':<25} ${cost['total']:.2f}\n")
                    f.write(f"{'Median Cost':<25} ${cost['median']:.4f}\n")
                    f.write(f"{'Records with Cost':<25} {cost['count']:,}\n")
                    f.write(f"\n")
                else:
                    f.write(f"LLM Cost Metrics: Not Available\n\n")
                
                # Process-wise tables
                if 'response_time_by_process' in metrics and metrics['response_time_by_process']:
                    f.write(f"RESPONSE TIME BY PROCESS\n")
                    f.write(f"=" * 27 + "\n")
                    f.write(f"{'Process Name':<40} {'Avg (s)':>10} {'P50 (s)':>10} {'Min (s)':>10} {'Max (s)':>10} {'Std':>10} {'N':>8}\n")
                    f.write(f"{'-'*100}\n")
                    for row in metrics['response_time_by_process']:
                        f.write(f"{str(row.get(self.column_mappings.get('process_name'), row.get('index', ''))):<40} "
                                f"{row.get('mean', 0):>10.2f} {row.get('median', 0):>10.2f} {row.get('min', 0):>10.2f} "
                                f"{row.get('max', 0):>10.2f} {row.get('std', 0):>10.2f} {row.get('count', 0):>8}\n")
                    f.write("\n")

                if 'llm_cost_by_process' in metrics and metrics['llm_cost_by_process']:
                    f.write(f"LLM COST BY PROCESS\n")
                    f.write(f"=" * 20 + "\n")
                    f.write(f"{'Process Name':<40} {'Avg ($)':>10} {'Median':>10} {'Min':>10} {'Max':>10} {'Total ($)':>12} {'N':>8}\n")
                    f.write(f"{'-'*110}\n")
                    for row in metrics['llm_cost_by_process']:
                        f.write(f"{str(row.get(self.column_mappings.get('process_name'), row.get('index', ''))):<40} "
                                f"{row.get('mean', 0):>10.4f} {row.get('median', 0):>10.4f} {row.get('min', 0):>10.4f} "
                                f"{row.get('max', 0):>10.4f} {row.get('total', 0):>12.2f} {row.get('count', 0):>8}\n")
                    f.write("\n")

                # Effective mode-wise tables
                if 'response_time_by_effective_mode' in metrics and metrics['response_time_by_effective_mode']:
                    f.write(f"RESPONSE TIME BY EFFECTIVE MODE\n")
                    f.write(f"=" * 32 + "\n")
                    f.write(f"{'Mode':<8} {'Mode Name':<30} {'Avg (s)':>10} {'P50 (s)':>10} {'Min (s)':>10} {'Max (s)':>10} {'Std':>10} {'N':>8}\n")
                    f.write(f"{'-'*120}\n")
                    for row in metrics['response_time_by_effective_mode']:
                        f.write(f"{int(row.get('effective_mode', -1)):>8} {row.get('mode_name',''): <30} "
                                f"{row.get('mean', 0):>10.2f} {row.get('median', 0):>10.2f} {row.get('min', 0):>10.2f} "
                                f"{row.get('max', 0):>10.2f} {row.get('std', 0):>10.2f} {row.get('count', 0):>8}\n")
                    f.write("\n")

                if 'llm_cost_by_effective_mode' in metrics and metrics['llm_cost_by_effective_mode']:
                    f.write(f"LLM COST BY EFFECTIVE MODE\n")
                    f.write(f"=" * 25 + "\n")
                    f.write(f"{'Mode':<8} {'Mode Name':<30} {'Avg ($)':>10} {'Median':>10} {'Min':>10} {'Max':>10} {'Total ($)':>12} {'N':>8}\n")
                    f.write(f"{'-'*125}\n")
                    for row in metrics['llm_cost_by_effective_mode']:
                        f.write(f"{int(row.get('effective_mode', -1)):>8} {row.get('mode_name',''): <30} "
                                f"{row.get('mean', 0):>10.4f} {row.get('median', 0):>10.4f} {row.get('min', 0):>10.4f} "
                                f"{row.get('max', 0):>10.4f} {row.get('total', 0):>12.2f} {row.get('count', 0):>8}\n")
                    f.write("\n")
                # Failure rate by effective mode
                if 'failure_by_effective_mode' in metrics and metrics['failure_by_effective_mode']:
                    f.write(f"FAILURE RATE (ERROR COUNTS) BY MODE\n")
                    f.write(f"=" * 35 + "\n")
                    f.write(f"{'Mode':<6} {'Name':<24} {'Error':>8} {'Success (Info)':>16} {'Total':>8} {'Failure %':>10}\n")
                    f.write(f"{'-'*70}\n")
                    overall_err = overall_info = 0
                    for row in metrics['failure_by_effective_mode']:
                        mode = int(row.get('effective_mode', -1))
                        name = row.get('mode_name', '')
                        err = int(row.get('error', 0))
                        info = int(row.get('info', 0))
                        total = err + info
                        failure_pct = row.get('failure_pct', 0.0)
                        overall_err += err; overall_info += info
                        f.write(f"{mode:<6} {name:<24} {err:>8} {info:>16} {total:>8} {failure_pct:>9.2f}%\n")
                    overall_total = overall_err + overall_info
                    overall_pct = (overall_err / overall_total * 100) if overall_total else 0
                    f.write(f"{'—':<6} {'Overall':<24} {overall_err:>8} {overall_info:>16} {overall_total:>8} {overall_pct:>9.2f}%\n\n")

                # Process-wise failure table
                if 'failure_by_process' in metrics and metrics['failure_by_process']:
                    f.write(f"FAILURE RATE (ERROR COUNTS) BY PROCESS\n")
                    f.write(f"=" * 38 + "\n")
                    f.write(f"{'Process Name':<40} {'Error':>8} {'Success (Info)':>16} {'Total':>8} {'Failure %':>10}\n")
                    f.write(f"{'-'*95}\n")
                    for row in metrics['failure_by_process']:
                        err = int(row.get('error', 0)); info = int(row.get('info', 0)); total = err + info
                        failure_pct = (err / total * 100) if total else 0
                        f.write(f"{str(row.get(self.column_mappings.get('process_name'), row.get('process_name',''))):<40} {err:>8} {info:>16} {total:>8} {failure_pct:>9.2f}%\n")
                    f.write("\n")

                # Process x Mode combined tables (if present)
                if 'response_time_by_process_mode' in metrics and metrics['response_time_by_process_mode']:
                    f.write(f"RESPONSE TIME BY PROCESS × MODE\n")
                    f.write(f"=" * 32 + "\n")
                    f.write(f"{'Process Name':<40} {'Mode':>6} {'Avg (s)':>10} {'P50 (s)':>10} {'Min (s)':>10} {'Max (s)':>10} {'Std':>10} {'N':>8}\n")
                    f.write(f"{'-'*120}\n")
                    for row in metrics['response_time_by_process_mode']:
                        f.write(f"{str(row.get(self.column_mappings.get('process_name'), '')):<40} {int(row.get('effective_mode', -1)):>6} "
                                f"{row.get('mean', 0):>10.2f} {row.get('median', 0):>10.2f} {row.get('min', 0):>10.2f} "
                                f"{row.get('max', 0):>10.2f} {row.get('std', 0):>10.2f} {row.get('count', 0):>8}\n")
                    f.write("\n")

                if 'llm_cost_by_process_mode' in metrics and metrics['llm_cost_by_process_mode']:
                    f.write(f"LLM COST BY PROCESS × MODE\n")
                    f.write(f"=" * 27 + "\n")
                    f.write(f"{'Process Name':<40} {'Mode':>6} {'Avg ($)':>10} {'Median':>10} {'Min':>10} {'Max':>10} {'Total ($)':>12} {'N':>8}\n")
                    f.write(f"{'-'*125}\n")
                    for row in metrics['llm_cost_by_process_mode']:
                        f.write(f"{str(row.get(self.column_mappings.get('process_name'), '')):<40} {int(row.get('effective_mode', -1)):>6} "
                                f"{row.get('mean', 0):>10.4f} {row.get('median', 0):>10.4f} {row.get('min', 0):>10.4f} "
                                f"{row.get('max', 0):>10.4f} {row.get('total', 0):>12.2f} {row.get('count', 0):>8}\n")
                    f.write("\n")

                if 'failure_by_process_mode' in metrics and metrics['failure_by_process_mode']:
                    f.write(f"FAILURE RATE (ERROR COUNTS) BY PROCESS × MODE\n")
                    f.write(f"=" * 45 + "\n")
                    f.write(f"{'Process Name':<40} {'Mode':>6} {'Error':>8} {'Success (Info)':>16} {'Total':>8} {'Failure %':>10}\n")
                    f.write(f"{'-'*135}\n")
                    for row in metrics['failure_by_process_mode']:
                        err = int(row.get('error', 0)); info = int(row.get('info', 0)); total = err + info
                        failure_pct = (err / total * 100) if total else 0
                        f.write(f"{str(row.get(self.column_mappings.get('process_name'), '')):<40} {int(row.get('effective_mode', -1)):>6} {err:>8} {info:>16} {total:>8} {failure_pct:>9.2f}%\n")
                    f.write("\n")
                    f.write(f"LLM COST BY EFFECTIVE MODE\n")
                    f.write(f"=" * 25 + "\n")
                    f.write(f"{'Mode':<8} {'Mode Name':<30} {'Avg ($)':>10} {'Median':>10} {'Min':>10} {'Max':>10} {'Total ($)':>12} {'N':>8}\n")
                    f.write(f"{'-'*125}\n")
                    for row in metrics['llm_cost_by_effective_mode']:
                        f.write(f"{int(row.get('effective_mode', -1)):>8} {row.get('mode_name',''): <30} "
                                f"{row.get('mean', 0):>10.4f} {row.get('median', 0):>10.4f} {row.get('min', 0):>10.4f} "
                                f"{row.get('max', 0):>10.4f} {row.get('total', 0):>12.2f} {row.get('count', 0):>8}\n")
                    f.write("\n")

                # Failure/Success Rate (from preprocessed data)
                f.write(f"FAILURE/SUCCESS RATE (After Preprocessing)\n")
                f.write(f"=" * 45 + "\n")
                
                if 'status_analysis' in metrics:
                    status = metrics['status_analysis']
                    f.write(f"{'Status':<20} {'Count':<10} {'% of Total':<12}\n")
                    f.write(f"{'-'*42}\n")
                    f.write(f"{'error (Failure)':<20} {status['processed_errors']:<10} {status['error_rate']:.2f}%\n")
                    f.write(f"{'info (Success)':<20} {status['processed_success']:<10} {status['success_rate']:.2f}%\n")
                    f.write(f"{'Total':<20} {status['processed_total']:<10} 100.00%\n")
                    f.write(f"\n")
                    
                    # Note about processing
                    f.write(f"Processing Summary:\n")
                    f.write(f"- Original records: {status['original_total']:,}\n")
                    f.write(f"- Records after preprocessing: {status['processed_total']:,}\n")
                    f.write(f"- Records removed: {status['original_total'] - status['processed_total']:,}\n")
                    f.write(f"\n")
                else:
                    f.write(f"Status analysis not available (no status column found)\n\n")
                
                # Error Categories (NEW)
                if 'error_categories' in metrics and metrics['error_categories']:
                    f.write(f"ERROR TYPE CATEGORIES\n")
                    f.write(f"=" * 25 + "\n")
                    f.write(f"{'Error Category':<35} {'Count':<8}\n")
                    f.write(f"{'-'*43}\n")
                    
                    for category, count in metrics['error_categories'].items():
                        f.write(f"{category:<35} {count:<8}\n")
                    
                    f.write(f"\n")
                    f.write(f"Total error categories: {len(metrics['error_categories'])}\n")
                    f.write(f"Total categorized errors: {sum(metrics['error_categories'].values())}\n")
                    f.write(f"\n")
                
                # Detailed Error Breakdown
                if 'error_breakdown' in metrics and metrics['error_breakdown']:
                    f.write(f"DETAILED ERROR BREAKDOWN\n")
                    f.write(f"=" * 30 + "\n")
                    f.write(f"{'Error Message':<105} {'Count':<8}\n")
                    f.write(f"{'-'*113}\n")
                    
                    for error_msg, count in metrics['error_breakdown'].items():
                        # Show more of the error message (increased from 55 to 100 chars)
                        display_msg = str(error_msg)[:100] + "..." if len(str(error_msg)) > 100 else str(error_msg)
                        f.write(f"{display_msg:<105} {count:<8}\n")
                    
                    f.write(f"\n")
                    f.write(f"Total unique error messages: {len(metrics['error_breakdown'])}\n")
                    f.write(f"Total error occurrences: {sum(metrics['error_breakdown'].values())}\n")
                    f.write(f"\n")
                
                # Charts Information
                f.write(f"GENERATED CHARTS\n")
                f.write(f"=" * 20 + "\n")
                f.write(f"1. DAU Chart: dau_chart.png\n")
                f.write(f"   - Daily Active Users (total activities per day)\n")
                f.write(f"2. DAUU Chart: dauu_chart.png\n")
                f.write(f"   - Daily Active Unique Users (unique users per day)\n")
                if 'effective_mode' in self.df.columns:
                    f.write(f"3. Mode-wise DAU Chart: mode_wise_dau_chart.png\n")
                    f.write(f"   - Daily Active Users split by effective mode\n")
                f.write(f"\n")
                
                # Footer
                f.write(f"=" * 50 + "\n")
                f.write(f"Analysis completed successfully!\n")
                f.write(f"All files saved in: {self.output_dir}\n")
            
            print(f"✓ Comprehensive metrics saved: {txt_path}")
            return True
            
        except Exception as e:
            print(f"Error saving metrics to TXT: {e}")
            traceback.print_exc()
            return False
    
    def run_analysis(self) -> bool:
        """Run complete individual analysis"""
        print(f"🚀 STARTING INDIVIDUAL ANALYSIS: {self.file_name}")
        print("=" * 60)
        
        # Step 1: Load data and detect columns
        if not self.load_and_detect_columns():
            return False
        
        # Step 2: Preprocess data
        if not self.preprocess_data():
            return False
        # Update output directory based on detected Service column/value
        self._maybe_update_output_dir_with_service()
        
        # Step 3: Calculate metrics
        metrics = self.calculate_metrics()
        if not metrics:
            print("❌ Failed to calculate metrics")
            return False
        
        # Step 4: Create charts (non-fatal)
        if not self.create_dau_dauu_charts():
            print("⚠️  Failed to create DAU/DAUU charts; continuing")
        
        # Step 4b: Create mode-wise DAU chart if effective_mode exists (non-fatal)
        if 'effective_mode' in self.df.columns:
            if not self.create_mode_wise_dau_chart():
                print("⚠️  Failed to create mode-wise DAU chart (skipping)")
        
        # Step 5: Create response time analysis charts (non-fatal)
        if not self.create_response_time_charts():
            print("⚠️  Failed to create response time charts; continuing")
        
        # Step 6: Save metrics to TXT (non-fatal)
        if not self.save_metrics_to_txt(metrics):
            print("⚠️  Failed to save metrics; continuing")
        
        # Step 7: Generate daily analysis (non-fatal)
        if not self.generate_daily_analysis():
            print("⚠️  Failed to generate daily analysis; continuing")

        print(f"\n✅ ANALYSIS COMPLETED SUCCESSFULLY!")
        print(f"📁 All outputs saved in: {self.output_dir}")
        print(f"📊 Generated files:")
        print(f"   - dau_chart.png (Daily Active Users)")
        print(f"   - dauu_chart.png (Daily Active Unique Users)")
        print(f"   - response_time_analysis.png (Comprehensive RT Analysis)")
        print(f"   - response_time_percentiles.png (RT Percentiles Summary)")
        print(f"   - daily_response_time_range.png (Daily Min/Max/Avg Chart)")
        print(f"   - metrics_analysis.txt (Complete metrics)")
        print(f"   - daily_analysis.txt (Day-over-Day Analysis with 5 Key Metrics)")
        
        return True

    def create_mode_wise_dau_chart(self) -> bool:
        """Create a mode-wise DAU chart if effective_mode exists."""
        try:
            if 'effective_mode' not in self.df.columns:
                return True
            # Map mode numbers to names
            mode_map = {
                1: 'isDocument', 2: 'isInternet', 3: 'isDatabase', 4: 'isDirectTaxCode', 5: 'isGlobal',
                6: 'isHarvey', 7: 'isDatabaseGeneric', 8: 'isNLP', 9: 'isDeepResearch', 10: 'isDraft',
                11: 'isAutoMode', 12: 'isMultipleDbGeneric', 13: 'isDatabaseGenericVersion2', 14: 'isDatabaseGenericLite',
                15: 'isDeepResearchWebSearch', 0: 'UnresolvedRedirect'
            }
            if 'formatted_date' not in self.df.columns:
                return True
            # Compute daily counts per effective_mode
            df = self.df.copy()
            df['effective_mode'] = pd.to_numeric(df['effective_mode'], errors='coerce')
            df = df.dropna(subset=['effective_mode'])
            if df.empty:
                return True
            grouped = df.groupby(['formatted_date', 'effective_mode']).size().reset_index(name='count')
            # Pivot to have modes as series over dates
            pivot = grouped.pivot(index='formatted_date', columns='effective_mode', values='count').fillna(0)
            # Prepare plot
            plt.figure(figsize=(16, 9))
            x_positions = range(len(pivot.index))
            date_labels = [pd.to_datetime(d).strftime('%m-%d') for d in pivot.index]
            # Plot each mode as a line
            for mode in pivot.columns:
                series = pivot[mode].values
                mode_name = mode_map.get(int(mode), str(int(mode)))
                plt.plot(x_positions, series, marker='o', linewidth=2, markersize=5, label=f"{mode_name} ({int(mode)})")
                # Add data label for the last point of each series to reduce clutter
                if len(series) > 0:
                    last_x = x_positions[-1]
                    last_y = series[-1]
                    plt.annotate(f"{int(last_y)}", xy=(last_x, last_y), xytext=(0, 6), textcoords='offset points',
                                 fontsize=8, ha='center', va='bottom')
            plt.title(f"{self.file_name} - Mode-wise Daily Active Users (DAU)\nWeekdays Only (Continuous Timeline)", fontsize=16, fontweight='bold', pad=20)
            plt.xlabel('Date (Weekdays Only)', fontsize=14, fontweight='bold')
            plt.ylabel('Daily Active Users', fontsize=14, fontweight='bold')
            plt.xticks(x_positions, date_labels, rotation=45, ha='right')
            plt.yticks(fontsize=12)
            plt.grid(True, alpha=0.3, linestyle='--')
            plt.legend(fontsize=9, ncol=2, loc='upper left')
            plt.tight_layout()
            out_path = f"{self.output_dir}/mode_wise_dau_chart.png"
            plt.savefig(out_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            print(f"✓ Mode-wise DAU chart saved: {out_path}")
            return True
        except Exception as e:
            print(f"Error creating mode-wise DAU chart: {e}")
            traceback.print_exc()
            return False

        # Step 4b: Create mode-wise DAU chart if effective_mode exists
        if 'effective_mode' in self.df.columns:
            if not self.create_mode_wise_dau_chart():
                print("⚠️  Failed to create mode-wise DAU chart (skipping)")
        
        # Step 5: Create response time analysis charts
        if not self.create_response_time_charts():
            print("❌ Failed to create response time charts")
            return False
        
        # Step 6: Save metrics to TXT
        if not self.save_metrics_to_txt(metrics):
            print("❌ Failed to save metrics")
            return False
        
        # Step 7: Generate daily analysis
        if not self.generate_daily_analysis():
            print("❌ Failed to generate daily analysis")
            return False
        
        print(f"\n✅ ANALYSIS COMPLETED SUCCESSFULLY!")
        print(f"📁 All outputs saved in: {self.output_dir}")
        print(f"📊 Generated files:")
        print(f"   - dau_chart.png (Daily Active Users)")
        print(f"   - dauu_chart.png (Daily Active Unique Users)")
        print(f"   - response_time_analysis.png (Comprehensive RT Analysis)")
        print(f"   - response_time_percentiles.png (RT Percentiles Summary)")
        print(f"   - daily_response_time_range.png (Daily Min/Max/Avg Chart)")
        print(f"   - metrics_analysis.txt (Complete metrics)")
        print(f"   - daily_analysis.txt (Day-over-Day Analysis with 5 Key Metrics)")
        
        return True


def analyze_file(file_path: str, compare: Optional[Tuple[str, str]] = None) -> bool:
    """Analyze a single file"""
    analyzer = SimpleIndividualAnalyzer(file_path, compare_dates=compare)
    return analyzer.run_analysis()


def analyze_all_source_files():
    """Analyze all files in source_data directory"""
    source_dir = "/Users/shtlpmac027/Documents/DataDog/source_data"
    
    if not os.path.exists(source_dir):
        print(f"❌ Source directory not found: {source_dir}")
        return
    
    # Find Excel files
    excel_files = []
    for file in os.listdir(source_dir):
        if file.endswith(('.xlsx', '.xls')):
            excel_files.append(os.path.join(source_dir, file))
    
    if not excel_files:
        print(f"❌ No Excel files found in {source_dir}")
        return
    
    print(f"Found {len(excel_files)} Excel files to analyze:")
    for file in excel_files:
        print(f"  - {os.path.basename(file)}")
    
    print(f"\n" + "=" * 80)
    print(f"STARTING INDIVIDUAL ANALYSIS FOR ALL FILES")
    print(f"=" * 80)
    
    successful = []
    failed = []
    
    for i, file_path in enumerate(excel_files, 1):
        file_name = os.path.basename(file_path)
        print(f"\n🔄 Analyzing file {i}/{len(excel_files)}: {file_name}")
        print("-" * 60)
        
        try:
            if analyze_file(file_path):
                successful.append(file_name)
                print(f"✅ Successfully analyzed: {file_name}")
            else:
                failed.append(file_name)
                print(f"❌ Failed to analyze: {file_name}")
        except Exception as e:
            failed.append(file_name)
            print(f"❌ Error analyzing {file_name}: {e}")
    
    # Final summary
    print(f"\n" + "=" * 80)
    print(f"ANALYSIS SUMMARY")
    print(f"=" * 80)
    print(f"✅ Successfully analyzed: {len(successful)} files")
    for file in successful:
        print(f"   - {file}")
    
    if failed:
        print(f"\n❌ Failed to analyze: {len(failed)} files")
        for file in failed:
            print(f"   - {file}")
    
    print(f"\n📁 All individual analyses saved in:")
    print(f"   /Users/shtlpmac027/Documents/DataDog/individual_analysis/")
    print(f"\nEach file folder contains:")
    print(f"   - dau_chart.png")
    print(f"   - dauu_chart.png") 
    print(f"   - metrics_analysis.txt")


if __name__ == "__main__":
    import sys
    import argparse
    parser = argparse.ArgumentParser(description='Run individual analysis for a file or all source files')
    parser.add_argument('file', nargs='?', help='Path to specific file to analyze (optional)')
    parser.add_argument('--compare', help='Date pair to compare in daily analysis, format dd/mm,dd/mm or yyyy-mm-dd,yyyy-mm-dd')
    args = parser.parse_args()

    compare_tuple: Optional[Tuple[str, str]] = None
    if args.compare:
        try:
            parts = [p.strip() for p in args.compare.split(',')]
            if len(parts) == 2:
                compare_tuple = (parts[0], parts[1])
            else:
                print('⚠️  --compare expects two dates separated by a comma')
        except Exception:
            print('⚠️  Failed parsing --compare; proceeding without it')
            compare_tuple = None

    if args.file:
        file_path = args.file
        if os.path.exists(file_path):
            analyze_file(file_path, compare_tuple)
        else:
            print(f"❌ File not found: {file_path}")
    else:
        analyze_all_source_files()
