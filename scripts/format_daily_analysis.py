#!/usr/bin/env python3
"""
Enhanced script to format daily analysis files with color highlighting and bold formatting
"""

import pandas as pd
import os
import glob
import re
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

def parse_daily_analysis_file(file_path):
    """Parse a daily analysis file and extract metrics in table format"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Extract service name from directory
    service_name = os.path.basename(os.path.dirname(file_path))
    
    # Extract dates from filename
    filename = os.path.basename(file_path)
    match = re.search(r'daily_analysis_(\d+-\d+)_vs_(\d+-\d+)\.txt', filename)
    if not match:
        return None
    
    date1, date2 = match.groups()
    
    # Parse metrics from content
    metrics = {}
    
    # Parse each metric section
    sections = content.split('\n\n')
    
    for section in sections:
        if 'Latency Metric' in section:
            metrics['Latency'] = parse_metric_section(section, date1, date2)
        elif 'Throughput Metric' in section:
            metrics['Throughput'] = parse_metric_section(section, date1, date2)
        elif 'LLM Cost Metric' in section:
            metrics['LLM Cost'] = parse_metric_section(section, date1, date2)
        elif 'Reliability Metric' in section:
            metrics['Reliability'] = parse_metric_section(section, date1, date2)
        elif 'User Activity Metric' in section:
            metrics['User Activity'] = parse_metric_section(section, date1, date2)
    
    return {
        'service': service_name,
        'date1': date1,
        'date2': date2,
        'metrics': metrics
    }

def parse_metric_section(section, date1, date2):
    """Parse a metric section and extract values"""
    
    lines = section.strip().split('\n')
    
    # Extract today's value
    today_value = None
    yesterday_value = None
    change_text = None
    status = None
    
    for line in lines:
        if "Today's" in line and ":" in line:
            # Extract numeric value (handle different formats)
            if "Cost" in line:
                # For cost: "Today's Total Cost ($): 0.59"
                match = re.search(r':\s*\$?([\d.]+)', line)
            elif "Rate" in line:
                # For success rate: "Today's Success Rate: 99.9%"
                match = re.search(r':\s*([\d.]+)', line)
            else:
                # For other metrics: "Today's Avg Response Time: 1.354ms"
                match = re.search(r':\s*([\d.]+)', line)
            
            if match:
                today_value = match.group(1)
                
        elif "Yesterday's" in line and ":" in line:
            # Extract numeric value (handle different formats)
            if "Cost" in line:
                # For cost: "Yesterday's Total Cost ($): 0.64"
                match = re.search(r':\s*\$?([\d.]+)', line)
            elif "Rate" in line:
                # For success rate: "Yesterday's Success Rate: 99.2%"
                match = re.search(r':\s*([\d.]+)', line)
            else:
                # For other metrics: "Yesterday's Avg Response Time: 1.210ms"
                match = re.search(r':\s*([\d.]+)', line)
            
            if match:
                yesterday_value = match.group(1)
                
        elif "Change:" in line or "Change ($):" in line:
            if "Change ($):" in line:
                change_text = line.split('Change ($):')[1].strip()
            else:
                change_text = line.split('Change:')[1].strip()
        elif "Status:" in line:
            status = line.split('Status:')[1].strip()
    
    return {
        'date1_value': today_value,
        'date2_value': yesterday_value,
        'change': change_text,
        'status': status
    }

def get_status_color(status):
    """Get color based on status"""
    if not status:
        return None
    
    status = status.upper()
    
    # Positive statuses - Green
    if status in ['IMPROVING', 'GROWING', 'EFFICIENT']:
        return PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    
    # Negative statuses - Red
    elif status in ['DEGRADING', 'DECLINING', 'EXPENSIVE']:
        return PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    
    # Neutral statuses - Yellow
    elif status == 'STABLE':
        return PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    return None

def get_change_color(change_text):
    """Get color based on change direction"""
    if not change_text:
        return None
    
    # Check for improvement indicators
    if any(indicator in change_text.upper() for indicator in ['↓', 'IMPROVEMENT', 'DECREASE']):
        return PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    
    # Check for degradation indicators
    elif any(indicator in change_text.upper() for indicator in ['↑', 'INCREASE', 'DEGRADATION']):
        return PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    
    # No change
    elif 'NO CHANGE' in change_text.upper() or '0.0%' in change_text:
        return PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    return None

def create_formatted_excel():
    """Create enhanced Excel with color highlighting and bold formatting"""
    
    base_dir = "/Users/shtlpmac027/Documents/DataDog/individual_analysis"
    daily_files = glob.glob(f"{base_dir}/**/daily_analysis_*.txt", recursive=True)
    
    # Group by date comparison
    date_groups = {}
    
    for file_path in daily_files:
        parsed_data = parse_daily_analysis_file(file_path)
        if parsed_data:
            date_key = f"{parsed_data['date1']}_vs_{parsed_data['date2']}"
            
            if date_key not in date_groups:
                date_groups[date_key] = []
            
            date_groups[date_key].append(parsed_data)
    
    # Create Excel file
    output_file = "/Users/shtlpmac027/Documents/DataDog/formatted_daily_analysis.xlsx"
    
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # Define styles
    bold_font = Font(bold=True, size=12)
    header_font = Font(bold=True, size=11)
    service_font = Font(bold=True, size=12, color='2F4F4F')
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    for date_key, services_data in date_groups.items():
        print(f"Processing date comparison: {date_key}")
        
        # Create new worksheet
        ws = wb.create_sheet(title=f"Daily_Analysis_{date_key.replace('-', '_')[:25]}")
        
        # Headers
        headers = ['Service', f'{date_key.split("_vs_")[0]}', f'{date_key.split("_vs_")[1]}', 'Change', 'Status']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
            cell.font = Font(bold=True, color='FFFFFF')
            cell.alignment = center_alignment
            cell.border = border
        
        current_row = 2
        
        for service_data in services_data:
            service_name = service_data['service']
            date1 = service_data['date1']
            date2 = service_data['date2']
            metrics = service_data['metrics']
            
            # Add service header row
            service_cell = ws.cell(row=current_row, column=1, value=service_name)
            service_cell.font = service_font
            service_cell.fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
            service_cell.alignment = center_alignment
            service_cell.border = border
            
            # Merge cells for service name
            ws.merge_cells(f'A{current_row}:E{current_row}')
            
            current_row += 1
            
            # Add metrics for this service
            for metric_name, metric_data in metrics.items():
                # Service column
                metric_cell = ws.cell(row=current_row, column=1, value=f'{metric_name} Metric')
                metric_cell.font = Font(bold=True)
                metric_cell.border = border
                
                # Date 1 value
                date1_cell = ws.cell(row=current_row, column=2, value=metric_data['date1_value'] or '')
                date1_cell.border = border
                date1_cell.alignment = center_alignment
                
                # Date 2 value
                date2_cell = ws.cell(row=current_row, column=3, value=metric_data['date2_value'] or '')
                date2_cell.border = border
                date2_cell.alignment = center_alignment
                
                # Change column with color highlighting
                change_cell = ws.cell(row=current_row, column=4, value=metric_data['change'] or '')
                change_cell.border = border
                change_cell.alignment = center_alignment
                change_fill = get_change_color(metric_data['change'])
                if change_fill:
                    change_cell.fill = change_fill
                
                # Status column with color highlighting
                status_cell = ws.cell(row=current_row, column=5, value=metric_data['status'] or '')
                status_cell.border = border
                status_cell.alignment = center_alignment
                status_fill = get_status_color(metric_data['status'])
                if status_fill:
                    status_cell.fill = status_fill
                
                current_row += 1
            
            # Add empty row between services
            current_row += 1
        
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Created sheet: Daily_Analysis_{date_key.replace('-', '_')} with {current_row-1} rows")
    
    # Save the workbook
    wb.save(output_file)
    print(f"\n🎉 Enhanced daily analysis created: {output_file}")

if __name__ == "__main__":
    create_formatted_excel()
