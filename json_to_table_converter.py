#!/usr/bin/env python3
"""
JSON to Table Converter for Seattle Opera Data

This script converts JSON data containing show information to CSV or Excel format
with the table structure: SHOW, DATES, ROLE, ARTIST, OTHER

Features:
- Process single JSON files or entire directories
- Append data to existing files
- Support for both CSV and Excel output formats
- Recursive directory processing
- Custom file patterns
"""

import json
import csv
import pandas as pd
import argparse
import glob
import os
from pathlib import Path
from typing import List, Dict, Any


def load_json_data(file_path: str) -> Dict[str, Any]:
    """Load JSON data from file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return json.load(file)
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        raise
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in file '{file_path}': {e}")
        raise


def extract_table_data(json_data: Dict[str, Any]) -> List[Dict[str, str]]:
    """
    Extract table data from JSON structure.
    Returns a list of dictionaries with keys: SHOW, DATES, ROLE, ARTIST, OTHER
    """
    table_data = []
    
    # Navigate through the JSON structure
    if 'result' in json_data and 'contents' in json_data['result']:
        contents = json_data['result']['contents']
        
        for content in contents:
            if 'fields' in content:
                fields = content['fields']
                
                # Extract show name
                show = ""
                if 'SHOW' in fields and 'valueString' in fields['SHOW']:
                    show = fields['SHOW']['valueString']
                
                # Extract dates
                dates = ""
                if 'DATES' in fields and 'valueString' in fields['DATES']:
                    dates = fields['DATES']['valueString']
                elif 'DATE' in fields and 'valueDate' in fields['DATE']:
                    dates = fields['DATE']['valueDate']
                
                # Extract roles and artists
                if 'ROLES' in fields and 'valueArray' in fields['ROLES']:
                    for role_entry in fields['ROLES']['valueArray']:
                        if 'valueObject' in role_entry:
                            role_obj = role_entry['valueObject']
                            
                            role = ""
                            artist = ""
                            other = ""
                            
                            if 'ROLE' in role_obj and 'valueString' in role_obj['ROLE']:
                                role = role_obj['ROLE']['valueString']
                            
                            if 'ARTIST' in role_obj and 'valueString' in role_obj['ARTIST']:
                                artist = role_obj['ARTIST']['valueString']
                            
                            if 'OTHER' in role_obj and 'valueString' in role_obj['OTHER']:
                                other = role_obj['OTHER']['valueString']
                            
                            # Add row to table data
                            if role and artist:  # Only add if both role and artist exist
                                table_data.append({
                                    'SHOW': show,
                                    'DATES': dates,
                                    'ROLE': role,
                                    'ARTIST': artist,
                                    'OTHER': other
                                })
    
    return table_data


def save_to_csv(data: List[Dict[str, str]], output_path: str, append_mode: bool = False):
    """Save data to CSV file."""
    if not data:
        print("No data to save.")
        return
    
    file_exists = os.path.exists(output_path)
    mode = 'a' if append_mode else 'w'
    
    with open(output_path, mode, newline='', encoding='utf-8') as csvfile:
        fieldnames = ['SHOW', 'DATES', 'ROLE', 'ARTIST', 'OTHER']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        # Write header only if file doesn't exist or not in append mode
        if not file_exists or not append_mode:
            writer.writeheader()
        
        # Write data rows
        for row in data:
            writer.writerow(row)
    
    action = "appended to" if append_mode and file_exists else "saved to"
    print(f"Data successfully {action} CSV: {output_path}")


def save_to_excel(data: List[Dict[str, str]], output_path: str, append_mode: bool = False):
    """Save data to Excel file."""
    if not data:
        print("No data to save.")
        return
    
    try:
        # Create DataFrame from new data
        new_df = pd.DataFrame(data)
        
        # Ensure columns are in the correct order
        new_df = new_df[['SHOW', 'DATES', 'ROLE', 'ARTIST', 'OTHER']]
        
        if append_mode and os.path.exists(output_path):
            # Read existing data and append new data
            existing_df = pd.read_excel(output_path, engine='openpyxl')
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            combined_df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"Data successfully appended to Excel: {output_path}")
        else:
            # Save new file or overwrite existing
            new_df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"Data successfully saved to Excel: {output_path}")
        
    except ImportError:
        print("Error: openpyxl library not found. Please install it with: pip install openpyxl")
        raise


def process_multiple_files(input_pattern: str, output_path: str, output_format: str, append_mode: bool = False):
    """Process multiple JSON files matching the given pattern."""
    json_files = glob.glob(input_pattern)
    
    if not json_files:
        print(f"No JSON files found matching pattern: {input_pattern}")
        return
    
    print(f"Found {len(json_files)} JSON files to process.")
    
    all_data = []
    processed_files = 0
    
    for json_file in json_files:
        try:
            print(f"Processing: {json_file}")
            json_data = load_json_data(json_file)
            table_data = extract_table_data(json_data)
            
            if table_data:
                all_data.extend(table_data)
                processed_files += 1
                print(f"  Extracted {len(table_data)} rows")
            else:
                print(f"  No valid data found in {json_file}")
                
        except Exception as e:
            print(f"  Error processing {json_file}: {e}")
            continue
    
    if not all_data:
        print("No valid data extracted from any files.")
        return
    
    print(f"\nTotal extracted {len(all_data)} rows from {processed_files} files.")
    
    # Save all data
    if output_format == 'excel':
        save_to_excel(all_data, output_path, append_mode)
    else:
        save_to_csv(all_data, output_path, append_mode)
    
    return all_data


def main():
    """Main function to handle command line arguments and conversion."""
    parser = argparse.ArgumentParser(description="Convert JSON data to CSV or Excel table format")
    parser.add_argument('input', help='Path to input JSON file or folder containing JSON files')
    parser.add_argument('-o', '--output', help='Output file path (default: auto-generated)')
    parser.add_argument('-f', '--format', choices=['csv', 'excel'], default='csv',
                       help='Output format: csv or excel (default: csv)')
    parser.add_argument('-a', '--append', action='store_true',
                       help='Append data to existing output file instead of overwriting')
    parser.add_argument('-r', '--recursive', action='store_true',
                       help='Process JSON files recursively in subdirectories')
    parser.add_argument('--pattern', default='*.json',
                       help='File pattern for JSON files (default: *.json)')
    
    args = parser.parse_args()
    
    # Check if input is a directory or a single file
    input_path = Path(args.input)
    
    if input_path.is_dir():
        # Process multiple files in directory
        if args.recursive:
            pattern = os.path.join(args.input, '**', args.pattern)
            input_pattern = glob.glob(pattern, recursive=True)
        else:
            pattern = os.path.join(args.input, args.pattern)
            input_pattern = pattern
        
        # Determine output file path for multiple files
        if args.output:
            output_path = args.output
        else:
            if args.format == 'excel':
                output_path = os.path.join(args.input, 'combined_data.xlsx')
            else:
                output_path = os.path.join(args.input, 'combined_data.csv')
        
        # Process multiple files
        table_data = process_multiple_files(input_pattern, output_path, args.format, args.append)
        
    elif input_path.is_file():
        # Process single file
        print(f"Loading JSON data from: {args.input}")
        json_data = load_json_data(args.input)
        
        # Extract table data
        print("Extracting table data...")
        table_data = extract_table_data(json_data)
        
        if not table_data:
            print("No valid data found in the JSON file.")
            return
        
        print(f"Extracted {len(table_data)} rows of data.")
        
        # Determine output file path
        if args.output:
            output_path = args.output
        else:
            if args.format == 'excel':
                output_path = input_path.with_suffix('.xlsx')
            else:
                output_path = input_path.with_suffix('.csv')
        
        # Save data
        if args.format == 'excel':
            save_to_excel(table_data, str(output_path), args.append)
        else:
            save_to_csv(table_data, str(output_path), args.append)
    
    else:
        print(f"Error: Input path '{args.input}' is neither a file nor a directory.")
        return
    
    # Display sample of the data
    if table_data:
        print("\nSample of converted data (first 5 rows):")
        print("-" * 90)
        print(f"{'SHOW':<20} {'DATES':<10} {'ROLE':<25} {'ARTIST':<20} {'OTHER'}")
        print("-" * 90)
        for i, row in enumerate(table_data[:5]):
            print(f"{row['SHOW']:<20} {row['DATES']:<10} {row['ROLE']:<25} {row['ARTIST']:<20} {row['OTHER']}")
        
        if len(table_data) > 5:
            print(f"... and {len(table_data) - 5} more rows")


if __name__ == "__main__":
    main()