#!/usr/bin/env python3
"""
Batch process all companies in the 饮料 (beverage) industry spreadsheet.
"""

import sys
import os
import json

sys.path.insert(0, os.path.expanduser('~/.hermes/hermes-agent'))
sys.path.insert(0, '/home/mike/Projects/gsupdater')

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from add_yoy_section import add_yoy_section

GOOGLE_TOKEN_PATH = os.path.expanduser('~/.hermes/google_token.json')

def main():
    # Load beverage spreadsheet ID
    with open('/home/mike/Projects/gsupdater/industry_spreadsheets.json', 'r') as f:
        data = json.load(f)
    spreadsheet_id = data['饮料']['spreadsheet_id']
    print(f"饮料 spreadsheet ID: {spreadsheet_id}")
    
    # Get credentials
    creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_PATH)
    service = build('sheets', 'v4', credentials=creds)
    
    # Get all company sheets from Summary
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="Summary!A1:AZ2"
    ).execute()
    rows = result.get('values', [])
    
    companies = []
    if len(rows) >= 2:
        codes = rows[0]
        names = rows[1]
        for j in range(len(codes)):
            code = str(codes[j]).strip() if j < len(codes) else ''
            name = str(names[j]).strip() if j < len(names) else ''
            if code and name:
                sheet_name = f"{name}财务"
                companies.append((code, name, sheet_name))
    
    print(f"\nFound {len(companies)} companies:")
    for code, name, sheet_name in companies:
        print(f"  {code} - {name} ({sheet_name})")
    
    # Process each company
    print("\n" + "="*60)
    print("Processing all companies...")
    print("="*60 + "\n")
    
    success_count = 0
    skip_count = 0
    error_count = 0
    
    for code, name, sheet_name in companies:
        print(f"\n{'='*60}")
        print(f"Processing: {code} - {name}")
        print(f"{'='*60}")
        
        try:
            result = add_yoy_section(service, spreadsheet_id, sheet_name, dry_run=False)
            if result:
                success_count += 1
            else:
                skip_count += 1
        except Exception as e:
            print(f"ERROR processing {name}: {e}")
            error_count += 1
    
    print("\n" + "="*60)
    print(f"BATCH COMPLETE: {success_count} success, {skip_count} skipped, {error_count} errors")
    print("="*60)

if __name__ == '__main__':
    main()
