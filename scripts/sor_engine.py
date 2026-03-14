#!/usr/bin/env python3
"""
Earthsoft SOR Processing Engine
================================
Parses SOR PDFs (IRUSSOR 2021, DSR 2023, etc.) and uploads them to Supabase.

Usage:
  python3 sor_engine.py upload --pdf /path/to/file.pdf --name "IRUSSOR 2021" --key irussor_2021
  python3 sor_engine.py upload --json /path/to/file.json --name "DSR 2023" --key dsr_2023
  python3 sor_engine.py list
  python3 sor_engine.py search --query "earthwork" --source irussor_2021

Requires: pdfplumber (for PDF parsing only)
"""

import json
import re
import sys
import urllib.request
import urllib.error
import argparse
import os

# ──────────────────────────────────────────────────────────────────────
# Supabase Configuration (same as supabase-config.js)
# ──────────────────────────────────────────────────────────────────────
SUPABASE_URL = os.environ.get(
    'SUPABASE_URL',
    'https://gpiujhtcuagyopinszbe.supabase.co'
)
SUPABASE_KEY = os.environ.get(
    'SUPABASE_KEY',
    'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdwaXVqaHRjdWFneW9waW5zemJlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMzMjI3NDAsImV4cCI6MjA4ODg5ODc0MH0.5VFWVFKuFvTbFbBJlpSl7c8U8cueb_uBWgIV5fSjrP4'
)

HEADERS = {
    'apikey': SUPABASE_KEY,
    'Authorization': f'Bearer {SUPABASE_KEY}',
    'Content-Type': 'application/json',
    'Prefer': 'return=representation'
}


def supabase_request(method, table, data=None, params=''):
    """Make a request to Supabase REST API."""
    url = f'{SUPABASE_URL}/rest/v1/{table}{params}'
    body = json.dumps(data).encode('utf-8') if data else None
    req = urllib.request.Request(url, data=body, headers=HEADERS, method=method)
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8')
        print(f'  Supabase API error ({e.code}): {error_body}')
        raise


# ──────────────────────────────────────────────────────────────────────
# PDF Parser
# ──────────────────────────────────────────────────────────────────────
def parse_sor_pdf(pdf_path):
    """Parse an SOR PDF and return list of {code, description, unit, rate}."""
    try:
        import pdfplumber
    except ImportError:
        print('Error: pdfplumber is required for PDF parsing.')
        print('Install it with: pip3 install pdfplumber')
        sys.exit(1)

    items = []
    item_pattern = re.compile(
        r'^(\d{4,6}[a-zA-Z]?)\s+(.*?)\s+'
        r'(Cum|Sqm|100\sSqm|Km|RM|Rm|Tonne|Ea|Kg|Pair|Rkm|Tkm|'
        r'Sqcm|10\sCum|100\sCum|1000\sCum|Litre|Day|Month|L\.S\.|'
        r'Quintal|Sq\.m\.|Cu\.m\.|KM|Set|No\.?|Each|Per\sKm|MT)\s+'
        r'([0-9,]+\.\d{2})',
        re.IGNORECASE
    )

    with pdfplumber.open(pdf_path) as pdf:
        print(f'  Opened PDF: {pdf_path} ({len(pdf.pages)} pages)')
        current_item = None

        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            for line in text.split('\n'):
                match = item_pattern.search(line)
                if match:
                    if current_item:
                        items.append(current_item)

                    current_item = {
                        'code': match.group(1).strip(),
                        'description': match.group(2).strip(),
                        'unit': match.group(3).strip(),
                        'rate': float(match.group(4).replace(',', ''))
                    }
                elif current_item and not re.match(r'^\d', line.strip()):
                    # Continuation of description
                    current_item['description'] += ' ' + line.strip()

        if current_item:
            items.append(current_item)

    print(f'  Extracted {len(items)} items from PDF')
    return items


def parse_sor_json(json_path):
    """Load SOR from a pre-formatted JSON file."""
    with open(json_path, 'r', encoding='utf-8') as f:
        items = json.load(f)
    if not isinstance(items, list):
        print('Error: JSON must be an array of items.')
        sys.exit(1)
    print(f'  Loaded {len(items)} items from JSON')
    return items


# ──────────────────────────────────────────────────────────────────────
# Upload to Supabase
# ──────────────────────────────────────────────────────────────────────
def upload_to_supabase(items, source_key, display_name, description=''):
    """Upload parsed SOR items to Supabase database."""

    print(f'\n  Uploading to Supabase...')
    print(f'  Source: {display_name} ({source_key})')
    print(f'  Items: {len(items)}')

    # Step 1: Check if source already exists
    try:
        existing = supabase_request('GET', 'sor_sources', params=f'?source_key=eq.{source_key}&select=id')
        if existing and len(existing) > 0:
            source_id = existing[0]['id']
            print(f'  Source already exists (id: {source_id}). Replacing items...')
            # Delete existing items
            supabase_request('DELETE', 'sor_items', params=f'?source_id=eq.{source_id}')
            # Update source metadata
            supabase_request('PATCH', 'sor_sources',
                data={'display_name': display_name, 'item_count': len(items), 'description': description},
                params=f'?source_key=eq.{source_key}')
        else:
            # Create new source
            result = supabase_request('POST', 'sor_sources', data={
                'source_key': source_key,
                'display_name': display_name,
                'description': description,
                'item_count': len(items)
            })
            source_id = result[0]['id'] if isinstance(result, list) else result['id']
            print(f'  Created source (id: {source_id})')
    except Exception as e:
        print(f'  Error creating/updating source: {e}')
        print('\n  ⚠  Make sure you have run the SQL schema first!')
        print('     See: scripts/supabase_sor_schema.sql')
        return False

    # Step 2: Upload items in batches of 100
    batch_size = 100
    total_uploaded = 0

    for i in range(0, len(items), batch_size):
        batch = items[i:i + batch_size]
        rows = []
        for item in batch:
            rows.append({
                'source_id': source_id,
                'item_code': str(item.get('code', '')).strip(),
                'description': str(item.get('description', '')).strip(),
                'unit': str(item.get('unit', '')).strip(),
                'rate': float(item.get('rate', 0)),
                'chapter': str(item.get('chapter', '')).strip() or None
            })

        try:
            supabase_request('POST', 'sor_items', data=rows)
            total_uploaded += len(rows)
            pct = round(total_uploaded / len(items) * 100)
            print(f'  Uploaded {total_uploaded}/{len(items)} items ({pct}%)', end='\r')
        except Exception as e:
            print(f'\n  Error uploading batch at index {i}: {e}')

    print(f'\n  ✅ Successfully uploaded {total_uploaded} items to Supabase!')
    return True


# ──────────────────────────────────────────────────────────────────────
# List / Search
# ──────────────────────────────────────────────────────────────────────
def list_sources():
    """List all SOR sources in the database."""
    try:
        sources = supabase_request('GET', 'sor_sources', params='?select=*&order=uploaded_at.desc')
        if not sources:
            print('No SOR sources found. Upload one first.')
            return
        print(f'\n{"Source Key":<20} {"Name":<25} {"Items":>8}  {"Date"}')
        print('-' * 70)
        for s in sources:
            dt = s.get('uploaded_at', '')[:10]
            print(f'{s["source_key"]:<20} {s["display_name"]:<25} {s["item_count"]:>8}  {dt}')
    except Exception as e:
        print(f'Error listing sources: {e}')


def search_items(query, source_key=None):
    """Search SOR items by code or description."""
    params = f'?select=item_code,description,unit,rate,sor_sources(display_name)'
    params += f'&or=(item_code.ilike.*{query}*,description.ilike.*{query}*)'
    if source_key:
        params += f'&sor_sources.source_key=eq.{source_key}'
    params += '&limit=20'

    try:
        results = supabase_request('GET', 'sor_items', params=params)
        if not results:
            print('No items found.')
            return
        print(f'\n{"Code":<10} {"Unit":<12} {"Rate":>12}  Description')
        print('-' * 80)
        for r in results:
            desc = (r['description'] or '')[:50]
            print(f'{r["item_code"]:<10} {r["unit"]:<12} {r["rate"]:>12,.2f}  {desc}')
        print(f'\n{len(results)} results shown.')
    except Exception as e:
        print(f'Error searching: {e}')


# ──────────────────────────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='Earthsoft SOR Processing Engine — Parse PDFs and upload to Supabase'
    )
    subparsers = parser.add_subparsers(dest='command')

    # upload
    upload_parser = subparsers.add_parser('upload', help='Parse and upload SOR to database')
    upload_parser.add_argument('--pdf', help='Path to SOR PDF file')
    upload_parser.add_argument('--json', help='Path to SOR JSON file')
    upload_parser.add_argument('--name', required=True, help='Display name (e.g. "IRUSSOR 2021")')
    upload_parser.add_argument('--key', required=True, help='Source key (e.g. "irussor_2021")')
    upload_parser.add_argument('--desc', default='', help='Optional description')

    # list
    subparsers.add_parser('list', help='List all SOR sources in database')

    # search
    search_parser = subparsers.add_parser('search', help='Search SOR items')
    search_parser.add_argument('--query', '-q', required=True, help='Search term')
    search_parser.add_argument('--source', '-s', help='Filter by source key')

    args = parser.parse_args()

    if args.command == 'upload':
        if args.pdf:
            items = parse_sor_pdf(args.pdf)
        elif args.json:
            items = parse_sor_json(args.json)
        else:
            print('Error: Provide either --pdf or --json')
            sys.exit(1)

        if not items:
            print('No items parsed. Aborting upload.')
            sys.exit(1)

        upload_to_supabase(items, args.key, args.name, args.desc)

    elif args.command == 'list':
        list_sources()

    elif args.command == 'search':
        search_items(args.query, args.source)

    else:
        parser.print_help()


if __name__ == '__main__':
    main()
