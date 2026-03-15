#!/usr/bin/env python3
import os
import sys
import re
import json
import uuid
from datetime import datetime

# ==============================================================================
# Earthsoft LAR Processing Engine
# Extracts LAR (Lowest Accepted Rates) from PDFs and syncs to Supabase.
# ==============================================================================

try:
    from supabase import create_client, Client
except ImportError:
    print("Error: 'supabase' Python package not found.")
    print("Install it with: pip3 install supabase")
    sys.exit(1)

# Default Supabase configuration (overridden by env vars)
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

def get_supabase_client():
    if not SUPABASE_URL or not SUPABASE_KEY:
        print("Error: SUPABASE_URL and SUPABASE_KEY environment variables must be set.")
        sys.exit(1)
    return create_client(SUPABASE_URL, SUPABASE_KEY)

# ── PDF Parser ──────────────────────────────────────────────────────────────
def parse_lar_pdf(pdf_path):
    """Parse an LAR PDF and return list of items."""
    try:
        import pdfplumber
    except ImportError:
        print('Error: pdfplumber is required for PDF parsing.')
        sys.exit(1)

    items = []
    # Pattern: Code Description Unit Rate
    # Note: LARs often have extra columns for contractors or contract refs.
    # This pattern is optimized for common Railway LAR formats.
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
            if not text: continue

            for line in text.split('\n'):
                match = item_pattern.search(line)
                if match:
                    if current_item: items.append(current_item)
                    current_item = {
                        'item_code': match.group(1).strip(),
                        'description': match.group(2).strip(),
                        'unit': match.group(3).strip(),
                        'rate': float(match.group(4).replace(',', '')),
                        'contract_ref': os.path.basename(pdf_path)
                    }
                elif current_item and not re.match(r'^\d', line.strip()):
                    current_item['description'] += ' ' + line.strip()

        if current_item: items.append(current_item)

    print(f'  Extracted {len(items)} items from PDF')
    return items

# ── Supabase Sync ────────────────────────────────────────────────────────────
def sync_to_supabase(items, display_name):
    supabase = get_supabase_client()
    source_key = display_name.lower().replace(' ', '_').replace('-', '_')
    
    print(f'Syncing source: {display_name}...')
    
    # 1. Upsert LAR source
    source_data = {
        "source_key": source_key,
        "display_name": display_name,
        "item_count": len(items)
    }
    
    res = supabase.table("lar_sources").upsert(source_data, on_conflict="source_key").execute()
    if not res.data:
        print("Error creating source.")
        return
    
    source_id = res.data[0]['id']
    print(f'Source ID: {source_id}')
    
    # 2. Clear existing items
    print('Clearing existing items for this source...')
    supabase.table("lar_items").delete().eq("source_id", source_id).execute()
    
    # 3. Batch insert items
    batch_size = 100
    total_items = len(items)
    print(f'Inserting {total_items} items in batches of {batch_size}...')
    
    for i in range(0, total_items, batch_size):
        batch = items[i:i + batch_size]
        payload = []
        for item in batch:
            payload.append({
                "source_id": source_id,
                "item_code": item['item_code'],
                "description": item['description'],
                "unit": item['unit'],
                "rate": item['rate'],
                "contract_ref": item.get('contract_ref', '')
            })
        
        supabase.table("lar_items").insert(payload).execute()
        print(f'  Uploaded {min(i + batch_size, total_items)}/{total_items}...')

    print('Upload Complete!')

# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print("Usage: python3 lar_engine.py <pdf_path> [display_name]")
        sys.exit(1)

    path = sys.argv[1]
    name = sys.argv[2] if len(sys.argv) > 2 else os.path.basename(path).replace('.pdf','').upper()

    if not os.path.exists(path):
        print(f"Error: Path {path} does not exist.")
        sys.exit(1)

    items = parse_lar_pdf(path)
    if items:
        sync_to_supabase(items, name)

if __name__ == "__main__":
    main()
