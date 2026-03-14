
import json
import re

def parse_nanonets_sor(json_path):
    with open(json_path, 'r') as f:
        data = json.load(f)
    
    items = []
    current_item = None
    
    # Common units to help identify unit blocks
    units = {"Cum", "Sqm", "100 Sqm", "Km", "RM", "Rm", "Tonne", "Kg", "Pair", "Rkm", "Tkm", "Sqcm", "Litre", "Day", "Month", "L.S.", "Quintal", "Sq.m.", "Cu.m.", "KM", "MT", "Each", "Set"}

    for page in data.get('pages', []):
        content = page.get('content', [])
        page_id = page.get('page_id')
        
        # Skip index pages or intro pages (approximate)
        if page_id < 8: continue
        
        for block in content:
            text = block.get('text', '').strip()
            block_type = block.get('type')
            
            if block_type in ['header', 'page_number', 'table_title']:
                continue
                
            # Check for 6-digit code (e.g., 011010)
            if re.match(r'^\d{6}$', text):
                if current_item:
                    items.append(current_item)
                current_item = {
                    'code': text,
                    'description': '',
                    'unit': '',
                    'rate': 0.0,
                    'stage': 'desc' # stage can be 'desc', 'unit', 'rate'
                }
            elif current_item:
                # If stage is desc and we found a unit, move to unit stage
                if current_item['stage'] == 'desc':
                    if text in units or re.match(r'^\d+\s+\w+', text): # e.g. "100 Sqm"
                        current_item['unit'] = text
                        current_item['stage'] = 'unit'
                    else:
                        current_item['description'] = (current_item['description'] + ' ' + text).strip()
                
                elif current_item['stage'] == 'unit':
                    # Try to parse rate
                    rate_text = text.replace(',', '')
                    try:
                        current_item['rate'] = float(rate_text)
                        current_item['stage'] = 'done'
                    except ValueError:
                        # Maybe still part of unit? or description?
                        # If it's not a number, and current unit is small, maybe it's still unit
                        if not current_item['unit']:
                            current_item['unit'] = text
                        else:
                            # Edge case fallback
                            pass
                            
    if current_item:
        items.append(current_item)
        
    # Final cleanup
    final_items = []
    for item in items:
        if item['code'] and item['description']:
            final_items.append({
                'code': item['code'],
                'description': item['description'],
                'unit': item['unit'],
                'rate': item['rate']
            })
            
    return final_items

if __name__ == "__main__":
    items = parse_nanonets_sor('/Users/moyukhroy/Downloads/0b639519-9e3c-456c-8389-094da67a99ac-1-50.json')
    print(json.dumps(items[:10], indent=2))
    print(f"Total items extracted: {len(items)}")
    
    with open('/tmp/parsed_irussor_2021.json', 'w') as f:
        json.dump(items, f, indent=2)
