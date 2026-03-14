import pdfplumber
import json
import re

def parse_irussor_2021(pdf_path, output_json_path):
    sor_items = []
    
    # We will try a basic heuristic: look for lines that start with a 6-digit number
    # (or 4/5 digits), followed by a description, a unit, and a rate.
    # IRUSSOR 2021 usually has columns like: Item No | Description | Unit | Rate
    item_pattern = re.compile(r'^(\d{4,6}[a-zA-Z]?)\s+(.*?)\s+(Cum|Sqm|100\sSqm|Km|Rm|Tonne|Ea|Kg|Pair|Rkm|Tkm|Sqcm|10\sCum|100\sCum|1000\sCum|Litre|Day|Month|L.S.|Quintal|Milli\sLitre|Sq.m.|Cu.m.|Sq.m|Cu.m)\s+([0-9,]+\.\d{2})', re.IGNORECASE)
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"Opened PDF: {pdf_path}. Pages: {len(pdf.pages)}")
            # For demonstration, we'll scan the first 50 pages to get started. 
            # In production, scan all pages, but it takes time.
            num_pages_to_scan = min(150, len(pdf.pages))
            
            for i in range(num_pages_to_scan):
                page = pdf.pages[i]
                text = page.extract_text()
                if not text:
                    continue
                
                # We can also use extract_table if the PDF has clear table geometry, but IRUSSOR is notoriously difficult.
                # Let's try parsing line by line first.
                lines = text.split('\n')
                current_item = None
                
                for line in lines:
                    match = item_pattern.search(line)
                    if match:
                        if current_item:
                            sor_items.append(current_item)
                        
                        code = match.group(1).strip()
                        desc = match.group(2).strip()
                        unit = match.group(3).strip()
                        rate_str = match.group(4).replace(',', '').strip()
                        rate = float(rate_str)
                        
                        current_item = {
                            "code": code,
                            "description": desc,
                            "unit": unit,
                            "rate": rate,
                            "source": "irussor_2021"
                        }
                    elif current_item and not re.match(r'^\d', line.strip()):
                        # Append to description if it's a continuation line (no item number at start)
                        current_item["description"] += " " + line.strip()
                        
                if current_item:
                    sor_items.append(current_item)
                    
            print(f"Extracted {len(sor_items)} items from the first {num_pages_to_scan} pages.")
            
    except Exception as e:
        print(f"Error parsing PDF: {e}")
        return

    if len(sor_items) > 0:
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(sor_items, f, indent=2, ensure_ascii=False)
        print(f"Saved to {output_json_path}")
    else:
        print("No items found using the heuristic pattern. The PDF structure might be complex and require table extraction.")

if __name__ == '__main__':
    parse_irussor_2021(
        "/Users/moyukhroy/Downloads/0b639519-9e3c-456c-8389-094da67a99ac.pdf", 
        "/Users/moyukhroy/Desktop/My Software/Earthsoft/Earthsoft anti/public/workspace/data/irussor_2021.json"
    )
