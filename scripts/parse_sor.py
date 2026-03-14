import pandas as pd
import json
import argparse
import os

def parse_sor(input_file, output_file, source_name):
    """
    Parses an Excel representation of an SOR (like IRUSSOR or DSR) and saves to JSON.
    Assumes standard columns: Item Code, Description, Unit, Rate.
    """
    try:
        df = pd.read_excel(input_file)
        
        # Lowercase all headers to make it easier to map
        df.columns = [str(c).lower().strip() for c in df.columns]
        
        # Finding potential column mappings
        col_code = next((c for c in df.columns if 'no' in c or 'code' in c or 'item' in c), None)
        col_desc = next((c for c in df.columns if 'desc' in c or 'particular' in c), None)
        col_unit = next((c for c in df.columns if 'unit' in c), None)
        col_rate = next((c for c in df.columns if 'rate' in c or 'amount' in c), None)

        if not all([col_code, col_desc, col_unit, col_rate]):
            print("Could not automatically detect standard headers: [Item Code, Description, Unit, Rate]")
            print(f"Found headers: {df.columns.tolist()}")
            return
        
        # Keep relevant columns and rename them cleanly
        df = df[[col_code, col_desc, col_unit, col_rate]]
        df.columns = ['code', 'description', 'unit', 'rate']
        
        # Clean data
        df = df.dropna(subset=['code', 'rate'])
        df['code'] = df['code'].astype(str).str.strip()
        df['description'] = df['description'].astype(str).str.strip()
        df['unit'] = df['unit'].astype(str).str.strip()
        # Ensure rate is float
        df['rate'] = pd.to_numeric(df['rate'], errors='coerce')
        df = df.dropna(subset=['rate'])
        df['source'] = source_name

        data = df.to_dict(orient='records')
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            
        print(f"Successfully processed {len(data)} items and saved to {output_file}")
        
    except Exception as e:
        print(f"Error parsing SOR file: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert SOR Excel to Earthsoft JSON Format")
    parser.add_argument("input", help="Path to input Excel file (e.g., IRUSSOR_2021.xlsx)")
    parser.add_argument("output", help="Path to output JSON file (e.g., sor_db.json)")
    parser.add_argument("--source", default="irussor_2021", help="Name of the SOR source")
    args = parser.parse_args()
    
    parse_sor(args.input, args.output, args.source)
