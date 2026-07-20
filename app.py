import pandas as pd
import sys
import re
from datetime import datetime

def clean_text(text):
    if pd.isna(text):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

def generate_zoho_invoice_csv(excel_path, zoho_items_csv_path, output_csv_path, default_customer="Walk-in Customer"):
    # 1. Load Zoho Items Master Map for exact matching
    zoho_items_df = pd.read_csv(zoho_items_csv_path)
    item_map = {}
    for _, row in zoho_items_df.iterrows():
        raw_name = row.get('Item Name')
        sku = row.get('SKU', '')
        rate = row.get('Rate', 0.0)
        if pd.notna(raw_name):
            item_map[clean_text(raw_name)] = {'exact_name': raw_name, 'sku': sku, 'rate': rate}

    # 2. Extract Metadata from Raw Excel Report
    raw_df = pd.read_excel(excel_path, header=None)
    date_range_str = str(raw_df.iloc[2, 1]) if len(raw_df) > 2 else "Unknown Date"
    branch_name = str(raw_df.iloc[3, 1]) if len(raw_df) > 3 else "Unknown Branch"

    # Attempt basic date parsing from string (e.g., "2026-07-01 to 2026-07-07")
    dates = re.findall(r'\d{4}-\d{2}-\d{2}', date_range_str)
    start_date = dates[0] if len(dates) > 0 else datetime.today().strftime('%Y-%m-%d')
    end_date = dates[1] if len(dates) > 1 else start_date

    # Unique invoice identifier so Zoho groups all these rows into ONE invoice
    timestamp = datetime.now().strftime('%Y%m%d%H%M')
    unique_invoice_num = f"INV-{clean_text(branch_name)[:5].upper()}-{timestamp}"

    # 3. Process Items Table from Excel
    data_df = pd.read_excel(excel_path, header=5)
    clean_df = data_df[['Product', 'Quantity']].dropna(subset=['Product', 'Quantity'])
    clean_df['Quantity'] = pd.to_numeric(clean_df['Quantity'], errors='coerce').fillna(0).astype(int)
    clean_df = clean_df[clean_df['Quantity'] > 0]

    # 4. Construct Final Zoho Invoice Layout Rows
    invoice_rows = []
    
    for _, row in clean_df.iterrows():
        excel_name = row['Product']
        quantity = row['Quantity']
        clean_excel_name = clean_text(excel_name)

        matched_name = excel_name
        matched_sku = ""
        matched_rate = 0.0

        # Run character-stripped match check
        if clean_excel_name in item_map:
            matched_name = item_map[clean_excel_name]['exact_name']
            matched_sku = item_map[clean_excel_name]['sku']
            matched_rate = item_map[clean_excel_name]['rate']
        else:
            # Fallback containment fuzzy logic
            for z_key, z_val in item_map.items():
                if clean_excel_name in z_key or z_key in clean_excel_name:
                    matched_name = z_val['exact_name']
                    matched_sku = z_val['sku']
                    matched_rate = z_val['rate']
                    break

        invoice_rows.append({
            'Invoice Date': start_date,
            'Invoice Number': unique_invoice_num,
            'Invoice Status': 'Draft',
            'Customer Name': default_customer,
            'Item Name': matched_name,
            'SKU': matched_sku,
            'Quantity': quantity,
            'Item Price': matched_rate,
            'cf_location': branch_name,
            'cf_start_date': start_date,
            'cf_end_date': end_date
        })

    # 5. Output cleanly formatted CSV
    output_df = pd.DataFrame(invoice_rows)
    output_df.to_csv(output_csv_path, index=False, encoding='utf-8')
    print(f"🎉 Pristine Zoho Invoice CSV created successfully with {len(output_df)} line items!")

if __name__ == "__main__":
    # Example execution usage: python invoice_generator.py report.xlsx zoho_items.csv ready_to_import.csv
    if len(sys.argv) >= 4:
        generate_zoho_invoice_csv(sys.argv[1], sys.argv[2], sys.argv[3])