import streamlit as st
import pandas as pd
import re
import io
import os
from datetime import datetime

st.set_page_config(page_title="Zoho Invoice CSV Generator", layout="centered")

st.title("👨‍🍳 Menu Engineering to Zoho Invoice Converter")
st.write("Upload your menu Excel report to instantly download a formatted Zoho Inventory Import CSV.")

menu_file = st.file_uploader("Upload Menu Excel Report (.xlsx)", type=["xlsx"])

def clean_text(text):
    if pd.isna(text):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

item_map = {}
if os.path.exists("zoho_items.csv"):
    try:
        zoho_items_df = pd.read_csv("zoho_items.csv")
        for _, row in zoho_items_df.iterrows():
            raw_name = row.get('Item Name')
            sku = row.get('SKU', '')
            rate = row.get('Rate', 0.0)
            if pd.notna(raw_name):
                item_map[clean_text(raw_name)] = {'exact_name': raw_name, 'sku': sku, 'rate': rate}
    except Exception as e:
        st.warning(f"Note: Found zoho_items.csv but couldn't parse it. Error: {e}")

if menu_file:
    try:
        with st.spinner("Processing report..."):
            raw_df = pd.read_excel(menu_file, header=None)
            date_range_str = str(raw_df.iloc[2, 1]) if len(raw_df) > 2 else "Unknown Date"
            branch_name = str(raw_df.iloc[3, 1]) if len(raw_df) > 3 else "Unknown Branch"

            dates = re.findall(r'\d{4}-\d{2}-\d{2}', date_range_str)
            start_date = dates[0] if len(dates) > 0 else datetime.today().strftime('%Y-%m-%d')
            end_date = dates[1] if len(dates) > 1 else start_date

            timestamp = datetime.now().strftime('%Y%m%d%H%M')
            unique_invoice_num = f"INV-{clean_text(branch_name)[:5].upper()}-{timestamp}"

            data_df = pd.read_excel(menu_file, header=5)
            clean_df = data_df[['Product', 'Quantity']].dropna(subset=['Product', 'Quantity'])
            clean_df['Quantity'] = pd.to_numeric(clean_df['Quantity'], errors='coerce').fillna(0).astype(int)
            clean_df = clean_df[clean_df['Quantity'] > 0]

            invoice_rows = []
            for _, row in clean_df.iterrows():
                excel_name = row['Product']
                quantity = row['Quantity']
                clean_excel_name = clean_text(excel_name)

                matched_name = excel_name
                matched_sku = ""
                matched_rate = 0.0

                if item_map:
                    if clean_excel_name in item_map:
                        matched_name = item_map[clean_excel_name]['exact_name']
                        matched_sku = item_map[clean_excel_name]['sku']
                        matched_rate = item_map[clean_excel_name]['rate']
                    else:
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
                    'Customer Name': 'Walk-in Customer',
                    'Item Name': matched_name,
                    'SKU': matched_sku,
                    'Quantity': quantity,
                    'Item Price': matched_rate,
                    'cf_location': branch_name,
                    'cf_start_date': start_date,
                    'cf_end_date': end_date
                })

            output_df = pd.DataFrame(invoice_rows)
            
            csv_buffer = io.StringIO()
            output_df.to_csv(csv_buffer, index=False, encoding='utf-8')
            csv_data = csv_buffer.getvalue()
            
            safe_filename = clean_text(branch_name)[:15]

        st.success(f"🎉 Processed {len(output_df)} items for {branch_name} successfully!")
        st.dataframe(output_df.head(5), use_container_width=True)
        
        st.download_button(
            label="📥 Download Zoho Invoice CSV",
            data=csv_data,
            file_name=f"zoho_invoice_{safe_filename}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
