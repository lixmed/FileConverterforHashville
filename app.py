import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import difflib

st.set_page_config(page_title="Zoho Invoice CSV Generator", layout="centered")

st.title("👨‍🍳 Menu Engineering to Zoho Invoice Converter")
st.write("Upload your menu Excel report. The matching engine will force-map every item to your Zoho master list to ensure correct pricing.")

ZOHO_MASTER = {
    "preparing pesto sauce": {"exact_name": "preparing pesto SAUCE", "rate": 113.75, "sku": "110"},
    "preparing travel mushroom": {"exact_name": "preparing travel mashroom", "rate": 135.00, "sku": "111"},
    "preparing travel mashroom": {"exact_name": "preparing travel mashroom", "rate": 135.00, "sku": "111"},
    "preparing coleslaw": {"exact_name": "preparing cloeslaw", "rate": 58.50, "sku": "149-COMP"},
    "preparing cloeslaw": {"exact_name": "preparing cloeslaw", "rate": 58.50, "sku": "149-COMP"},
    "preparing flour دقيق متبل جاهز": {"exact_name": "preparing flour دقيق متبل جاهز", "rate": 23.49, "sku": "121"},
    "استربس مجهز": {"exact_name": "استربس مجهز", "rate": 26.53, "sku": "145"},
    "بانيه مجهز": {"exact_name": "بانيه مجهز", "rate": 53.05, "sku": "146"},
    "preparing honey butter": {"exact_name": "preparing honey butter", "rate": 125.93, "sku": "143"},
    "hashville oil sauce": {"exact_name": "HASHVILLE oil SAUCE", "rate": 115.98, "sku": "114"},
    "دجاجة بالعضم": {"exact_name": "دجاجة بالعضم", "rate": 77.00, "sku": "249"},
    "coleslaw": {"exact_name": "coleslaw", "rate": 40.00, "sku": "SIDE-05"},
    "creamy onion sandwiches": {"exact_name": "creamy onyion sandwiches", "rate": 235.00, "sku": ""},
    "creamy onyion sandwiches": {"exact_name": "creamy onyion sandwiches", "rate": 235.00, "sku": ""},
    "sriracha honey": {"exact_name": "sriracha honey", "rate": 210.00, "sku": ""},
    "chikchen slaw": {"exact_name": "chikchen slaw", "rate": 210.00, "sku": ""},
    "chicken slaw": {"exact_name": "chikchen slaw", "rate": 210.00, "sku": ""},
    "burnout butter": {"exact_name": "burnout butter", "rate": 210.00, "sku": ""},
    "truffle chicken": {"exact_name": "Truffel chicken", "rate": 210.00, "sku": ""},
    "truffel chicken": {"exact_name": "Truffel chicken", "rate": 210.00, "sku": ""},
    "the hashville": {"exact_name": "the hashville", "rate": 210.00, "sku": ""},
    "mangoana": {"exact_name": "mangoana", "rate": 210.00, "sku": ""},
    "burrata chicken": {"exact_name": "burrata chicken", "rate": 270.00, "sku": ""},
    "crispy cheese sandwiches": {"exact_name": "crispy cheese sandwiches", "rate": 265.00, "sku": ""},
    "sweetie nashville sandwiches": {"exact_name": "sweetie nashville sandwiches", "rate": 240.00, "sku": ""},
    "shrimp sandwiche": {"exact_name": "shrimp sandwiche", "rate": 230.00, "sku": ""},
    "shrimp sandwich": {"exact_name": "shrimp sandwiche", "rate": 230.00, "sku": ""},
    "sugar on my tongue": {"exact_name": "Sugar on my tongue", "rate": 230.00, "sku": ""},
    "creamy onion 3 pieces": {"exact_name": "creamy onyion 3 pieces", "rate": 260.00, "sku": "MEAL-01"},
    "creamy onyion 3 pieces": {"exact_name": "creamy onyion 3 pieces", "rate": 260.00, "sku": "MEAL-01"},
    "فراخ بلعضم 4قطع": {"exact_name": "فراخ بلعضم 4قطع", "rate": 280.00, "sku": "MEAL-02"},
    "crispy cheese 3 pieces": {"exact_name": "crispy cheese 3 pieces", "rate": 370.00, "sku": "MEAL-03"},
    "hashvill 3 pieces": {"exact_name": "hashvill 3 pieces", "rate": 340.00, "sku": "MEAL-04"},
    "hashville 3 pieces": {"exact_name": "hashvill 3 pieces", "rate": 340.00, "sku": "MEAL-04"},
    "hashvill 2 pieces": {"exact_name": "hashvill 2 pieces", "rate": 265.00, "sku": "MEAL-05"},
    "hashville 2 pieces": {"exact_name": "hashvill 2 pieces", "rate": 265.00, "sku": "MEAL-05"},
    "hashvill 5 pieces": {"exact_name": "hashvill 5 pieces", "rate": 480.00, "sku": "MEAL-06"},
    "hashville 5 pieces": {"exact_name": "hashvill 5 pieces", "rate": 480.00, "sku": "MEAL-06"},
    "mangoana burata": {"exact_name": "mangoana burata", "rate": 370.00, "sku": "MEAL-07"},
    "mangoana burrata": {"exact_name": "mangoana burata", "rate": 370.00, "sku": "MEAL-07"},
    "honey spicy burata": {"exact_name": "honey spicy burata", "rate": 370.00, "sku": "MEAL-08"},
    "honey spicy burrata": {"exact_name": "honey spicy burata", "rate": 370.00, "sku": "MEAL-08"},
    "chicken pieces 4": {"exact_name": "chicken piecec 4", "rate": 240.00, "sku": "MEAL-09"},
    "chicken piecec 4": {"exact_name": "chicken piecec 4", "rate": 240.00, "sku": "MEAL-09"},
    "shrimp meal 10psc": {"exact_name": "shrimp meal 10psc", "rate": 390.00, "sku": "MEAL-10"},
    "shrimp meal 10pcs": {"exact_name": "shrimp meal 10psc", "rate": 390.00, "sku": "MEAL-10"},
    "lemon del 3psc": {"exact_name": "lemon del 3psc", "rate": 350.00, "sku": "MEAL-11"},
    "lemon del 3pcs": {"exact_name": "lemon del 3psc", "rate": 350.00, "sku": "MEAL-11"},
    "shrimp meal 7psc": {"exact_name": "shrimp meal 7psc", "rate": 360.00, "sku": "MEAL-12"},
    "shrimp meal 7pcs": {"exact_name": "shrimp meal 7psc", "rate": 360.00, "sku": "MEAL-12"},
    "truffle nashville chicken": {"exact_name": "truffle nashville chicken", "rate": 155.00, "sku": "SIDE-01"},
    "cheese hashville fries": {"exact_name": "cheese hashville fries", "rate": 155.00, "sku": "SIDE-02"},
    "combo": {"exact_name": "combo", "rate": 65.00, "sku": "SIDE-03"},
    "french fries": {"exact_name": "franch fries", "rate": 35.00, "sku": "SIDE-04"},
    "franch fries": {"exact_name": "franch fries", "rate": 35.00, "sku": "SIDE-04"},
    "soft drink": {"exact_name": "soft drink", "rate": 35.00, "sku": "SIDE-07"},
    "sriracha honey side": {"exact_name": "sriracha honey side", "rate": 30.00, "sku": "SIDE-08"},
    "honey spicy": {"exact_name": "honey spicy", "rate": 30.00, "sku": "SIDE-09"},
    "nashville cheese": {"exact_name": "nashvile cheese", "rate": 30.00, "sku": "SIDE-10"},
    "nashvile cheese": {"exact_name": "nashvile cheese", "rate": 30.00, "sku": "SIDE-10"},
    "smoke honey mustard": {"exact_name": "smoke honey mustard", "rate": 30.00, "sku": "SIDE-11"},
    "hashville sauce": {"exact_name": "hashville sauce", "rate": 30.00, "sku": "SIDE-12"},
    "largest coleslaw sauce": {"exact_name": "largest coleslaw sauce", "rate": 140.00, "sku": "SIDE-13"},
    "largest sriracha sauce": {"exact_name": "largest sriracha sauce", "rate": 140.00, "sku": "SIDE-14"},
    "largest honey spicy sauce": {"exact_name": "largest honey spicy sauce", "rate": 140.00, "sku": "SIDE-15"},
    "largest nashville cheese sauce": {"exact_name": "largest nashvile cheese sauce", "rate": 140.00, "sku": "SIDE-16"},
    "largest nashvile cheese sauce": {"exact_name": "largest nashvile cheese sauce", "rate": 140.00, "sku": "SIDE-16"},
    "largest smoke honey mustard sauce": {"exact_name": "largest smoke honey mustard sauce", "rate": 140.00, "sku": "SIDE-17"},
    "largest hashville sauce": {"exact_name": "largest hashville sauce", "rate": 140.00, "sku": "SIDE-18"}
}

menu_file = st.file_uploader("Upload Menu Excel Report (.xlsx)", type=["xlsx"])

def clean_lookup_key(text):
    if pd.isna(text):
        return ""
    t = str(text).strip().lower()
    t = re.sub(r'\s+[tT]$', '', t)
    t = re.sub(r'\s+[eE][aA][cC][hH]$', '', t)
    return re.sub(r'[^a-zA-Z0-9\u0600-\u06FF]', '', t)

clean_master = {clean_lookup_key(k): v for k, v in ZOHO_MASTER.items()}
master_keys = list(clean_master.keys())

if menu_file:
    try:
        with st.spinner("Processing report and force-matching items..."):
            raw_df = pd.read_excel(menu_file, header=None)
            date_range_str = str(raw_df.iloc[2, 1]) if len(raw_df) > 2 else "Unknown Date"
            branch_name = str(raw_df.iloc[3, 1]) if len(raw_df) > 3 else "Unknown Branch"

            dates = re.findall(r'\d{4}-\d{2}-\d{2}', date_range_str)
            start_date = dates[0] if len(dates) > 0 else datetime.today().strftime('%Y-%m-%d')
            end_date = dates[1] if len(dates) > 1 else start_date

            timestamp = datetime.now().strftime('%Y%m%d%H%M')
            unique_invoice_num = f"INV-{re.sub(r'[^a-zA-Z0-9]', '', branch_name)[:5].upper()}-{timestamp}"

            data_df = pd.read_excel(menu_file, header=5)
            clean_df = data_df[['Product', 'Quantity']].dropna(subset=['Product', 'Quantity'])
            clean_df['Quantity'] = pd.to_numeric(clean_df['Quantity'], errors='coerce').fillna(0).astype(int)
            clean_df = clean_df[clean_df['Quantity'] > 0]

            invoice_rows = []
            
            for _, row in clean_df.iterrows():
                excel_name = str(row['Product']).strip()
                quantity = row['Quantity']
                clean_excel_key = clean_lookup_key(excel_name)

                # 1. Direct match check
                if clean_excel_key in clean_master:
                    target_item = clean_master[clean_excel_key]
                else:
                    # 2. Containment check fallback
                    found_match = None
                    for m_key in master_keys:
                        if clean_excel_key in m_key or m_key in clean_excel_key:
                            found_match = clean_master[m_key]
                            break
                    
                    if found_match:
                        target_item = found_match
                    else:
                        # 3. Maximum-Probability Match Fallback (Guarantees every line gets a product mapping)
                        best_key = None
                        best_score = -1.0
                        for m_key in master_keys:
                            score = difflib.SequenceMatcher(None, clean_excel_key, m_key).ratio()
                            if score > best_score:
                                best_score = score
                                best_key = m_key
                        target_item = clean_master[best_key]

                invoice_rows.append({
                    'Invoice Date': start_date,
                    'Invoice Number': unique_invoice_num,
                    'Invoice Status': 'Draft',
                    'Customer Name': 'Walk-in Customer',
                    'Item Name': target_item['exact_name'],
                    'SKU': target_item['sku'],
                    'Quantity': quantity,
                    'Item Price': target_item['rate'],
                    'cf_location': branch_name,
                    'cf_start_date': start_date,
                    'cf_end_date': end_date
                })

            output_df = pd.DataFrame(invoice_rows)
            
            csv_buffer = io.StringIO()
            output_df.to_csv(csv_buffer, index=False, encoding='utf-8')
            csv_data = csv_buffer.getvalue()
            
            safe_filename = re.sub(r'[^a-zA-Z0-9]', '', branch_name)[:15].lower()

        st.success(f"🎉 Processed {len(output_df)} items for {branch_name} successfully!")
        st.dataframe(output_df, use_container_width=True)
        
        st.download_button(
            label="📥 Download Zoho Invoice CSV",
            data=csv_data,
            file_name=f"zoho_invoice_{safe_filename}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
