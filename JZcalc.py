import streamlit as st
import pandas as pd
import numpy as np
import base64
from io import BytesIO
from PIL import Image
import os

# Set page config
st.set_page_config(
    page_title="Margin Calculator",
    page_icon="🧮",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling and larger fonts
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stButton button {
        width: 100%;
        font-size: 1.2rem;
    }
    .metric-container {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 15px;
    }
    .metric-label {
        font-weight: bold;
        font-size: 1.1rem;
    }
    .metric-value {
        font-size: 1.3rem;
        padding: 8px 0;
    }
    .header {
        text-align: center;
        margin-bottom: 30px;
        font-size: 2.8rem;
    }
    .logo-container {
        text-align: center;
        margin: 0 auto;
        display: block;
    }
    .quadrant {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
    }
    .quadrant-title {
        font-weight: bold;
        font-size: 1.4rem;
        margin-bottom: 15px;
        text-align: center;
    }
    [data-testid="stImage"] {
        display: block;
        margin-left: auto;
        margin-right: auto;
        text-align: center;
    }
    .stTextInput input, .stNumberInput input, .stSelectbox > div > div[role="listbox"] {
        font-size: 1.4rem !important;
    }
    .stTextInput label, .stNumberInput label, .stSelectbox label {
        font-size: 1.5rem !important;
        font-weight: 600 !important;
    }
    .stTab [data-baseweb="tab"] {
        font-size: 1.2rem !important;
    }
    .stTab [data-baseweb="tab-list"] {
        gap: 1rem;
    }
    .stSubheader {
        font-size: 1.5rem !important;
    }
    p, div {
        font-size: 1.2rem;
    }
    .sidebar .sidebar-content {
        padding: 2rem 1rem;
    }
    button[data-baseweb="tab"] {
        font-size: 1.3rem !important;
    }
    .pricing-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.4rem;
        margin-bottom: 20px;
    }
    .pricing-table th {
        background-color: #f0f2f6;
        padding: 12px;
        text-align: center;
        font-weight: bold;
        border: 1px solid #ddd;
    }
    .pricing-table td {
        padding: 12px;
        border: 1px solid #ddd;
        text-align: center;
    }
    .pricing-table td:first-child {
        text-align: left;
    }
    .pricing-table tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    .pricing-table tr:hover {
        background-color: #e9ecef;
    }
    .delete-button {
        background-color: #ff6b6b;
        color: white;
        border: none;
        padding: 8px 12px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 1.1rem;
        margin: 5px;
    }
    .delete-button:hover {
        background-color: #e74c3c;
    }
    .everyday-row {
        background-color: #d4edda !important;
    }
    .editable-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.2rem;
    }
    .editable-table th {
        background-color: #f0f2f6;
        padding: 10px;
        text-align: center;
        font-weight: bold;
        border: 1px solid #ddd;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    .editable-table td {
        padding: 8px;
        border: 1px solid #ddd;
        text-align: center;
    }
    .editable-table tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    .editable-table tr:hover {
        background-color: #e9ecef;
    }
    html {
        zoom: 67%;
    }
</style>

<script>
    document.addEventListener('DOMContentLoaded', function() {
        document.body.style.zoom = "67%";
    });
</script>
""", unsafe_allow_html=True)

# Initialize session state for storing scan scenarios
if 'scan_scenarios' not in st.session_state:
    st.session_state.scan_scenarios = pd.DataFrame()
    
# For managing scenario deletion
if 'delete_scenario_indices' not in st.session_state:
    st.session_state.delete_scenario_indices = []

# For tracking edited values
if 'edited_cells' not in st.session_state:
    st.session_state.edited_cells = {}

# Add logo and title to the main area
try:
    image = Image.open('image.png')
    col1, col2, col3 = st.columns([1, 3, 1])
    with col2:
        st.image(image, width=600, use_container_width=True)
except:
    st.warning("Logo 'image.png' not found. Please add it to the same directory as the app.")

st.markdown("<h1 class='header'>Pricing & Margin Calculator</h1>", unsafe_allow_html=True)

# Create sidebar for inputs
with st.sidebar:
    # Brand and Size inputs
    brand = st.text_input("Brand")
    size_options = ["1.75L", "1L", "750mL", "375mL", "355mL", "200mL", "100mL", "50mL"]
    size = st.selectbox("Size", size_options)
    
    # Cost inputs
    case_cost = st.number_input("Case Cost ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    bottles_per_case = st.number_input("# Bottles per Case", min_value=1, value=12, step=1)
    
    # Calculate bottle cost automatically
    if bottles_per_case > 0:
        bottle_cost = case_cost / bottles_per_case
    else:
        bottle_cost = 0.0
    
    st.text(f"Bottle Cost: ${bottle_cost:.2f}")
    
    # Scan and coupon inputs
    base_scan = st.number_input("Base Scan ($)", min_value=0.0, value=0.0, step=0.25, format="%.2f")
    deep_scan = st.number_input("Deep Scan ($)", min_value=0.0, value=0.0, step=0.25, format="%.2f")
    coupon = st.number_input("Coupon ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    
    # Pricing inputs
    edlp_price = st.number_input("Everyday Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    tpr_base_price = st.number_input("TPR Base Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    tpr_deep_price = st.number_input("TPR Deep Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    ad_base_price = st.number_input("Ad/Feature Base Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    ad_deep_price = st.number_input("Ad/Feature Deep Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    
    # Optional customer/state input
    customer_state = st.text_input("Customer / State", "")

# Helper function to calculate margins (similar to BuildMarginString in VBA)
def calculate_margin(price, cost, scan, coupon):
    if price > 0 and cost > 0:
        gm = price - (cost - scan)
        gm_per = gm / price
        
        gm_coupon = price - (cost - scan - coupon)
        gm_coupon_per = gm_coupon / price
        
        return {
            "gm_percent": gm_per * 100,
            "gm_dollars": gm,
            "gm_coupon_percent": gm_coupon_per * 100,
            "gm_coupon_dollars": gm_coupon
        }
    else:
        return {
            "gm_percent": 0,
            "gm_dollars": 0,
            "gm_coupon_percent": 0,
            "gm_coupon_dollars": 0
        }

# Calculate margins for Everyday Price
edlp_margins = calculate_margin(edlp_price, bottle_cost, 0, coupon)

# Calculate margins for other pricing scenarios
tpr_base_margins = calculate_margin(tpr_base_price, bottle_cost, base_scan, coupon)
tpr_deep_margins = calculate_margin(tpr_deep_price, bottle_cost, deep_scan, coupon)
ad_base_margins = calculate_margin(ad_base_price, bottle_cost, base_scan, coupon)
ad_deep_margins = calculate_margin(ad_deep_price, bottle_cost, deep_scan, coupon)

# Create data for the pricing comparison table (including Everyday Price)
pricing_data = {
    "Pricing Scenario": ["Everyday Price", "TPR (Base Scan)", "TPR (Deep Scan)", "Ad/Feature (Base Scan)", "Ad/Feature (Deep Scan)"],
    "Price": [f"${edlp_price:.2f}", f"${tpr_base_price:.2f}", f"${tpr_deep_price:.2f}", f"${ad_base_price:.2f}", f"${ad_deep_price:.2f}"],
    "Gross Margin %": [f"{edlp_margins['gm_percent']:.1f}%", f"{tpr_base_margins['gm_percent']:.1f}%", 
                      f"{tpr_deep_margins['gm_percent']:.1f}%", f"{ad_base_margins['gm_percent']:.1f}%", 
                      f"{ad_deep_margins['gm_percent']:.1f}%"],
    "Gross Margin $": [f"${edlp_margins['gm_dollars']:.2f}", f"${tpr_base_margins['gm_dollars']:.2f}", 
                      f"${tpr_deep_margins['gm_dollars']:.2f}", f"${ad_base_margins['gm_dollars']:.2f}", 
                      f"${ad_deep_margins['gm_dollars']:.2f}"],
    "With Coupon %": [f"{edlp_margins['gm_coupon_percent']:.1f}%", f"{tpr_base_margins['gm_coupon_percent']:.1f}%", 
                     f"{tpr_deep_margins['gm_coupon_percent']:.1f}%", f"{ad_base_margins['gm_coupon_percent']:.1f}%", 
                     f"{ad_deep_margins['gm_coupon_percent']:.1f}%"],
    "With Coupon $": [f"${edlp_margins['gm_coupon_dollars']:.2f}", f"${tpr_base_margins['gm_coupon_dollars']:.2f}", 
                     f"${tpr_deep_margins['gm_coupon_dollars']:.2f}", f"${ad_base_margins['gm_coupon_dollars']:.2f}", 
                     f"${ad_deep_margins['gm_coupon_dollars']:.2f}"]
}

# Create the HTML table
table_html = "<table class='pricing-table'><thead><tr>"
# Add headers
for col in pricing_data.keys():
    table_html += f"<th>{col}</th>"
table_html += "</tr></thead><tbody>"

# Add rows
for i in range(len(pricing_data["Pricing Scenario"])):
    # Add special highlighting class for Everyday Price row
    if i == 0:  # Everyday Price is the first row
        table_html += "<tr class='everyday-row'>"
    else:
        table_html += "<tr>"
    
    for col in pricing_data.keys():
        table_html += f"<td>{pricing_data[col][i]}</td>"
    table_html += "</tr>"
table_html += "</tbody></table>"

# Display the table
st.markdown(table_html, unsafe_allow_html=True)

# Add Save Scan Scenario button below the main table
if st.button("Save Scan Scenario"):
    # Create a row with all the data in the flat format
    scenario_data = {
        "Brand": brand,
        "Size": size,
        "Case Cost": case_cost,
        "# Bottles/Cs": bottles_per_case,
        "Bottle Cost": bottle_cost,
        "Base Scan": base_scan,
        "Deep Scan": deep_scan,
        "Coupon": coupon,
        "Customer/State": customer_state,
        "Everyday Shelf Price": edlp_price,
        "Everyday GM %": edlp_margins["gm_percent"],
        "Everyday GM $": edlp_margins["gm_dollars"],
        "Everyday GM % (With Coupon)": edlp_margins["gm_coupon_percent"],
        "TPR Price (Base Scan)": tpr_base_price,
        "TPR GM % (Base Scan)": tpr_base_margins["gm_percent"],
        "TPR GM $ (Base Scan)": tpr_base_margins["gm_dollars"],
        "TPR GM % (With Coupon)": tpr_base_margins["gm_coupon_percent"],
        "TPR Price (Deep Scan)": tpr_deep_price,
        "TPR GM % (Deep Scan)": tpr_deep_margins["gm_percent"],
        "TPR GM $ (Deep Scan)": tpr_deep_margins["gm_dollars"],
        "TPR GM % (With Coupon)": tpr_deep_margins["gm_coupon_percent"],
        "Ad/Feature Price (Base Scan)": ad_base_price,
        "Ad GM % (Base Scan)": ad_base_margins["gm_percent"],
        "Ad GM $ (Base Scan)": ad_base_margins["gm_dollars"],
        "Ad GM % (With Coupon)": ad_base_margins["gm_coupon_percent"],
        "Ad/Feature Price (Deep Scan)": ad_deep_price,
        "Ad GM % (Deep Scan)": ad_deep_margins["gm_percent"],
        "Ad GM $ (Deep Scan)": ad_deep_margins["gm_dollars"],
        "Ad GM % (With Coupon)": ad_deep_margins["gm_coupon_percent"]
    }
    
    # Create a DataFrame with the scenario data
    new_scenario = pd.DataFrame([scenario_data])
    
    # Append to scan scenarios
    if st.session_state.scan_scenarios.empty:
        st.session_state.scan_scenarios = new_scenario
    else:
        st.session_state.scan_scenarios = pd.concat([st.session_state.scan_scenarios, new_scenario], ignore_index=True)
    
    st.success("Scan scenario saved successfully!")

# Helper function to recalculate margins in the scenarios table
def recalculate_margins(df, row_idx):
    # Get relevant values
    bottle_cost = df.at[row_idx, 'Bottle Cost']
    base_scan = df.at[row_idx, 'Base Scan']
    deep_scan = df.at[row_idx, 'Deep Scan']
    coupon = df.at[row_idx, 'Coupon']
    
    # Everyday price
    edlp_price = df.at[row_idx, 'Everyday Shelf Price']
    edlp_margins = calculate_margin(edlp_price, bottle_cost, 0, coupon)
    df.at[row_idx, 'Everyday GM %'] = edlp_margins["gm_percent"]
    df.at[row_idx, 'Everyday GM $'] = edlp_margins["gm_dollars"]
    df.at[row_idx, 'Everyday GM % (With Coupon)'] = edlp_margins["gm_coupon_percent"]
    
    # TPR Base
    tpr_base_price = df.at[row_idx, 'TPR Price (Base Scan)']
    tpr_base_margins = calculate_margin(tpr_base_price, bottle_cost, base_scan, coupon)
    df.at[row_idx, 'TPR GM % (Base Scan)'] = tpr_base_margins["gm_percent"]
    df.at[row_idx, 'TPR GM $ (Base Scan)'] = tpr_base_margins["gm_dollars"]
    df.at[row_idx, 'TPR GM % (With Coupon)'] = tpr_base_margins["gm_coupon_percent"]
    
    # TPR Deep
    tpr_deep_price = df.at[row_idx, 'TPR Price (Deep Scan)']
    tpr_deep_margins = calculate_margin(tpr_deep_price, bottle_cost, deep_scan, coupon)
    df.at[row_idx, 'TPR GM % (Deep Scan)'] = tpr_deep_margins["gm_percent"]
    df.at[row_idx, 'TPR GM $ (Deep Scan)'] = tpr_deep_margins["gm_dollars"]
    df.at[row_idx, 'TPR GM % (With Coupon)'] = tpr_deep_margins["gm_coupon_percent"]
    
    # Ad Base
    ad_base_price = df.at[row_idx, 'Ad/Feature Price (Base Scan)']
    ad_base_margins = calculate_margin(ad_base_price, bottle_cost, base_scan, coupon)
    df.at[row_idx, 'Ad GM % (Base Scan)'] = ad_base_margins["gm_percent"]
    df.at[row_idx, 'Ad GM $ (Base Scan)'] = ad_base_margins["gm_dollars"]
    df.at[row_idx, 'Ad GM % (With Coupon)'] = ad_base_margins["gm_coupon_percent"]
    
    # Ad Deep
    ad_deep_price = df.at[row_idx, 'Ad/Feature Price (Deep Scan)']
    ad_deep_margins = calculate_margin(ad_deep_price, bottle_cost, deep_scan, coupon)
    df.at[row_idx, 'Ad GM % (Deep Scan)'] = ad_deep_margins["gm_percent"]
    df.at[row_idx, 'Ad GM $ (Deep Scan)'] = ad_deep_margins["gm_dollars"]
    df.at[row_idx, 'Ad GM % (With Coupon)'] = ad_deep_margins["gm_coupon_percent"]
    
    # Calculate bottle cost if case cost or bottles per case changes
    if 'Case Cost' in st.session_state.edited_cells.get(row_idx, {}) or '# Bottles/Cs' in st.session_state.edited_cells.get(row_idx, {}):
        case_cost = df.at[row_idx, 'Case Cost']
        bottles_per_case = df.at[row_idx, '# Bottles/Cs']
        if bottles_per_case > 0:
            df.at[row_idx, 'Bottle Cost'] = case_cost / bottles_per_case
    
    return df

# List of columns that trigger recalculation
calculated_columns = [
    'Everyday GM %', 'Everyday GM $', 'Everyday GM % (With Coupon)',
    'TPR GM % (Base Scan)', 'TPR GM $ (Base Scan)', 'TPR GM % (With Coupon)',
    'TPR GM % (Deep Scan)', 'TPR GM $ (Deep Scan)', 'TPR GM % (With Coupon)',
    'Ad GM % (Base Scan)', 'Ad GM $ (Base Scan)', 'Ad GM % (With Coupon)',
    'Ad GM % (Deep Scan)', 'Ad GM $ (Deep Scan)', 'Ad GM % (With Coupon)'
]

# Display Saved Scan Scenarios section
st.markdown("<h2 style='margin-top: 30px;'>Saved Scan Scenarios</h2>", unsafe_allow_html=True)

if st.session_state.scan_scenarios.empty:
    st.info("No scan scenarios saved yet. Use the 'Save Scan Scenario' button above to save scenarios.")
else:
    # Process deletion of scenarios
    if st.button("Delete Selected Scenarios") and st.session_state.delete_scenario_indices:
        st.session_state.scan_scenarios = st.session_state.scan_scenarios.drop(index=st.session_state.delete_scenario_indices).reset_index(drop=True)
        st.session_state.delete_scenario_indices = []
        st.success("Selected scenarios deleted successfully!")
        st.experimental_rerun()
    
    # Handle column name compatibility between "Market" and "Customer/State"
    if 'Customer/State' not in st.session_state.scan_scenarios.columns and 'Market' in st.session_state.scan_scenarios.columns:
        st.session_state.scan_scenarios = st.session_state.scan_scenarios.rename(columns={'Market': 'Customer/State'})
    
    # Create editable dataframe
    with st.container():
        st.write("Click on any cell to edit (except calculated values)")
        
        # Identify non-calculated columns that can be edited
        all_columns = st.session_state.scan_scenarios.columns.tolist()
        editable_columns = [col for col in all_columns if col not in calculated_columns]
        
        # Create a copy of the dataframe to display and edit
        edit_df = st.session_state.scan_scenarios.copy()
        
        # Display checkboxes for deletion with unique keys
        checkboxes = []
        for i in range(len(edit_df)):
            cb = st.checkbox(f"Select #{i+1}", key=f"select_scenario_{i}")
            checkboxes.append(cb)
            if cb and i not in st.session_state.delete_scenario_indices:
                st.session_state.delete_scenario_indices.append(i)
            elif not cb and i in st.session_state.delete_scenario_indices:
                st.session_state.delete_scenario_indices.remove(i)
        
        # Create tabs for different views of the table
        tab1, tab2, tab3 = st.tabs(["Basic Info", "Pricing Details", "All Data"])
        
        with tab1:
            # Define columns for basic view
            basic_cols = ['Brand', 'Size', 'Case Cost', '# Bottles/Cs', 'Bottle Cost', 'Base Scan', 'Deep Scan', 'Coupon', 'Customer/State', 'Everyday Shelf Price']
            basic_cols = [col for col in basic_cols if col in edit_df.columns]
            
            for i, row in edit_df[basic_cols].iterrows():
                st.markdown(f"**Scenario #{i+1}**")
                cols = st.columns(len(basic_cols))
                for j, col_name in enumerate(basic_cols):
                    with cols[j]:
                        # Determine if this field should be editable
                        is_editable = col_name in editable_columns
                        is_numeric = pd.api.types.is_numeric_dtype(edit_df[col_name].dtype)
                        
                        st.write(f"**{col_name}**")
                        
                        if is_editable:
                            if is_numeric:
                                # Handle numeric inputs
                                step = 0.01 if "Price" in col_name or "Cost" in col_name or "Scan" in col_name or "Coupon" in col_name else 1
                                format_str = "%.2f" if step == 0.01 else None
                                min_val = 0.0 if step == 0.01 else 1
                                
                                new_val = st.number_input(
                                    "",
                                    value=float(row[col_name]),
                                    key=f"edit_{i}_{col_name}",
                                    min_value=float(min_val) if "# Bottles" in col_name else 0.0,
                                    step=float(step),
                                    format=format_str
                                )
                                
                                if new_val != edit_df.at[i, col_name]:
                                    edit_df.at[i, col_name] = new_val
                                    if i not in st.session_state.edited_cells:
                                        st.session_state.edited_cells[i] = {}
                                    st.session_state.edited_cells[i][col_name] = True
                                    # Recalculate margins if needed
                                    if col_name in ['Case Cost', '# Bottles/Cs', 'Bottle Cost', 'Base Scan', 'Deep Scan', 'Coupon', 'Everyday Shelf Price', 
                                                   'TPR Price (Base Scan)', 'TPR Price (Deep Scan)', 'Ad/Feature Price (Base Scan)', 'Ad/Feature Price (Deep Scan)']:
                                        edit_df = recalculate_margins(edit_df, i)
                            else:
                                # Handle text inputs
                                if col_name == 'Size' and 'Size' in edit_df.columns:
                                    new_val = st.selectbox(
                                        "",
                                        options=size_options,
                                        index=size_options.index(row[col_name]) if row[col_name] in size_options else 0,
                                        key=f"edit_{i}_{col_name}"
                                    )
                                else:
                                    new_val = st.text_input("", value=row[col_name], key=f"edit_{i}_{col_name}")
                                
                                if new_val != edit_df.at[i, col_name]:
                                    edit_df.at[i, col_name] = new_val
                                    if i not in st.session_state.edited_cells:
                                        st.session_state.edited_cells[i] = {}
                                    st.session_state.edited_cells[i][col_name] = True
                        else:
                            # Show calculated values as text
                            if is_numeric:
                                if "GM %" in col_name:
                                    st.write(f"{row[col_name]:.1f}%")
                                elif "Cost" in col_name or "Price" in col_name or "Scan" in col_name or "GM $" in col_name or "Coupon" in col_name:
                                    st.write(f"${row[col_name]:.2f}")
                                else:
                                    st.write(f"{row[col_name]}")
                            else:
                                st.write(f"{row[col_name]}")
                
                st.markdown("---")
        
        with tab2:
            # Define columns for pricing view
            pricing_cols = ['Brand', 'Everyday Shelf Price', 'Everyday GM %', 'Everyday GM $',
                           'TPR Price (Base Scan)', 'TPR GM % (Base Scan)', 'TPR GM $ (Base Scan)',
                           'TPR Price (Deep Scan)', 'TPR GM % (Deep Scan)', 'TPR GM $ (Deep Scan)',
                           'Ad/Feature Price (Base Scan)', 'Ad GM % (Base Scan)', 'Ad GM $ (Base Scan)',
                           'Ad/Feature Price (Deep Scan)', 'Ad GM % (Deep Scan)', 'Ad GM $ (Deep Scan)']
            pricing_cols = [col for col in pricing_cols if col in edit_df.columns]
            
            for i, row in edit_df[pricing_cols].iterrows():
                st.markdown(f"**Scenario #{i+1}: {row['Brand']}**")
                
                # Create three columns for the three types of pricing
                price_col1, price_col2, price_col3, price_col4 = st.columns(4)
                
                with price_col1:
                    st.markdown("**Everyday Price**")
                    
                    # Price
                    st.write("Price:")
                    new_val = st.number_input(
                        "",
                        value=float(row['Everyday Shelf Price']),
                        key=f"edit_{i}_everyday_price",
                        min_value=0.0,
                        step=0.01,
                        format="%.2f"
                    )
                    if new_val != edit_df.at[i, 'Everyday Shelf Price']:
                        edit_df.at[i, 'Everyday Shelf Price'] = new_val
                        if i not in st.session_state.edited_cells:
                            st.session_state.edited_cells[i] = {}
                        st.session_state.edited_cells[i]['Everyday Shelf Price'] = True
                        edit_df = recalculate_margins(edit_df, i)
                    
                    # Display calculated values
                    st.write(f"GM %: {row['Everyday GM %']:.1f}%")
                    st.write(f"GM $: ${row['Everyday GM $']:.2f}")
                
                with price_col2:
                    st.markdown("**TPR (Base Scan)**")
                    
                    # Price
                    st.write("Price:")
                    new_val = st.number_input(
                        "",
                        value=float(row['TPR Price (Base Scan)']),
                        key=f"edit_{i}_tpr_base_price",
                        min_value=0.0,
                        step=0.01,
                        format="%.2f"
                    )
                    if new_val != edit_df.at[i, 'TPR Price (Base Scan)']:
                        edit_df.at[i, 'TPR Price (Base Scan)'] = new_val
                        if i not in st.session_state.edited_cells:
                            st.session_state.edited_cells[i] = {}
                        st.session_state.edited_cells[i]['TPR Price (Base Scan)'] = True
                        edit_df = recalculate_margins(edit_df, i)
                    
                    # Display calculated values
                    st.write(f"GM %: {row['TPR GM % (Base Scan)']:.1f}%")
                    st.write(f"GM $: ${row['TPR GM $ (Base Scan)']:.2f}")
                
                with price_col3:
                    st.markdown("**TPR (Deep Scan)**")
                    
                    # Price
                    st.write("Price:")
                    new_val = st.number_input(
                        "",
                        value=float(row['TPR Price (Deep Scan)']),
                        key=f"edit_{i}_tpr_deep_price",
                        min_value=0.0,
                        step=0.01,
                        format="%.2f"
                    )
                    if new_val != edit_df.at[i, 'TPR Price (Deep Scan)']:
                        edit_df.at[i, 'TPR Price (Deep Scan)'] = new_val
                        if i not in st.session_state.edited_cells:
                            st.session_state.edited_cells[i] = {}
                        st.session_state.edited_cells[i]['TPR Price (Deep Scan)'] = True
                        edit_df = recalculate_margins(edit_df, i)
                    
                    # Display calculated values
                    st.write(f"GM %: {row['TPR GM % (Deep Scan)']:.1f}%")
                    st.write(f"GM $: ${row['TPR GM $ (Deep Scan)']:.2f}")
                
                with price_col4:
                    st.markdown("**Ad/Feature**")
                    
                    # Base Price
                    st.write("Base Price:")
                    new_val = st.number_input(
                        "",
                        value=float(row['Ad/Feature Price (Base Scan)']),
                        key=f"edit_{i}_ad_base_price",
                        min_value=0.0,
                        step=0.01,
                        format="%.2f"
                    )
                    if new_val != edit_df.at[i, 'Ad/Feature Price (Base Scan)']:
                        edit_df.at[i, 'Ad/Feature Price (Base Scan)'] = new_val
                        if i not in st.session_state.edited_cells:
                            st.session_state.edited_cells[i] = {}
                        st.session_state.edited_cells[i]['Ad/Feature Price (Base Scan)'] = True
                        edit_df = recalculate_margins(edit_df, i)
                    
                    # Deep Price
                    st.write("Deep Price:")
                    new_val = st.number_input(
                        "",
                        value=float(row['Ad/Feature Price (Deep Scan)']),
                        key=f"edit_{i}_ad_deep_price",
                        min_value=0.0,
                        step=0.01,
                        format="%.2f"
                    )
                    if new_val != edit_df.at[i, 'Ad/Feature Price (Deep Scan)']:
                        edit_df.at[i, 'Ad/Feature Price (Deep Scan)'] = new_val
                        if i not in st.session_state.edited_cells:
                            st.session_state.edited_cells[i] = {}
                        st.session_state.edited_cells[i]['Ad/Feature Price (Deep Scan)'] = True
                        edit_df = recalculate_margins(edit_df, i)
                
                st.markdown("---")
        
        with tab3:
            # Show all data in a table format
            st.dataframe(edit_df.style.format({
                col: "${:.2f}" if "Price" in col or "Cost" in col or "Scan" in col or "GM $" in col or "Coupon" in col else "{:.1f}%" if "GM %" in col else "{}"
                for col in edit_df.columns
            }), height=400)
        
        # Update the session state with the edited dataframe
        st.session_state.scan_scenarios = edit_df.copy()
        
        # Add a button to apply all changes
        if st.button("Apply All Changes"):
            st.session_state.edited_cells = {}
            st.success("All changes applied successfully!")
            st.experimental_rerun()
    
    # Export button
    if st.button("Export Scan Scenarios to Excel"):
        # Function to create a downloadable Excel file
        def to_excel_scenarios(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Scan Scenarios', index=False)
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Scan Scenarios']
            
            # Set column width and formats
            format_currency = workbook.add_format({'num_format': '$#,##0.00'})
            format_percent = workbook.add_format({'num_format': '0.0%'})
            
            # Format currency columns
            for i, col in enumerate(df.columns):
                if "Price" in col or "Cost" in col or "Scan" in col or "Coupon" in col or "GM $" in col:
                    worksheet.set_column(i, i, 12, format_currency)
                elif "GM %" in col:
                    worksheet.set_column(i, i, 12, format_percent)
                else:
                    worksheet.set_column(i, i, 15)
            
            # Freeze top row and make it bold
            worksheet.freeze_panes(1, 0)
            header_format = workbook.add_format({'bold': True})
            worksheet.set_row(0, None, header_format)
            
            writer.close()
            return output.getvalue()
        
        excel_file = to_excel_scenarios(st.session_state.scan_scenarios)
        b64 = base64.b64encode(excel_file).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="saved_scan_scenarios.xlsx">Download Scan Scenarios Excel File</a>'
        st.markdown(href, unsafe_allow_html=True)
        st.success("Export complete! Click the link above to download.")
    
    # Clear button
    if st.button("Clear All Scan Scenarios"):
        st.session_state.scan_scenarios = pd.DataFrame()
        st.session_state.delete_scenario_indices = []
        st.session_state.edited_cells = {}
        st.success("All scan scenarios cleared!")
        st.experimental_rerun()