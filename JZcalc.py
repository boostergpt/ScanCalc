import streamlit as st
import pandas as pd
import numpy as np
import base64
from io import BytesIO
from PIL import Image
import os

# Set page config
st.set_page_config(
    page_title="Margin Calculator MAYBE",
    page_icon="🧮",
    layout="wide"
)

# Custom CSS for better styling and uniform fonts
st.markdown("""
<style>
    /* Global font settings for uniformity */
    * {
        font-family: 'Arial', sans-serif !important;
    }
    
    /* Streamlit elements */
    .main {
        padding: 2rem;
    }
    .stButton button {
        width: 100%;
        font-size: 1.2rem;
        font-family: 'Arial', sans-serif !important;
    }
    .stTextInput input, .stNumberInput input, .stSelectbox > div > div[role="listbox"] {
        font-size: 1.2rem !important;
        font-family: 'Arial', sans-serif !important;
    }
    .stTextInput label, .stNumberInput label, .stSelectbox label {
        font-size: 1.2rem !important;
        font-weight: 500 !important;
        font-family: 'Arial', sans-serif !important;
    }
    .stTab [data-baseweb="tab"] {
        font-size: 1.2rem !important;
        font-family: 'Arial', sans-serif !important;
    }
    .stTab [data-baseweb="tab-list"] {
        gap: 1rem;
    }
    .stSubheader {
        font-size: 1.5rem !important;
        font-family: 'Arial', sans-serif !important;
    }
    p, div, span, button, a {
        font-size: 1.2rem;
        font-family: 'Arial', sans-serif !important;
    }
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Arial', sans-serif !important;
    }
    
    /* Component styling */
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
        font-size: 2.2rem;
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
    .sidebar .sidebar-content {
        padding: 2rem 1rem;
    }
    button[data-baseweb="tab"] {
        font-size: 1.3rem !important;
        font-family: 'Arial', sans-serif !important;
    }
    
    /* Pricing table styling */
    .pricing-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.4rem;
        margin-bottom: 20px;
        font-family: 'Arial', sans-serif !important;
    }
    .pricing-table th {
        background-color: #f0f2f6;
        padding: 12px;
        text-align: center;
        font-weight: bold;
        border: 1px solid #ddd;
        font-family: 'Arial', sans-serif !important;
    }
    .pricing-table td {
        padding: 12px;
        border: 1px solid #ddd;
        text-align: center;
        font-family: 'Arial', sans-serif !important;
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
        font-family: 'Arial', sans-serif !important;
    }
    .delete-button:hover {
        background-color: #e74c3c;
    }
    .everyday-row {
        background-color: #d4edda !important;
    }
    
    /* Excel-like Scenarios table styling with alternating blue/white colors */
    .scenarios-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.2rem;
        margin-bottom: 20px;
        font-family: 'Arial', sans-serif !important;
        border: 1px solid #bdd7ee; /* Excel-like border */
        box-shadow: 0 0 5px rgba(0,0,0,0.1); /* Subtle shadow for depth */
    }
    .scenarios-table th {
        background-color: #4472c4; /* Excel blue header */
        color: white;
        padding: 10px;
        text-align: center;
        font-weight: bold;
        border: 1px solid #bdd7ee;
        position: sticky;
        top: 0;
        z-index: 10;
        font-family: 'Arial', sans-serif !important;
    }
    .scenarios-table td {
        padding: 8px;
        border: 1px solid #bdd7ee; /* Excel-like cell borders */
        text-align: center;
        font-family: 'Arial', sans-serif !important;
    }
    /* Brand column width */
    .scenarios-table th:nth-child(2),
    .scenarios-table td:nth-child(2) {
        width: 300px; /* Make Brand column wider */
        min-width: 300px;
        max-width: 300px;
        text-align: left;
        white-space: normal;
    }
    /* Grid lines */
    .scenarios-table th, .scenarios-table td {
        border: 1px solid #bdd7ee; /* Stronger grid lines */
    }
    .scenarios-table tr:nth-child(odd) {
        background-color: #ffffff; /* White rows */
    }
    .scenarios-table tr:nth-child(even) {
        background-color: #deebf7; /* Light blue rows - Excel style */
    }
    .scenarios-table tr:hover {
        background-color: #c5e0b4; /* Excel-like selection color */
    }
    .checkbox-col {
        width: 50px;
        min-width: 50px;
        text-align: center;
        background-color: #f7f7f7; /* Slight gray for checkbox column */
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for storing scan scenarios
if 'scan_scenarios' not in st.session_state:
    st.session_state.scan_scenarios = pd.DataFrame()
    
# For managing scenario deletion
if 'delete_scenario_indices' not in st.session_state:
    st.session_state.delete_scenario_indices = []

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

# Display Saved Scan Scenarios section
st.markdown("<h2 style='margin-top: 30px;'>Saved Scan Scenarios</h2>", unsafe_allow_html=True)

if st.session_state.scan_scenarios.empty:
    st.info("No scan scenarios saved yet. Use the 'Save Scan Scenario' button above to save scenarios.")
else:
    # Handle column name compatibility between "Market" and "Customer/State"
    if 'Customer/State' not in st.session_state.scan_scenarios.columns and 'Market' in st.session_state.scan_scenarios.columns:
        st.session_state.scan_scenarios = st.session_state.scan_scenarios.rename(columns={'Market': 'Customer/State'})
    
    # Create display dataframe with selected columns (with fallback columns if needed)
    display_columns = ['Brand', 'Size', 'Case Cost', 'Bottle Cost']
    # Add Customer/State if it exists, otherwise skip it
    if 'Customer/State' in st.session_state.scan_scenarios.columns:
        display_columns.append('Customer/State')
    # Add the pricing columns
    display_columns.extend(['Everyday Shelf Price', 'Everyday GM %', 'Everyday GM $', 
                          'TPR Price (Base Scan)', 'TPR GM % (Base Scan)', 'TPR GM $ (Base Scan)',
                          'TPR Price (Deep Scan)', 'TPR GM % (Deep Scan)', 'TPR GM $ (Deep Scan)',
                          'Ad/Feature Price (Base Scan)', 'Ad GM % (Base Scan)', 'Ad GM $ (Base Scan)',
                          'Ad/Feature Price (Deep Scan)', 'Ad GM % (Deep Scan)', 'Ad GM $ (Deep Scan)'])
    
    # Process deletion of scenarios if button was clicked
    if st.button("Delete Selected Scenarios") and st.session_state.delete_scenario_indices:
        st.session_state.scan_scenarios = st.session_state.scan_scenarios.drop(index=st.session_state.delete_scenario_indices).reset_index(drop=True)
        st.session_state.delete_scenario_indices = []
        st.success("Selected scenarios deleted successfully!")
        st.experimental_rerun()
    
    # Create Excel-like table with checkboxes and alternating row colors with grid
    table_html = "<div style='overflow-x: auto;'><table class='scenarios-table'><thead><tr>"
    # Add checkbox column header
    table_html += "<th class='checkbox-col'>Select</th>"
    
    # Add other column headers
    for col in display_columns:
        table_html += f"<th>{col}</th>"
    table_html += "</tr></thead><tbody>"
    
    # Add rows with data
    for i in range(len(st.session_state.scan_scenarios)):
        # Row class is handled by CSS for alternating colors
        table_html += "<tr>"
        
        # Add checkbox column
        checkbox_id = f"checkbox_{i}"
        is_checked = "checked" if i in st.session_state.delete_scenario_indices else ""
        table_html += f"<td class='checkbox-col'><input type='checkbox' id='{checkbox_id}' {is_checked}></td>"
        
        # Add data columns
        for j, col in enumerate(display_columns):
            value = st.session_state.scan_scenarios.iloc[i][col]
            
            # Format values based on type
            if isinstance(value, (int, float)):
                if "Price" in col or "Cost" in col or "Scan" in col or "GM $" in col:
                    formatted_value = f"${value:.2f}"
                elif "GM %" in col:
                    formatted_value = f"{value:.1f}%"
                else:
                    formatted_value = f"{value}"
            else:
                formatted_value = f"{value}"
            
            # Special cell class if this is the Brand column (index 0 in display_columns)
            if j == 0 and col == "Brand":
                table_html += f"<td class='brand-column'>{formatted_value}</td>"
            else:
                table_html += f"<td>{formatted_value}</td>"
        
        table_html += "</tr>"
    
    table_html += "</tbody></table></div>"
    
    # JavaScript to handle checkbox selections
    checkbox_script = """
    <script>
    document.addEventListener('DOMContentLoaded', function() {
        const checkboxes = document.querySelectorAll('input[type="checkbox"][id^="checkbox_"]');
        checkboxes.forEach(function(checkbox) {
            checkbox.addEventListener('change', function() {
                const index = this.id.split('_')[1];
                if (this.checked) {
                    // Add to selected indices
                    window.parent.postMessage({
                        'type': 'streamlit:setComponentValue',
                        'value': index
                    }, '*');
                } else {
                    // Remove from selected indices
                    window.parent.postMessage({
                        'type': 'streamlit:removeComponentValue',
                        'value': index
                    }, '*');
                }
            });
        });
    });
    </script>
    """
    
    # Display the table
    st.components.v1.html(table_html + checkbox_script, height=400, scrolling=True)
    
    # Create a form to collect checkbox states
    with st.form("scenario_selection_form"):
        # Create hidden checkboxes to store state
        for i in range(len(st.session_state.scan_scenarios)):
            # These are invisible but maintain state
            if st.checkbox(f"Scenario {i}", value=i in st.session_state.delete_scenario_indices, key=f"select_{i}", label_visibility="collapsed"):
                if i not in st.session_state.delete_scenario_indices:
                    st.session_state.delete_scenario_indices.append(i)
            else:
                if i in st.session_state.delete_scenario_indices:
                    st.session_state.delete_scenario_indices.remove(i)
        
        # Submit button
        st.form_submit_button("Update Selection")
    
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
        st.success("All scan scenarios cleared!")
        st.experimental_rerun()