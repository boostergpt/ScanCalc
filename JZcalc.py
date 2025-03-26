import streamlit as st
import pandas as pd
import numpy as np
import base64
from io import BytesIO
from PIL import Image
import os
import datetime
import math

# Import lists from lists.py
from lists import brand_list, segment_list, size_options, segment_mapping, segment_size_combinations

# Set page config
st.set_page_config(
    page_title="Pricing & Margin Calculator",
    page_icon="🧮",
    layout="wide"
)

# Custom CSS for styling
custom_css = """
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
    
    /* Make scenarios table match the pricing table styling */
    .scenarios-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.4rem;
        margin-bottom: 20px;
        font-family: 'Arial', sans-serif !important;
    }
    
    /* Index ratio table styling */
    .index-ratio-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.2rem;
        margin-bottom: 20px;
        font-family: 'Arial', sans-serif !important;
    }
    .index-ratio-table th {
        background-color: #f0f2f6;
        padding: 8px;
        text-align: center;
        font-weight: bold;
        border: 1px solid #ddd;
        font-family: 'Arial', sans-serif !important;
    }
    .index-ratio-table td {
        padding: 8px;
        border: 1px solid #ddd;
        text-align: center;
        font-family: 'Arial', sans-serif !important;
    }
    .above-average {
        color: #228B22 !important; /* Forest green */
    }
    .below-average {
        color: #8B0000 !important; /* Deep red */
    }
    .high-value-cell {
        background-color: #90EE90 !important; /* Light green */
    }
    .low-value-cell {
        background-color: #FFB6C1 !important; /* Light pink */
    }
</style>
"""

# Apply custom CSS
st.markdown(custom_css, unsafe_allow_html=True)

# Initialize session state variables
if 'scan_scenarios' not in st.session_state:
    st.session_state.scan_scenarios = pd.DataFrame()
    
if 'h1_data' not in st.session_state:
    # Create a placeholder H1Data dataframe
    # In real use, this would be loaded from a file or database
    st.session_state.h1_data = pd.DataFrame()

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

# Function to generate H1 data if not available
def generate_sample_h1_data():
    """Generate sample H1 data for demo purposes"""
    segments = ["Cordials", "Gin", "NAW", "RTD", "RTS", "Rum", "Scotch", "Tequila", "Vodka"]
    sizes = ["1.75L", "1L", "750mL", "375mL", "355mL"]
    
    data = []
    
    for segment in segments:
        available_sizes = segment_size_combinations.get(segment, ["1.75L", "750mL"])
        for size in available_sizes:
            # Base value varies by segment and size to create realistic patterns
            base_value = np.random.uniform(50000, 500000)
            
            # Create 27 weeks of data with some trends and randomness
            weekly_values = []
            for week in range(1, 28):
                # Add seasonality and randomness
                seasonal_factor = 1 + 0.2 * math.sin(week / 4)  # Creates a wave pattern
                random_factor = np.random.uniform(0.8, 1.2)     # Random variation
                
                weekly_value = base_value * seasonal_factor * random_factor
                weekly_values.append(weekly_value)
            
            row = [segment, size] + weekly_values
            data.append(row)
    
    # Create column names
    columns = ["Segment", "Size"] + [f"Week{i}" for i in range(1, 28)]
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=columns)
    return df

# Generate H1 data if needed
if st.session_state.h1_data.empty:
    st.session_state.h1_data = generate_sample_h1_data()

# Function to calculate index ratios from H1 data
def get_index_ratios(segment, size):
    """Calculate index ratios for a given segment and size"""
    # Look up the correct segment name for data lookup
    data_segment = segment_mapping.get(segment, segment)
    
    # Find the matching row in the H1 data
    h1_data = st.session_state.h1_data
    
    if h1_data.empty:
        return None, None, None
        
    matching_row = h1_data[(h1_data['Segment'].str.upper() == data_segment.upper()) & 
                           (h1_data['Size'].str.upper() == size.upper())]
    
    if matching_row.empty:
        return None, None, None
        
    # Extract weekly data
    weekly_data = matching_row.iloc[0, 2:].values.astype(float)
    
    # Calculate average weekly RSV
    avg_weekly_rsv = weekly_data.mean()
    
    # Calculate standard deviation
    std_dev = weekly_data.std()
    
    # Calculate threshold for highlighting
    threshold = std_dev / avg_weekly_rsv if avg_weekly_rsv > 0 else 0
    
    # Calculate index ratios
    index_ratios = [val / avg_weekly_rsv if avg_weekly_rsv > 0 else 0 for val in weekly_data]
    
    return index_ratios, avg_weekly_rsv, threshold

# Function to export to Excel
def to_excel(df):
    """Convert dataframe to Excel file"""
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

# Function to export segment size index ratios
def export_segment_size_index_ratios():
    """Create a dataframe with segment size index ratios for all combinations"""
    # Format date for week labels
    current_year = datetime.datetime.now().year
    week_start_date = datetime.datetime(current_year, 6, 30)
    
    # Create week labels
    week_labels = []
    for i in range(1, 28):
        week_end_date = week_start_date + datetime.timedelta(days=6)
        if i == 1:
            week_label = f"{i}: Jun 30-Jul 6"
        else:
            week_label = f"{i}: {week_start_date.strftime('%b %d')}-{week_end_date.strftime('%b %d')}"
        week_labels.append(week_label)
        week_start_date += datetime.timedelta(days=7)
    
    # Create dataframe with week labels as index
    df = pd.DataFrame(index=week_labels)
    
    # Add columns for each segment/size combination
    for segment, sizes in segment_size_combinations.items():
        display_segment = next((k for k, v in segment_mapping.items() if v == segment), segment)
        
        for size in sizes:
            # Get index ratios for this segment/size combo
            index_ratios, avg_weekly_rsv, threshold = get_index_ratios(display_segment, size)
            
            if index_ratios:
                # Add the column to the dataframe
                column_name = f"{display_segment} | {size}"
                df[column_name] = index_ratios
    
    return df

def create_calculator_ui():
    # Create sidebar for inputs
    with st.sidebar:
        # Brand, Segment and Size inputs
        brand = st.selectbox("Brand", brand_list, index=0)
        segment = st.selectbox("Segment", segment_list, index=0)
        size = st.selectbox("Size", size_options, index=0)
        
        # Add a separator
        st.markdown("---")
        
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
        
        # Add a separator
        st.markdown("---")
        
        # Pricing inputs
        edlp_price = st.number_input("Everyday Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        tpr_base_price = st.number_input("TPR Base Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        tpr_deep_price = st.number_input("TPR Deep Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        ad_base_price = st.number_input("Ad/Feature Base Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        ad_deep_price = st.number_input("Ad/Feature Deep Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        
        # Optional customer/state input
        customer_state = st.text_input("Customer / State", "")

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
            "Segment": segment,
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
        
        # Display the scenarios using Streamlit's dataframe
        st.dataframe(st.session_state.scan_scenarios)
        
        # Export button
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Export Scan Scenarios to Excel"):
                # Convert dataframe to excel
                excel_file = to_excel(st.session_state.scan_scenarios)
                b64 = base64.b64encode(excel_file).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="saved_scan_scenarios.xlsx">Download Scan Scenarios Excel File</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("Export complete! Click the link above to download.")
        
        with col2:
            # Clear button
            if st.button("Clear All Scan Scenarios"):
                st.session_state.scan_scenarios = pd.DataFrame()
                st.rerun()
                

def create_h1_data_analysis_ui():
    """Create the UI for H1 Data Analysis tab"""
    st.markdown("<h2>H1 Data Analysis</h2>", unsafe_allow_html=True)
    
    # Create a two-column layout
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # User selects segment and size for analysis
        segment = st.selectbox("Select Segment", segment_list, index=0, key='h1_segment')
        
        # Filter size options based on selected segment
        data_segment = segment_mapping.get(segment, segment)
        if data_segment in segment_size_combinations:
            valid_sizes = ["-- Select Size --"] + segment_size_combinations[data_segment]
        else:
            valid_sizes = size_options
            
        size = st.selectbox("Select Size", valid_sizes, index=0, key='h1_size')
    
    with col2:
        # Display average weekly RSV
        if segment != "-- Select Segment --" and size != "-- Select Size --":
            index_ratios, avg_weekly_rsv, threshold = get_index_ratios(segment, size)
            
            if avg_weekly_rsv:
                st.metric("H1 Average Weekly RSV", f"${avg_weekly_rsv/1000000:.2f}M")
            else:
                st.info("No data available for the selected segment and size.")
    
    # Display the index ratios table if data is available
    if segment != "-- Select Segment --" and size != "-- Select Size --" and index_ratios is not None:
        # Create a weekly table with dates and index ratios
        st.markdown("<h3>Weekly Index Ratios</h3>", unsafe_allow_html=True)
        
        # Format date for week labels
        current_year = datetime.datetime.now().year
        week_start_date = datetime.datetime(current_year, 6, 30)
        
        # Create a table to display the data
        week_table = "<table class='index-ratio-table'><thead><tr><th>F'25 Week</th><th>Index Ratio</th></tr></thead><tbody>"
        
        for i, ratio in enumerate(index_ratios):
            week_end_date = week_start_date + datetime.timedelta(days=6)
            
            if i == 0:
                week_label = f"1: Jun 30-Jul 6"
            else:
                week_label = f"{i+1}: {week_start_date.strftime('%b %d')}-{week_end_date.strftime('%b %d')}"
            
            # Determine cell class based on ratio value compared to threshold
            cell_class = ""
            if avg_weekly_rsv > 0:
                if ratio > 1 + threshold:
                    cell_class = "class='above-average'"
                elif ratio < 1 - (0.5 * threshold):
                    cell_class = "class='below-average'"
            
            week_table += f"<tr><td>{week_label}</td><td {cell_class}>{ratio:.0%}</td></tr>"
            
            week_start_date += datetime.timedelta(days=7)
            
        week_table += "</tbody></table>"
        
        # Display the table
        st.markdown(week_table, unsafe_allow_html=True)
    
    # Add a separator
    st.markdown("---")
    
    # Export button for all segment/size combinations
    st.markdown("<h3>Export Segment/Size Index Ratios</h3>", unsafe_allow_html=True)
    
    if st.button("Export All Segment/Size Index Ratios"):
        # Create a dataframe with all segment/size index ratios
        index_ratios_df = export_segment_size_index_ratios()
        
        # Convert to Excel
        excel_file = BytesIO()
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        index_ratios_df.to_excel(writer, sheet_name='SegSizeIndexRatio_NielsenxAOC', index=True)
        
        # Get the workbook and worksheet objects for formatting
        workbook = writer.book
        worksheet = writer.sheets['SegSizeIndexRatio_NielsenxAOC']
        
        # Format index as dates
        worksheet.set_column(0, 0, 18)  # Width for week column
        
        # Format data columns as percentages
        for col_num in range(1, len(index_ratios_df.columns) + 1):
            worksheet.set_column(col_num, col_num, 6)  # Narrow columns
            
        # Add conditional formatting
        for col_num in range(1, len(index_ratios_df.columns) + 1):
            # Light green for values > 1 + threshold
            worksheet.conditional_format(1, col_num, len(index_ratios_df) + 1, col_num, {
                'type': '3_color_scale',
                'min_color': '#FFB6C1',  # Light pink
                'mid_color': '#FFFFFF',  # White
                'max_color': '#90EE90',  # Light green
                'min_type': 'num',
                'min_value': 0.7,
                'mid_type': 'num',
                'mid_value': 1.0,
                'max_type': 'num',
                'max_value': 1.3
            })
            
        # Set header row format
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#C8C8C8',  # Light gray
            'border': 1
        })
        for col_num, value in enumerate(index_ratios_df.columns.values):
            worksheet.write(0, col_num + 1, value, header_format)
        worksheet.write(0, 0, 'F\'25 Week', header_format)
            
        # Freeze panes
        worksheet.freeze_panes(1, 1)
        
        # Page setup
        worksheet.set_page_setup({
            'orientation': 'landscape',
            'fit_to_width': 1,
            'fit_to_height': 1
        })
        
        # Close the writer
        writer.close()
        
        # Create download link
        b64 = base64.b64encode(excel_file.getvalue()).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="SegSizeIndexRatio_NielsenxAOC.xlsx">Download SegSizeIndexRatio Excel File</a>'
        st.markdown(href, unsafe_allow_html=True)
        st.success("Export complete! Click the link above to download.")
        
    # Optionally, display a preview of the index ratios for all segment/size combinations
    if st.checkbox("Show preview of all segment/size index ratios"):
        preview_df = export_segment_size_index_ratios()
        if not preview_df.empty:
            st.dataframe(preview_df.style.format("{:.0%}"))
        else:
            st.info("No data available for preview.")


def main():
    # Add logo and title to the main area
    try:
        image = Image.open('image.png')
        col1, col2, col3 = st.columns([1, 3, 1])
        with col2:
            st.image(image, width=600, use_column_width=True)
    except Exception:
        st.warning("Logo 'image.png' not found. Please add it to the same directory as the app.")

    st.markdown("<h1 class='header'>Pricing & Margin Calculator</h1>", unsafe_allow_html=True)

    # Create tabs for Calculator and H1 Data Analysis
    tab1, tab2 = st.tabs(["Calculator", "H1 Data Analysis"])

    with tab1:
        # Main calculator section
        create_calculator_ui()

    with tab2:
        # H1 Data Analysis section
        create_h1_data_analysis_ui()


# Execute the app
if __name__ == "__main__":
    main()