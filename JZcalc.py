import streamlit as st
import pandas as pd
import numpy as np
import base64
from io import BytesIO
from PIL import Image
import os
import datetime
import math
import re

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
    
    /* Additional styles for sidebar columns */
    .stSidebar .stColumn > div {
        padding-left: 0 !important;
        padding-right: 0 !important;
    }
    .stSidebar .stNumberInput label, .stSidebar .stSelectbox label, .stSidebar .stTextInput label {
        font-size: 0.9rem !important;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .stSidebar .stSelectbox > div > div[role="listbox"] {
        font-size: 0.9rem !important;
    }
    .stSidebar .stNumberInput input, .stSidebar .stTextInput input {
        font-size: 0.9rem !important;
    }
    
    /* Pricing table styling - MADE SMALLER */
    .pricing-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.1rem; /* Reduced from 1.4rem */
        margin-bottom: 15px; /* Reduced from 20px */
        font-family: 'Arial', sans-serif !important;
    }
    .pricing-table th {
        background-color: #f0f2f6;
        padding: 8px; /* Reduced from 12px */
        text-align: center;
        font-weight: bold;
        border: 1px solid #ddd;
        font-family: 'Arial', sans-serif !important;
    }
    .pricing-table td {
        padding: 8px; /* Reduced from 12px */
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
        padding: 6px 10px; /* Reduced from 8px 12px */
        border-radius: 4px;
        cursor: pointer;
        font-size: 1rem; /* Reduced from 1.1rem */
        margin: 4px; /* Reduced from 5px */
        font-family: 'Arial', sans-serif !important;
    }
    .delete-button:hover {
        background-color: #e74c3c;
    }
    .everyday-row {
        background-color: #d4edda !important;
    }
    
    /* Make scenarios table match the pricing table styling - MADE SMALLER */
    .scenarios-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1.1rem; /* Reduced from 1.4rem */
        margin-bottom: 15px; /* Reduced from 20px */
        font-family: 'Arial', sans-serif !important;
    }
    
    /* Index ratio table styling - MADE SMALLER */
    .index-ratio-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 1rem; /* Reduced from 1.2rem */
        margin-bottom: 15px; /* Reduced from 20px */
        font-family: 'Arial', sans-serif !important;
    }
    .index-ratio-table th {
        background-color: #f0f2f6;
        padding: 6px; /* Reduced from 8px */
        text-align: center;
        font-weight: bold;
        border: 1px solid #ddd;
        font-family: 'Arial', sans-serif !important;
    }
    .index-ratio-table td {
        padding: 6px; /* Reduced from 8px */
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
    
    /* Added compact layout for combined view */
    .compact-section {
        margin-bottom: 1rem;
    }
    .compact-section h2, .compact-section h3 {
        margin-top: 0.8rem;
        margin-bottom: 0.5rem;
        font-size: 1.3rem;
    }
    .compact-divider {
        margin: 0.8rem 0;
        border-top: 1px solid #eee;
    }
    
    /* Added compact layout for combined view */
    .compact-section {
        margin-bottom: 0.5rem;
    }
    .compact-section h2, .compact-section h3, .compact-section h4 {
        margin-top: 0.5rem;
        margin-bottom: 0.3rem;
        font-size: 1.1rem;
        font-weight: bold;
    }
    .compact-divider {
        margin: 0.8rem 0;
        border-top: 1px solid #eee;
    }
    
    /* Side panel for H1 Data Analysis */
    .side-panel {
        background-color: #f8f9fa;
        padding: 10px;
        border-radius: 5px;
        border-left: 3px solid #4e73df;
        height: 100%;
    }
    
    /* Scrollable area */
    .scrollable-area {
        max-height: 350px;
        overflow-y: auto;
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

# Enhanced input validation function
def format_numeric_input(value, decimal_places=2):
    """Format numeric input to specified decimal places"""
    if value is None or value == "":
        return 0.0
    
    try:
        # Convert to float and round
        numeric_value = float(value)
        formatted_value = round(numeric_value, decimal_places)
        return formatted_value
    except ValueError:
        # If conversion fails, return 0.0
        return 0.0

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

# New function to calculate weighted averages (similar to CalculateWeightedAverage in VBA)
def calculate_weighted_average(tpr_price, ad_price, bottle_cost, scan, coupon, ad_percentage):
    """Calculate weighted average margins based on ad percentage"""
    if ad_percentage is None or ad_percentage == "":
        return None, None
    
    # Calculate TPR and Ad margins
    tpr_margins = calculate_margin(tpr_price, bottle_cost, scan, coupon)
    ad_margins = calculate_margin(ad_price, bottle_cost, scan, coupon)
    
    # Convert ad_percentage to decimal
    ad_percentage_decimal = ad_percentage / 100
    tpr_percentage = 1 - ad_percentage_decimal
    
    # Calculate weighted averages
    weighted_gm_percent = (tpr_margins["gm_percent"] * tpr_percentage) + (ad_margins["gm_percent"] * ad_percentage_decimal)
    weighted_gm_dollars = (tpr_margins["gm_dollars"] * tpr_percentage) + (ad_margins["gm_dollars"] * ad_percentage_decimal)
    
    return weighted_gm_percent, weighted_gm_dollars

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
    columns = ["Segment", "Size"] + [f"Week {i}" for i in range(1, 28)]
    
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
        
    # Make case-insensitive comparison
    matching_row = h1_data[(h1_data['Segment'].str.upper() == data_segment.upper()) & 
                           (h1_data['Size'].str.upper() == size.upper())]
    
    if matching_row.empty:
        return None, None, None
        
    # Get week columns (all columns except the first two)
    week_cols = h1_data.columns[2:]
    
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
        elif "GM %" in col or "% on Ad" in col:
            worksheet.set_column(i, i, 12, format_percent)
        else:
            worksheet.set_column(i, i, 15)
    
    # Freeze top row and make it bold
    worksheet.freeze_panes(1, 0)
    header_format = workbook.add_format({'bold': True})
    worksheet.set_row(0, None, header_format)
    
    writer.close()
    return output.getvalue()

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
        # Create two columns for brand/segment/size inputs
        cols = st.sidebar.columns(2)
        
        # Left column
        brand = cols[0].selectbox("Brand", brand_list, index=0)
        segment = cols[0].selectbox("Segment", segment_list, index=0)
        
        # Right column
        size = cols[1].selectbox("Size", size_options, index=0)
        customer_state = cols[1].text_input("Customer", "")
        
        # Add a separator
        st.markdown("---")
        
        # Cost inputs (full width for visibility)
        cols = st.sidebar.columns(2)
        
        case_cost_input = cols[0].number_input("Case Cost ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        case_cost = format_numeric_input(case_cost_input, 2)
        
        bottles_per_case_input = cols[1].number_input("Bottles/Case", min_value=1, value=12, step=1)
        bottles_per_case = int(format_numeric_input(bottles_per_case_input, 0))
        
        # Calculate bottle cost automatically
        if bottles_per_case > 0:
            bottle_cost = case_cost / bottles_per_case
        else:
            bottle_cost = 0.0
        
        st.text(f"Bottle Cost: ${bottle_cost:.2f}")
        
        # Scan and coupon inputs in columns
        cols = st.sidebar.columns(2)
        
        base_scan_input = cols[0].number_input("Base Scan ($)", min_value=0.0, value=0.0, step=0.25, format="%.2f")
        base_scan = format_numeric_input(base_scan_input, 2)
        
        deep_scan_input = cols[1].number_input("Deep Scan ($)", min_value=0.0, value=0.0, step=0.25, format="%.2f")
        deep_scan = format_numeric_input(deep_scan_input, 2)
        
        # Coupon input (full width)
        coupon_input = st.number_input("Coupon ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        coupon = format_numeric_input(coupon_input, 2)
        
        # Add a separator
        st.markdown("---")
        
        # Pricing inputs
        cols = st.sidebar.columns(2)
        
        edlp_price_input = cols[0].number_input("EDLP ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        edlp_price = format_numeric_input(edlp_price_input, 2)
        
        # Empty space in right column to align
        cols[1].text("")
        
        cols = st.sidebar.columns(2)
        
        tpr_base_price_input = cols[0].number_input("TPR Base ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        tpr_base_price = format_numeric_input(tpr_base_price_input, 2)
        
        tpr_deep_price_input = cols[1].number_input("TPR Deep ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        tpr_deep_price = format_numeric_input(tpr_deep_price_input, 2)
        
        cols = st.sidebar.columns(2)
        
        ad_base_price_input = cols[0].number_input("Ad Base ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        ad_base_price = format_numeric_input(ad_base_price_input, 2)
        
        ad_deep_price_input = cols[1].number_input("Ad Deep ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        ad_deep_price = format_numeric_input(ad_deep_price_input, 2)
        
        # Add a separator before ad percentages
        st.markdown("---")
        st.markdown("### Ad Percentages")
        
        cols = st.sidebar.columns(2)
        
        ad_percentage_base_input = cols[0].number_input("% on Ad (Base)", min_value=0, max_value=100, value=0, step=1)
        ad_percentage_base = format_numeric_input(ad_percentage_base_input, 0)
        
        ad_percentage_deep_input = cols[1].number_input("% on Ad (Deep)", min_value=0, max_value=100, value=0, step=1)
        ad_percentage_deep = format_numeric_input(ad_percentage_deep_input, 0)

    # Calculate margins for Everyday Price
    edlp_margins = calculate_margin(edlp_price, bottle_cost, 0, coupon)

    # Calculate margins for other pricing scenarios
    tpr_base_margins = calculate_margin(tpr_base_price, bottle_cost, base_scan, coupon)
    tpr_deep_margins = calculate_margin(tpr_deep_price, bottle_cost, deep_scan, coupon)
    ad_base_margins = calculate_margin(ad_base_price, bottle_cost, base_scan, coupon)
    ad_deep_margins = calculate_margin(ad_deep_price, bottle_cost, deep_scan, coupon)
    
    # Calculate weighted averages
    weighted_base_percent, weighted_base_dollars = calculate_weighted_average(
        tpr_base_price, ad_base_price, bottle_cost, base_scan, coupon, ad_percentage_base)
    
    weighted_deep_percent, weighted_deep_dollars = calculate_weighted_average(
        tpr_deep_price, ad_deep_price, bottle_cost, deep_scan, coupon, ad_percentage_deep)

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
    
    # Add weighted average information if ad percentages are provided
    if ad_percentage_base > 0 or ad_percentage_deep > 0:
        pricing_data["% on Ad"] = ["N/A", f"{ad_percentage_base:.0f}%", f"{ad_percentage_deep:.0f}%", "N/A", "N/A"]
        
        if weighted_base_percent is not None:
            pricing_data["Weighted Avg GM %"] = ["N/A", f"{weighted_base_percent:.1f}%", 
                                               f"{weighted_deep_percent:.1f}%" if weighted_deep_percent is not None else "N/A", 
                                               "N/A", "N/A"]
            
        if weighted_base_dollars is not None:
            pricing_data["Weighted Avg GM $"] = ["N/A", f"${weighted_base_dollars:.2f}", 
                                               f"${weighted_deep_dollars:.2f}" if weighted_deep_dollars is not None else "N/A",
                                               "N/A", "N/A"]

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
            "Ad GM % (With Coupon)": ad_deep_margins["gm_coupon_percent"],
            "% on Ad (Base)": ad_percentage_base,
            "Weighted Avg GM % (Base)": weighted_base_percent if weighted_base_percent is not None else 0,
            "Weighted Avg GM $ (Base)": weighted_base_dollars if weighted_base_dollars is not None else 0,
            "% on Ad (Deep)": ad_percentage_deep,
            "Weighted Avg GM % (Deep)": weighted_deep_percent if weighted_deep_percent is not None else 0,
            "Weighted Avg GM $ (Deep)": weighted_deep_dollars if weighted_deep_dollars is not None else 0
        }
        
        # Create a DataFrame with the scenario data
        new_scenario = pd.DataFrame([scenario_data])
        
        # Append to scan scenarios
        if st.session_state.scan_scenarios.empty:
            st.session_state.scan_scenarios = new_scenario
        else:
            st.session_state.scan_scenarios = pd.concat([st.session_state.scan_scenarios, new_scenario], ignore_index=True)
        
        st.success("Scan scenario saved successfully!")

def create_h1_data_analysis_ui(container):
    """Create the UI for H1 Data Analysis section"""
    container.markdown("<h4 class='compact-section'>H1 Data Analysis</h4>", unsafe_allow_html=True)
    
    # User selects segment and size for analysis
    segment = container.selectbox("Select Segment", segment_list, index=0, key='h1_segment', label_visibility="collapsed")
    
    # Filter size options based on selected segment
    data_segment = segment_mapping.get(segment, segment)
    if data_segment in segment_size_combinations:
        valid_sizes = ["-- Select Size --"] + segment_size_combinations[data_segment]
    else:
        valid_sizes = size_options
        
    size = container.selectbox("Select Size", valid_sizes, index=0, key='h1_size', label_visibility="collapsed")
    
    # Display the index ratios in a colored list if data is available
    if segment != "-- Select Segment --" and size != "-- Select Size --":
        index_ratios, avg_weekly_rsv, threshold = get_index_ratios(segment, size)
        
        if avg_weekly_rsv:
            container.markdown(f"<p style='margin-bottom:5px;'>Average Weekly RSV: <strong>${avg_weekly_rsv/1000000:.2f}M</strong></p>", unsafe_allow_html=True)
            
            if index_ratios is not None:
                # Format date for week labels
                current_year = datetime.datetime.now().year
                week_start_date = datetime.datetime(current_year, 6, 30)
                
                # Create a colored list to display the data
                list_html = "<div style='max-height:400px; overflow-y:auto; background-color:#f8f9fa; padding:8px; border-radius:5px; font-size:0.9rem;'>"
                
                for i, ratio in enumerate(index_ratios):
                    week_end_date = week_start_date + datetime.timedelta(days=6)
                    
                    if i == 0:
                        week_label = f"1: Jun 30-Jul 6"
                    else:
                        week_label = f"{i+1}: {week_start_date.strftime('%b %d')}-{week_end_date.strftime('%b %d')}"
                    
                    # Determine color based on ratio value compared to threshold
                    color = "#333333"  # Default dark gray
                    bg_color = ""
                    if avg_weekly_rsv > 0:
                        if ratio > 1 + threshold:
                            color = "#228B22"  # Forest green
                            bg_color = "background-color:#e6f4ea;"
                        elif ratio < 1 - (0.5 * threshold):
                            color = "#8B0000"  # Deep red
                            bg_color = "background-color:#fae9e8;"
                    
                    list_html += f"<div style='padding:3px 6px; margin-bottom:2px; {bg_color} border-radius:3px;'><span style='display:inline-block; width:110px; font-size:0.8rem;'>{week_label}</span> <span style='font-weight:bold; color:{color};'>{ratio:.0%}</span></div>"
                    
                    week_start_date += datetime.timedelta(days=7)
                    
                list_html += "</div>"
                
                # Display the list
                container.markdown(list_html, unsafe_allow_html=True)
        else:
            container.info("No data available for this selection.")

def create_scenarios_ui():
    """Create the UI for Saved Scan Scenarios tab"""
    st.markdown("<h2>Saved Scan Scenarios</h2>", unsafe_allow_html=True)

    if st.session_state.scan_scenarios.empty:
        st.info("No scan scenarios saved yet. Use the 'Save Scan Scenario' button in the Calculator tab to save scenarios.")
    else:
        # Handle column name compatibility between "Market" and "Customer/State"
        if 'Customer/State' not in st.session_state.scan_scenarios.columns and 'Market' in st.session_state.scan_scenarios.columns:
            st.session_state.scan_scenarios = st.session_state.scan_scenarios.rename(columns={'Market': 'Customer/State'})
        
        # Display the scenarios using Streamlit's dataframe
        st.dataframe(st.session_state.scan_scenarios)
        
        # Export buttons
        col1, col2, col3 = st.columns(3)
        
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
    
    # Add a separator
    st.markdown("<hr class='compact-divider'>", unsafe_allow_html=True)
    
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
        worksheet.set_landscape()         # Set landscape orientation
        worksheet.fit_to_pages(1, 0)      # Fit to 1 page wide, as many pages tall as needed
        worksheet.set_margins(left=0.33, right=0.33, top=0.33, bottom=0.33)  # Set margins

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
            st.image(image, width=600, use_container_width=True)
    except Exception:
        st.warning("Logo 'image.png' not found. Please add it to the same directory as the app.")

    st.markdown("<h1 class='header'>Pricing & Margin Calculator</h1>", unsafe_allow_html=True)

    # Create tabs for Calculator and Scenarios
    tab1, tab2 = st.tabs(["Calculator and Analysis", "Saved Scenarios"])

    with tab1:
        # Create two columns for calculator and H1 data
        calc_col, h1_col = st.columns([3, 1])
        
        with calc_col:
            # Main calculator section
            create_calculator_ui()
        
        with h1_col:
            # Create a container with styling for the H1 Data Analysis
            with st.container():
                st.markdown("<div class='side-panel'>", unsafe_allow_html=True)
                # H1 Data Analysis section to the right
                create_h1_data_analysis_ui(st)
                st.markdown("</div>", unsafe_allow_html=True)

    with tab2:
        # Saved Scenarios section
        create_scenarios_ui()

# Execute the app
if __name__ == "__main__":
    main()