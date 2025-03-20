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
    layout="wide"
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
    .stTextInput input, .stNumberInput input, .stSelectbox > div > div[role="listbox"] {
        font-size: 1.2rem !important;
    }
    .stTextInput label, .stNumberInput label, .stSelectbox label {
        font-size: 1.2rem !important;
        font-weight: 500 !important;
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
</style>
""", unsafe_allow_html=True)

# Initialize session state for storing historical data
if 'historical_data' not in st.session_state:
    st.session_state.historical_data = pd.DataFrame()

# Add logo and title to the main area
try:
    image = Image.open('image.png')
    col1, col2, col3 = st.columns([1, 3, 1])
    with col2:
        st.image(image, width=600, use_column_width=True)
except:
    st.warning("Logo 'image.png' not found. Please add it to the same directory as the app.")

st.markdown("<h1 class='header'>Pricing & Margin Calculator</h1>", unsafe_allow_html=True)

# Create sidebar for inputs
with st.sidebar:
    st.subheader("Product Information")
    
    # Brand and Size inputs
    brand = st.text_input("Brand")
    size_options = ["1.75L", "1L", "750mL", "375mL", "355mL", "200mL", "100mL", "50mL"]
    size = st.selectbox("Size", size_options)
    
    st.subheader("Cost Information")
    
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
    st.subheader("Promotions")
    base_scan = st.number_input("Base Scan ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    deep_scan = st.number_input("Deep Scan ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    coupon = st.number_input("Coupon ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    
    # Pricing inputs
    st.subheader("Pricing")
    edlp_price = st.number_input("Everyday Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    tpr_base_price = st.number_input("TPR Base Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    tpr_deep_price = st.number_input("TPR Deep Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    ad_base_price = st.number_input("Ad/Feature Base Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    ad_deep_price = st.number_input("Ad/Feature Deep Scan Price ($)", min_value=0.0, value=0.0, step=0.01, format="%.2f")
    
    # Optional market input for AI insights
    market = st.text_input("Market (State/Region)", "")

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

# Function to display margin metrics in a consistent format
def display_margin_metrics(container, title, price, cost, scan, coupon):
    margin_data = calculate_margin(price, cost, scan, coupon)
    
    with container:
        st.markdown(f"<div class='quadrant-title'>{title}</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-label'>Price</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>${price:.2f}</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-label'>Gross Margin %</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>{margin_data['gm_percent']:.1f}%</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-label'>Gross Margin $</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>${margin_data['gm_dollars']:.2f}</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-label'>With Coupon %</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>{margin_data['gm_coupon_percent']:.1f}%</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='metric-container'>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-label'>With Coupon $</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>${margin_data['gm_coupon_dollars']:.2f}</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        return margin_data

# Calculate margins
edlp_margins = calculate_margin(edlp_price, bottle_cost, 0, coupon)
tpr_base_margins = calculate_margin(tpr_base_price, bottle_cost, base_scan, coupon)
tpr_deep_margins = calculate_margin(tpr_deep_price, bottle_cost, deep_scan, coupon)
ad_base_margins = calculate_margin(ad_base_price, bottle_cost, base_scan, coupon)
ad_deep_margins = calculate_margin(ad_deep_price, bottle_cost, deep_scan, coupon)

# Create data for the pricing comparison table
pricing_data = {
    "Pricing Scenario": ["TPR (Base Scan)", "TPR (Deep Scan)", "Ad/Feature (Base Scan)", "Ad/Feature (Deep Scan)"],
    "Price": [f"${tpr_base_price:.2f}", f"${tpr_deep_price:.2f}", f"${ad_base_price:.2f}", f"${ad_deep_price:.2f}"],
    "Gross Margin %": [f"{tpr_base_margins['gm_percent']:.1f}%", f"{tpr_deep_margins['gm_percent']:.1f}%", 
                      f"{ad_base_margins['gm_percent']:.1f}%", f"{ad_deep_margins['gm_percent']:.1f}%"],
    "Gross Margin $": [f"${tpr_base_margins['gm_dollars']:.2f}", f"${tpr_deep_margins['gm_dollars']:.2f}", 
                      f"${ad_base_margins['gm_dollars']:.2f}", f"${ad_deep_margins['gm_dollars']:.2f}"],
    "With Coupon %": [f"{tpr_base_margins['gm_coupon_percent']:.1f}%", f"{tpr_deep_margins['gm_coupon_percent']:.1f}%", 
                     f"{ad_base_margins['gm_coupon_percent']:.1f}%", f"{ad_deep_margins['gm_coupon_percent']:.1f}%"],
    "With Coupon $": [f"${tpr_base_margins['gm_coupon_dollars']:.2f}", f"${tpr_deep_margins['gm_coupon_dollars']:.2f}", 
                     f"${ad_base_margins['gm_coupon_dollars']:.2f}", f"${ad_deep_margins['gm_coupon_dollars']:.2f}"]
}

# Display the table with custom styling
st.markdown("""
<style>
.pricing-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 1.2rem;
    margin-bottom: 20px;
}
.pricing-table th {
    background-color: #f0f2f6;
    padding: 12px;
    text-align: left;
    font-weight: bold;
    border: 1px solid #ddd;
}
.pricing-table td {
    padding: 12px;
    border: 1px solid #ddd;
}
.pricing-table tr:nth-child(even) {
    background-color: #f8f9fa;
}
.pricing-table tr:hover {
    background-color: #e9ecef;
}
</style>
""", unsafe_allow_html=True)

# Create the HTML table
table_html = "<table class='pricing-table'><thead><tr>"
# Add headers
for col in pricing_data.keys():
    table_html += f"<th>{col}</th>"
table_html += "</tr></thead><tbody>"

# Add rows
for i in range(len(pricing_data["Pricing Scenario"])):
    table_html += "<tr>"
    for col in pricing_data.keys():
        table_html += f"<td>{pricing_data[col][i]}</td>"
    table_html += "</tr>"
table_html += "</tbody></table>"

# Display the table
st.markdown(table_html, unsafe_allow_html=True)

# Create tabs for additional features
tab1, tab2 = st.tabs(["Everyday Price Metrics", "Historical Data"])

with tab1:
    # Add Everyday Price to the table
    everyday_data = {
        "Pricing Scenario": ["Everyday Price"],
        "Price": [f"${edlp_price:.2f}"],
        "Gross Margin %": [f"{edlp_margins['gm_percent']:.1f}%"],
        "Gross Margin $": [f"${edlp_margins['gm_dollars']:.2f}"],
        "With Coupon %": [f"{edlp_margins['gm_coupon_percent']:.1f}%"],
        "With Coupon $": [f"${edlp_margins['gm_coupon_dollars']:.2f}"]
    }
    
    # Create HTML table for Everyday Price
    ep_table_html = "<table class='pricing-table'><thead><tr>"
    # Add headers
    for col in everyday_data.keys():
        ep_table_html += f"<th>{col}</th>"
    ep_table_html += "</tr></thead><tbody><tr>"
    # Add the single row
    for col in everyday_data.keys():
        ep_table_html += f"<td>{everyday_data[col][0]}</td>"
    ep_table_html += "</tr></tbody></table>"
    
    # Display the Everyday Price table
    st.markdown(ep_table_html, unsafe_allow_html=True)

with tab2:
    st.subheader("Historical Data")
    
    # Add current data to historical dataframe
    if st.button("Save Current Data"):
        # Create a dataframe with all the data
        data = {
            "Brand": [brand],
            "Size": [size],
            "Case Cost": [case_cost],
            "# Bottles/Cs": [bottles_per_case],
            "Bottle Cost": [bottle_cost],
            "Base Scan": [base_scan],
            "Deep Scan": [deep_scan],
            "Coupon": [coupon],
            "Everyday Shelf Price": [edlp_price],
            "Everyday GM %": [edlp_margins["gm_percent"]],
            "Everyday GM $": [edlp_margins["gm_dollars"]],
            "Everyday GM % (With Coupon)": [edlp_margins["gm_coupon_percent"]],
            "TPR Price (Base Scan)": [tpr_base_price],
            "TPR GM % (Base Scan)": [tpr_base_margins["gm_percent"]],
            "TPR GM $ (Base Scan)": [tpr_base_margins["gm_dollars"]],
            "TPR GM % (With Coupon)": [tpr_base_margins["gm_coupon_percent"]],
            "TPR Price (Deep Scan)": [tpr_deep_price],
            "TPR GM % (Deep Scan)": [tpr_deep_margins["gm_percent"]],
            "TPR GM $ (Deep Scan)": [tpr_deep_margins["gm_dollars"]],
            "TPR GM % (With Coupon)": [tpr_deep_margins["gm_coupon_percent"]],
            "Ad/Feature Price (Base Scan)": [ad_base_price],
            "Ad GM % (Base Scan)": [ad_base_margins["gm_percent"]],
            "Ad GM $ (Base Scan)": [ad_base_margins["gm_dollars"]],
            "Ad GM % (With Coupon)": [ad_base_margins["gm_coupon_percent"]],
            "Ad/Feature Price (Deep Scan)": [ad_deep_price],
            "Ad GM % (Deep Scan)": [ad_deep_margins["gm_percent"]],
            "Ad GM $ (Deep Scan)": [ad_deep_margins["gm_dollars"]],
            "Ad GM % (With Coupon)": [ad_deep_margins["gm_coupon_percent"]]
        }
        
        new_data = pd.DataFrame(data)
        
        # Append to historical data
        if st.session_state.historical_data.empty:
            st.session_state.historical_data = new_data
        else:
            st.session_state.historical_data = pd.concat([st.session_state.historical_data, new_data], ignore_index=True)
        
        st.success("Data saved to historical record!")
    
    # Export button for historical data
    if not st.session_state.historical_data.empty:
        st.dataframe(st.session_state.historical_data)
        
        # Function to create a downloadable Excel file
        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Calculator Output', index=False)
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Calculator Output']
            
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
        
        if st.button("Export Historical Data to Excel"):
            excel_file = to_excel(st.session_state.historical_data)
            b64 = base64.b64encode(excel_file).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="margin_calculator_history.xlsx">Download Excel File</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("Export complete! Click the link above to download.")
        
        if st.button("Clear Historical Data"):
            st.session_state.historical_data = pd.DataFrame()
            st.success("Historical data cleared!")
    else:
        st.info("No historical data saved yet. Use the 'Save Current Data' button to start building your dataset.")