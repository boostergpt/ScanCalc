import streamlit as st
# Set page config FIRST
st.set_page_config(
    page_title="Margin Calculator",
    page_icon="🧮",
    layout="wide"
)

# Then import other libraries
import pandas as pd
import numpy as np
import base64
from io import BytesIO
import openai
import json
from PIL import Image
import os

# OpenAI API Configuration - Move the st.warning after set_page_config
try:
    OPENAI_API_KEY = st.secrets["api_keys"]["openai"]
except KeyError:
    OPENAI_API_KEY = "YOUR_OPENAI_API_KEY_HERE"  # Fallback value
    st.warning("OpenAI API key is not set in secrets. AI features will not work.")

# Custom CSS for better styling
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stButton button {
        width: 100%;
    }
    .metric-container {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px;
        margin-bottom: 10px;
    }
    .metric-label {
        font-weight: bold;
        font-size: 0.9rem;
    }
    .metric-value {
        font-size: 1.1rem;
        padding: 5px 0;
    }
    .header {
        text-align: center;
        margin-bottom: 20px;
    }
    .centered-logo {
        display: flex;
        justify-content: center;
        align-items: center;
        flex-direction: column;
        margin-bottom: 20px;
    }
    .quadrant {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 15px;
    }
    .quadrant-title {
        font-weight: bold;
        font-size: 1.2rem;
        margin-bottom: 10px;
        text-align: center;
    }
    .edlp-section {
        background-color: #e9ecef;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 25px;
        display: flex;
        flex-direction: row;
        align-items: center;
    }
    .edlp-metric {
        flex: 1;
        margin: 0 10px;
        text-align: center;
    }
    .edlp-value {
        font-size: 1.2rem;
        font-weight: bold;
    }
    .edlp-label {
        font-size: 0.9rem;
        color: #666;
    }
    .save-button {
        margin-top: 20px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for storing historical data and calculated scans
if 'historical_data' not in st.session_state:
    st.session_state.historical_data = pd.DataFrame()
    
if 'calculated_scans' not in st.session_state:
    st.session_state.calculated_scans = pd.DataFrame()
    
if 'user_question' not in st.session_state:
    st.session_state.user_question = ""

# Add logo and title with better alignment
try:
    st.markdown("""
    <div style="display: flex; align-items: center; justify-content: center; margin-bottom: 20px;">
        <div style="text-align: center;">
            <img src="data:image/png;base64,{}" width="500">
            <h1 style="margin-top: 10px;">Pricing & Margin Calculator</h1>
        </div>
    </div>
    """.format(base64.b64encode(open('image.png', 'rb').read()).decode()), unsafe_allow_html=True)
except Exception as e:
    st.error(f"Error loading image: {str(e)}")
    st.warning("Logo 'image.png' not found. Make sure it's in the same directory as this app.")
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

# Create a current data row in the requested format
current_data = {
    "Brand": brand,
    "Size": size,
    "Case Cost": case_cost,
    "# Bottles/Cs": bottles_per_case,
    "Bottle Cost": bottle_cost,
    "Base Scan": base_scan,
    "Deep Scan": deep_scan,
    "Coupon": coupon,
    "Everyday Shelf Price": edlp_price,
    "Everyday GM %": edlp_margins["gm_percent"] / 100,  # Convert to decimal for matching format
    "Everyday GM $": edlp_margins["gm_dollars"],
    "Everyday GM % (With Coupon)": edlp_margins["gm_coupon_percent"] / 100,  # Convert to decimal
    "TPR Price (Base Scan)": tpr_base_price,
    "TPR GM % (Base Scan)": tpr_base_margins["gm_percent"] / 100,  # Convert to decimal
    "TPR GM $ (Base Scan)": tpr_base_margins["gm_dollars"],
    "TPR GM % (With Coupon)": tpr_base_margins["gm_coupon_percent"] / 100,  # Convert to decimal
    "TPR Price (Deep Scan)": tpr_deep_price,
    "TPR GM % (Deep Scan)": tpr_deep_margins["gm_percent"] / 100,  # Convert to decimal
    "TPR GM $ (Deep Scan)": tpr_deep_margins["gm_dollars"],
    "TPR GM % (With Coupon)": tpr_deep_margins["gm_coupon_percent"] / 100,  # Convert to decimal
    "Ad/Feature Price (Base Scan)": ad_base_price,
    "Ad GM % (Base Scan)": ad_base_margins["gm_percent"] / 100,  # Convert to decimal
    "Ad GM $ (Base Scan)": ad_base_margins["gm_dollars"],
    "Ad GM % (With Coupon)": ad_base_margins["gm_coupon_percent"] / 100,  # Convert to decimal
    "Ad/Feature Price (Deep Scan)": ad_deep_price,
    "Ad GM % (Deep Scan)": ad_deep_margins["gm_percent"] / 100,  # Convert to decimal
    "Ad GM $ (Deep Scan)": ad_deep_margins["gm_dollars"],
    "Ad GM % (With Coupon)": ad_deep_margins["gm_coupon_percent"] / 100  # Convert to decimal
}

# EDLP Metrics section at the top as a single row
st.markdown("<div class='edlp-section'>", unsafe_allow_html=True)

# EDLP Price
st.markdown(f"""
<div class='edlp-metric'>
    <div class='edlp-label'>EDLP Price</div>
    <div class='edlp-value'>${edlp_price:.2f}</div>
</div>
""", unsafe_allow_html=True)

# EDLP Metrics section at the top as a single horizontal row
st.markdown("""
<div style="display: flex; flex-direction: row; background-color: #e9ecef; border-radius: 10px; padding: 15px; margin-bottom: 25px;">
    <div style="flex: 1; text-align: center; margin: 0 10px;">
        <div style="font-size: 0.9rem; color: #666;">EDLP Price</div>
        <div style="font-size: 1.2rem; font-weight: bold;">${:.2f}</div>
    </div>
    <div style="flex: 1; text-align: center; margin: 0 10px;">
        <div style="font-size: 0.9rem; color: #666;">EDLP Margin %</div>
        <div style="font-size: 1.2rem; font-weight: bold;">{:.1f}%</div>
    </div>
    <div style="flex: 1; text-align: center; margin: 0 10px;">
        <div style="font-size: 0.9rem; color: #666;">EDLP Margin $</div>
        <div style="font-size: 1.2rem; font-weight: bold;">${:.2f}</div>
    </div>
    <div style="flex: 1; text-align: center; margin: 0 10px;">
        <div style="font-size: 0.9rem; color: #666;">With Coupon %</div>
        <div style="font-size: 1.2rem; font-weight: bold;">{:.1f}%</div>
    </div>
    <div style="flex: 1; text-align: center; margin: 0 10px;">
        <div style="font-size: 0.9rem; color: #666;">With Coupon $</div>
        <div style="font-size: 1.2rem; font-weight: bold;">${:.2f}</div>
    </div>
</div>
""".format(
    edlp_price, 
    edlp_margins["gm_percent"],
    edlp_margins["gm_dollars"],
    edlp_margins["gm_coupon_percent"],
    edlp_margins["gm_coupon_dollars"]
), unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)

# Create the four quadrants layout
col1, col2 = st.columns(2)
col3, col4 = st.columns(2)

# Display metrics in each quadrant
with col1:
    st.markdown("<div class='quadrant'>", unsafe_allow_html=True)
    display_margin_metrics(col1, "TPR (Base Scan)", tpr_base_price, bottle_cost, base_scan, coupon)
    st.markdown("</div>", unsafe_allow_html=True)

with col2:
    st.markdown("<div class='quadrant'>", unsafe_allow_html=True)
    display_margin_metrics(col2, "Ad/Feature (Base Scan)", ad_base_price, bottle_cost, base_scan, coupon)
    st.markdown("</div>", unsafe_allow_html=True)

with col3:
    st.markdown("<div class='quadrant'>", unsafe_allow_html=True)
    display_margin_metrics(col3, "TPR (Deep Scan)", tpr_deep_price, bottle_cost, deep_scan, coupon)
    st.markdown("</div>", unsafe_allow_html=True)

with col4:
    st.markdown("<div class='quadrant'>", unsafe_allow_html=True)
    display_margin_metrics(col4, "Ad/Feature (Deep Scan)", ad_deep_price, bottle_cost, deep_scan, coupon)
    st.markdown("</div>", unsafe_allow_html=True)

# Save button in the center
st.markdown('<div class="save-button">', unsafe_allow_html=True)
if st.button("Save", key="save_button"):
    # Create a new data frame from the current data
    new_data = pd.DataFrame([current_data])
    
    # Add to calculated scans
    if st.session_state.calculated_scans.empty:
        st.session_state.calculated_scans = new_data
    else:
        st.session_state.calculated_scans = pd.concat([st.session_state.calculated_scans, new_data], ignore_index=True)
    
    st.success("Data saved to Calculated Scans!")
st.markdown('</div>', unsafe_allow_html=True)
    
# Create tabs for additional features - now without EDLP tab
tab1, tab2 = st.tabs(["Calculated Scans", "Historical Data"])

with tab1:
    st.subheader("Calculated Scans")
    
    # Display the data in the required format
    if not st.session_state.calculated_scans.empty:
        st.dataframe(st.session_state.calculated_scans)
        
        # Export to Excel functionality
        if st.button("Export Calculated Scans to Excel"):
            # Function to create a downloadable Excel file
            def to_excel(df):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Calculated Scans', index=False)
                
                # Get the workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Calculated Scans']
                
                # Set column width and formats
                format_currency = workbook.add_format({'num_format': '$#,##0.00'})
                format_percent = workbook.add_format({'num_format': '0.00%'})
                
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
            
            excel_file = to_excel(st.session_state.calculated_scans)
            b64 = base64.b64encode(excel_file).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="calculated_scans.xlsx">Download Excel File</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("Export complete! Click the link above to download.")
            
        if st.button("Clear Calculated Scans"):
            st.session_state.calculated_scans = pd.DataFrame()
            st.success("Calculated scans cleared!")
    else:
        st.info("No data in Calculated Scans yet. Use the 'Save' button to add data.")

with tab2:
    st.subheader("Historical Data")
    
    # Add current data to historical dataframe
    if st.button("Save Current Data to History"):
        # Create a new data frame from the current data
        new_data = pd.DataFrame([current_data])
        
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
            format_percent = workbook.add_format({'num_format': '0.00%'})
            
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
        st.info("No historical data saved yet. Use the 'Save Current Data to History' button to start building your dataset.")



# Custom Questions Section
 
# Add some instructions
with st.expander("How to use this calculator"):
    st.write("""
    1. Enter product information in the sidebar (Brand, Size, Case Cost, etc.)
    2. Enter promotion information (Base Scan, Deep Scan, Coupon)
    3. Enter pricing information for each scenario
    4. The calculator will automatically show the margin metrics for each scenario
    5. Click the 'Save' button to save current data to the Calculated Scans tab
    6. Use the 'Save Current Data to History' button in the Historical Data tab to keep a record of previous calculations
    7. Use the AI Insights tab to get customized recommendations and answer specific questions
    
    **Key Terminology:**
    - **Base Scan**: Standard scan amount applied to the product
    - **Deep Scan**: Enhanced scan amount for deeper promotions
    - **TPR**: Temporary Price Reduction
    - **EDLP**: Everyday Low Price
    - **GM**: Gross Margin
    """)

# Note about API key for users
st.caption("Note: AI features require an OpenAI API key. Make sure to replace the placeholder in the code with your actual key.")
