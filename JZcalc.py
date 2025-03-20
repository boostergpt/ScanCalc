import streamlit as st
import pandas as pd
import numpy as np
import base64
from io import BytesIO
import openai
import json
from PIL import Image
import os

# OpenAI API Configuration
# IMPORTANT: Replace this with your actual API key when running the app
OPENAI_API_KEY = "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"  # You'll replace this when running

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

# Create tabs for additional features
tab1, tab2, tab3 = st.tabs(["Everyday Price Metrics", "Historical Data", "AI Insights"])

with tab1:
    st.markdown("<div class='quadrant'>", unsafe_allow_html=True)
    display_margin_metrics(st, "Everyday Price", edlp_price, bottle_cost, 0, coupon)
    st.markdown("</div>", unsafe_allow_html=True)

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

# GenAI Insights Section
with tab3:
    st.subheader("💡 AI-Powered Insights & Action Plan")

    # Initialize OpenAI client
    openai.api_key = OPENAI_API_KEY

    def generate_insights(brand, size, pricing_data, market=""):
        """
        Generate insights and action plan using OpenAI API based on calculator inputs
        """
        # Check if API key has been set
        if openai.api_key == "YOUR_OPENAI_API_KEY_HERE":
            return {
                "market_insights": "Please set your OpenAI API key to enable AI-powered insights.",
                "competitive_analysis": "",
                "action_plan": "",
                "promotion_strategy": ""
            }
        
        # Create context for the AI
        prompt = f"""
        You are an expert beverage industry analyst and pricing strategist. Generate insights and an action plan for:
        
        Brand: {brand}
        Size: {size}
        Market: {market if market else "Not specified"}
        
        Current pricing structure:
        - Everyday Price: ${pricing_data['everyday_price']:.2f} (Margin: {pricing_data['everyday_margin']:.1f}%)
        - TPR with Base Scan: ${pricing_data['tpr_base_price']:.2f} (Margin: {pricing_data['tpr_base_margin']:.1f}%)
        - TPR with Deep Scan: ${pricing_data['tpr_deep_price']:.2f} (Margin: {pricing_data['tpr_deep_margin']:.1f}%)
        - Ad/Feature with Base Scan: ${pricing_data['ad_base_price']:.2f} (Margin: {pricing_data['ad_base_margin']:.1f}%)
        - Ad/Feature with Deep Scan: ${pricing_data['ad_deep_price']:.2f} (Margin: {pricing_data['ad_deep_margin']:.1f}%)
        
        Base Cost: ${pricing_data['bottle_cost']:.2f}
        
        Provide these four sections:
        1. Market Insights: Brief analysis of the current market for this product
        2. Competitive Analysis: How this pricing strategy compares to competitors
        3. Recommended Action Plan: 3-4 bullet point action items
        4. Promotional Strategy: Specific recommendation on which pricing scenario to focus on
        
        Your response should be concise, specific to this brand and size, and focused on actionable insights.
        Format response as JSON with keys: market_insights, competitive_analysis, action_plan, promotion_strategy.
        """
        
        try:
            # Call OpenAI API
            response = openai.chat.completions.create(
                model="gpt-4o",  # Using GPT-4o but can be changed to other models
                messages=[
                    {"role": "system", "content": "You are an AI assistant specializing in beverage industry pricing and strategy."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7
            )
            
            # Extract and parse the response
            content = response.choices[0].message.content
            
            # Extract JSON from the response (in case there's additional text)
            try:
                # Try to parse the entire response as JSON
                insights = json.loads(content)
            except json.JSONDecodeError:
                # If that fails, look for JSON within the text (common with LLM outputs)
                import re
                json_match = re.search(r'```json\n(.*?)\n```', content, re.DOTALL)
                if json_match:
                    insights = json.loads(json_match.group(1))
                else:
                    # Fallback if JSON parsing fails
                    insights = {
                        "market_insights": "Could not generate structured insights. Please check your API key and try again.",
                        "competitive_analysis": "",
                        "action_plan": "",
                        "promotion_strategy": ""
                    }
            
            return insights
            
        except Exception as e:
            # Handle errors gracefully
            return {
                "market_insights": f"Error generating insights: {str(e)}",
                "competitive_analysis": "Please check your API key and internet connection.",
                "action_plan": "",
                "promotion_strategy": ""
            }

    # Button to generate insights
    if st.button("Generate AI Insights", key="generate_insights"):
        # Check if brand is provided
        if not brand:
            st.warning("Please enter a brand name in the sidebar to generate insights")
        else:
            with st.spinner("Generating insights and action plan..."):
                # Gather pricing data to send to the API
                pricing_data = {
                    "everyday_price": edlp_price,
                    "everyday_margin": edlp_margins["gm_percent"],
                    "tpr_base_price": tpr_base_price,
                    "tpr_base_margin": tpr_base_margins["gm_percent"],
                    "tpr_deep_price": tpr_deep_price,
                    "tpr_deep_margin": tpr_deep_margins["gm_percent"],
                    "ad_base_price": ad_base_price,
                    "ad_base_margin": ad_base_margins["gm_percent"],
                    "ad_deep_price": ad_deep_price,
                    "ad_deep_margin": ad_deep_margins["gm_percent"],
                    "bottle_cost": bottle_cost,
                    "ad_base_margin": ad_base_margins["gm_percent"],
                    "ad_deep_price": ad_deep_price,
                    "ad_deep_margin": ad_deep_margins["gm_percent"],
                    "bottle_cost": bottle_cost
                }
                
                # Generate insights
                insights = generate_insights(brand, size, pricing_data, market)
                
                # Display insights in nicely formatted sections
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("### Market Insights")
                    st.markdown(f"{insights['market_insights']}")
                    
                    st.markdown("### Recommended Action Plan")
                    st.markdown(f"{insights['action_plan']}")
                    
                with col2:
                    st.markdown("### Competitive Analysis")
                    st.markdown(f"{insights['competitive_analysis']}")
                    
                    st.markdown("### Promotional Strategy")
                    st.markdown(f"{insights['promotion_strategy']}")
    else:
        st.info("Enter your product details and click 'Generate AI Insights' to receive customized recommendations and market analysis.")

    # Custom Questions Section
    st.markdown("---")
    st.subheader("💬 Ask Questions About Your Strategy")

    def ask_question(question, brand, size, pricing_data, market=""):
        """
        Allow users to ask specific questions about their pricing strategy
        """
        # Check if API key has been set
        if openai.api_key == "YOUR_OPENAI_API_KEY_HERE":
            return "Please set your OpenAI API key to enable this feature."
        
        # Create context for the AI
        prompt = f"""
        You are an expert beverage industry analyst and pricing strategist. Answer the following question specifically for:
        
        Brand: {brand}
        Size: {size}
        Market: {market if market else "Not specified"}
        
        Current pricing structure:
        - Everyday Price: ${pricing_data['everyday_price']:.2f} (Margin: {pricing_data['everyday_margin']:.1f}%)
        - TPR with Base Scan: ${pricing_data['tpr_base_price']:.2f} (Margin: {pricing_data['tpr_base_margin']:.1f}%)
        - TPR with Deep Scan: ${pricing_data['tpr_deep_price']:.2f} (Margin: {pricing_data['tpr_deep_margin']:.1f}%)
        - Ad/Feature with Base Scan: ${pricing_data['ad_base_price']:.2f} (Margin: {pricing_data['ad_base_margin']:.1f}%)
        - Ad/Feature with Deep Scan: ${pricing_data['ad_deep_price']:.2f} (Margin: {pricing_data['ad_deep_margin']:.1f}%)
        
        Base Cost: ${pricing_data['bottle_cost']:.2f}
        
        Question: {question}
        
        Provide a direct, specific answer based on beverage industry expertise. Be helpful, concise, and practical.
        If the question cannot be answered with the information provided, suggest what additional data would be needed.
        """
        
        try:
            # Call OpenAI API
            response = openai.chat.completions.create(
                model="gpt-4o",  # Using GPT-4o but can be changed to other models
                messages=[
                    {"role": "system", "content": "You are an AI assistant specializing in beverage industry pricing and strategy."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7
            )
            
            # Extract the response
            return response.choices[0].message.content
            
        except Exception as e:
            # Handle errors gracefully
            return f"Error generating response: {str(e)}\nPlease check your API key and internet connection."

    # User question input
    user_question = st.text_area("Enter your specific question about pricing strategy:", 
                                placeholder="Example: How does my TPR Deep strategy compare to industry standards? What would be the impact of increasing my base scan by $1?")

    # Button to submit question
    if st.button("Ask Question", key="ask_question"):
        # Check if brand and question are provided
        if not brand:
            st.warning("Please enter a brand name in the sidebar before asking a question")
        elif not user_question:
            st.warning("Please enter a question")
        else:
            with st.spinner("Generating response..."):
                # Gather pricing data to send to the API
                pricing_data = {
                    "everyday_price": edlp_price,
                    "everyday_margin": edlp_margins["gm_percent"],
                    "tpr_base_price": tpr_base_price,
                    "tpr_base_margin": tpr_base_margins["gm_percent"],
                    "tpr_deep_price": tpr_deep_price,
                    "tpr_deep_margin": tpr_deep_margins["gm_percent"],
                    "ad_base_price": ad_base_price,
                    "ad_base_margin": ad_base_margins["gm_percent"],
                    "ad_deep_price": ad_deep_price,
                    "ad_deep_margin": ad_deep_margins["gm_percent"],
                    "bottle_cost": bottle_cost
                }
                
                # Ask the question
                answer = ask_question(user_question, brand, size, pricing_data, market)
                
                # Display the answer
                st.markdown("### Answer")
                st.markdown(answer)