import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
import glob
import re
import requests
import time
from bs4 import BeautifulSoup

def load_portfolio_data(file):
    """Load and process the portfolio data from Excel file."""
    try:
        # Read the Excel file without headers
        df = pd.read_excel(file, header=None)
        
        # Find the row containing column headers (looking for 'Company Name')
        header_row = None
        for idx, row in df.iterrows():
            if row.astype(str).str.contains('Company Name').any():
                header_row = idx
                break
        
        if header_row is None:
            raise ValueError("Could not find column headers in the Excel file")
        
        # Read the Excel file again with the correct header
        df = pd.read_excel(file, header=header_row)
        
        # Clean up the data
        df = df[df['Company Name'].notna()]  # Remove rows without company names
        
        # Convert numeric columns
        for col in ['Balance', 'Rate (Rs.)', 'Value (Rs.)']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Remove rows with NaN values in important columns
        df = df.dropna(subset=['Company Name', 'Balance', 'Rate (Rs.)', 'Value (Rs.)'])
        
        return df
    
    except Exception as e:
        st.error(f"Error in load_portfolio_data: {str(e)}")
        raise e

def get_demat_info(file):
    """Extract demat account information from filename or file content."""
    if isinstance(file, str):
        # If it's a file path, get the filename
        filename = os.path.basename(file)
        name = os.path.splitext(filename)[0]
    else:
        # If it's an uploaded file, get the filename
        name = os.path.splitext(file.name)[0]
    
    # Default values
    dp_id = None
    person_name = None
    
    # Try to extract person name from filename
    # Common patterns: "name demat", "name's demat", etc.
    name_patterns = [
        r"^([A-Za-z\s]+)\s+demat",
        r"^([A-Za-z\s]+)'s\s+demat",
        r"^([A-Za-z\s]+)\s+portfolio"
    ]
    
    for pattern in name_patterns:
        match = re.search(pattern, name, re.IGNORECASE)
        if match:
            person_name = match.group(1).strip()
            break
    
    # If no name found in filename, try to extract from file content
    if not person_name:
        try:
            # Read the file to find name and DP ID
            df = pd.read_excel(file, header=None)
            for idx, row in df.iterrows():
                for col in row:
                    if isinstance(col, str):
                        # Look for DP ID
                        if 'DP ID' in col:
                            dp_id = col.split(':')[1].strip()
                        
                        # Look for name patterns in the content
                        for pattern in name_patterns:
                            match = re.search(pattern, col, re.IGNORECASE)
                            if match:
                                person_name = match.group(1).strip()
                                break
                        
                        if person_name:
                            break
                if person_name:
                    break
        except:
            pass
    
    # If still no name found, use the filename
    if not person_name:
        person_name = name
    
    # Create a dictionary with the extracted information
    info = {
        "person_name": person_name,
        "dp_id": dp_id,
        "display_name": f"{person_name} ({dp_id})" if dp_id else person_name
    }
    
    return info

def fetch_current_price(isin):
    """Fetch current market price for a given ISIN."""
    try:
        # Using NSE India website to fetch current price
        # This is a simplified approach and may need adjustments based on actual API availability
        url = f"https://www.nseindia.com/get-quotes/equity?symbol={isin}"
        
        # Add headers to mimic a browser request
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9',
        }
        
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            # This is a simplified approach - actual implementation would depend on the website structure
            price_element = soup.select_one('.last-price')
            if price_element:
                return float(price_element.text.strip().replace(',', ''))
        
        # If we can't get the price, return None
        return None
    except Exception as e:
        st.warning(f"Could not fetch price for ISIN {isin}: {str(e)}")
        return None

def get_current_prices(df):
    """Get current prices for all stocks in the dataframe."""
    # Create a cache for prices to avoid duplicate API calls
    price_cache = {}
    
    # Add a new column for current price
    df['Current Price (Rs.)'] = None
    
    # Show a progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Process each row
    for i, row in df.iterrows():
        isin = row['ISIN']
        status_text.text(f"Fetching price for {row['Company Name']} ({isin})...")
        
        # Check if we already have this price in the cache
        if isin in price_cache:
            df.at[i, 'Current Price (Rs.)'] = price_cache[isin]
        else:
            # Fetch the current price
            current_price = fetch_current_price(isin)
            
            # Store in cache and dataframe
            price_cache[isin] = current_price
            df.at[i, 'Current Price (Rs.)'] = current_price
            
            # Add a small delay to avoid rate limiting
            time.sleep(0.0)
        
        # Update progress
        progress_bar.progress((i + 1) / len(df))
    
    # Clear the progress indicators
    progress_bar.empty()
    status_text.empty()
    
    # Calculate current value
    df['Current Value (Rs.)'] = df['Balance'] * df['Current Price (Rs.)'].fillna(df['Rate (Rs.)'])
    
    return df

def display_portfolio(df, title, tab_id, show_current_prices=False):
    """Display portfolio information for a given dataframe."""
    if df is None or len(df) == 0:
        st.error(f"No valid data found for {title}")
        return
    
    # Calculate total portfolio value
    total_value = df['Value (Rs.)'].sum()
    
    # Calculate current value if available
    current_value = None
    if 'Current Value (Rs.)' in df.columns:
        current_value = df['Current Value (Rs.)'].sum()
    
    # Top level metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Portfolio Value", f"₹{total_value:,.2f}")
    with col2:
        st.metric("Number of Stocks", len(df))
    with col3:
        st.metric("Average Investment per Stock", f"₹{(total_value/len(df)):,.2f}")
    
    # Show current value if available
    if current_value is not None:
        st.metric("Current Portfolio Value", f"₹{current_value:,.2f}", 
                 f"{(current_value - total_value):,.2f} ({(current_value/total_value - 1)*100:.2f}%)")
    
    # Holdings Table with Sorting
    st.subheader("Holdings Details")
    
    # Select columns to display
    display_cols = ['Company Name', 'Balance', 'Rate (Rs.)', 'Value (Rs.)', 'Scrip Type', 'ISIN']
    if show_current_prices and 'Current Price (Rs.)' in df.columns:
        display_cols.extend(['Current Price (Rs.)', 'Current Value (Rs.)'])
    
    sortable_df = df[display_cols].copy()
    sortable_df = sortable_df.sort_values('Value (Rs.)', ascending=False)
    sortable_df = sortable_df.reset_index(drop=True)  # Reset the index
    
    # Format the numeric columns
    for col in ['Rate (Rs.)', 'Value (Rs.)', 'Current Price (Rs.)', 'Current Value (Rs.)']:
        if col in sortable_df.columns:
            sortable_df[col] = sortable_df[col].apply(lambda x: f"₹{x:,.2f}" if pd.notna(x) else "N/A")
    
    st.dataframe(sortable_df, use_container_width=True)
    
    # Top Holdings Pie Chart
    st.subheader("Top 5 Holdings")
    top_5 = df.nlargest(5, 'Value (Rs.)')
    fig_pie = px.pie(
        top_5,
        values='Value (Rs.)',
        names='Company Name',
        title='Top 5 Holdings Distribution'
    )
    st.plotly_chart(fig_pie, use_container_width=True, key=f"pie_{tab_id}")
    
    # Stock Price Distribution
    st.subheader("Stock Price Distribution")
    fig_hist = px.histogram(
        df,
        x='Rate (Rs.)',
        title='Distribution of Stock Prices',
        nbins=20
    )
    st.plotly_chart(fig_hist, use_container_width=True, key=f"hist_{tab_id}")

def main():
    st.set_page_config(page_title="Portfolio Dashboard", layout="wide")
    
    # Add custom CSS
    st.markdown("""
        <style>
        .main {
            padding: 2rem;
        }
        .stPlotlyChart {
            background-color: #ffffff;
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .upload-text {
            text-align: center;
            padding: 2rem;
            border: 2px dashed #ccc;
            border-radius: 5px;
            margin-bottom: 2rem;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.title("📊 Portfolio Dashboard")
    
    # Add a sidebar for options and file upload
    with st.sidebar:
        st.header("Options")
        show_current_prices = st.checkbox("Show Current Market Prices", value=False)
        if show_current_prices:
            st.info("This will fetch current market prices for all stocks. This may take some time.")
        
        st.header("Upload Portfolio Files")
        uploaded_files = st.file_uploader(
            "Upload your portfolio Excel files",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Upload one or more Excel files containing your portfolio data. " +
                 "File name format: PersonName_DPID.xlsx (e.g., John-Doe_IN123456.xlsx)"
        )
    
    if not uploaded_files:
        st.markdown("""
            <div class="upload-text">
                <h3>👋 Welcome to the Portfolio Dashboard!</h3>
                <p>Please upload your portfolio Excel files using the sidebar to get started.</p>
                <p>Each file should contain the following columns:</p>
                <ul>
                    <li>Company Name</li>
                    <li>Balance (number of shares)</li>
                    <li>Rate (Rs.)</li>
                    <li>Value (Rs.)</li>
                    <li>ISIN</li>
                </ul>
                <p>Optional columns:</p>
                <ul>
                    <li>Scrip Type (defaults to EQ if not provided)</li>
                </ul>
            </div>
        """, unsafe_allow_html=True)
        return
    
    try:
        # Create a dictionary to store dataframes for each demat
        demat_data = {}
        demat_info = {}
        
        # Load data from each uploaded file
        for uploaded_file in uploaded_files:
            info = get_demat_info(uploaded_file)
            demat_name = info["display_name"]
            demat_info[demat_name] = info
            
            try:
                df = load_portfolio_data(uploaded_file)
                demat_data[demat_name] = df
            except Exception as e:
                st.error(f"Error loading {demat_name}: {str(e)}")
        
        if not demat_data:
            st.error("No valid data could be loaded from any of the uploaded files")
            return
        
        # Fetch current prices if requested
        if show_current_prices:
            st.info("Fetching current market prices for all stocks...")
            for demat_name, df in demat_data.items():
                demat_data[demat_name] = get_current_prices(df)
        
        # Create tabs for each demat and a consolidated tab
        tab_names = list(demat_data.keys()) + ["Consolidated"]
        tabs = st.tabs(tab_names)
        
        # Display data for each demat in its own tab
        for i, (demat_name, df) in enumerate(demat_data.items()):
            with tabs[i]:
                info = demat_info[demat_name]
                st.header(f"{info['person_name']}'s Portfolio")
                if info['dp_id']:
                    st.subheader(f"DP ID: {info['dp_id']}")
                display_portfolio(df, demat_name, i, show_current_prices)
        
        # Create consolidated data in the last tab
        with tabs[-1]:
            st.header("Consolidated Portfolio")
            
            # Combine all dataframes
            consolidated_df = pd.concat(demat_data.values(), ignore_index=True)
            
            # Define the aggregation dictionary based on whether current prices are available
            agg_dict = {
                'Balance': 'sum',
                'Rate (Rs.)': 'mean',  # Use mean for rate
                'Value (Rs.)': 'sum',
                'Scrip Type': 'first',  # Take the first value
                'ISIN': 'first'  # Take the first value
            }
            
            # Add current price columns to aggregation if they exist
            if show_current_prices and 'Current Price (Rs.)' in consolidated_df.columns:
                agg_dict.update({
                    'Current Price (Rs.)': 'mean',
                    'Current Value (Rs.)': 'sum'
                })
            
            # Group by Company Name and aggregate
            consolidated_df = consolidated_df.groupby('Company Name').agg(agg_dict).reset_index()
            
            # Display consolidated portfolio
            display_portfolio(consolidated_df, "Consolidated Portfolio", "consolidated", show_current_prices)
        
    except Exception as e:
        st.error(f"Error processing the data: {str(e)}")
        st.info("Please make sure the Excel files are in the correct format.")

if __name__ == "__main__":
    main() 