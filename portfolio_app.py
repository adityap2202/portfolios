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
import yfinance as yf
from concurrent.futures import ThreadPoolExecutor, as_completed

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
    """Fetch current market price for a given ISIN using Yahoo Finance."""
    max_retries = 3
    retry_delay = 2  # seconds
    
    for attempt in range(max_retries):
        try:
            # First search for the ticker using Yahoo Finance search API
            search_url = f"https://query2.finance.yahoo.com/v1/finance/search?q={isin}"
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'application/json'
            }
            
            response = requests.get(search_url, headers=headers)
            
            if response.status_code == 429:  # Too Many Requests
                if attempt < max_retries - 1:
                    time.sleep(retry_delay * (attempt + 1))  # Exponential backoff
                    continue
                else:
                    st.warning(f"Rate limited by Yahoo Finance for ISIN {isin}. Please try again later.")
                    return None
            
            if response.status_code != 200:
                st.warning(f"Failed to search for ISIN {isin}. Status code: {response.status_code}")
                return None
            
            data = response.json()
            quotes = data.get('quotes', [])
            
            if not quotes:
                st.warning(f"No quotes found for ISIN {isin}")
                return None
            
            # Look for NSE ticker in the search results
            nse_ticker = None
            for quote in quotes:
                if quote.get('exchange') == 'NSI' and quote.get('symbol', '').endswith('.NS'):
                    nse_ticker = quote.get('symbol')
                    break
            
            if nse_ticker:
                # Get the price using the found ticker
                stock = yf.Ticker(nse_ticker)
                current_price = stock.info.get('regularMarketPrice')
                if current_price:
                    return current_price
            
            # If NSE ticker not found or price not available, try BSE
            for quote in quotes:
                if quote.get('exchange') == 'BSE' and quote.get('symbol', '').endswith('.BO'):
                    bse_ticker = quote.get('symbol')
                    stock = yf.Ticker(bse_ticker)
                    current_price = stock.info.get('regularMarketPrice')
                    if current_price:
                        return current_price
            
            st.warning(f"No price found for ISIN {isin} in either NSE or BSE")
            return None
            
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(retry_delay * (attempt + 1))
                continue
            else:
                st.warning(f"Error fetching price for ISIN {isin}: {str(e)}")
                return None
    
    return None

def get_current_prices(df):
    """Get current prices for all stocks in the dataframe using parallel processing."""
    # Create a cache for prices to avoid duplicate API calls
    price_cache = {}
    
    # Add a new column for current price
    df['Current Price (Rs.)'] = None
    
    # Show a progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Get unique ISINs to avoid duplicate API calls
    unique_isins = df['ISIN'].unique()
    total_isins = len(unique_isins)
    
    # Create a mapping of ISIN to row indices for faster updates
    isin_to_rows = {isin: df[df['ISIN'] == isin].index.tolist() for isin in unique_isins}
    
    # Process ISINs in parallel with reduced number of workers to avoid rate limiting
    with ThreadPoolExecutor(max_workers=3) as executor:
        # Submit all tasks
        future_to_isin = {executor.submit(fetch_current_price, isin): isin for isin in unique_isins}
        
        # Process completed tasks
        completed = 0
        for future in as_completed(future_to_isin):
            isin = future_to_isin[future]
            try:
                current_price = future.result()
                # Update all rows with this ISIN
                for row_idx in isin_to_rows[isin]:
                    df.at[row_idx, 'Current Price (Rs.)'] = current_price
                price_cache[isin] = current_price
            except Exception as e:
                st.warning(f"Error fetching price for {isin}: {str(e)}")
            
            # Update progress
            completed += 1
            progress_bar.progress(completed / total_isins)
            status_text.text(f"Processed {completed} of {total_isins} stocks")
            
            # Add a small delay between batches to avoid rate limiting
            if completed % 5 == 0:  # Every 5 stocks
                time.sleep(1)
    
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