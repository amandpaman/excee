import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Set page configuration
st.set_page_config(
    page_title="Telecom Data Analyzer",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main {
        padding-top: 1rem;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #007bff;
        margin: 0.5rem 0;
    }
    .filter-section {
        background-color: #ffffff;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #dee2e6;
        margin-bottom: 1rem;
    }
    .status-green { color: #28a745; font-weight: bold; }
    .status-red { color: #dc3545; font-weight: bold; }
    .status-yellow { color: #ffc107; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# Helper functions
@st.cache_data
def load_excel_data(uploaded_file):
    """Load and cache Excel data with multi-level headers"""
    try:
        # Try to read with multi-level headers first
        excel_file = pd.ExcelFile(uploaded_file)
        sheets = {}
        
        for sheet_name in excel_file.sheet_names:
            try:
                # Try reading with multi-level headers
                df_multi = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=[0, 1])
                
                # Check if it's actually multi-level
                if any(isinstance(col, tuple) for col in df_multi.columns):
                    # Flatten multi-level columns by combining level 0 and level 1
                    df_multi.columns = [f"{col[0]}_{col[1]}" if col[1] not in ['Unnamed', 'nan', ''] and not pd.isna(col[1]) 
                                      else col[0] for col in df_multi.columns]
                    sheets[sheet_name] = df_multi
                else:
                    # Fall back to single header
                    df_single = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                    sheets[sheet_name] = df_single
                    
            except Exception as e:
                # Fall back to single header if multi-level fails
                try:
                    df_single = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                    sheets[sheet_name] = df_single
                except Exception as e2:
                    st.error(f"Error reading sheet {sheet_name}: {e2}")
                    continue
                    
        return sheets
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def clean_column_names(df):
    """Clean column names and map common variations"""
    df = df.copy()
    
    # Remove extra spaces and standardize
    df.columns = df.columns.str.strip()
    
    # Create column mapping for common variations
    column_mapping = {}
    
    for col in df.columns:
        clean_col = str(col).strip().upper()
        
        # Map variations to standard names
        if any(x in clean_col for x in ['CIRCUIT', 'ID']):
            if 'CIRCUIT' in clean_col and 'ID' in clean_col:
                column_mapping[col] = 'CIRCUIT_ID'
        elif any(x in clean_col for x in ['CUSTOMER', 'NAME']):
            if 'CUSTOMER' in clean_col and 'NAME' in clean_col:
                column_mapping[col] = 'CUSTOMER_NAME'
        elif 'LATITUDE' in clean_col or 'LAT' in clean_col:
            column_mapping[col] = 'LATITUDE'
        elif 'LONGITUDE' in clean_col or 'LON' in clean_col:
            column_mapping[col] = 'LONGITUDE'
        elif 'ZONE' in clean_col:
            column_mapping[col] = 'ZONE'
        elif 'CIRCLE' in clean_col and 'END' in clean_col:
            column_mapping[col] = 'CIRCLE'
        elif 'STATUS' in clean_col and 'CONNECT' in clean_col:
            column_mapping[col] = 'CONNECTION_STATUS'
        elif 'RAG' in clean_col and 'STATUS' in clean_col:
            column_mapping[col] = 'RAG_STATUS'
        elif 'AGING' in clean_col:
            column_mapping[col] = 'AGING'
        elif 'CAPEX' in clean_col:
            if 'LM' in clean_col or 'LAST' in clean_col:
                column_mapping[col] = 'LM_CAPEX'
            elif 'NETWORK' in clean_col:
                column_mapping[col] = 'NETWORK_CAPEX'
            elif 'CPE' in clean_col:
                column_mapping[col] = 'CPE_CAPEX'
        elif 'OTC' in clean_col:
            column_mapping[col] = 'OTC'
        elif 'ARC' in clean_col:
            column_mapping[col] = 'ARC'
    
    # Rename columns
    df = df.rename(columns=column_mapping)
    
    return df

def get_column_categories(df):
    """Categorize columns based on their names"""
    categories = {
        'Order Information': [],
        'Location Information': [],
        'Commercial Information': [],
        'Technical Information': [],
        'Status Information': [],
        'Timeline Information': [],
        'Other': []
    }
    
    for col in df.columns:
        col_upper = str(col).upper()
        
        if any(x in col_upper for x in ['CIRCUIT', 'OSM', 'FR', 'ORDER']):
            categories['Order Information'].append(col)
        elif any(x in col_upper for x in ['LATITUDE', 'LONGITUDE', 'ADDRESS', 'AREA', 'DISTRICT', 'ZONE']):
            categories['Location Information'].append(col)
        elif any(x in col_upper for x in ['OTC', 'ARC', 'CAPEX', 'VALUE', 'COST']):
            categories['Commercial Information'].append(col)
        elif any(x in col_upper for x in ['DEVICE', 'EQUIPMENT', 'ROUTER', 'CISCO', 'PROVIDER', 'MEDIA']):
            categories['Technical Information'].append(col)
        elif any(x in col_upper for x in ['STATUS', 'RAG', 'CONNECTED']):
            categories['Status Information'].append(col)
        elif any(x in col_upper for x in ['DATE', 'AGING', 'DAYS', 'TIMELINE']):
            categories['Timeline Information'].append(col)
        else:
            categories['Other'].append(col)
    
    # Remove empty categories
    categories = {k: v for k, v in categories.items() if v}
    
    return categories

def clean_numeric_columns(df, columns):
    """Clean and convert columns to numeric"""
    for col in columns:
        if col in df.columns:
            # Handle text values that should be numeric
            df[col] = df[col].astype(str).str.replace(',', '').str.replace('₹', '').str.replace('Rs', '')
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def parse_dates(df, date_columns):
    """Parse date columns"""
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def get_safe_column_values(df, column_name, default_value='N/A'):
    """Safely get unique values from a column"""
    if column_name in df.columns:
        values = df[column_name].dropna().unique().tolist()
        # Handle different data types
        try:
            return sorted([str(v) for v in values if v is not None])
        except:
            return [str(v) for v in values if v is not None]
    else:
        return [default_value]

def find_best_column_match(df, search_terms):
    """Find the best matching column for given search terms"""
    for col in df.columns:
        col_upper = str(col).upper()
        if any(term.upper() in col_upper for term in search_terms):
            return col
    return None

# Main App
def main():
    st.title("📡 Advanced Telecom Data Analyzer")
    st.markdown("*Designed for Multi-Level Excel Data Structure*")
    st.markdown("---")
    
    # Sidebar for file upload
    with st.sidebar:
        st.header("📂 Data Upload")
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls', 'xlsm'],
            help="Upload your telecom data Excel file with hierarchical columns"
        )
        
        if uploaded_file is not None:
            sheets = load_excel_data(uploaded_file)
            
            if sheets:
                # Sheet selection
                sheet_names = list(sheets.keys())
                selected_sheet = st.selectbox("Select Sheet", sheet_names)
                df_raw = sheets[selected_sheet].copy()
                
                # Clean column names
                df = clean_column_names(df_raw)
                
                st.success(f"✅ Loaded {len(df)} records from '{selected_sheet}'")
                st.info(f"📊 {len(df.columns)} columns available")
            else:
                st.error("❌ Failed to load Excel file")
                return
        else:
            st.info("👆 Please upload an Excel file to begin analysis")
            return
    
    # Data preprocessing
    with st.spinner("Processing data..."):
        # Show original vs cleaned column structure
        with st.expander("📋 Column Structure Analysis", expanded=False):
            st.write("**Original Columns:**")
            st.write(df_raw.columns.tolist())
            
            st.write("**Cleaned/Mapped Columns:**")
            st.write(df.columns.tolist())
            
            # Show column categories
            categories = get_column_categories(df)
            st.write("**Column Categories:**")
            for category, cols in categories.items():
                if cols:
                    st.write(f"**{category}:** {', '.join(cols)}")
        
        # Identify and clean numeric columns
        potential_numeric_cols = []
        for col in df.columns:
            col_upper = str(col).upper()
            if any(x in col_upper for x in ['CAPEX', 'OTC', 'ARC', 'VALUE', 'COST', 'AGING', 'DAYS', 'LATITUDE', 'LONGITUDE']):
                potential_numeric_cols.append(col)
        
        df = clean_numeric_columns(df, potential_numeric_cols)
        
        # Parse potential date columns
        potential_date_cols = []
        for col in df.columns:
            col_upper = str(col).upper()
            if any(x in col_upper for x in ['DATE', 'TIME']):
                potential_date_cols.append(col)
        
        df = parse_dates(df, potential_date_cols)
    
    # Dynamic column detection
    zone_col = find_best_column_match(df, ['ZONE'])
    circle_col = find_best_column_match(df, ['CIRCLE', 'A END CIRCLE'])
    status_col = find_best_column_match(df, ['CONNECTED STATUS', 'CONNECTION STATUS', 'STATUS'])
    customer_col = find_best_column_match(df, ['CUSTOMER NAME', 'CUSTOMER'])
    lat_col = find_best_column_match(df, ['LATITUDE', 'LAT'])
    lon_col = find_best_column_match(df, ['LONGITUDE', 'LON', 'LONG'])
    aging_col = find_best_column_match(df, ['OVERALL AGING', 'AGING'])
    rag_col = find_best_column_match(df, ['RAG STATUS', 'RAG'])
    
    # Main dashboard
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Overview", "🗺️ Geographic", "💰 Financial", "⏱️ Timeline", "🔧 Technical", "📈 Advanced Analytics"
    ])
    
    # Tab 1: Overview Dashboard
    with tab1:
        st.header("📊 Executive Dashboard")
        
        # Dynamic Filters
        with st.expander("🔍 Dynamic Filters", expanded=True):
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                if zone_col:
                    zones = ['All'] + get_safe_column_values(df, zone_col)
                    selected_zone = st.selectbox(f"Zone ({zone_col})", zones)
                else:
                    selected_zone = 'All'
                    st.write("❌ Zone column not detected")
                
            with col2:
                if circle_col:
                    circles = ['All'] + get_safe_column_values(df, circle_col)
                    selected_circle = st.selectbox(f"Circle ({circle_col})", circles)
                else:
                    selected_circle = 'All'
                    st.write("❌ Circle column not detected")
                
            with col3:
                if status_col:
                    statuses = ['All'] + get_safe_column_values(df, status_col)
                    selected_status = st.selectbox(f"Status ({status_col})", statuses)
                else:
                    selected_status = 'All'
                    st.write("❌ Status column not detected")
                
            with col4:
                if customer_col:
                    customers = ['All'] + get_safe_column_values(df, customer_col)[:20]  # Limit for performance
                    selected_customer = st.selectbox(f"Customer ({customer_col})", customers)
                else:
                    selected_customer = 'All'
                    st.write("❌ Customer column not detected")
        
        # Apply filters
        filtered_df = df.copy()
        if selected_zone != 'All' and zone_col:
            filtered_df = filtered_df[filtered_df[zone_col].astype(str) == selected_zone]
        if selected_circle != 'All' and circle_col:
            filtered_df = filtered_df[filtered_df[circle_col].astype(str) == selected_circle]
        if selected_status != 'All' and status_col:
            filtered_df = filtered_df[filtered_df[status_col].astype(str) == selected_status]
        if selected_customer != 'All' and customer_col:
            filtered_df = filtered_df[filtered_df[customer_col].astype(str) == selected_customer]
        
        # Key Metrics
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            total_orders = len(filtered_df)
            st.metric("Total Orders", f"{total_orders:,}")
        
        with col2:
            if status_col:
                connected = len(filtered_df[filtered_df[status_col].astype(str).str.contains('Connected', case=False, na=False)])
                connection_rate = (connected / total_orders * 100) if total_orders > 0 else 0
                st.metric("Connected", f"{connected:,}", f"{connection_rate:.1f}%")
            else:
                st.metric("Connected", "N/A", "Status col missing")
        
        with col3:
            # Find CAPEX columns
            capex_cols = [col for col in df.columns if 'CAPEX' in str(col).upper()]
            if capex_cols:
                total_capex = filtered_df[capex_cols].sum().sum()
                st.metric("Total CAPEX", f"₹{total_capex:,.0f}")
            else:
                st.metric("Total CAPEX", "N/A", "CAPEX cols missing")
        
        with col4:
            if aging_col:
                avg_aging = filtered_df[aging_col].mean()
                st.metric("Avg Aging (Days)", f"{avg_aging:.0f}" if not pd.isna(avg_aging) else "N/A")
            else:
                st.metric("Avg Aging", "N/A", "Aging col missing")
        
        with col5:
            if status_col:
                on_hold = len(filtered_df[filtered_df[status_col].astype(str).str.contains('Hold', case=False, na=False)])
                st.metric("On Hold", f"{on_hold:,}")
            else:
                st.metric("On Hold", "N/A", "Status col missing")
        
        # Charts
        col1, col2 = st.columns(2)
        
        with col1:
            # Status Distribution
            if status_col:
                status_counts = filtered_df[status_col].value_counts()
                fig_status = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title=f"Status Distribution ({status_col})",
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                st.plotly_chart(fig_status, use_container_width=True)
            else:
                st.warning("Status column not found for visualization")
        
        with col2:
            # Zone-wise Orders
            if zone_col:
                zone_counts = filtered_df[zone_col].value_counts()
                fig_zone = px.bar(
                    x=zone_counts.index,
                    y=zone_counts.values,
                    title=f"Orders by Zone ({zone_col})",
                    labels={'x': 'Zone', 'y': 'Order Count'},
                    color=zone_counts.values,
                    color_continuous_scale='Blues'
                )
                st.plotly_chart(fig_zone, use_container_width=True)
            else:
                st.warning("Zone column not found for visualization")
        
        # Additional Charts
        col1, col2 = st.columns(2)
        
        with col1:
            if rag_col:
                rag_counts = filtered_df[rag_col].value_counts()
                fig_rag = px.bar(
                    x=rag_counts.index,
                    y=rag_counts.values,
                    title=f"RAG Status Distribution ({rag_col})",
                    labels={'x': 'RAG Status', 'y': 'Count'},
                    color=rag_counts.index
                )
                st.plotly_chart(fig_rag, use_container_width=True)
            else:
                st.info("RAG Status column not detected")
        
        with col2:
            if customer_col:
                customer_counts = filtered_df[customer_col].value_counts().head(10)
                fig_customer = px.bar(
                    x=customer_counts.values,
                    y=customer_counts.index,
                    orientation='h',
                    title=f"Top 10 Customers ({customer_col})",
                    labels={'x': 'Order Count', 'y': 'Customer'},
                    color=customer_counts.values,
                    color_continuous_scale='Greens'
                )
                st.plotly_chart(fig_customer, use_container_width=True)
            else:
                st.info("Customer column not detected")
    
    # Tab 2: Geographic Analysis
    with tab2:
        st.header("🗺️ Geographic Analysis")
        
        if not (lat_col and lon_col):
            st.error("❌ Geographic analysis requires both Latitude and Longitude columns")
            st.info(f"Detected: Latitude='{lat_col}', Longitude='{lon_col}'")
            
            # Show available columns that might be coordinates
            st.write("**Available columns that might contain coordinates:**")
            coord_candidates = [col for col in df.columns if any(x in str(col).upper() for x in ['LAT', 'LON', 'COORD', 'GPS'])]
            if coord_candidates:
                st.write(coord_candidates)
            else:
                st.write("No potential coordinate columns found")
            return
        
        # Map configuration
        col1, col2, col3 = st.columns(3)
        
        with col1:
            color_options = [col for col in [status_col, rag_col, zone_col, circle_col, customer_col] if col]
            if color_options:
                map_color_by = st.selectbox("Color Points By", ['None'] + color_options)
            else:
                map_color_by = 'None'
                st.warning("No suitable columns for coloring")
        
        with col2:
            size_options = [col for col in df.columns if any(x in str(col).upper() for x in ['CAPEX', 'VALUE', 'AGING', 'COST'])]
            if size_options:
                map_size_by = st.selectbox("Size Points By", ['None'] + size_options)
            else:
                map_size_by = 'None'
        
        with col3:
            show_valid_coords = st.checkbox("Show Only Valid Coordinates", value=True)
        
        # Prepare map data
        map_df = filtered_df.copy()
        
        if show_valid_coords:
            map_df = map_df.dropna(subset=[lat_col, lon_col])
            map_df = map_df[(map_df[lat_col] != 0) & (map_df[lon_col] != 0)]
        
        if len(map_df) > 0:
            # Create hover data
            hover_cols = [col for col in [customer_col, zone_col, status_col] if col and col in map_df.columns]
            
            # Create map
            fig_map = px.scatter_mapbox(
                map_df,
                lat=lat_col,
                lon=lon_col,
                color=map_color_by if map_color_by != 'None' else None,
                size=map_size_by if map_size_by != 'None' else None,
                hover_data=hover_cols,
                title=f"Geographic Distribution ({len(map_df)} points)",
                mapbox_style='open-street-map',
                height=600,
                zoom=6
            )
            st.plotly_chart(fig_map, use_container_width=True)
            
            # Geographic summary
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Points on Map", len(map_df))
            with col2:
                if zone_col and zone_col in map_df.columns:
                    unique_zones = map_df[zone_col].nunique()
                    st.metric("Unique Zones", unique_zones)
            with col3:
                area_col = find_best_column_match(df, ['AREA', 'REGION'])
                if area_col and area_col in map_df.columns:
                    unique_areas = map_df[area_col].nunique()
                    st.metric("Unique Areas", unique_areas)
        else:
            st.warning("No valid coordinates found in filtered data")
    
    # Tab 3: Financial Analysis  
    with tab3:
        st.header("💰 Financial Analysis")
        
        # Find all financial columns
        capex_cols = [col for col in df.columns if 'CAPEX' in str(col).upper()]
        cost_cols = [col for col in df.columns if any(x in str(col).upper() for x in ['OTC', 'ARC', 'VALUE', 'COST'])]
        financial_cols = capex_cols + cost_cols
        
        if not financial_cols:
            st.error("❌ No financial columns detected")
            st.info("Looking for columns containing: CAPEX, OTC, ARC, VALUE, COST")
            return
        
        st.success(f"✅ Found {len(financial_cols)} financial columns: {', '.join(financial_cols)}")
        
        # CAPEX Analysis
        if capex_cols:
            st.subheader("💼 CAPEX Analysis")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # CAPEX by type
                capex_totals = filtered_df[capex_cols].sum()
                capex_totals = capex_totals[capex_totals > 0]
                
                if len(capex_totals) > 0:
                    fig_capex = px.pie(
                        values=capex_totals.values,
                        names=capex_totals.index,
                        title="CAPEX Distribution by Type"
                    )
                    st.plotly_chart(fig_capex, use_container_width=True)
            
            with col2:
                # CAPEX by Zone/Circle
                group_col = zone_col or circle_col
                if group_col:
                    group_capex = filtered_df.groupby(group_col)[capex_cols].sum().sum(axis=1)
                    group_capex = group_capex[group_capex > 0]
                    
                    if len(group_capex) > 0:
                        fig_group_capex = px.bar(
                            x=group_capex.index,
                            y=group_capex.values,
                            title=f"Total CAPEX by {group_col}",
                            labels={'x': group_col, 'y': 'Total CAPEX (₹)'}
                        )
                        st.plotly_chart(fig_group_capex, use_container_width=True)
        
        # Financial Summary Table
        st.subheader("📊 Financial Summary")
        
        financial_summary = []
        for col in financial_cols:
            if col in filtered_df.columns:
                col_data = filtered_df[col].dropna()
                if len(col_data) > 0:
                    financial_summary.append({
                        'Column': col,
                        'Total (₹)': col_data.sum(),
                        'Average (₹)': col_data.mean(),
                        'Max (₹)': col_data.max(),
                        'Count': len(col_data)
                    })
        
        if financial_summary:
            summary_df = pd.DataFrame(financial_summary)
            st.dataframe(summary_df, use_container_width=True)
    
    # Tab 4: Timeline Analysis
    with tab4:
        st.header("⏱️ Timeline & Aging Analysis")
        
        # Find date and aging columns
        date_cols = [col for col in df.columns if any(x in str(col).upper() for x in ['DATE', 'TIME'])]
        aging_cols = [col for col in df.columns if any(x in str(col).upper() for x in ['AGING', 'DAYS', 'HOLD'])]
        
        if not (date_cols or aging_cols):
            st.error("❌ No timeline or aging columns detected")
            return
        
        st.success(f"✅ Found timeline columns: {', '.join(date_cols + aging_cols)}")
        
        # Aging Analysis
        if aging_cols:
            st.subheader("📅 Aging Analysis")
            
            selected_aging_col = st.selectbox("Select Aging Column", aging_cols)
            
            if selected_aging_col in filtered_df.columns:
                aging_data = filtered_df[selected_aging_col].dropna()
                
                if len(aging_data) > 0:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig_aging_hist = px.histogram(
                            aging_data,
                            title=f"Aging Distribution ({selected_aging_col})",
                            nbins=30,
                            labels={'value': 'Days', 'count': 'Count'}
                        )
                        st.plotly_chart(fig_aging_hist, use_container_width=True)
                    
                    with col2:
                        if status_col:
                            aging_by_status = filtered_df.groupby(status_col)[selected_aging_col].mean()
                            fig_aging_status = px.bar(
                                x=aging_by_status.index,
                                y=aging_by_status.values,
                                title=f"Average {selected_aging_col} by Status",
                                labels={'x': 'Status', 'y': 'Average Days'}
                            )
                            st.plotly_chart(fig_aging_status, use_container_width=True)
        
        # Date Analysis
        if date_cols:
            st.subheader("📆 Date Analysis")
            
            selected_date_col = st.selectbox("Select Date Column", date_cols)
            
            if selected_date_col in filtered_df.columns:
                date_data = pd.to_datetime(filtered_df[selected_date_col], errors='coerce').dropna()
                
                if len(date_data) > 0:
                    # Timeline chart
                    monthly_counts = date_data.dt.to_period('M').value_counts().sort_index()
                    
                    fig_timeline = px.line(
                        x=monthly_counts.index.astype(str),
                        y=monthly_counts.values,
                        title=f"Orders Over Time ({selected_date_col})",
                        labels={'x': 'Month', 'y': 'Order Count'}
                    )
                    st.plotly_chart(fig_timeline, use_container_width=True)
    
    # Tab 5: Technical Analysis
    with tab5:
        st.header("🔧 Technical Analysis")
        
        # Find technical columns
        tech_cols = [col for col in df.columns if any(x in str(col).upper() for x in 
                    ['DEVICE', 'EQUIPMENT', 'ROUTER', 'PROVIDER', 'MEDIA', 'TYPE', 'MAKE', 'MODEL', 'SERVICE'])]
        
        if not tech_cols:
            st.error("❌ No technical columns detected")
            st.info("Looking for columns containing: DEVICE, EQUIPMENT, ROUTER, PROVIDER, MEDIA, TYPE, MAKE, MODEL, SERVICE")
            return
        
        st.success(f"✅ Found {len(tech_cols)} technical columns: {', '.join(tech_cols)}")
        
        # Technical Analysis
        col1, col2 = st.columns(2)
        
        with col1:
            # Device analysis
            device_cols = [col for col in tech_cols if any(x in str(col).upper() for x in ['DEVICE', 'MAKE', 'MODEL'])]
            if device_cols:
                selected_device_col = st.selectbox("Select Device Column", device_cols)
                
                if selected_device_col in filtered_df.columns:
                    device_counts = filtered_df[selected_device_col].value_counts().head(10)
                    
                    if len(device_counts) > 0:
                        fig_devices = px.bar(
                            x=device_counts.values,
                            y=device_counts.index,
                            orientation='h',
                            title=f"Top Devices ({selected_device_col})",
                            labels={'x': 'Count', 'y': 'Device'}
                        )
                        st.plotly_chart(fig_devices, use_container_width=True)
        
        with col2:
            # Service provider analysis
            provider_cols = [col for col in tech_cols if 'PROVIDER' in str(col).upper()]
            if provider_cols:
                selected_provider_col = st.selectbox("Select Provider Column", provider_cols)
                
                if selected_provider_col in filtered_df.columns:
                    provider_counts = filtered_df[selected_provider_col].value_counts()
                    
                    if len(provider_counts) > 0:
                        fig_providers = px.pie(
                            values=provider_counts.values,
                            names=provider_counts.index,
                            title=f"Service Providers ({selected_provider_col})"
                        )
                        st.plotly_chart(fig_providers, use_container_width=True)
        
        # Technical Configuration Summary
        st.subheader("📋 Technical Configuration Summary")
        
        tech_summary = []
        for col in tech_cols:
            if col in filtered_df.columns:
                unique_values = filtered_df[col].nunique()
                most_common = filtered_df[col].mode().iloc[0] if len(filtered_df[col].mode()) > 0 else 'N/A'
                data_availability = f"{filtered_df[col].notna().sum()}/{len(filtered_df)}"
                
                tech_summary.append({
                    'Column': col,
                    'Unique Values': unique_values,
                    'Most Common': str(most_common),
                    'Data Availability': data_availability
                })
        
        if tech_summary:
            tech_df = pd.DataFrame(tech_summary)
            st.dataframe(tech_df, use_container_width=True)
    
    # Tab 6: Advanced Analytics
    with tab6:
        st.header("📈 Advanced Analytics")
        
        # Column Selection for Analysis
        st.subheader("🔍 Custom Analysis")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Available Numeric Columns:**")
            numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
            if numeric_columns:
                selected_numeric_cols = st.multiselect(
                    "Select columns for correlation analysis",
                    numeric_columns,
                    default=numeric_columns[:5] if len(numeric_columns) >= 5 else numeric_columns
                )
            else:
                st.warning("No numeric columns found")
                selected_numeric_cols = []
        
        with col2:
            st.write("**Available Categorical Columns:**")
            categorical_columns = df.select_dtypes(include=['object']).columns.tolist()
            if categorical_columns:
                selected_cat_col = st.selectbox(
                    "Select categorical column for analysis",
                    categorical_columns
                )
            else:
                st.warning("No categorical columns found")
                selected_cat_col = None
        
        # Correlation Analysis
        if len(selected_numeric_cols) >= 2:
            st.subheader("🔗 Correlation Analysis")
            
            corr_matrix = filtered_df[selected_numeric_cols].corr()
            
            fig_corr = px.imshow(
                corr_matrix,
                title="Correlation Matrix",
                color_continuous_scale='RdBu',
                aspect='auto',
                text_auto=True
            )
            st.plotly_chart(fig_corr, use_container_width=True)
        
        # Performance Analysis
        if aging_col and status_col:
            st.subheader("⚡ Performance Analysis")
            
            # Create performance score based on aging
            performance_df = filtered_df.copy()
            if aging_col in performance_df.columns:
                max_aging = performance_df[aging_col].max()
                if max_aging > 0:
                    performance_df['Performance_Score'] = 100 - (performance_df[aging_col] / max_aging * 100)
                    
                    # Performance by groups
                    if zone_col or circle_col:
                        group_col = zone_col or circle_col
                        perf_by_group = performance_df.groupby(group_col)['Performance_Score'].mean().sort_values(ascending=False)
                        
                        if len(perf_by_group) > 0:
                            fig_perf = px.bar(
                                x=perf_by_group.index,
                                y=perf_by_group.values,
                                title=f"Performance Score by {group_col}",
                                labels={'x': group_col, 'y': 'Performance Score'},
                                color=perf_by_group.values,
                                color_continuous_scale='RdYlGn'
                            )
                            st.plotly_chart(fig_perf, use_container_width=True)
        
        # Data Quality Analysis
        st.subheader("🔍 Data Quality Analysis")
        
        # Missing data analysis
        missing_data = df.isnull().sum()
        missing_data = missing_data[missing_data > 0].sort_values(ascending=False)
        
        if len(missing_data) > 0:
            col1, col2 = st.columns(2)
            
            with col1:
                fig_missing = px.bar(
                    x=missing_data.values,
                    y=missing_data.index,
                    orientation='h',
                    title="Missing Data by Column",
                    labels={'x': 'Missing Count', 'y': 'Column'}
                )
                st.plotly_chart(fig_missing, use_container_width=True)
            
            with col2:
                # Data completeness percentage
                completeness = ((len(df) - missing_data) / len(df) * 100).sort_values(ascending=True)
                fig_completeness = px.bar(
                    x=completeness.values,
                    y=completeness.index,
                    orientation='h',
                    title="Data Completeness %",
                    labels={'x': 'Completeness %', 'y': 'Column'},
                    color=completeness.values,
                    color_continuous_scale='RdYlGn'
                )
                st.plotly_chart(fig_completeness, use_container_width=True)
        
        # Data Export Section
        st.subheader("💾 Data Export & Summary")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📊 Export Filtered Data"):
                csv = filtered_df.to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csv,
                    file_name=f"telecom_filtered_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
        
        with col2:
            if st.button("📈 Export Column Mapping"):
                # Create mapping report
                mapping_data = []
                for orig_col, clean_col in zip(df_raw.columns, df.columns):
                    mapping_data.append({
                        'Original_Column': orig_col,
                        'Cleaned_Column': clean_col,
                        'Data_Type': str(df[clean_col].dtype),
                        'Non_Null_Count': df[clean_col].notna().sum(),
                        'Sample_Value': str(df[clean_col].iloc[0]) if len(df) > 0 else 'N/A'
                    })
                
                mapping_df = pd.DataFrame(mapping_data)
                csv = mapping_df.to_csv(index=False)
                st.download_button(
                    label="Download Column Mapping",
                    data=csv,
                    file_name=f"column_mapping_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
        
        with col3:
            if st.button("📋 Generate Summary Report"):
                # Create comprehensive summary
                summary_data = {
                    'Metric': [
                        'Total Records',
                        'Total Columns',
                        'Numeric Columns',
                        'Date Columns',
                        'Technical Columns',
                        'Financial Columns',
                        'Data Completeness (%)',
                        'Records After Filters'
                    ],
                    'Value': [
                        len(df),
                        len(df.columns),
                        len(numeric_columns),
                        len([col for col in df.columns if any(x in str(col).upper() for x in ['DATE', 'TIME'])]),
                        len(tech_cols) if 'tech_cols' in locals() else 0,
                        len(financial_cols) if 'financial_cols' in locals() else 0,
                        f"{((df.notna().sum().sum() / (len(df) * len(df.columns))) * 100):.1f}%",
                        len(filtered_df)
                    ]
                }
                
                summary_df = pd.DataFrame(summary_data)
                csv = summary_df.to_csv(index=False)
                st.download_button(
                    label="Download Summary",
                    data=csv,
                    file_name=f"data_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
        
        # Current filter status
        st.info(f"📋 Current Filter Status: Showing {len(filtered_df):,} of {len(df):,} records")
        
        # Column detection summary
        with st.expander("🔍 Column Detection Summary", expanded=False):
            detection_summary = {
                'Data Type': ['Zone', 'Circle', 'Status', 'Customer', 'Latitude', 'Longitude', 'Aging', 'RAG Status'],
                'Detected Column': [zone_col, circle_col, status_col, customer_col, lat_col, lon_col, aging_col, rag_col],
                'Status': ['✅' if col else '❌' for col in [zone_col, circle_col, status_col, customer_col, lat_col, lon_col, aging_col, rag_col]]
            }
            
            detection_df = pd.DataFrame(detection_summary)
            st.dataframe(detection_df, use_container_width=True)

    # Footer
    st.markdown("---")
    st.markdown(
        f"""
        <div style='text-align: center; color: #666;'>
            📡 Advanced Telecom Data Analyzer | Processed {len(df):,} records from {len(df.columns)} columns<br>
            Built with Streamlit & Plotly | Handles Multi-Level Excel Headers
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
