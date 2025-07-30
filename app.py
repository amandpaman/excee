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
    page_title="Dynamic Excel Data Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main { padding-top: 1rem; }
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
</style>
""", unsafe_allow_html=True)

# Helper functions
@st.cache_data
def load_excel_data(uploaded_file):
    """Load Excel data with automatic header detection"""
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheets = {}
        
        for sheet_name in excel_file.sheet_names:
            try:
                # Try different header configurations
                df_single = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
                
                # Try multi-level headers if single level has unnamed columns
                if any('Unnamed' in str(col) for col in df_single.columns):
                    try:
                        df_multi = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=[0, 1])
                        # Flatten multi-level columns
                        df_multi.columns = [f"{col[0]}_{col[1]}" if col[1] not in ['Unnamed', 'nan', ''] and not pd.isna(col[1]) 
                                          else str(col[0]) for col in df_multi.columns]
                        # Clean up column names
                        df_multi.columns = [col.strip().replace('_nan', '').replace('nan_', '') for col in df_multi.columns]
                        sheets[sheet_name] = df_multi
                    except:
                        sheets[sheet_name] = df_single
                else:
                    sheets[sheet_name] = df_single
                    
            except Exception as e:
                st.warning(f"Could not read sheet {sheet_name}: {e}")
                continue
                
        return sheets
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def analyze_column_types(df):
    """Automatically analyze and categorize columns"""
    column_info = {}
    
    for col in df.columns:
        col_data = df[col].dropna()
        
        if len(col_data) == 0:
            column_info[col] = {
                'type': 'empty',
                'unique_values': 0,
                'sample_values': [],
                'is_numeric': False,
                'is_date': False,
                'is_categorical': False
            }
            continue
        
        # Check if numeric
        numeric_count = 0
        try:
            pd.to_numeric(col_data, errors='raise')
            is_numeric = True
            numeric_count = len(col_data)
        except:
            try:
                # Try to convert some values to numeric
                numeric_series = pd.to_numeric(col_data, errors='coerce')
                numeric_count = numeric_series.notna().sum()
                is_numeric = numeric_count > len(col_data) * 0.7  # If >70% are numeric
            except:
                is_numeric = False
        
        # Check if date
        date_count = 0
        try:
            pd.to_datetime(col_data, errors='raise')
            is_date = True
            date_count = len(col_data)
        except:
            try:
                date_series = pd.to_datetime(col_data, errors='coerce')
                date_count = date_series.notna().sum()
                is_date = date_count > len(col_data) * 0.7  # If >70% are dates
            except:
                is_date = False
        
        # Determine if categorical (reasonable number of unique values)
        unique_values = col_data.nunique()
        is_categorical = unique_values <= min(50, len(col_data) * 0.5)
        
        # Get sample values
        sample_values = col_data.unique()[:5].tolist()
        
        column_info[col] = {
            'type': 'numeric' if is_numeric else 'date' if is_date else 'categorical' if is_categorical else 'text',
            'unique_values': unique_values,
            'sample_values': [str(v) for v in sample_values],
            'is_numeric': is_numeric,
            'is_date': is_date,
            'is_categorical': is_categorical,
            'data_count': len(col_data),
            'total_count': len(df),
            'completeness': len(col_data) / len(df) * 100
        }
    
    return column_info

def create_dynamic_filters(df, column_info):
    """Create filters based on actual data"""
    filters = {}
    
    # Only create filters for categorical columns with reasonable number of unique values
    categorical_cols = [col for col, info in column_info.items() 
                       if info['is_categorical'] and info['unique_values'] > 1 and info['unique_values'] <= 30]
    
    if not categorical_cols:
        return filters
    
    st.subheader("🔍 Dynamic Filters")
    
    # Create filters in columns
    num_cols = min(4, len(categorical_cols))
    if num_cols > 0:
        cols = st.columns(num_cols)
        
        for i, col in enumerate(categorical_cols[:4]):  # Limit to 4 filters for UI
            with cols[i % num_cols]:
                unique_values = sorted(df[col].dropna().astype(str).unique().tolist())
                selected = st.multiselect(
                    f"{col} ({len(unique_values)} options)",
                    options=unique_values,
                    default=[],
                    key=f"filter_{col}"
                )
                if selected:
                    filters[col] = selected
    
    # Additional filters if more categorical columns exist
    if len(categorical_cols) > 4:
        with st.expander(f"Additional Filters ({len(categorical_cols) - 4} more columns)"):
            remaining_cols = categorical_cols[4:]
            for col in remaining_cols:
                unique_values = sorted(df[col].dropna().astype(str).unique().tolist())
                selected = st.multiselect(
                    f"{col}",
                    options=unique_values,
                    default=[],
                    key=f"filter_extra_{col}"
                )
                if selected:
                    filters[col] = selected
    
    return filters

def apply_filters(df, filters):
    """Apply selected filters to dataframe"""
    filtered_df = df.copy()
    
    for col, values in filters.items():
        if values and col in filtered_df.columns:
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(values)]
    
    return filtered_df

def create_summary_metrics(original_df, filtered_df, column_info):
    """Create dynamic summary metrics"""
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Total Records", f"{len(original_df):,}")
    
    with col2:
        st.metric("Filtered Records", f"{len(filtered_df):,}", 
                 f"{len(filtered_df) - len(original_df):+,}")
    
    with col3:
        st.metric("Total Columns", len(original_df.columns))
    
    with col4:
        numeric_cols = [col for col, info in column_info.items() if info['is_numeric']]
        st.metric("Numeric Columns", len(numeric_cols))
    
    with col5:
        categorical_cols = [col for col, info in column_info.items() if info['is_categorical']]
        st.metric("Categorical Columns", len(categorical_cols))

def create_dynamic_visualizations(df, column_info, chart_type, selected_columns):
    """Create visualizations based on user selection"""
    
    if not selected_columns:
        st.warning("Please select at least one column for visualization")
        return
    
    if chart_type == "Bar Chart":
        if len(selected_columns) >= 1:
            col = selected_columns[0]
            if column_info[col]['is_categorical']:
                value_counts = df[col].value_counts().head(20)
                fig = px.bar(
                    x=value_counts.index,
                    y=value_counts.values,
                    title=f"Distribution of {col}",
                    labels={'x': col, 'y': 'Count'}
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning(f"Column '{col}' is not suitable for bar chart. Try histogram instead.")
    
    elif chart_type == "Pie Chart":
        if len(selected_columns) >= 1:
            col = selected_columns[0]
            if column_info[col]['is_categorical'] and column_info[col]['unique_values'] <= 10:
                value_counts = df[col].value_counts().head(10)
                fig = px.pie(
                    values=value_counts.values,
                    names=value_counts.index,
                    title=f"Distribution of {col}"
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning(f"Column '{col}' has too many unique values for pie chart. Try bar chart instead.")
    
    elif chart_type == "Histogram":
        if len(selected_columns) >= 1:
            col = selected_columns[0]
            if column_info[col]['is_numeric']:
                fig = px.histogram(
                    df,
                    x=col,
                    title=f"Histogram of {col}",
                    nbins=30
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning(f"Column '{col}' is not numeric. Try bar chart instead.")
    
    elif chart_type == "Scatter Plot":
        if len(selected_columns) >= 2:
            x_col, y_col = selected_columns[0], selected_columns[1]
            if column_info[x_col]['is_numeric'] and column_info[y_col]['is_numeric']:
                color_col = selected_columns[2] if len(selected_columns) > 2 and column_info[selected_columns[2]]['is_categorical'] else None
                fig = px.scatter(
                    df,
                    x=x_col,
                    y=y_col,
                    color=color_col,
                    title=f"Scatter Plot: {x_col} vs {y_col}",
                    hover_data=selected_columns[:5]
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Both X and Y columns must be numeric for scatter plot.")
        else:
            st.warning("Please select at least 2 columns for scatter plot.")
    
    elif chart_type == "Line Chart":
        if len(selected_columns) >= 2:
            x_col, y_col = selected_columns[0], selected_columns[1]
            if column_info[y_col]['is_numeric']:
                # Sort by x column for better line chart
                df_sorted = df.sort_values(x_col)
                fig = px.line(
                    df_sorted,
                    x=x_col,
                    y=y_col,
                    title=f"Line Chart: {y_col} over {x_col}"
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Y column must be numeric for line chart.")
        else:
            st.warning("Please select at least 2 columns for line chart.")
    
    elif chart_type == "Box Plot":
        if len(selected_columns) >= 1:
            y_col = selected_columns[0]
            if column_info[y_col]['is_numeric']:
                x_col = selected_columns[1] if len(selected_columns) > 1 and column_info[selected_columns[1]]['is_categorical'] else None
                fig = px.box(
                    df,
                    x=x_col,
                    y=y_col,
                    title=f"Box Plot of {y_col}" + (f" by {x_col}" if x_col else "")
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("First column must be numeric for box plot.")
    
    elif chart_type == "Heatmap":
        numeric_cols = [col for col in selected_columns if column_info[col]['is_numeric']]
        if len(numeric_cols) >= 2:
            corr_matrix = df[numeric_cols].corr()
            fig = px.imshow(
                corr_matrix,
                title="Correlation Heatmap",
                color_continuous_scale='RdBu',
                aspect='auto',
                text_auto=True
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Please select at least 2 numeric columns for heatmap.")

def create_geographic_map(df, column_info):
    """Create geographic visualization if coordinate columns exist"""
    
    # Look for potential coordinate columns
    potential_lat_cols = [col for col in df.columns if any(term in col.upper() for term in ['LAT', 'LATITUDE'])]
    potential_lon_cols = [col for col in df.columns if any(term in col.upper() for term in ['LON', 'LONG', 'LONGITUDE'])]
    
    if not (potential_lat_cols and potential_lon_cols):
        st.info("No geographic coordinates detected. Looking for columns containing 'LAT', 'LATITUDE', 'LON', 'LONGITUDE'")
        return
    
    lat_col = st.selectbox("Select Latitude Column", potential_lat_cols)
    lon_col = st.selectbox("Select Longitude Column", potential_lon_cols)
    
    if lat_col and lon_col:
        # Clean coordinate data
        map_df = df.copy()
        map_df[lat_col] = pd.to_numeric(map_df[lat_col], errors='coerce')
        map_df[lon_col] = pd.to_numeric(map_df[lon_col], errors='coerce')
        
        # Remove invalid coordinates
        map_df = map_df.dropna(subset=[lat_col, lon_col])
        map_df = map_df[(map_df[lat_col] != 0) | (map_df[lon_col] != 0)]
        
        if len(map_df) > 0:
            # Options for map customization
            col1, col2 = st.columns(2)
            
            with col1:
                categorical_cols = [col for col, info in column_info.items() if info['is_categorical'] and info['unique_values'] <= 20]
                color_col = st.selectbox("Color by (optional)", ['None'] + categorical_cols)
                color_col = color_col if color_col != 'None' else None
            
            with col2:
                numeric_cols = [col for col, info in column_info.items() if info['is_numeric']]
                size_col = st.selectbox("Size by (optional)", ['None'] + numeric_cols)
                size_col = size_col if size_col != 'None' else None
            
            # Create map
            fig = px.scatter_mapbox(
                map_df,
                lat=lat_col,
                lon=lon_col,
                color=color_col,
                size=size_col,
                hover_data=list(map_df.columns)[:5],
                title=f"Geographic Distribution ({len(map_df)} points)",
                mapbox_style='open-street-map',
                height=600,
                zoom=6
            )
            st.plotly_chart(fig, use_container_width=True)
            
            st.success(f"✅ Mapped {len(map_df)} points with valid coordinates")
        else:
            st.warning("No valid coordinate data found")

# Main App
def main():
    st.title("📊 Dynamic Excel Data Analyzer")
    st.markdown("*Automatically adapts to any Excel file structure*")
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("📂 Upload Data")
        uploaded_file = st.file_uploader(
            "Choose Excel File",
            type=['xlsx', 'xls', 'xlsm'],
            help="Upload any Excel file - the app will automatically analyze its structure"
        )
        
        if uploaded_file is not None:
            with st.spinner("Loading Excel file..."):
                sheets = load_excel_data(uploaded_file)
            
            if sheets:
                # Sheet selection
                selected_sheet = st.selectbox("Select Sheet", list(sheets.keys()))
                df = sheets[selected_sheet].copy()
                
                # Remove completely empty columns and rows
                df = df.dropna(axis=1, how='all').dropna(axis=0, how='all')
                
                st.success(f"✅ Loaded {len(df):,} rows × {len(df.columns)} columns")
                
                # Analyze column types
                with st.spinner("Analyzing data structure..."):
                    column_info = analyze_column_types(df)
                
            else:
                st.error("❌ Could not load Excel file")
                return
        else:
            st.info("👆 Upload an Excel file to start")
            st.markdown("""
            **This app will automatically:**
            - Detect your column structure
            - Create appropriate filters
            - Suggest suitable visualizations
            - Handle any Excel format
            """)
            return
    
    # Main content
    if 'df' in locals():
        # Data Structure Overview
        with st.expander("📋 Data Structure Overview", expanded=False):
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Column Information:**")
                structure_data = []
                for col, info in column_info.items():
                    structure_data.append({
                        'Column': col,
                        'Type': info['type'].title(),
                        'Unique Values': info['unique_values'],
                        'Completeness (%)': f"{info['completeness']:.1f}%",
                        'Sample Values': ', '.join(info['sample_values'][:3])
                    })
                
                structure_df = pd.DataFrame(structure_data)
                st.dataframe(structure_df, use_container_width=True, height=300)
            
            with col2:
                st.write("**Data Type Distribution:**")
                type_counts = {}
                for info in column_info.values():
                    type_counts[info['type']] = type_counts.get(info['type'], 0) + 1
                
                fig_types = px.pie(
                    values=list(type_counts.values()),
                    names=list(type_counts.keys()),
                    title="Column Types Distribution"
                )
                st.plotly_chart(fig_types, use_container_width=True)
        
        # Dynamic Filters
        filters = create_dynamic_filters(df, column_info)
        filtered_df = apply_filters(df, filters)
        
        # Summary Metrics
        st.subheader("📊 Data Summary")
        create_summary_metrics(df, filtered_df, column_info)
        
        # Show filter effect
        if filters:
            st.info(f"🔍 Applied {len(filters)} filters. Showing {len(filtered_df):,} of {len(df):,} records.")
        
        # Tabs for different analyses
        tab1, tab2, tab3, tab4 = st.tabs(["📈 Visualizations", "🗺️ Geographic", "📊 Data Analysis", "💾 Export"])
        
        with tab1:
            st.subheader("📈 Create Custom Visualizations")
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                # Chart type selection
                chart_type = st.selectbox(
                    "Select Chart Type",
                    ["Bar Chart", "Pie Chart", "Histogram", "Scatter Plot", "Line Chart", "Box Plot", "Heatmap"]
                )
                
                # Column selection based on chart type
                if chart_type in ["Heatmap"]:
                    numeric_cols = [col for col, info in column_info.items() if info['is_numeric']]
                    selected_columns = st.multiselect(
                        "Select Numeric Columns",
                        numeric_cols,
                        default=numeric_cols[:3] if len(numeric_cols) >= 3 else numeric_cols
                    )
                elif chart_type in ["Scatter Plot", "Line Chart"]:
                    all_cols = list(df.columns)
                    selected_columns = st.multiselect(
                        "Select Columns (X, Y, Color)",
                        all_cols,
                        default=all_cols[:2] if len(all_cols) >= 2 else all_cols
                    )
                else:
                    all_cols = list(df.columns)
                    selected_columns = st.multiselect(
                        "Select Columns",
                        all_cols,
                        default=[all_cols[0]] if all_cols else []
                    )
            
            with col2:
                if selected_columns:
                    create_dynamic_visualizations(filtered_df, column_info, chart_type, selected_columns)
        
        with tab2:
            st.subheader("🗺️ Geographic Analysis")
            create_geographic_map(filtered_df, column_info)
        
        with tab3:
            st.subheader("📊 Statistical Analysis")
            
            # Correlation analysis for numeric columns
            numeric_cols = [col for col, info in column_info.items() if info['is_numeric']]
            
            if len(numeric_cols) >= 2:
                st.write("**Correlation Analysis:**")
                selected_numeric = st.multiselect(
                    "Select numeric columns for correlation",
                    numeric_cols,
                    default=numeric_cols[:5] if len(numeric_cols) >= 5 else numeric_cols
                )
                
                if len(selected_numeric) >= 2:
                    corr_matrix = filtered_df[selected_numeric].corr()
                    
                    fig_corr = px.imshow(
                        corr_matrix,
                        title="Correlation Matrix",
                        color_continuous_scale='RdBu',
                        aspect='auto',
                        text_auto=True
                    )
                    st.plotly_chart(fig_corr, use_container_width=True)
            
            # Descriptive statistics
            st.write("**Descriptive Statistics:**")
            if numeric_cols:
                desc_stats = filtered_df[numeric_cols].describe()
                st.dataframe(desc_stats, use_container_width=True)
            else:
                st.info("No numeric columns found for statistical analysis")
        
        with tab4:
            st.subheader("💾 Export Data")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.write("**Export Filtered Data:**")
                if st.button("📊 Export Current View"):
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(
                        label="Download CSV",
                        data=csv,
                        file_name=f"filtered_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
            
            with col2:
                st.write("**Export Data Summary:**")
                if st.button("📋 Export Summary"):
                    summary_data = []
                    for col, info in column_info.items():
                        summary_data.append({
                            'Column': col,
                            'Type': info['type'],
                            'Unique_Values': info['unique_values'],
                            'Completeness_Percent': info['completeness'],
                            'Sample_Values': '; '.join(info['sample_values'])
                        })
                    
                    summary_df = pd.DataFrame(summary_data)
                    csv = summary_df.to_csv(index=False)
                    st.download_button(
                        label="Download Summary",
                        data=csv,
                        file_name=f"data_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
            
            with col3:
                st.write("**Current Status:**")
                st.info(f"📊 Original: {len(df):,} rows")
                st.info(f"🔍 Filtered: {len(filtered_df):,} rows")
                st.info(f"📈 Ready for export")
        
        # Raw data view
        with st.expander("👀 View Raw Data", expanded=False):
            st.dataframe(filtered_df, use_container_width=True, height=400)

if __name__ == "__main__":
    main()
