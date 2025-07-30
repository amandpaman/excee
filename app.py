import streamlit as st
import pandas as pd
import plotly.express as px

# App title
st.set_page_config(page_title="Excel Data Visualizer", layout="wide")
st.title("📊 Excel Data Visualization Dashboard")

# Sidebar for file upload
with st.sidebar:
    st.header("Configuration")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    st.markdown("---")
    st.markdown("**How to use:**")
    st.markdown("1. Upload your Excel file")
    st.markdown("2. Select filters")
    st.markdown("3. View interactive visualizations")

# Main content
if uploaded_file is not None:
    try:
        # Read Excel file
        excel_data = pd.ExcelFile(uploaded_file)
        sheet_names = excel_data.sheet_names
        
        # Select sheet
        selected_sheet = st.selectbox("Select Sheet", sheet_names)
        df = excel_data.parse(selected_sheet)
        
        # Clean column names (remove special characters)
        df.columns = df.columns.str.replace('[^a-zA-Z0-9]', '_', regex=True)
        
        # Store original dataframe
        if 'original_df' not in st.session_state:
            st.session_state.original_df = df
        
        st.success("Data loaded successfully!")
        
        # Show raw data
        with st.expander("View Raw Data"):
            st.dataframe(df)

        # Filter section
        st.subheader("Apply Filters")
        selected_columns = st.multiselect("Select columns to display", df.columns, default=df.columns.tolist()[:5])
        
        filtered_df = df[selected_columns]
        
        # Numeric filters
        numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
        for col in numeric_cols:
            if col in filtered_df.columns:
                min_val = float(df[col].min())
                max_val = float(df[col].max())
                step = (max_val - min_val) / 100
                selected_range = st.slider(
                    f"Range for {col}",
                    min_val,
                    max_val,
                    (min_val, max_val),
                    step=step
                )
                filtered_df = filtered_df[(filtered_df[col] >= selected_range[0]) & (filtered_df[col] <= selected_range[1])]
        
        # Categorical filters
        categorical_cols = df.select_dtypes(include=['object', 'category']).columns
        for col in categorical_cols:
            if col in filtered_df.columns:
                unique_values = df[col].unique()
                selected_values = st.multiselect(f"Filter {col}", unique_values, default=unique_values)
                filtered_df = filtered_df[filtered_df[col].isin(selected_values)]
        
        # Display filtered data
        st.subheader("Filtered Data")
        st.dataframe(filtered_df)
        
        # Visualization section
        st.subheader("Data Visualization")
        plot_col1, plot_col2 = st.columns(2)
        
        with plot_col1:
            st.markdown("### Scatter Plot")
            x_axis = st.selectbox("X-axis", numeric_cols, index=0)
            y_axis = st.selectbox("Y-axis", numeric_cols, index=1 if len(numeric_cols) > 1 else 0)
            color_by = st.selectbox("Color by", [None] + categorical_cols.tolist())
            
            if len(numeric_cols) >= 2:
                fig_scatter = px.scatter(
                    filtered_df,
                    x=x_axis,
                    y=y_axis,
                    color=color_by,
                    hover_data=filtered_df.columns,
                    title=f"{y_axis} vs {x_axis}"
                )
                st.plotly_chart(fig_scatter, use_container_width=True)
        
        with plot_col2:
            st.markdown("### Bar Chart")
            bar_col = st.selectbox("Column for bars", categorical_cols, index=0)
            agg_col = st.selectbox("Value column", numeric_cols)
            
            agg_method = st.selectbox("Aggregation method", ["sum", "mean", "count"])
            
            if bar_col in filtered_df.columns and agg_col in filtered_df.columns:
                agg_df = filtered_df.groupby(bar_col).agg({agg_col: agg_method}).reset_index()
                
                fig_bar = px.bar(
                    agg_df,
                    x=bar_col,
                    y=agg_col,
                    title=f"{agg_method.title()} of {agg_col} by {bar_col}"
                )
                st.plotly_chart(fig_bar, use_container_width=True)
        
        # Additional visualizations
        st.markdown("### Correlation Matrix")
        
        if len(numeric_cols) > 1:
            corr_matrix = filtered_df[numeric_cols].corr()
            fig_corr = px.imshow(
                corr_matrix,
                text_auto=True,
                title="Correlation Between Numeric Columns"
            )
            st.plotly_chart(fig_corr, use_container_width=True)
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")

else:
    st.info("Please upload an Excel file to begin")
    st.markdown("""
    ### Features:
    - View and filter Excel data interactively
    - Create scatter plots comparing numeric columns
    - Generate bar charts with aggregation
    - Visualize correlations between columns
    - Responsive layout for all screen sizes
    """)

# Add custom CSS for better styling
st.markdown("""
<style>
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 4px;
        padding: 0.5rem 1rem;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .sidebar .sidebar-content {
        background-color: #f5f5f5;
    }
    div[data-testid="stExpander"] div[role="button"] p {
        font-size: 1rem;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)
