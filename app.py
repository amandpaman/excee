import streamlit as st
import pandas as pd

# Load Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file:
    sheet_names = pd.ExcelFile(uploaded_file).sheet_names
    selected_sheet = st.selectbox("Select Sheet", sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    st.markdown("### Data Preview")
    st.dataframe(df)

    st.markdown("---")
    st.markdown("### Apply Filters")

    # Auto-detect common columns
    columns_to_filter = ["Partner", "Zone", "Status", "Circle", "CAPEX", "Feasibility"]

    filters = {}
    for col in columns_to_filter:
        if col in df.columns:
            unique_vals = df[col].dropna().unique()
            selected_vals = st.multiselect(f"Select {col}", options=sorted(unique_vals))
            if selected_vals:
                filters[col] = selected_vals

    # Apply filters
    filtered_df = df.copy()
    for col, vals in filters.items():
        filtered_df = filtered_df[filtered_df[col].isin(vals)]

    st.markdown("### 🔍 Filtered Results")
    st.dataframe(filtered_df)

    st.download_button("📥 Download Filtered Data", filtered_df.to_csv(index=False), "filtered_data.csv")
