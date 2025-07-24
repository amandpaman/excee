import streamlit as st
import pandas as pd
import plotly.express as px

# --- App UI Settings ---
st.set_page_config(page_title="Excel Dashboard Viewer", layout="wide")

st.markdown("""
    <style>
        .main { background: linear-gradient(to bottom right, #f0f9ff, #cbebff); }
        .stApp { font-family: 'Segoe UI', sans-serif; }
        .block-container { padding: 2rem 2rem 2rem 2rem; }
        .css-1d391kg { background-color: #f5f7fa !important; border-radius: 10px; padding: 20px; box-shadow: 2px 2px 8px rgba(0,0,0,0.1); }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Excel Dashboard Viewer")
st.markdown("Upload your complex Excel data to view, explore, and visualize it interactively.")

# --- File Upload ---
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, header=1)
        df.dropna(how="all", inplace=True)
        df.dropna(axis=1, how="all", inplace=True)

        st.success("✅ File uploaded and parsed successfully!")
        
        with st.expander("📑 View Raw Data"):
            st.dataframe(df, use_container_width=True)

        # --- Filter Sidebar ---
        st.sidebar.header("🔎 Data Filters")
        selected_col = st.sidebar.selectbox("Select column to filter", df.columns)
        unique_vals = df[selected_col].dropna().unique()
        selected_val = st.sidebar.multiselect(f"Choose {selected_col}", unique_vals, default=unique_vals[:5])

        filtered_df = df[df[selected_col].isin(selected_val)]

        st.subheader(f"🔍 Filtered Data by `{selected_col}`")
        st.dataframe(filtered_df, use_container_width=True)

        # --- Visualizations ---
        st.subheader("📈 Visualizations")

        col1, col2 = st.columns(2)

        # Pie Chart
        with col1:
            pie_col = st.selectbox("🧩 Pie Chart Column", df.columns)
            pie_data = df[pie_col].value_counts().reset_index()
            pie_chart = px.pie(pie_data, names='index', values=pie_col, title=f"Distribution of {pie_col}")
            st.plotly_chart(pie_chart, use_container_width=True)

        # Bar Chart
        with col2:
            bar_col = st.selectbox("📊 Bar Chart Column", df.columns, index=1)
            bar_data = df[bar_col].value_counts().reset_index()
            bar_chart = px.bar(bar_data, x='index', y=bar_col, title=f"{bar_col} Frequency", color='index')
            st.plotly_chart(bar_chart, use_container_width=True)

        st.markdown("---")

        st.subheader("📤 Export Cleaned Data")
        csv = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download CSV", data=csv, file_name="filtered_data.csv", mime='text/csv')

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👈 Please upload an Excel file to continue.")
