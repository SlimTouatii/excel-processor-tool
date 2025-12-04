import streamlit as st
import pandas as pd
import io

# --- App Configuration ---
st.set_page_config(page_title="Excel Column Extractor", layout="wide")

st.title("📊 Excel Table Processor")
st.markdown("""
This tool allows you to upload an Excel file, select specific data, 
and generate a summary report with two side-by-side tables.
""")

# --- Step 1: File Upload ---
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully!")
        st.write("Preview of your data:", df.head())

        # --- Step 2: User Inputs ---
        st.divider()
        st.subheader("Configuration")

        col1, col2 = st.columns(2)
        
        with col1:
            user_name = st.text_input("Enter your Name:")
        
        # Select columns to keep
        all_columns = df.columns.tolist()
        columns_to_extract = st.multiselect(
            "Select the columns you want to keep (Table 1):", 
            options=all_columns
        )

        with col2:
            # Select specific operational columns
            person_id_col = st.selectbox(
                "Which column represents the Person ID?", 
                options=all_columns
            )
            
            amount_col = st.selectbox(
                "Which column represents the Amount of Money?", 
                options=all_columns,
                index=len(all_columns)-1 if len(all_columns) > 0 else 0
            )

        # --- Step 3: Processing ---
        if st.button("Generate New Excel File"):
            if not user_name:
                st.error("Please enter your name first.")
            elif not columns_to_extract:
                st.error("Please select at least one column to extract.")
            else:
                # --- Prepare Table 1 (Left Side) ---
                df_table_1 = df[columns_to_extract].copy()
                df_table_1.insert(0, 'name', user_name)
                
                # --- CLEANING DATA FOR TABLE 2 ---
                clean_df = df.copy()
                
                # 1. Convert column to string so we can run text replacement
                clean_series = clean_df[amount_col].astype(str)
                
                # 2. Remove ALL spaces (including the hidden \xa0 non-breaking space)
                # '\s+' is a regex code that means "any white space character"
                clean_series = clean_series.str.replace(r'\s+', '', regex=True)
                
                # 3. Replace comma with dot (Python needs dots for decimals)
                clean_series = clean_series.str.replace(',', '.')
                
                # 4. Convert to proper numbers
                clean_df[amount_col] = pd.to_numeric(clean_series, errors='coerce').fillna(0)

                # --- Prepare Table 2 (Right Side - Summary) ---
                df_table_2 = clean_df.groupby(person_id_col)[amount_col].sum().reset_index()
                df_table_2.columns = [person_id_col, f"Total {amount_col}"]

                # --- Step 4: Write to Excel Buffer ---
                buffer = io.BytesIO()
                
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    # Write Table 1
                    df_table_1.to_excel(writer, sheet_name='Report', startrow=0, startcol=0, index=False)
                    
                    # Write Table 2
                    start_col_table_2 = len(df_table_1.columns) + 1
                    df_table_2.to_excel(writer, sheet_name='Report', startrow=0, startcol=start_col_table_2, index=False)

                # --- Step 5: Download Button ---
                st.divider()
                st.subheader("Download Result")
                
                st.download_button(
                    label="📥 Download Processed Excel",
                    data=buffer,
                    file_name=f"{user_name}_processed_report.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
                st.success("File generated! The hidden spaces have been removed.")
                st.write("Preview of the generated Summary (Table 2):")
                st.dataframe(df_table_2)

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")