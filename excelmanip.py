import streamlit as st
import pandas as pd
import io

# --- App Configuration ---
st.set_page_config(page_title="Excel Table Processor", layout="wide")

st.title("📊 Excel Table Processor")

# --- Step 1: Report Mode Selection ---
report_mode = st.selectbox(
    "Select Report Destination:",
    ["For mister Ahmed's office", "For cnss"]
)

st.divider()

# --- Step 2: File Upload ---
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # --- FIX: Handle Duplicate Columns ---
        # If Excel has duplicate column names (e.g. 'المدين' appearing twice),
        # Pandas can get confused. We rename duplicates here (Col, Col.1, Col.2).
        if len(df.columns) != len(set(df.columns)):
            st.warning("⚠️ Duplicate column names detected! We have renamed them automatically to prevent errors (e.g., 'Name' and 'Name.1').")
            new_cols = []
            seen_cols = {}
            for col in df.columns:
                col_str = str(col)
                if col_str in seen_cols:
                    seen_cols[col_str] += 1
                    new_cols.append(f"{col_str}.{seen_cols[col_str]}")
                else:
                    seen_cols[col_str] = 0
                    new_cols.append(col_str)
            df.columns = new_cols
            
        st.success("File uploaded successfully!")
        
        with st.expander("See raw data preview"):
            st.dataframe(df.head())

        # --- Step 3: User Inputs ---
        st.subheader("Configuration")

        col1, col2 = st.columns(2)
        
        with col1:
            user_name = st.text_input("Enter your Name:")
            
            # Select columns to extract
            all_columns = df.columns.tolist()
            columns_to_extract = st.multiselect(
                "Select the columns you want to extract:", 
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

        # --- Step 4: Processing Logic ---
        if st.button("Generate Excel File"):
            if not user_name:
                st.error("Please enter your name first.")
            elif not columns_to_extract:
                st.error("Please select at least one column to extract.")
            else:
                # --- A. CLEANING (Common to both modes) ---
                # We do the cleaning ONCE here to fix the "space" and "comma" issues
                clean_df = df.copy()
                
                # Force to string, remove ALL spaces (regex), replace comma with dot
                clean_series = clean_df[amount_col].astype(str)
                clean_series = clean_series.str.replace(r'\s+', '', regex=True)
                clean_series = clean_series.str.replace(',', '.')
                
                # Convert to number
                clean_df[amount_col] = pd.to_numeric(clean_series, errors='coerce').fillna(0)

                # Initialize buffer for the file
                buffer = io.BytesIO()
                file_suffix = ""

                # ===================================================
                # MODE 1: FOR MISTER AHMED'S OFFICE (Original Logic)
                # ===================================================
                if report_mode == "For mister Ahmed's office":
                    file_suffix = "ahmed_office"
                    
                    # Table 1: Detailed Data with User Name
                    df_table_1 = df[columns_to_extract].copy() # Use original df for text data
                    df_table_1.insert(0, 'name', user_name)
                    
                    # Table 2: Summary Data (Grouped)
                    df_table_2 = clean_df.groupby(person_id_col)[amount_col].sum().reset_index()
                    df_table_2.columns = [person_id_col, f"Total {amount_col}"]
                    
                    # SORTING: Sort Table 2 by Total Amount (Descending)
                    df_table_2 = df_table_2.sort_values(by=f"Total {amount_col}", ascending=False)

                    # Write to Excel (Side by Side) - Left to Right (Default)
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        # Write Table 1
                        df_table_1.to_excel(writer, sheet_name='Report', startrow=0, startcol=0, index=False)
                        
                        # Write Table 2 (Left a gap of 1 column)
                        start_col_table_2 = len(df_table_1.columns) + 1
                        df_table_2.to_excel(writer, sheet_name='Report', startrow=0, startcol=start_col_table_2, index=False)

                # ===================================================
                # MODE 2: FOR CNSS (New Logic)
                # ===================================================
                else:
                    file_suffix = "cnss"

                    # 1. Group by Person ID to make it unique
                    df_cnss = clean_df.groupby(person_id_col)[amount_col].sum().reset_index()
                    
                    # Get text info
                    df_text_info = df.drop_duplicates(subset=[person_id_col])[columns_to_extract + [person_id_col]]
                    
                    # Merge them
                    final_df = pd.merge(df_text_info, df_cnss, on=person_id_col, how='left')

                    # 2. Insert User Name at the start
                    final_df.insert(0, 'name', user_name)

                    # 3. Add the Empty Columns
                    final_df['CIN'] = ""
                    final_df['Tel'] = ""
                    final_df['Remarque'] = ""

                    # 4. Rename the money column to be clear it's a Total
                    final_df.rename(columns={amount_col: f"Total {amount_col}"}, inplace=True)

                    # 5. SORTING: Sort by Total Money (Biggest to Smallest)
                    final_df = final_df.sort_values(by=f"Total {amount_col}", ascending=False)

                    # 6. Reorder columns
                    # Added person_id_col explicitly after name
                    cols_order = ['name', person_id_col] + columns_to_extract + ['CIN', 'Tel', 'Remarque', f"Total {amount_col}"]
                    cols_order = list(dict.fromkeys(cols_order)) # Remove dupes
                    final_df = final_df[cols_order]

                    # Write to Excel (Single Table) - Left to Right (Default)
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        final_df.to_excel(writer, sheet_name='CNSS_Report', index=False)
                        
                        # Optional: Add formatting to make headers bold
                        workbook = writer.book
                        worksheet = writer.sheets['CNSS_Report']
                        header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
                        for col_num, value in enumerate(final_df.columns.values):
                            worksheet.write(0, col_num, value, header_fmt)

                # --- Step 5: Download Button ---
                st.divider()
                st.subheader("Result Ready")
                st.success(f"Generated report for: {report_mode}")
                
                st.download_button(
                    label="📥 Download Excel File",
                    data=buffer,
                    file_name=f"{user_name}_{file_suffix}.xlsx",
                    mime="application/vnd.ms-excel"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
