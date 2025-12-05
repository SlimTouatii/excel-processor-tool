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
        if len(df.columns) != len(set(df.columns)):
            st.warning("⚠️ Duplicate column names detected! We have renamed them automatically to prevent errors.")
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
                clean_df = df.copy()
                clean_series = clean_df[amount_col].astype(str)
                clean_series = clean_series.str.replace(r'\s+', '', regex=True)
                clean_series = clean_series.str.replace(',', '.')
                clean_df[amount_col] = pd.to_numeric(clean_series, errors='coerce').fillna(0)

                # Initialize buffer for the file
                buffer = io.BytesIO()
                file_suffix = ""

                # ===================================================
                # MODE 1: FOR MISTER AHMED'S OFFICE
                # ===================================================
                if report_mode == "For mister Ahmed's office":
                    file_suffix = "ahmed_office"
                    
                    df_table_1 = df[columns_to_extract].copy()
                    df_table_1.insert(0, 'name', user_name)
                    
                    df_table_2 = clean_df.groupby(person_id_col)[amount_col].sum().reset_index()
                    df_table_2.columns = [person_id_col, f"Total {amount_col}"]
                    df_table_2 = df_table_2.sort_values(by=f"Total {amount_col}", ascending=False)

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet('Report')
                        writer.sheets['Report'] = worksheet
                        
                        # Define Formats
                        # border: 1 adds the border
                        # text_wrap: True keeps text inside the cell (no spill)
                        # valign: top makes it look better if rows have different heights
                        cell_fmt = workbook.add_format({
                            'border': 1, 
                            'text_wrap': True,
                            'valign': 'top'
                        })
                        header_fmt = workbook.add_format({
                            'bold': True, 
                            'border': 1, 
                            'align': 'center', 
                            'valign': 'vcenter', 
                            'bg_color': '#D3D3D3',
                            'text_wrap': True
                        })

                        # Write Table 1
                        df_table_1.to_excel(writer, sheet_name='Report', startrow=0, startcol=0, index=False)
                        
                        # Apply formatting to Table 1 Columns
                        # set_column(first_col, last_col, width, cell_format)
                        for i, col in enumerate(df_table_1.columns):
                            # Set column width to 20 and apply border/wrap format
                            worksheet.set_column(i, i, 20, cell_fmt)
                            # Write Header manually to apply header format
                            worksheet.write(0, i, col, header_fmt)

                        # Write Table 2
                        start_col_table_2 = len(df_table_1.columns) + 1
                        df_table_2.to_excel(writer, sheet_name='Report', startrow=0, startcol=start_col_table_2, index=False)

                        # Apply formatting to Table 2 Columns
                        for i, col in enumerate(df_table_2.columns):
                            col_idx = start_col_table_2 + i
                            worksheet.set_column(col_idx, col_idx, 20, cell_fmt)
                            worksheet.write(0, col_idx, col, header_fmt)

                # ===================================================
                # MODE 2: FOR CNSS
                # ===================================================
                else:
                    file_suffix = "cnss"

                    df_cnss = clean_df.groupby(person_id_col)[amount_col].sum().reset_index()
                    df_text_info = df.drop_duplicates(subset=[person_id_col])[columns_to_extract + [person_id_col]]
                    final_df = pd.merge(df_text_info, df_cnss, on=person_id_col, how='left')

                    final_df.insert(0, 'name', user_name)
                    final_df['CIN'] = ""
                    final_df['Tel'] = ""
                    final_df['Remarque'] = ""
                    final_df.rename(columns={amount_col: f"Total {amount_col}"}, inplace=True)
                    final_df = final_df.sort_values(by=f"Total {amount_col}", ascending=False)

                    cols_order = ['name', person_id_col] + columns_to_extract + ['CIN', 'Tel', 'Remarque', f"Total {amount_col}"]
                    cols_order = list(dict.fromkeys(cols_order))
                    final_df = final_df[cols_order]

                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet('CNSS_Report')
                        writer.sheets['CNSS_Report'] = worksheet
                        
                        # Define Formats
                        cell_fmt = workbook.add_format({
                            'border': 1, 
                            'text_wrap': True,
                            'valign': 'top'
                        })
                        header_fmt = workbook.add_format({
                            'bold': True, 
                            'border': 1, 
                            'align': 'center', 
                            'valign': 'vcenter', 
                            'bg_color': '#D3D3D3',
                            'text_wrap': True
                        })

                        # Write Table
                        final_df.to_excel(writer, sheet_name='CNSS_Report', index=False)
                        
                        # Apply formatting to all columns
                        for i, col in enumerate(final_df.columns):
                            worksheet.set_column(i, i, 20, cell_fmt)
                            worksheet.write(0, i, col, header_fmt)

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
