import pandas as pd
import streamlit as st
from io import BytesIO

st.title('Concatenate Multiple Excel Files (Optimized)')

uploaded_files = st.file_uploader("Choose Excel files", accept_multiple_files=True, type=["xlsx"])

if uploaded_files:
    all_data = []

    # Iterate over uploaded files
    for file in uploaded_files:
        # Read only the necessary sheet and columns if applicable
        df = pd.read_excel(file, engine='openpyxl')  # openpyxl is often faster for xlsx
        all_data.append(df)

    # Efficiently concatenate
    combined_df = pd.concat(all_data, ignore_index=True)

    st.write('Combined Data Preview:', combined_df)

    # Write to an in-memory buffer instead of disk
    output_buffer = BytesIO()
    combined_df.to_excel(output_buffer, index=False, engine='openpyxl')
    output_buffer.seek(0)

    st.download_button(
        label="Download Combined Excel File",
        data=output_buffer,
        file_name='combined_output.xlsx',
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
