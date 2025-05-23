import pandas as pd
import streamlit as st

# Streamlit app interface to upload files
st.title('Concatenate Multiple Excel Files')

  
#streamlit run concatenar_xls.py

# Allow user to upload multiple Excel files
uploaded_files = st.file_uploader("Choose Excel files", accept_multiple_files=True)

if uploaded_files:
    all_data = []  # List to store data from all files

    for file in uploaded_files:
        # Read each uploaded Excel file
        df = pd.read_excel(file)
        
        # Append the data from the current file to the list
        all_data.append(df)

    # Concatenate all the DataFrames from the uploaded files
    combined_df = pd.concat(all_data, ignore_index=True)

    # Show the combined data in the app
    st.write('Combined Data Preview:', combined_df)

    # Allow the user to download the combined file
    output = 'combined_output.xlsx'
    combined_df.to_excel(output, index=False)

    with open(output, 'rb') as f:
        st.download_button(
            label="Download Combined Excel File",
            data=f,
            file_name=output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
