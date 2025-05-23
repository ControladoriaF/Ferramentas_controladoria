import streamlit as st
import pandas as pd
from dbfread import DBF
import io
from io import BytesIO

st.title("Download Arquivo .DBF ")

uploaded_file = st.file_uploader("Importe", type=["dbf"])


#streamlit run Importação_de_arquivo.py

if uploaded_file is not None:
    # Read the uploaded .dbf file using dbfread
    with open("temp_file.dbf", "wb") as f:
        f.write(uploaded_file.read())

    table = DBF("temp_file.dbf", load=True, encoding='latin-1')
    df = pd.DataFrame(iter(table))
    st.write(f"Quantidade de Linhas {len(df)}")
    st.dataframe(df)
    modified_excel = BytesIO()

    with pd.ExcelWriter(modified_excel, engine="xlsxwriter") as writer:
                        df.to_excel(writer, index=False, sheet_name="Tranferências")
            
    modified_excel.seek(0)

    download_geral = st.download_button(
                    label="Debito",
                    data=modified_excel,
                    file_name=F"Banco.xlsx",
                    mime="text/csv"
                )



    st.success("Arquivo Carregado")
    
