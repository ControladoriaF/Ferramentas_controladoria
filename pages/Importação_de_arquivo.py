import streamlit as st
import os
import sqlite3
import datetime as dt
import time
import pandas as pd
from io import BytesIO
import streamlit.components.v1 as components
from datetime import datetime
import numpy as np
import json
import io
import re
from reportlab.lib.pagesizes import letter,landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,Image
from PyPDF2 import PdfMerger
import plotly.express as px
import os

from streamlit_extras.stylable_container import stylable_container 

import zipfile
import plotly.io as pio
#pio.kaleido.scope.default_format = "png"
# Connect to SQLite database (or create it)
conn = sqlite3.connect('files.db')
cursor = conn.cursor()

# Create a table to store file metadata (if not exists)
cursor.execute("""
    CREATE TABLE IF NOT EXISTS files (
        store_name TEXT,
        file_path TEXT
              
    )
""")

def convert_to_int(value):
        try:
            # Replace commas with periods for decimal values
            value = str(value).replace(',', '.')
            
            # Attempt to convert to float first, then to int
            return float(value)
        except (ValueError, TypeError):
            # If it fails, return NaN or the original value
            return np.nan  # You can return np.nan or the original value if needed

# Check if the column 'list_ativos' already exists in the 'files' table

columns = cursor.fetchall()
conn.commit()

loja_map = {
    1: "01 - Irmã Dulce",
    2: "02 - Lourival Parente",
    3: "03 - Porto Alegre",
    4: "04 - Água Mineral",
    5: "05 - CD Chapadinha ",
    6: "06 - Renascença",
    7: "07 - Parque Piauí",
    8: "08 - São Joaquim",
    9: "09 - Mocambinho",
    10: "10 - Morada do Sol",
    11: "11 - Dirceu",
    12: "12 - Centro",
    13: "13 - CD frios",
    14: "14 - Água Branca",
    16: "16 - Emp. Dom Severino",
    17: "17 - Cristo Rei",
    18: "18 - Emp. Mocambinho",
    22: "22 - Noé Mendes",
    23: "23 - Kennedy",
    101 : "101 - CD Areias"

    }
# Check if 'list_ativos' is already present
column_names = [column[1] for column in columns]

st.set_page_config(layout="wide",page_title="Importação de arquivo",
                   )

def example():
    with stylable_container(
        key="green_button",
        css_styles="""
            button {
                background-color: green;
                color: white;
                border-radius: 20px;
            }
            """,
    ):
        st.button("Green button")

    st.button("Normal button")
    


#example() 


with stylable_container(
    key="container_with_border",
    css_styles="""
        {   background-color: rgba(67, 202, 90, 0.3);
            border: 5px solid rgba(67, 202, 90, 0.2);
            border-radius: 0.5rem;
            padding: calc(1em - 1px)
        }
        """,
):

    #st.markdown("This is a container with a border.")
    st.title("Relatorio 150052(Ciss) -> Isscolector")
    st.header("ARQUIVO DE ESTOQUE DA CISS PARA ALIMENTAR O ISSCOLECTOR")


    relatorio = st.file_uploader("Importe o Arquivo de estoque da CISS",accept_multiple_files = True)

    if relatorio is not None:

        all_df = []

        d = st.date_input("dia do inventário", value=dt.date.today())
        d = d.strftime("%d-%m-%y")
        
        for file in relatorio:
            # Read the Excel file
            df_estoque_geral = pd.read_excel(file)

            ###########
            
            ############
            df_estoque_geral["loja"] = df_estoque_geral["estoque_sintetico_idempresa"]
            df_estoque_geral["loja"].replace(loja_map,inplace=True)
            
            nome_loja ="".join(str(df_estoque_geral["loja"].unique()))
            nome_loja = nome_loja.strip("[]")
            nome_loja = nome_loja.strip("'")
            
            #st.write(len(df_estoque_geral["descrdivisao"].unique()) )
            
            if len(df_estoque_geral["descrdivisao"].unique()) < 8:
                nome_divisão ="".join(str(df_estoque_geral["descrdivisao"].unique()))
                
            else:
                nome_divisão ="".join(str(df_estoque_geral["iddivisao"].unique()))

            
            
            nome_divisão = nome_divisão.strip("´")

            nome_divisão = re.sub(r"[\[\]'\" ]+", " ", nome_divisão).strip()
            nome_divisão = nome_divisão.replace("/", " e ")

            unico_list = df_estoque_geral["produtos_view_idsecao"].unique()

            nome_seção ="".join(str(unico_list))
            nome_seção = nome_seção.strip("[]")

            #df_estoque_geral["Codigo_novo"] = df_estoque_geral["idcodbarprodtrib"].fillna(df_estoque_geral["produtos_view_idcodbarprod"])
            #target_sections = [] # vão usar o codigo antigo
            target_sections = ["CONGELADOS","AÇOUGUE","FRIOS E RESFRIADOS","HORTIFRUTI","SALGADOS"] 
            df_estoque_geral["Codigo_novo"] = df_estoque_geral.apply(
                lambda row: row["produtos_view_idcodbarprod"]  # fallback to old if main is missing and not in target section
                if (
                    row["descrdivisao"]  in target_sections and (pd.isna(row["idcodbarprodtrib"]) or row["idcodbarprodtrib"] != row["produtos_view_idcodbarprod"] ) or pd.isna(row["idcodbarprodtrib"])
                )
                else row["idcodbarprodtrib"],  # otherwise, use main
                axis=1
            )
                        

            # Check the columns
            # Select relevant columns
            df_estoque = df_estoque_geral[["Codigo_novo", "estoque_sintetico_qtdatualestoque","customedioun"]]

            df_produtos = df_estoque_geral[["Codigo_novo", "produtos_view_descricaoproduto","secao_descrsecao","produtos_view_idsubproduto"]]
            df_produtos_2 = df_estoque_geral[["Codigo_novo", "produtos_view_idsubproduto"]]
            
            df_estoque['Codigo_novo'] =  df_estoque['Codigo_novo'].round(0)
            df_estoque.rename(columns = {"Codigo_novo":"CODIGO","estoque_sintetico_qtdatualestoque":"QTD","customedioun":"VALORUNIT"}, inplace=True)
            df_produtos.rename(columns = {"Codigo_novo":"CODIGO","produtos_view_descricaoproduto":"DESCRICAO","secao_descrsecao":"DIVISÃO","produtos_view_idsubproduto":"CODIGO_INTERNO"}, inplace=True)
            df_produtos_2.rename(columns = {"Codigo_novo":"CODIGO_BARRA","produtos_view_idsubproduto":"CODIGO_INTERNO"}, inplace=True)
            
            # Convert quantities to integer
            # First, replace NaN values with a placeholder, like 0, or another appropriate value
            df_estoque['CODIGO'] = df_estoque['CODIGO'].apply(lambda x: int(x) if pd.notna(x) else 0)  # Replace NaN with 0
            df_estoque['CODIGO'] = df_estoque['CODIGO'].round(0).astype(int)

            df_produtos['CODIGO'] = df_produtos['CODIGO'].apply(lambda x: int(x) if pd.notna(x) else 0)  # Replace NaN with 0
            df_produtos['CODIGO'] = df_produtos['CODIGO'].round(0).astype(int)

            df_produtos_2['CODIGO_BARRA'] = df_produtos_2['CODIGO_BARRA'].apply(lambda x: int(x) if pd.notna(x) else 0)  # Replace NaN with 0
            df_produtos_2['CODIGO_BARRA'] = df_produtos_2['CODIGO_BARRA'].round(0).astype(int)

            df_produtos_2['CODIGO_INTERNO'] = df_produtos_2['CODIGO_INTERNO'].apply(lambda x: int(x) if pd.notna(x) else 0)  # Replace NaN with 0
            df_produtos_2['CODIGO_INTERNO'] = df_produtos_2['CODIGO_INTERNO'].round(0).astype(int)
            
            # Ensure 'QTD' is treated as integer (replace NaN with 0 for 'QTD' too, or another placeholder if necessary)
            df_estoque['QTD'] = df_estoque['QTD'].apply(lambda x: float(x) if pd.notna(x) else 0)
            df_estoque['QTD'] = df_estoque['QTD'].round(2)
            
            # Save to CSV (in memory)
            csv_estoque = df_estoque.to_csv(sep=';', index=False, header=True)
            csv_produtos = df_produtos.to_csv(sep=';', index=False, header=True)
            csv_produtos_2 = df_produtos_2.to_csv(sep=';', index=False, header=True)
            
            # Check if the CSV data is empty
            if not csv_estoque:
                st.error("Error: Arquivo estoque está vazio.")
            if not csv_produtos:
                st.error("Error: Arquivo produtos está vazio.")
            if not csv_produtos_2:
                st.error("Error: Arquivo produtos 2 está vazio.")
            
            
            dataframes = [
            (df_estoque, "Estoque"),
            (df_produtos, "Produtos"),
            (df_produtos_2, "Produtos Internos")
                ]
            
            def generate_zip():
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    # Writing each dataframe to the zip file
                    for df, name in [(df_estoque, f'Estoque  {nome_loja} - {nome_divisão} {d}'), (df_produtos, f'Produtos {nome_loja} - {nome_divisão} {d}'), (df_produtos_2, f'Produtos Internos {nome_loja} - {nome_divisão} {d}')]:
                        csv_buffer = io.StringIO()
                        df.to_csv(csv_buffer, sep=';', index=False, header=True)
                        zip_file.writestr(f'{name}.txt', csv_buffer.getvalue())
                zip_buffer.seek(0)
                return zip_buffer.getvalue()

            # Go back to the beginning of the buffer
            container = st.container(border=True)
            
            co1,co2=st.columns([0.3,1])

            with container:

                st.header(f"Inventário {nome_loja} ({d}) {nome_divisão} {nome_seção} ")
                
                st.header(f"Soma Itens: {round(df_estoque['QTD'].sum(),2)}")

                # Use this for the download button
                if st.button(f"Gerar Arquivo ZIP {nome_loja} {nome_divisão}"):
                    zip_bytes = generate_zip()
                    st.session_state["zip_bytes"] = zip_bytes
                    st.success("Arquivo ZIP gerado com sucesso!")

                    if "zip_bytes" in st.session_state:
                        st.download_button(
                            label=f"Baixar Arquivo ZIP {nome_divisão}",
                            data=st.session_state["zip_bytes"],
                            file_name=f"Inventario Loja {nome_loja} {d}.zip",
                            mime="application/zip",
                            key=f"download_zip{nome_loja}"
                        )


                with st.expander("Modo antigo"):
                    # Only show download buttons if the CSV data is not empty
                    if csv_estoque:
                        # Create a download button for the 'estoque' CSV
                        st.download_button(
                            label=f"Estoque {nome_loja} {nome_divisão}",
                            data=csv_estoque,
                            file_name=f"Inventario_loja_{datetime.today().strftime('%d-%m-%Y')}_estoque.txt",
                            mime="text/csv"
                        )

                    if csv_produtos:
                        # Create a download button for the 'produtos' CSV
                        st.download_button(
                            label=f"Produtos {nome_loja} {nome_divisão}",
                            data=csv_produtos,
                            file_name=f"Inventario_loja_{datetime.today().strftime('%d-%m-%Y')}_produtos.txt",
                            mime="text/csv"
                        )

                    if csv_produtos_2:
                        st.download_button(
                            label=f"Produtos Internos {nome_loja} {nome_divisão}",
                            data=csv_produtos_2,
                            file_name=f"Inventario_loja_{datetime.today().strftime('%d-%m-%Y')}_produtos Internos (opcional) .txt",
                            mime="text/csv"
                    )
        

with stylable_container(
    key="container_with_border_2",
    css_styles="""
        {   background-color: rgba(202, 193, 67, 0.2);
            border: 5px solid rgba(67, 202, 90, 0.2);
            border-radius: 0.5rem;
            padding: calc(1em - 1px)
        }
        """,
):
    st.title("Divergência Zerados/Negativos/Positivos para cada Loja")
    st.header("IMPORTE A DIVERGêNCIA ISSCOLECTOR (DIVERGENCIA -> ATIVO)")

    uploaded_file = st.file_uploader("Importe a divergência ISSCOLECTOR", type=["xlsx", "xls"])

    nomes = [
        "Anderson Roberto Barros da Cruz",
        "Arivaldo Silva Santos",
        "Leandro Pereira da Silva",
        "Rudyson Rafael Alves Paulo",
        "Raimundo Nonato dos Santos Nascimento",
        "Rayrton Oliveira Guedes",
        "Roberto Cesar Ribeiro Vieira",
        "Welisson Fideles Marinho",
        "Luís Fernando Rabelo Magalhães",
        "Lucas Marto da Silva",
        "Shesley Werik Gomes de Carvalho",
        "Roger Williams dos Santos e Silva",
        "Luís Mendes da Silva",
        "Douglas Natanyel de Sousa Dias",
        "Mario Jorge Borges Dutra",
        "Francisco Victor Oliveira da Silva"
    ]


    categorias = [
        "AÇOUGUE",
        "ALMOXARIFADO",
        "BAZAR",
        "COMMODITIES",
        "CONFEITARIA e PADARIA",
        "CONGELADOS",
        "FRIOS E RESFRIADOS",
        "HIGIENE E LIMPEZA",
        "HORTIFRUTI",
        "MERCEARIA LÍQUIDA",
        "MERCEARIA SECA DOCE",
        "MERCEARIA SECA SALGADA",
        "PADARIA INDUSTRIALIZADOS",
        "PERFUMARIA",
        "PRODUTOS NATURAIS",
        "RESFRIADOS LÁCTEOS",
        "ROTISSERIE",
        "SALGADOS"
    ]

    loja_list = list(loja_map.values())
    loja_list.insert(0,"Escolha uma loja")

    ativo_list = nomes
    ativo_list.insert(0,"Escolha uma Ativo")

    divi_list = categorias
    divi_list.insert(0,"Escolha uma Divisão")

    col1,col2,col3,col4 = st.columns(4)

    download_is = False
    download_button = None

    lista_ativo = []
    # Streamlit uploader widget
    if uploaded_file:

        with col1:
            list_select = st.selectbox("escolha uma loja",options=loja_list)

        with col2:
            divi_select = st.selectbox("escolha uma Divisão",options=divi_list)
        with col3:
            Tamanho_fonte = st.number_input("Tamanho da fonte",value=7.5)
            
        #download_button = st.button("Carregar o arquivo")

        with col2:
            
            if "Escolha uma loja" not in list_select and "Escolha uma Divisão" not in divi_select :
                file_data = uploaded_file.getvalue()
                
                # Create a unique identifier for the file (could be timestamp or any unique identifier)
                file_name = uploaded_file.name
                file_metadata = f"{list_select} {divi_select}"

                # Insert file metadata and file data into the database
                cursor.execute("INSERT INTO files (store_name, file_path) VALUES (?, ?)", (file_metadata, file_data))
                conn.commit()
                
                # Read the file as binary

                df_puro = pd.read_excel(uploaded_file)  # Read the file into a DataFrame

                df_puro["DESCRICAO"] = df_puro["DESCRICAO"].fillna('Unknown')
                
                df_puro['DIFERENCA'] = df_puro['DIFERENCA'].apply(convert_to_int)
                df_puro['QTD ARQUIVO ESTOQUE'] = df_puro['QTD ARQUIVO ESTOQUE'].apply(convert_to_int)
                df_puro['QTD CONTADA'] = df_puro['QTD CONTADA'].apply(convert_to_int)
                df_puro = df_puro.rename(columns={"DIFERENCA":"DIFERENCA_Iss"})

                df_puro["DIFERENCA"] = df_puro["QTD CONTADA"] - df_puro["QTD ARQUIVO ESTOQUE"] 
                df_puro["DIFERENCA"] = df_puro["DIFERENCA"].round(2)
        
               # st.write(f"estoque {df_puro["QTD ARQUIVO ESTOQUE"].sum()}")
                #st.write(f"Contada {df_puro["QTD CONTADA"].sum()}")
                #st.write(f"Diferenca {df_puro["DIFERENCA"].sum()}")

                df_zerados_True = df_puro[(df_puro["STATUS"] == "DIVERGENTE") & (df_puro["QTD CONTADA"] == 0)]

                df_zerados = df_puro.loc[(df_puro["DIFERENCA"] > 0) & (df_puro["QTD CONTADA"] != 0) ]
                df = df_puro.loc[(df_puro["DIFERENCA"] < 0) & (df_puro["QTD CONTADA"] != 0) ]

                #df = df_puro[["CODIGO","DESCRICAO","DIFERENCA","DESC INFO1","DESC INFO2","VALOR ITEM","VALOR DIF","STATUS","AREAS COLETA","QTD CONTADA"]]
                #df_zerados = df_puro[(df_puro["STATUS"] == "DIVERGENTE") & (df_puro["QTD CONTADA"] == 0)]
                
                df = df[["CODIGO","DESCRICAO","DIFERENCA","DESC INFO1","DESC INFO2","VALOR DIF","AREAS COLETA"]]
                df_zerados = df_zerados[["CODIGO","DESCRICAO","DIFERENCA","DESC INFO1","DESC INFO2","VALOR DIF","AREAS COLETA"]]
                df_zerados_True = df_zerados_True[["CODIGO","DESCRICAO","DIFERENCA","DESC INFO1","DESC INFO2","VALOR DIF","AREAS COLETA"]]

                df_zerados["CODIGO"] = df_zerados["CODIGO"].astype('int')

                df['VALOR DIF'] = df['VALOR DIF'].apply(convert_to_int)
                df_zerados["VALOR DIF"] = df_zerados["VALOR DIF"].apply(convert_to_int)
                df_zerados_True["VALOR DIF"] = df_zerados_True["VALOR DIF"].apply(convert_to_int)

                #st.write(f"Sobras {len(df_zerados["DIFERENCA"])}/ soma {df_zerados["DIFERENCA"].sum()}")
                #st.write(f"Zerados {len(df_zerados_True["DIFERENCA"])}/ soma {df_zerados_True["DIFERENCA"].sum()}")
                #st.write(f"Faltas {len(df["DIFERENCA"])} /soma {df["DIFERENCA"].sum()}")

                df = df.sort_values(by=["VALOR DIF",'DIFERENCA'], ascending=True)
                df_zerados = df_zerados.sort_values(by=["VALOR DIF",'DIFERENCA'], ascending=False)
                df_zerados_True = df_zerados_True.sort_values(by=["VALOR DIF",'DIFERENCA'], ascending=False)
                
                st.write(f"A len do Zerados {len(df_zerados_True)} Soma {df_zerados_True['DIFERENCA'].sum()}")
                st.write(f"A len do Sobras {len(df_zerados)} Soma {df_zerados['DIFERENCA'].sum()}")
                st.write(f"A len das Faltas {len(df)} Soma {df['DIFERENCA'].sum()}")

                st.write(f"{df['DIFERENCA'].sum()+df_zerados['DIFERENCA'].sum()+df_zerados_True['DIFERENCA'].sum()}")
                st.write(f"{df['DIFERENCA'].sum()+df_zerados['DIFERENCA'].sum()}")
                st.write(f"{df['DIFERENCA'].sum()+df_zerados['DIFERENCA'].sum()}")

                df
                df_zerados
                df_zerados_True

                df_faltas_area_top = df_zerados.nlargest(10,"DIFERENCA")
                
                #Titulo
                styles = getSampleStyleSheet()
                titulo_style = styles["Title"]

                #adcionar ao TIle
                titulo = Paragraph('Top Faltas - Divergências' , titulo_style)
                titulo_sobra = Paragraph('Top Sobras - Divergências' , titulo_style)
                titulo_Zerados = Paragraph('Zerados - Divergências' , titulo_style)

                if st.button("Gerar PDF"):

                    def generate_pdf_2(df):
                        pdf_file = "Top Sobras.pdf"
                        data_for_pdf = [df.columns.to_list()] + df.values.tolist()
                        
                        doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
                        elements = []

                        elements.append(titulo)
                        elements.append(Spacer(1,12))

                        table = Table(data_for_pdf,repeatRows=1)

                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.red),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                            ('TOPPADDING', (0, 0), (-1, -1), 0),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                            ('FONTSIZE', (0, 0), (-1, -1), Tamanho_fonte),
                            ('GRID',(0,0),(-1,-1),0.5,colors.black),
                            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.grey)
                        ]))

                        elements.append(table)
                        doc.build(elements)
                        return pdf_file
                    
                    def generate_pdf_Falta(df):

                        pdf_file = "Top Faltas.pdf"
                        data_for_pdf = [df.columns.to_list()] + df.values.tolist()

                        doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
                        elements = []

                        elements.append(titulo_sobra)
                        elements.append(Spacer(1,12))

                        table = Table(data_for_pdf,repeatRows=1)

                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.green),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                            ('TOPPADDING', (0, 0), (-1, -1), 0),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                            ('FONTSIZE', (0, 0), (-1, -1), Tamanho_fonte),
                            ('GRID',(0,0),(-1,-1),0.5,colors.black),
                            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.grey)
                        ]))

                        elements.append(table)
                        doc.build(elements)
                        return pdf_file

                    def generate_pdf_zerados(df):

                        pdf_file = "Top Zerados.pdf"
                        data_for_pdf = [df.columns.to_list()] + df.values.tolist()

                        doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
                        elements = []
                        #x="AREAS COLETA",y=df_zerados_True["DIFERENCA"].abs()
                        grafico_zerado = px.bar(df_faltas_area_top,x='AREAS COLETA', y=df_faltas_area_top["DIFERENCA"].abs(),title ="teste")

                        grafico_zerado.update_traces(marker=dict(line=dict(color='white', width=1)))
                    
                        #grafico_zerado
                        chart_img = "sobras_chart.png"
                        grafico_zerado.write_image(chart_img,scale=2)

                        #elements.append(Image(chart_img,width=600,height=400))
                        elements.append(titulo_Zerados)
                        elements.append(Spacer(1,12))

                        table = Table(data_for_pdf,repeatRows=1)

                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.blue),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                            ('TOPPADDING', (0, 0), (-1, -1), 0),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                            ('FONTSIZE', (0, 0), (-1, -1), Tamanho_fonte),
                            ('GRID',(0,0),(-1,-1),0.5,colors.black),
                            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.grey)
                        ]))

                        elements.append(table)
                        doc.build(elements)
                        return pdf_file

                    # Generate individual PDFs
                    pdf_sobras = generate_pdf_2(df) #ok
                    pdf_faltas = generate_pdf_Falta(df_zerados)
                    pdf_zerados = generate_pdf_zerados(df_zerados_True)#ok?

                    df
                    df_zerados
                    df_zerados_True

                    # Merge the PDFs
                    merger = PdfMerger()

                    for pdf in [pdf_faltas,pdf_sobras,pdf_zerados]:
                        merger.append(pdf)

                    final_pdf = f"Top Perdas e Sobras {file_metadata}.pdf"
                    merger.write(final_pdf)
                    merger.close()

                    # Confirm file exists before trying to open
                    if os.path.exists(final_pdf):
                        with open(final_pdf, "rb") as f:
                            st.download_button("Baixar PDF Unificado", f, file_name=final_pdf, mime="application/pdf")

                        # Optional cleanup
                        os.remove(pdf_zerados)
                        os.remove(pdf_sobras)
                        os.remove(pdf_faltas)
                        os.remove(final_pdf)
                    else:
                        st.error("Erro ao gerar o PDF final. Verifique os arquivos.")

                
with st.expander(":green[ADMINISTRATIVO]"):
                         
            st.title(":green[ARQUIVO A DIVERGÊNCIA DO ISSCOLECTOR PARA SUBIR PARA A CISS (SOMENTE ADMINISTRATIVO)]")
            st.header("IMPORTE A DIVERGÊNCIA DO ISSCOLECTOR (ISSCOLECTOR -> CISS)(SOMENTE ADMINISTRATIVO)")

            Divergencia_is = st.file_uploader("Importe A divergência do ISSCOLECTOR",accept_multiple_files=True)

            if Divergencia_is is not None:
                
                all_dfs = []

                all_dfs.append(Divergencia_is)

                st.write("Cuidado ao renomear os arquivos com caracteres invalidos => | , > , < , * , ? , '(aspas), / , `\` , :")

                for file in Divergencia_is:

                    nome = file.name
                    nome = nome.strip(".xlsx")
                    nome = nome.strip(".")
                    nome = nome.strip(" - ")


                    dfs_Iss = pd.read_excel(file)
                    
                    #df_IsColle = pd.read_excel(Divergencia_is)
                    #df_IsColle
                    # Extract the relevant columns
                    dfs_Iss_txt = dfs_Iss[["CODIGO", "QTD CONTADA"]]
                    
                    # Apply string formatting
                    dfs_Iss_txt["CODIGO"] = dfs_Iss_txt["CODIGO"].apply(lambda x: str(x).zfill(13))
                    
                    # Function to remove double quotes
                    def remove_hyphen(column):
                        return column.str.replace('"', '', regex=False)
                    
                    # Apply the function to remove double quotes
                    dfs_Iss_txt["CODIGO"] = remove_hyphen(dfs_Iss_txt["CODIGO"])
                    
                    # Convert the cleaned DataFrame to CSV format
                    df_IsColle_txt_csv = dfs_Iss_txt.to_csv(sep=" ", index=False, header=True)
                    
                    # Provide the cleaned CSV as a download link
                    st.download_button(
                        label=f"Baixar o Arquivo {nome}",
                        data=df_IsColle_txt_csv,
                        file_name=f"{nome}.txt",
                        mime="text/csv"
                    )
            
            st.title(":blue[ARQUIVO DE ESTOQUE DO CISS (SOMENTE ADMINISTRATIVO)]")
            st.header("IMPORTE O ESTOQUE DO CISS PARA TRATAR O ARQUIVO(SOMENTE ADMINISTRATIVO)")

            #loja_estoque_novo = st.selectbox("escolha uma loja ",options=loja_list)

            #divisão_estoque_novo = st.selectbox("escolha uma Divisão ",options=divi_list)

            #st.title(":green[ARQUIVO A DIVERGÊNCIA DO ISSCOLECTOR PARA SUBIR PARA A CISS (SOMENTE ADMINISTRATIVO)]")

            Estoque_novo = st.file_uploader("Importe o estoque ")
            Tamanho_fonte_2 = st.number_input("Tamanho da fonte",value=9)

            if Estoque_novo is not None:
                
                def generate_pdf(df):

                    pdf_file = "output.pdf"
                    data_for_pdf = [df.columns.to_list()] + df.values.tolist()

                    doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
                    elements = []

                    table = Table(data_for_pdf)
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), -2),
                        ('TOPPADDING', (0, 0), (-1, -1), 0),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('FONTSIZE', (0, 0), (-1, -1), Tamanho_fonte_2)
                    ]))

                    elements.append(table)
                    doc.build(elements)
                    return pdf_file
                
                df_estoque_novo = pd.read_excel(Estoque_novo)
                df_estoque_novo["emp.nome"] = df_estoque_novo["estoque_sintetico_idempresa"]
                df_estoque_novo["emp.nome"].replace(loja_map,inplace=True)

                divisãolist = df_estoque_novo["descrdivisao"].unique()
                divisãolist = ",".join(divisãolist)

                lojalist = df_estoque_novo["estoque_sintetico_idempresa"].unique()
                lojalist = "".join(str(lojalist))
                lojalist = lojalist.strip("[]")

                #df_estoque_geral["loja"].replace(loja_map,inplace=True)

                lojaname = df_estoque_novo["emp.nome"].unique()
                lojaname = "".join(str(lojaname))
                lojaname = lojaname.strip("[]")
            
                df_estoque_novo= df_estoque_novo[["produtos_view_idsubproduto","produtos_view_idcodbarprod","produtos_view_descricaoproduto","produtos_view_embalagementrada","secao_descrsecao"]]
                df_estoque_novo = df_estoque_novo.rename(columns={"produtos_view_idsubproduto":"Cod.Int","produtos_view_idcodbarprod":"Barra","produtos_view_descricaoproduto":"descr.","produtos_view_embalagementrada" :"Emb.compr","secao_descrsecao":"Seção"})
                
                sort_value = ["descr.","Seção"]

                list_box = st.selectbox("Ordenar por ",options=sort_value)

                df_estoque_novo = df_estoque_novo.sort_values(by=list_box)

                df_estoque_novo

                if st.button("Gerar PDF "):
                
                    pdf_file = generate_pdf(df_estoque_novo)

                    st.success(f"PDF Gerado: {pdf_file}")

                    # Provide download link for PDF
                    with open(pdf_file, "rb") as f:
                        st.download_button("Baixar PDF", f, file_name=f"Esto. {lojaname} Div - {divisãolist}.pdf", mime="application/pdf")
                    
                    # Clean up by removing the generated PDF file
                    os.remove(pdf_file)



            Estoque_seg_master = st.file_uploader("Arquivo Sg master Apos ter copiado em um excel", type=["xls", "xlsx","csv"])


            if Estoque_seg_master is not None:
                df_sg_master = pd.read_excel(Estoque_seg_master,header=None)
                df_sg_master.columns = ["Código", "Produto", "Cód.barras", "Qtd","Preço de Custo","Perc. Lucro","Preço de Venda","Ncm","Preço de Revenda "]

                        #df_estoque = df_estoque_geral[["produtos_view_idcodbarprod", "estoque_sintetico_qtdatualestoque","custonotafiscal"]]

                df_estoque_sg = df_sg_master[["Cód.barras","Qtd","Preço de Custo"]]
                df_produtos_sg = df_sg_master[["Cód.barras","Produto"]]
                df_produtos_Internos_sg = df_sg_master[["Cód.barras","Código"]]

                df_estoque_sg['Cód.barras'] = df_estoque_sg['Cód.barras'].apply(lambda x: int(x) if pd.notna(x) else 0)  # Replace NaN with 0
                df_estoque_sg['Cód.barras'] = df_estoque_sg['Cód.barras'].round(0).astype(int)

                df_produtos_sg['Cód.barras'] = df_produtos_sg['Cód.barras'].apply(lambda x: int(x) if pd.notna(x) else 0)  # Replace NaN with 0
                df_produtos_sg['Cód.barras'] = df_produtos_sg['Cód.barras'].round(0).astype(int)

                df_produtos_Internos_sg['Cód.barras'] = df_produtos_Internos_sg['Cód.barras'].apply(lambda x: int(x) if pd.notna(x) else 0)  # Replace NaN with 0
                df_produtos_Internos_sg['Cód.barras'] = df_produtos_Internos_sg['Cód.barras'].round(0).astype(int)


                csv_estoque_sg = df_estoque_sg.to_csv(sep=';', index=False, header=True)
                csv_produtos_sg = df_produtos_sg.to_csv(sep=';', index=False, header=True)
                csv_produtos_2_sg = df_produtos_Internos_sg.to_csv(sep=';', index=False, header=True)


                if csv_estoque_sg:
                    # Create a download button for the 'estoque' CSV
                    st.download_button(
                        label=f"Estoque Posto",
                        data=csv_estoque_sg,
                        file_name=f"Inventario_loja_{datetime.today().strftime('%d-%m-%Y')}_estoque.txt",
                        mime="text/csv"
                    )

                if csv_produtos_sg:
                    # Create a download button for the 'produtos' CSV
                    st.download_button(
                        label=f"Produtos Posto",
                        data=csv_produtos_sg,
                        file_name=f"Inventario_loja_{datetime.today().strftime('%d-%m-%Y')}_produtos.txt",
                        mime="text/csv"
                    )

                if csv_produtos_2_sg:
                    st.download_button(
                        label=f"Produtos Internos Posto",
                        data=csv_produtos_2_sg,
                        file_name=f"Inventario_loja_{datetime.today().strftime('%d-%m-%Y')}_produtos Internos (opcional) .txt",
                        mime="text/csv"
                    )
                df_sg_master