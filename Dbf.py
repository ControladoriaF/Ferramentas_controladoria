import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import pymysql
import uuid
from datetime import datetime
from dbfread import DBF
from urllib.parse import quote_plus
import hashlib
from io import BytesIO


st.title("Upload and View .DBF File")

uploaded_file = st.file_uploader("Choose a .dbf file", type=["dbf"])


host = st.text_input("Host", value="localhost")
user = st.text_input("User", value="root")
password = st.text_input("Password", type="password")
database = st.text_input("Database name", value="ferreiradados")
table_name = "DataDash"
password_encoded = quote_plus("Ferreira321@")
   
if uploaded_file is not None:
    with open("temp_file.dbf", "wb") as f:
        f.write(uploaded_file.read())

    table = DBF("temp_file.dbf", load=True, encoding='latin-1')
    df = pd.DataFrame(iter(table))


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

    uploaded_file.seek(0)
    file_hash = hashlib.md5(uploaded_file.read()).hexdigest()
    uploaded_file.seek(0)  # Reset again for reading

    # Clean column names
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    st.success("File loaded successfully!")
    st.dataframe(df)

    if st.button("Execute SQL Script"):
        # Add timestamp and file_id
        df["upload_time"] = pd.Timestamp.now()
        df["file_id"] = file_hash  # Use consistent hash as file ID

        st.subheader("Data to Upload")
        st.dataframe(df)

        try:
    # Create engine
            conn_str = f"mysql+pymysql://root:{password_encoded}@localhost/ferreiradados"
            engine = create_engine(conn_str)

            st.write("🔌 Connecting to:", conn_str)
            # Upload and create table if it doesn't exist
            df.to_sql(name=table_name, con=engine, if_exists='replace', index=False)

            st.success(f"✅ Table '{table_name}' created and data uploaded successfully!")

        except Exception as e:
            st.error(f"❌ Upload failed: {e}")


            with engine.connect() as conn:
                # Check for duplicate uploads
                if engine.dialect.has_table(conn, table_name):
                    existing_ids = pd.read_sql(f"SELECT DISTINCT file_id FROM {table_name}", conn)
                    if df["file_id"].iloc[0] in existing_ids["file_id"].values:
                        st.warning("This file has already been uploaded.")
                    else:
                        df.to_sql(name=table_name, con=engine, if_exists='append', index=False)
                        st.success("✅ Data uploaded successfully!")
                else:
                    # If table doesn't exist, create it and insert
                    df.to_sql(name=table_name, con=engine, if_exists='replace', index=False)
                    st.success("✅ Table created and data uploaded successfully!")

        except Exception as e:
            st.error(f"❌ Upload failed: {e}")
