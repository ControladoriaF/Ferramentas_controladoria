import streamlit as st
from ofxparse import OfxParser
import io
import pandas as pd
from io import BytesIO
import numpy as np

def convert_to_int(value):
        try:
            # Replace commas with periods for decimal values
            value = str(value).replace(',', '.')
            
            # Attempt to convert to float first, then to int
            return float(value)
        except (ValueError, TypeError):
            # If it fails, return NaN or the original value
            return np.nan  # You can return np.nan or the original value if needed

st.title("OFX File Reader")

uploaded_file = st.file_uploader("Upload your .ofx file", type=["ofx"])

if uploaded_file is not None:
    try:
        ofx = OfxParser.parse(io.StringIO(uploaded_file.read().decode('ISO-8859-1')))

        account = ofx.account
        st.subheader("Account Information")

        st.write(f"Bank ID: {getattr(account, 'bank_id', 'N/A')}")
        st.write(f"Account ID: {getattr(account, 'account_id', 'N/A')}")
        st.write(f"Account Type: {getattr(account, 'account_type', 'N/A')}")
        st.write(f"Currency: {getattr(ofx, 'currency', 'N/A')}")

        # Show transactions
        transactions = account.statement.transactions
        if transactions:
            data = [{
                "Date": txn.date,
                "Amount": txn.amount,
                "Type": txn.type,
                "Memo": txn.memo
            } for txn in transactions]

            df = pd.DataFrame(data)

            df['Amount'] = df['Amount'].apply(convert_to_int)

            st.subheader("Transactions")
            st.dataframe(df)
            st.write(df["Type"].unique())
            csv = df.to_csv(index=False).encode("utf-8")

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
            

            st.download_button("Download CSV", csv, "transactions.csv", "text/csv")
        else:
            st.warning("No transactions found in the OFX file.")

    except Exception as e:
        st.error(f"Failed to parse OFX file: {e}")
