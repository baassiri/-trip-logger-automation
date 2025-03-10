import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from invoice_automation import force_update_trip_log
from config import FILE_PATH

# Set Streamlit page config
st.set_page_config(page_title="Trip Logger", layout="wide")

# Title
st.title("🚛 Trip Logger Automation")

# Input Fields
client_name = st.text_input("Enter Client Name")
client_address = st.text_input("Enter Client Address")

# Button to log the trip
if st.button("Log Trip"):
    if client_name and client_address:
        detected_clients = [client_name]
        detected_addresses = {client_name: client_address}
        force_update_trip_log(detected_clients, detected_addresses)
        st.success(f"✅ Trip logged for {client_name}")
    else:
        st.warning("⚠️ Please enter both Client Name and Address.")

# Load and display the trip logs
st.subheader("📋 Current Trip Logs")

try:
    wb = load_workbook(FILE_PATH, data_only=True)
    ws = wb["TRIP LOGS"]

    data = []
    for row in ws.iter_rows(min_row=7, values_only=True):
        if any(row):
            data.append(row)

    df = pd.DataFrame(data, columns=["Date", "Client", "Base", "Home", "Destination 1", "Destination 2", "Destination 3", "Destination 4", "Destination 5"])
    st.dataframe(df)

    # Export options
    st.download_button("📥 Download as Excel", df.to_csv(index=False), "trip_logs.csv", "text/csv")
except Exception as e:
    st.error(f"⚠️ Could not load trip logs: {e}")

# Run with: streamlit run src/streamlit_app.py
