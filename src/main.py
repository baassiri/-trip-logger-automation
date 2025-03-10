# src/main.py
from invoice_automation import force_update_trip_log

if __name__ == "__main__":
    # Ask for client name and address
    client_name = input("Enter client name: ").strip()
    client_address = input("Enter client address: ").strip()

    if not client_name or not client_address:
        print("⚠️ Client name and address cannot be empty.")
    else:
        detected_clients = [client_name]
        detected_addresses = {client_name: client_address}

        # Write data to Excel
        force_update_trip_log(detected_clients, detected_addresses)
