from invoice_automation import force_update_trip_log

if __name__ == "__main__":
    # Ask for client name
    client_name = input("Enter client name: ").strip()

    addresses = []

    # Continuously ask for addresses until the user types 'done'
    while True:
        print("\n🔹 Enter Address Details (or type 'done' to finish):")
        address1 = input("Address Line 1: ").strip()
        if address1.lower() == "done":
            break

        address2 = input("Address Line 2 (Optional): ").strip()
        city = input("City: ").strip()
        state = input("State (2-letter code): ").strip().upper()
        zip_code = input("ZIP Code: ").strip()

        if not address1 or not city or not state or not zip_code:
            print("⚠️ Address, City, State, and ZIP Code are required.")
            continue  # Ask again if any field is missing

        # Format the full address properly
        full_address = f"{address1}, {address2 + ', ' if address2 else ''}{city}, {state} {zip_code}"
        addresses.append(full_address)

    if not client_name or not addresses:
        print("⚠️ Client name and at least one valid address are required.")
    else:
        detected_clients = [client_name]
        detected_addresses = {client_name: addresses}  # Store addresses in a dictionary

        # Write data to Excel
        force_update_trip_log(detected_clients, detected_addresses)

        print("\n✅ All destinations have been logged successfully!")
