from invoice_automation import force_update_trip_log

test_clients = ["Test Client"]
test_addresses = {"Test Client": ["123 Test St, Test City, TC 12345"]}

print("Running force_update_trip_log()...")
force_update_trip_log(test_clients, test_addresses)

print("✅ Function executed! Now check if the data appeared in TRIP LOGS.")
