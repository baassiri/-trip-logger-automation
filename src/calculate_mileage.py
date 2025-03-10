import requests

def get_distance(api_key, origin, destination):
    """Fetch distance from Google Maps API."""
    url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origin}&destinations={destination}&key={api_key}"
    response = requests.get(url).json()

    try:
        distance_miles = response["rows"][0]["elements"][0]["distance"]["value"] / 1609.34
        return round(distance_miles, 1)
    except KeyError:
        print("⚠️ Error retrieving distance.")
        return 0
