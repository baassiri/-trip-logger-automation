import json
from datetime import datetime

def json_serial(obj):
    """
    JSON serializer for objects not serializable by default json code.
    Converts datetime objects to ISO formatted strings.
    """
    if isinstance(obj, datetime):
        return obj.isoformat()
    raise TypeError("Type not serializable")

def load_json(json_string):
    """
    Parse a JSON string and return a Python object (dict or list).
    
    :param json_string: A JSON formatted string.
    :return: Python dict or list.
    """
    try:
        return json.loads(json_string)
    except json.JSONDecodeError as e:
        print(f"❌ JSON parse error: {e}")
        return None

def to_json(data, indent=4):
    """
    Convert a Python object (dict, list, etc.) into a JSON formatted string.
    
    :param data: Python object to serialize.
    :param indent: Indentation level for pretty printing (default 4).
    :return: A JSON string.
    """
    try:
        return json.dumps(data, indent=indent, default=json_serial)
    except TypeError as e:
        print(f"❌ JSON conversion error: {e}")
        return None

if __name__ == "__main__":
    # Example usage:
    sample_data = {
        "name": "John Doe",
        "registered": datetime.now(),
        "active": True,
        "numbers": [1, 2, 3]
    }
    json_str = to_json(sample_data)
    print("Converted JSON:", json_str)
    
    parsed = load_json(json_str)
    print("Parsed JSON:", parsed)
