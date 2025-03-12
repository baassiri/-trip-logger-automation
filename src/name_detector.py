import re
from typing import List, Tuple

class NameDetector:
    def __init__(self) -> None:
        """
        Initialize the NameDetector with regex patterns for detecting names and addresses.
        """
        # Pattern for names: matches one or more words starting with an uppercase letter followed by lowercase letters.
        self.name_pattern = re.compile(r"\b[A-Z][a-z]+(?:\s[A-Z][a-z]+)*\b")
        # Pattern for addresses: matches typical USA address formats like "123 Main St, City, ST 12345"
        self.address_pattern = re.compile(r"\d{1,5}\s[\w\s]+,\s[\w\s]+,\s[A-Z]{2}\s?\d{5}")

    def extract_names_and_addresses(self, message: str) -> Tuple[List[str], List[str]]:
        """
        Extracts names and addresses from a provided message.

        Args:
            message (str): The input text containing potential names and addresses.

        Returns:
            Tuple[List[str], List[str]]: A tuple where the first element is a list of names
            and the second element is a list of addresses found in the message.
        """
        names = self.name_pattern.findall(message)
        addresses = self.address_pattern.findall(message)
        return names, addresses
