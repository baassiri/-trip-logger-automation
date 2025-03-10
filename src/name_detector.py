# src/name_detector.py
import re

class NameDetector:
    def __init__(self):
        # Example regex patterns for names/addresses
        self.name_pattern = re.compile(r"\b[A-Z][a-z]+(?:\s[A-Z][a-z]+)*\b")
        self.address_pattern = re.compile(r"\d{1,5}\s[\w\s]+,\s[\w\s]+,\s[A-Z]{2}\s?\d{5}")

    def extract_names_and_addresses(self, message):
        """
        Extract names and addresses from a message.
        Return: (list_of_names, list_of_addresses)
        """
        names = self.name_pattern.findall(message)
        addresses = self.address_pattern.findall(message)
        return names, addresses
