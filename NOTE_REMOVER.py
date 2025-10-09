import logging
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import time as time_module
from datetime import datetime, timedelta
import xlwings as xw
import time
import random
from functools import wraps
import os
import argparse

# Import all config and variables from ATP_common_config.py
import ATP_common_config as config

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def delete_note_events(year, rip_word, verbose=False):
    """Deletes NOTE events containing `rip_word` for the specified year."""
    start_date = datetime(year, 1, 1).strftime("%Y-%m-%dT00:00:00")
    end_date = datetime(year, 12, 31).strftime("%Y-%m-%dT23:59:59")
    url_get = f"{config.url_base}/events.json"
    params = {"oldest": start_date, "newest": end_date, "category": "NOTE"}
    headers = config.API_headers

    try:
        resp = requests.get(url_get, headers=headers, params=params, auth=HTTPBasicAuth(config.username, config.api_key))
        resp.raise_for_status()
        events = resp.json()
        if verbose:
            logging.info(f"Fetched {len(events)} NOTE events for {year}")
        deleted = 0
        for event in events:
            if rip_word.lower() in event['name'].lower():
                event_id = event['id']
                url_del = f"{config.url_base}/events/{event_id}"
                del_resp = requests.delete(url_del, headers=headers, auth=HTTPBasicAuth(config.username, config.api_key))
                if del_resp.ok:
                    deleted += 1
                    logging.info(f"Deleted event ID={event_id} - Name: {event['name']}")
                else:
                    logging.error(f"Failed to delete event ID={event_id} - Status: {del_resp.status_code}")
        print(f"Deleted {deleted} events containing '{rip_word}' for year {year}.")
    except Exception as e:
        logging.error(f"Error processing notes: {e}")
        print("Failed to process/delete events.")

def main():
    parser = argparse.ArgumentParser(description="Delete NOTES containing a specific word for a given year.")
    parser.add_argument("--year", type=int, help="Year to check (e.g., 2026). If not provided, prompts interactively.")
    parser.add_argument("--rip_word", type=str, help="Word to match in NOTES to delete. If not provided, prompts interactively.")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging.")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO if args.verbose else logging.WARNING, format='%(asctime)s - %(levelname)s - %(message)s')

    # Prompt interactively if not supplied
    year = args.year if args.year else int(input("Year to check for NOTE events to delete? "))
    rip_word = args.rip_word if args.rip_word else input("Word to search for in NOTE events to delete (rip_word)? ")

    delete_note_events(year, rip_word, args.verbose)

if __name__ == "__main__":
    main()
