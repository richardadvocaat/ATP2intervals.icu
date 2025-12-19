from ATP_common_config import *

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
# map API categories to short labels used in sheet/sorting
CATEGORY_MAP = {"RACE_A": "A", "RACE_B": "B", "RACE_C": "C"}
API_RACE_CATEGORIES = list(CATEGORY_MAP.keys())

API_HEADERS = {"Content-Type": "application/json"}


def read_user_data(ATP_file_path, sheet_name="User_Data"):
    df = pd.read_excel(ATP_file_path, sheet_name=sheet_name)
    return df.set_index("Key")["Value"].to_dict()


def get_race_events(athlete_id: str, username: str, api_key: str, oldest: str, newest: str):
    """
    Fetch events for all categories and return a flat list of event dicts.
    Each event will have a 'category' key (taken from API or the requested category).
    """
    url = f"https://intervals.icu/api/v1/athlete/{athlete_id}/eventsjson"
    all_events = []
    for cat in API_RACE_CATEGORIES:
        params = {"oldest": oldest, "newest": newest, "category": cat}
        logging.info("Requesting %s params=%s", url, params)
        resp = requests.get(url, headers=API_HEADERS, params=params, auth=HTTPBasicAuth(username, api_key))
        logging.info("Status %s", resp.status_code)
        if resp.status_code == 200:
            try:
                events = resp.json()
                for e in events:
                    # ensure category is present so we can map to short label later
                    e.setdefault("category", cat)
                    all_events.append(e)
                logging.info("Fetched %d events for %s", len(events), cat)
            except Exception as ex:
                logging.error("JSON parse error for %s: %s", cat, ex)
        else:
            logging.error("Failed to fetch %s: %s %s", cat, resp.status_code, resp.text[:200])
    return all_events


def events_to_dataframe(events):
    if not events:
        return pd.DataFrame(columns=["date", "racename", "racetype", "racecategory"])
    df = pd.DataFrame(events)
    # pick/rename expected fields (end_date_local -> date)
    df = df.rename(columns={"end_date_local": "date", "name": "racename", "type": "racetype", "category": "racecategory"})
    # map API categories to short labels (if present)
    df["racecategory"] = df["racecategory"].astype(str).map(CATEGORY_MAP).fillna(df["racecategory"])
    # parse date to pandas datetime (keep datetime so Excel sees date type)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    # keep only columns we want in the desired order
    df = df[["date", "racename", "racetype", "racecategory"]]
    return df


def save_all_races_sheet(df: pd.DataFrame, output_file: str, sheet_name: str = "All_Races"):
    """Write a single combined sheet sorted by racecategory and date, format date as short Dutch style."""
    app = xw.App(visible=False)
    try:
        wb = xw.Book()
        # sort
        if not df.empty:
            df = df.sort_values(by=["racecategory", "date", "racename"]).reset_index(drop=True)
        # remove existing sheet if present, then add new
        if sheet_name in [s.name for s in wb.sheets]:
            wb.sheets[sheet_name].delete()
        sht = wb.sheets.add(sheet_name)
        # write header and values
        sht.range("A1").value = df.columns.tolist()
        sht.range("A2").value = df.values.tolist()
        # apply Excel short date / Dutch style (dd-mm-yyyy) to the date column
        nrows = len(df)
        if nrows > 0:
            # A2:A{nrows+1}
            date_range = sht.range(f"A2:A{nrows+1}")
            try:
                # try locale-aware short date first (may depend on environment)
                date_range.number_format = "Short Date"
            except Exception:
                date_range.api.NumberFormat = "dd-mm-yyyy"
            # enforce explicit dd-mm-yyyy as a guarantee
            date_range.api.NumberFormat = "dd-mm-yyyy"
        # remove default Sheet1 if still present
        if "Sheet1" in [s.name for s in wb.sheets] and len(wb.sheets) > 1:
            wb.sheets["Sheet1"].delete()
        wb.save(output_file)
        logging.info("Saved combined races to %s", output_file)
    finally:
        wb.close()
        app.quit()


def main():
    user_data = read_user_data(ATP_file_path)
    api_key = user_data.get("API_KEY")
    username = user_data.get("USERNAME")
    athlete_id = user_data.get("ATHLETE_ID")
    oldest_date = f"{ATP_year}-01-01T00:00:00"
    newest_date = f"{ATP_year}-12-31T23:59:59"

    events = get_race_events(athlete_id, username, api_key, oldest_date, newest_date)
    df = events_to_dataframe(events)
    if df.empty:
        print("No RACE events found.")
        return

    save_all_races_sheet(df, RACE_file_path)
    print(f"All races (combined) saved to {RACE_file_path}")


if __name__ == "__main__":
    main()
