from ATP_common_config import *
import os
from pathlib import Path

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


def save_all_races_sheet(df: pd.DataFrame, output_file: str, sheet_name: str = "Races"):
    """Write a single combined sheet sorted by racecategory and date, but preserve all other sheets.
    If the sheet already exists, overwrite its contents in-place and clear any leftover rows below.
    If the sheet does not exist, create it (without deleting other sheets).
    If no changes are detected, do nothing.
    """
    app = xw.App(visible=False)
    wb = None
    try:
        output_path = str(Path(output_file).resolve())

        # Open existing workbook if present; otherwise create a new workbook and save it.
        if os.path.exists(output_path):
            # opening by full path avoids creating a new workbook and overwriting existing file
            wb = xw.Book(output_path)
        else:
            wb = xw.Book()  # new workbook in current app
            # Save immediately so the workbook has a path (but only when file doesn't exist)
            try:
                wb.save(output_path)
            except Exception as e:
                logging.warning("Could not save new workbook to %s immediately: %s", output_path, e)

        # sort incoming DataFrame for a deterministic layout
        if not df.empty:
            df_to_write = df.sort_values(by=["racecategory", "date", "racename"]).reset_index(drop=True)
        else:
            df_to_write = pd.DataFrame(columns=["date", "racename", "racetype", "racecategory"])

        headers = df_to_write.columns.tolist()
        values = df_to_write.values.tolist()
        new_nrows = len(df_to_write)

        # get or create sheet (do NOT delete any existing sheets)
        sheet_names = [s.name for s in wb.sheets]
        if sheet_name in sheet_names:
            sht = wb.sheets[sheet_name]
        else:
            # add after last sheet to avoid reordering default sheets
            try:
                sht = wb.sheets.add(sheet_name, after=wb.sheets[-1])
            except Exception:
                sht = wb.sheets.add(sheet_name)

        # Read existing sheet contents (if any) to detect changes and to obtain old row count
        existing_data = None
        try:
            top_left = sht.range("A1").value
            if top_left is None:
                existing_data = None
            else:
                table = sht.range("A1").expand("table").value
                if table and len(table) >= 1:
                    existing_headers = table[0]
                    existing_rows = table[1:] if len(table) > 1 else []
                    # If rows are a single row, ensure we wrap it properly
                    existing_data = pd.DataFrame(existing_rows, columns=existing_headers) if existing_rows else pd.DataFrame(columns=existing_headers)
                else:
                    existing_data = None
        except Exception:
            existing_data = None

        # Normalize existing_data to match df_to_write column names
        if existing_data is not None and not existing_data.empty:
            existing_data = existing_data.rename(columns=lambda c: c.strip())
            if "date" in existing_data.columns:
                existing_data["date"] = pd.to_datetime(existing_data["date"], errors="coerce")
            for col in headers:
                if col not in existing_data.columns:
                    existing_data[col] = pd.NA
            existing_data = existing_data[headers]
        else:
            existing_data = pd.DataFrame(columns=headers)

        # Compare existing_data with df_to_write; if equal, skip writing
        def df_for_cmp(dframe):
            if dframe.empty:
                return pd.DataFrame()
            comp = dframe.copy()
            if "date" in comp.columns:
                comp["date"] = comp["date"].apply(lambda x: pd.NaT if pd.isna(x) else pd.to_datetime(x))
                comp["date"] = comp["date"].dt.strftime("%Y-%m-%dT%H:%M:%S").fillna("")
            comp = comp.fillna("").astype(str)
            return comp.reset_index(drop=True)

        existing_cmp = df_for_cmp(existing_data)
        new_cmp = df_for_cmp(df_to_write)

        if existing_cmp.equals(new_cmp):
            logging.info("No changes detected in %s sheet; nothing to update.", sheet_name)
            try:
                wb.save(output_path)
            except Exception:
                pass
            return

        # Overwrite header and new data starting at A1 / A2
        sht.range("A1").value = headers
        sht.range("A2").value = values

        # Clear leftover rows/columns if the previous sheet had more content than we now have
        old_nrows = 0
        old_ncols = 0
        try:
            used = sht.api.UsedRange
            if used is not None:
                old_total_rows = int(used.Rows.Count)
                old_total_cols = int(used.Columns.Count)
                old_nrows = max(0, old_total_rows - 1)  # exclude header
                old_ncols = old_total_cols
        except Exception:
            old_nrows = 0
            old_ncols = len(headers)

        # Clear leftover rows beyond new_nrows (only within the used columns)
        if old_nrows > new_nrows:
            start_row = new_nrows + 2
            end_row = old_nrows + 1
            # Determine number of columns to clear: use max(existing cols, new cols)
            cols_to_clear = max(old_ncols, len(headers))
            # Build column letter range for clearing (A..)
            # safe-guard for many columns
            last_col_letter = xw.utils.col_name(cols_to_clear) if cols_to_clear > 0 else "D"
            clear_range_str = f"A{start_row}:{last_col_letter}{end_row}"
            try:
                sht.range(clear_range_str).clear_contents()
            except Exception:
                # fallback: write empty strings row-by-row
                empty_row = [""] * cols_to_clear
                for r in range(start_row, end_row + 1):
                    try:
                        sht.range(f"A{r}").options(transpose=False).value = empty_row
                    except Exception:
                        pass

        # Also clear any leftover columns to the right of our headers in the header row (if present)
        try:
            used = sht.api.UsedRange
            if used is not None:
                existing_cols = int(used.Columns.Count)
                if existing_cols > len(headers):
                    start_col = xw.utils.col_name(len(headers) + 1)
                    end_col = xw.utils.col_name(existing_cols)
                    clear_cols_range = f"{start_col}1:{end_col}{old_nrows + 1 if old_nrows > 0 else 1}"
                    try:
                        sht.range(clear_cols_range).clear_contents()
                    except Exception:
                        pass
        except Exception:
            pass

        # apply Excel short date / Dutch style (dd-mm-yyyy) to the date column
        if new_nrows > 0:
            date_range = sht.range(f"A2:A{new_nrows+1}")
            try:
                date_range.number_format = "Short Date"
            except Exception:
                try:
                    date_range.api.NumberFormat = "dd-mm-yyyy"
                except Exception:
                    pass
            try:
                date_range.api.NumberFormat = "dd-mm-yyyy"
            except Exception:
                pass

        # Save workbook (do not delete any sheets)
        try:
            wb.save(output_path)
        except Exception as e:
            logging.warning("Could not save workbook to %s: %s", output_path, e)

        logging.info("Saved combined races to %s (updated in-place).", output_path)
    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass


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
