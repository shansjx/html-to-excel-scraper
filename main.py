import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
from dotenv import load_dotenv
from datetime import datetime
import json

def output_result(status, scraped_rows=0, updated_rows=0, output_file="", message=""):
    """Output structured result for UiPath integration using JSON format"""
    result = {
        "status": status,
        "scraped_rows": scraped_rows,
        "updated_rows": updated_rows,
        "output_file": output_file,
        "message": message,
        "timestamp": datetime.now().isoformat()
    }
    
    # Save JSON result to file for UiPath to read
    json_output_file = "scraped_data_result.json"
    try:
        with open(json_output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
        print(f"JSON_RESULT_FILE={json_output_file}")
    except Exception as e:
        print(f"ERROR: Could not write JSON result to {json_output_file}: {e}")

    # Original output format
    if scraped_rows > 0:
        print(f"SCRAPED_ROWS={scraped_rows}")
    if updated_rows > 0:
        print(f"UPDATED_ROWS={updated_rows}")
    if output_file:
        print(f"OUTPUT_FILE={output_file}")
    if message:
        print(message)

def scrape_data_from_html(html, master_excel_path=None):
    # print("[DEBUG] Entered scrape_data_from_html")
    # print(f"[DEBUG] master_excel_path: {master_excel_path}")
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        print("No table found to scrape")
        return pd.DataFrame(), False
    rows = table.find_all("tr")
    data = []
    if len(rows) <= 1:
        print("No data to scrape in table")
        return pd.DataFrame(), False
    # print(f"[DEBUG] Rows found in HTML: {len(rows)}")
    latest_timestamp = None
    master_df = None
    appended = False
    if master_excel_path and os.path.exists(master_excel_path):
        # print(f"[DEBUG] Master Excel exists at {master_excel_path}")
        master_df = pd.read_excel(master_excel_path)
        # print(f"[DEBUG] Master DataFrame columns: {master_df.columns.tolist()}")
        if not master_df.empty and "FirstCol" in master_df.columns:
            master_df["FirstCol"] = pd.to_datetime(master_df["FirstCol"], errors="coerce")
            latest_timestamp = master_df["FirstCol"].max()
            # print(f"Latest timestamp in master: {latest_timestamp}")
        else:
            print("[DEBUG] Master DataFrame is empty or missing 'FirstCol'")
    else:
        print(f"Master Excel does not exist at {master_excel_path}")

    current_year = datetime.now().year
    for row in rows:
        columns_list = row.find_all("td")
        num_of_columns = len(columns_list)
        if num_of_columns <= 1:
            continue
        timestamp_str = columns_list[1].text.strip() if num_of_columns > 1 else ""
        try:
            row_timestamp = datetime.strptime(f"{current_year} {timestamp_str}", "%Y %d %b, %H:%M")
            if latest_timestamp and row_timestamp <= latest_timestamp:
                continue
            item = {
                "FirstCol": row_timestamp,
                "SecondCol": columns_list[2].text.strip() if num_of_columns > 2 else "",
                "ThirdCol": columns_list[3].text.strip() if num_of_columns > 3 else "",
                "FourthCol": columns_list[4].text.strip() if num_of_columns > 4 else "",
            }
            if any([item["FirstCol"], item["SecondCol"], item["ThirdCol"], item["FourthCol"]]):
                data.append(item)
        except ValueError:
            continue
    print(f"SCRAPED_ROWS={len(data)}")
    if data:
        new_df = pd.DataFrame(data)
        # Reverse the new data rows so oldest entries at the top (since scraped data is most recent first)
        new_df = new_df.iloc[::-1].reset_index(drop=True)

        # Fill new data into the available empty slots in the 4 main columns, in the rows below the latest timestamp
        if master_df is not None and not master_df.empty:
            # Find the index of the row with the latest timestamp and insert after it
            insert_idx = master_df[master_df["FirstCol"] == latest_timestamp].index
            if len(insert_idx) > 0:
                insert_idx = insert_idx[-1] + 1
            else:
                insert_idx = len(master_df)

            # Only update as many rows as there are new data items, and only if those cells are empty
            num_newrows = len(new_df)
            num_rows = len(master_df)
            updated = 0
            for i in range(num_newrows):
                target_idx = insert_idx + i
                if target_idx >= num_rows:
                    print(f"[DEBUG] Not enough empty rows to fill new data at index {target_idx}, skipping.")
                    break

                slot_filled = 0
                total_cols = 0
                # Check if the target row has missing value or empty string slots in any of the 4 columns
                for col in ["FirstCol", "SecondCol", "ThirdCol", "FourthCol"]:
                    total_cols += 1
                    # to avoid dtype issues by Pandas
                    if col in master_df.columns:
                        master_df[col] = master_df[col].astype('object')
                    if pd.isna(master_df.at[target_idx, col]) or master_df.at[target_idx, col].strip() == "":
                        master_df.at[target_idx, col] = new_df.iloc[i][col]
                        slot_filled += 1
                
                # Count this row as updated if we filled all slots in it
                if slot_filled == total_cols:
                    updated += 1
            if updated > 0:
                try:
                    master_df.to_excel(master_excel_path, index=False)
                    # print(f"SUCCESSFULLY FILLED {updated} ROWS AFTER LATEST TIMESTAMP IN {master_excel_path}")
                    print(f"UPDATED_ROWS={updated}")
                    appended = True
                    # print("[DEBUG] Fill successful, wrote to Excel.")
                except PermissionError:
                    print(f"ERROR: Could not write to {master_excel_path}. File is currently either open or locked")
                except Exception as e:
                    print(f"ERROR: Could not write to {master_excel_path}: {e}")
            else:
                print("NO_EMPTY_SLOTS")
        return new_df, appended
    else:
        return pd.DataFrame(), False

def save_df_to_excel(df, output_file="scraped_data.xlsx"):
    if not df.empty:
        try:
            df.to_excel(output_file, index=False)
            print(f"OUTPUT_FILE={output_file}")
        except PermissionError:
            print(f"ERROR: Could not write to {output_file}. File is currently either open or locked")
        except Exception as e:
            print(f"ERROR: Could not write to {output_file}: {e}")

    else:
        print("NO_DATA")

def cleanup_scraped_file(output_file="scraped_data.xlsx"):
    if os.path.exists(output_file):
        os.remove(output_file)
        print(f"Deleted {output_file}")
    else:
        print(f"{output_file} does not exist.")

if __name__ == "__main__":
    import sys
    if "--cleanup" in sys.argv:
        cleanup_scraped_file()
        output_result("cleanup_completed", message="Cleaned up scraped_data.xlsx")
        sys.exit(0)
    
    html_file = sys.argv[1] if len(sys.argv) > 1 else None
    master_excel = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not html_file or not os.path.exists(html_file):
        output_result("error", message="NO_HTML_FILE")
        sys.exit(1)
    
    with open(html_file, encoding="utf-8") as f:
        html = f.read()
    
    if master_excel:
        df, appended = scrape_data_from_html(html, master_excel)
        if appended: # only save if new data was appended
            save_df_to_excel(df)
            output_result("success", 
                         scraped_rows=len(df), 
                         updated_rows=len(df), 
                         output_file="scraped_data.xlsx",
                         message="Data successfully scraped and updated")
        else:
            output_result("no_new_data", message="NO_NEW_DATA")
    else:
        output_result("error", message="NO_MASTER_EXCEL")
        sys.exit(1)
