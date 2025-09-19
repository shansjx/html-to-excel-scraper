# 📝automation-web-scraper

A Python Script to scrape data from a HTML File's Table, update an existing Excel file only filling empty cells in the main data columns.and return the newest data into a new Excel file.

## How It Works✨

- The script finds the latest timestamp in the Excel file.
- It scrapes new rows from the HTML table with newer timestamps.
- For each new row, it fills the next available empty cell in the main columns (`FirstCol`, `SecondCol`, `ThirdCol`, `FourthCol`) in the rows immediately after the latest timestamp.
- No existing data in other columns or rows is overwritten or shifted.

## Notes:
- Script only works on tables with timestamps in second column of html file table, skip over first column (assuming its an index column).
- New data file ("scraped_data.xlsx") only if there is new data not found in the master excel file.

- Scrape data from a downloaded HTML file, in cases where webpage requires login and authorization to access
- Can integrate into UIPath workflow.


### Local Setup

1. Clone the repository:
   ```
   git clone
    ```
2. Navigate to the project directory:
    ```
    cd automation-web-scraper
    ```
3. Create and activate a virtual environment (Windows):
    ```
    python -m venv venv

    .\venv\Scripts\activate 
    ```
4. Install the required packages:
    ```
    pip install -r requirements.txt
    ```

5. Run the script: (Ensure HTML file and master Excel file are closed before running)
    ```
    python main.py path/to/your/htmlfile.html path/to/your/masterexcel.xlsx
    ```
6. There will be a new file created in the same directory as the script, named `scraped_data.xlsx`, containing only the new rows not present in the master Excel file.
7. Clean up the file after use:
    ```
    python main.py --cleanup
    ```
8. Deactivate the virtual environment when done:
    ```
    deactivate
    ```

## Troubleshooting

- **File is open or locked:** Close the Excel file before running the script.