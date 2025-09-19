# automation-web-scraper
To scrape data from a HTML File, compare with an existing excel sheet, and return only the newest data into a new Excel file ("scraped_data.xlsx") created for UIPath to pick up and use in further automation.
- Script only works on tables with timestamps in second column, skip over first column (assuming its an index column).
- A Python Script that can integrate into UIPath workflow.
### Application case: 
- Scrape data from a downloaded HTML file, in cases where webpage requires login and authorization to access data


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

5. Run the script:
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

### Improvements 
- updating to a existing master Excel file instead of a new file each time
