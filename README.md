# automation-web-scraper
To scrape data from a webpage and save it into an Excel file, then email the file as an attachment.
- Implemented with Github Actions to run the script daily at a certain time (8:00am (UTC+8) by default)
- Add only the most recent items (From 24 hours ago to now (time of script execution))
- Script only works on tables with timestamp in second column, skip over first column (index column)

## Local Setup

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
5. Copy the `.env.sample` as an `.env` file in the project root directory and add your email credentials, and URL to scrape:
GMAIL_PASSWORD is a 16 character App Password generated from your Google Account settings.
https://myaccount.google.com/apppasswords

6. Run the script:
    ```
    python main.py
    ```

### Improvements 
- for specific use case: Enable clicking into a tab (with no subpath in URL)
- updating to a existing master Excel file instead of a new file each time
