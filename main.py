import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta

def scrape_data_into_excel(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")

    # Only find the first table element in the page
    table = soup.find("table")
    if not table:
        print("No table found")
        return
    
    rows = table.find_all("tr")
    data = []
    if len(rows) <= 1:
        print("No data to scrape.")
        return
    
    # Time window (24 hours ago to now)
    now = datetime.now()
    end_time = now
    start_time = now - timedelta(days=1)
    print(f"Scraping rows between {start_time} and {end_time}")

    for row in rows:
        item = {}
        columns_list = row.find_all("td")
        num_of_columns = len(columns_list)
        # Skip index 0 column, assuming its an index/checkbox column

        timestamp_str = columns_list[1].text.strip() if num_of_columns > 1 else "" # Timestamp column
        try:
            # Put as current year to the timestamp string (specifically only for current example website, which is missing year)
            current_year = now.year
            row_timestamp = datetime.strptime(f"{current_year} {timestamp_str}", "%Y %d %b, %H:%M")
            # Skip this row only if timestamp is outside time window
            if not (start_time <= row_timestamp < end_time):
                continue
            
            item = {
                "FirstCol": row_timestamp,
                "SecondCol": columns_list[2].text.strip() if num_of_columns > 2 else "",
                "ThirdCol": columns_list[3].text.strip() if num_of_columns > 3 else "",
                "FourthCol": columns_list[4].text.strip() if num_of_columns > 4 else "",
            }
            # Only append if at least one of the kept columns is not empty, ignore all empty rows
            if any([item["FirstCol"], item["SecondCol"], item["ThirdCol"], item["FourthCol"]]):
                data.append(item)

        # Skip row with invalid timestamp format
        except ValueError:
            continue

    if data:
        # Save to Excel
        df = pd.DataFrame(data)
        df.to_excel("scraped_data.xlsx", index=False)

        # Send the Excel file via gmail
        send_excel_via_email(
            file_path = "scraped_data.xlsx",
            subject = datetime.today().strftime('%d-%m-%Y') + " - New Scraped Data Excel File",
            body = "The attached Excel file has the scraped data",
            to_email = os.environ.get("TO_EMAIL"),
            from_email = os.environ.get("GMAIL_EMAIL"),
            password = os.environ.get("GMAIL_PASSWORD")
        )
    else:
        print("No new data found within the last 24 hours.")

def send_excel_via_email(file_path, subject, body, to_email, from_email, password):
    import smtplib
    from email.message import EmailMessage
    import os

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email
    msg.set_content(body)

    # Attach the excel file
    with open(file_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(file_path)
    msg.add_attachment(file_data, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=file_name)

    # Connect to gmail SMTP server
    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login(from_email, password)
        smtp.send_message(msg)
    print(f"Email sent to {to_email} with attachment {file_name}")

if __name__ == "__main__":
    load_dotenv()
    scrape_data_into_excel(os.environ.get("URL"))
    # Clean the generated excel file after sending
    if os.path.exists("scraped_data.xlsx"):
        os.remove("scraped_data.xlsx")
