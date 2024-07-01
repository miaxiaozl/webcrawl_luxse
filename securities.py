from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time

# Path to your ChromeDriver (ensure this matches the path where you placed ChromeDriver)
chrome_driver_path = '/usr/local/bin/chromedriver'  # e.g., '/usr/local/bin/chromedriver'

# Set up Selenium WebDriver with the path to Chrome executable
options = webdriver.ChromeOptions()
options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"  # Path to Chrome

service = ChromeService(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)

# URL of the website containing the tables
url = 'https://www.luxse.com/issuer/DeutscheBank/23571'

# Load the page
driver.get(url)

# Wait for the page to load completely
time.sleep(10)  # Increase this if needed to allow JavaScript to load content

# Click the securities tab
programmes_tab = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "div.css-32nwno"))
)
programmes_tab.click()

time.sleep(5)  # Wait for the Securities tab content to load

securities_tab = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "div.css-32nwno"))
)
securities_tab.click()

time.sleep(5)

# Function to extract data from the current page
def extract_data_from_page():
    html_content = driver.page_source
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table')
    if not table:
        raise ValueError("Securities table not found.")

    headers = [header.text.strip() for header in table.find_all('th')]
    rows = []
    for row in table.find_all('tr')[1:]:  # Skip the header row
        cells = row.find_all('td')
        rows.append([cell.text.strip() for cell in cells])
    
    return headers, rows

# Initialize data storage
all_rows = []
headers = []

# Loop to navigate through all pages
while True:
    # Extract data from the current page
    try:
        headers, rows = extract_data_from_page()
        all_rows.extend(rows)
    except ValueError as e:
        print(e)
        break

    # Try to find and click the next page button
    try:
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'li.next'))
        )
        next_button.click()
        time.sleep(5)  # Wait for the next page to load
    except:
        print("No more pages to navigate.")
        break

# Close the Selenium WebDriver
driver.quit()

# Create a DataFrame from the collected data
df = pd.DataFrame(all_rows, columns=headers)

# Save the DataFrame to a CSV file
csv_file = 'securities.csv'
df.to_csv(csv_file, index=False)
print(f"Data saved to {csv_file}")

# Convert the CSV to Excel
excel_file = 'securities.xlsx'
df.to_excel(excel_file, index=False)
print(f"{csv_file} has been converted to {excel_file}")

print("All tasks completed successfully.")
