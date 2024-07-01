from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time
import openpyxl

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

# Optionally, wait for a specific element to ensure the content has loaded
# wait = WebDriverWait(driver, 20)
# wait.until(EC.presence_of_element_located((By.XPATH, "//h2[text()='Programmes']")))

# Get the page source after JavaScript has rendered the content
html_content = driver.page_source

# Save the HTML content to a file for inspection
with open('webpage_content_dynamic.html', 'w', encoding='utf-8') as file:
    file.write(html_content)

print("HTML content saved to 'webpage_content_dynamic.html'.")

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

def scrape_table_by_section_title(soup, section_title):
    # Find the section containing the section_title
    section = soup.find(string=section_title)
    if not section:
        raise ValueError(f"Section titled '{section_title}' not found.")
    
    # Navigate to the parent element that likely contains the table
    parent_div = section.find_parent('div')
    if not parent_div:
        raise ValueError(f"Parent div for section '{section_title}' not found.")
    
    # Find the first table within this parent div
    table = parent_div.find_next('table')
    if not table:
        raise ValueError(f"Table for section '{section_title}' not found.")
    
    # Extract table headers
    headers = [header.text.strip() for header in table.find_all('th')]
    
    # Extract table rows
    rows = []
    for row in table.find_all('tr')[1:]:  # Skip the header row
        cells = row.find_all('td')
        rows.append([cell.text.strip() for cell in cells])
    
    # Create a DataFrame from the scraped data
    df = pd.DataFrame(rows, columns=headers)
    return df

def save_to_csv(dataframe, filename):
    dataframe.to_csv(filename, index=False)

# Define section titles we are looking for
sections = ['Programmes', 'Securities', 'Notices']

# Scrape each table and save to CSV
for section_title in sections:
    try:
        df = scrape_table_by_section_title(soup, section_title)
        filename = section_title.lower() + '.csv'
        save_to_csv(df, filename)
        print(f"{section_title} table scraped and saved successfully.")
    except ValueError as e:
        print(e)

# Define the CSV file names
csv_files = ['programmes.csv', 'securities.csv', 'notices.csv']

# Convert each CSV file to an Excel file
for csv_file in csv_files:
    # Read the CSV file into a DataFrame
    df = pd.read_csv(csv_file)
    
    # Define the Excel file name
    excel_file = csv_file.replace('.csv', '.xlsx')
    
    # Save the DataFrame to an Excel file
    df.to_excel(excel_file, index=False)
    
    print(f"{csv_file} has been converted to {excel_file}")

print("All files have been converted successfully.")

# Close the Selenium WebDriver
driver.quit()