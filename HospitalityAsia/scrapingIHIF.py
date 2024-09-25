import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import time
import csv

# Initialization
url = 'https://questexhospitality.app.swapcard.com/widget/event/ihif-asia/people/RXZlbnRWaWV3Xzc2Njg5NQ==?showActions=true'

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode

# Open Chrome driver
driver = webdriver.Chrome(options=chrome_options)

# Navigate to the URL
driver.get(url)

# Wait for the content to load
wait = WebDriverWait(driver, 20)
wait.until(EC.presence_of_element_located((By.CLASS_NAME, "sc-a7ac5ca1-0")))

# Function to handle scrolling
def scroll_to_bottom():
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

# Scroll to load all content
scroll_to_bottom()

# Get the source and parse it
source = driver.page_source
soup = BeautifulSoup(source, 'html.parser')

# Processing
scraped_data = []
main_div = soup.find('div', class_="sc-a7ac5ca1-0 etPmlQ")
if main_div:
    spans = main_div.find_all('span')
    for i in range(0, len(spans), 3):
        if i + 2 < len(spans):
            name = spans[i].text.strip()
            title = spans[i+1].text.strip()
            company = spans[i+2].text.strip()
            if name and title and company:  # Only add if all fields are non-empty
                scraped_data.append([name, title, company])

# Write to CSV
with open('questex_hospitality_data.csv', mode='w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Name', 'Title', 'Company'])  # Header
    writer.writerows(scraped_data)

# Clean up
print(f"Data saved successfully! {len(scraped_data)} people extracted.")