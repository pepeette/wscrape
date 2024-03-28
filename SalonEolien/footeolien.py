import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
import time
import csv

# Initialization
url = 'https://www.france-renouvelables.fr/annuaire-des-membres/'

# Open Chrome driver
driver = webdriver.Chrome()
#driver = webdriver.Firefox() 

# Navigate to the URL
driver.get(url)

# Click on the acceptance button only once
try:
    cookie_banner = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'tarteaucitronPersonalize2'))
    )
    cookie_banner.click()
except Exception as e:
    print(f"Unable to click the 'Accept All Cookies' button for {url}: {str(e)}")

# # Wait until the elements are loaded
# time.sleep(5)

# # Function to handle scrolling
# def custom_scroll():
#     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#     time.sleep(3)
#     driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.66);")

# # Main loop
# last_height = driver.execute_script("return document.body.scrollHeight")
# while True:
#     custom_scroll()
#     new_height = driver.execute_script("return document.body.scrollHeight")
#     if round(new_height * 0.9) >= last_height:
#         break
#     last_height = new_height

# Get the source and parse it
source = driver.page_source
soup = BeautifulSoup(source, 'html.parser')

# Processing
scraped_data = []
for element in soup.findAll('div', {'class': 'col-lg-6 col-md-6 mb-4'}):
    name = element.find('div', {'class': 'nom_membre'}).text
    phone = element.find('div', {'class': 'numero_telephone'})['title'].replace("Call with Ringover", "")
    mail = element.find('div', {'class': 'email'})['title']
    scraped_data.append([name, phone, mail])

# Write to CSV
with open('footeolien.csv', mode='w', newline='', encoding='utf-8') as csvfile:
    fieldnames = ['Name', 'Phone', 'Mail']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    writer.writeheader()
    for row in scraped_data:
        writer.writerow({'Name': row[0], 'Phone': row[1], 'Mail': row[2]})

# Clean up
print("Data saved successfully!")
driver.quit()