import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Load LinkedIn URLs from Excel
df = pd.read_excel('url.xlsx')  # Adjust path as needed

# Initialize WebDriver
driver = webdriver.Chrome()

# Log into LinkedIn (replace with your LinkedIn credentials)
driver.get('https://www.linkedin.com/login')
time.sleep(2)

username = driver.find_element(By.ID, 'username')
username.send_keys('nitinsai.avirneni@gmail.com')  # Replace with your LinkedIn username
password = driver.find_element(By.ID, 'password')
password.send_keys('2019@Padma')  # Replace with your LinkedIn password
password.send_keys(Keys.RETURN)
time.sleep(3)

# List to store company locations
company_locations = []

# Iterate through each URL in the Excel file
for index, row in df.iterrows():
    base_url = row['LinkedIn URL']

    # Skip if URL is missing or not a string
    if not isinstance(base_url, str) or pd.isna(base_url):
        company_locations.append(None)
        continue

    # Navigate to the LinkedIn company page
    driver.get(base_url)
    time.sleep(5)  # Wait for the page to load completely

    # Retry mechanism for loading location
    retries = 3
    company_location = None
    for attempt in range(retries):
        try:
            # Locate the company location within the "inline-block" div
            location_elements = driver.find_elements(
                By.XPATH, '//div[@class="inline-block"]/div[@class="org-top-card-summary-info-list__info-item"]'
            )

            # Assuming the location is the first element in the list
            if location_elements:
                company_location = location_elements[0].text
                print(f"Location for {base_url}: {company_location}")
            else:
                print(f"Location not found for {base_url}")

            break  # Exit retry loop if successful

        except Exception as e:
            print(f"Attempt {attempt + 1} - Could not retrieve location for {base_url}: {e}")
            time.sleep(2)  # Wait before retrying

    # Append results
    company_locations.append(company_location)

# Close the driver after processing all URLs
driver.quit()

# Save results back to Excel
df['Company Location'] = company_locations
df.to_excel('updated_companies_with_location.xlsx', index=False)
print("Data has been saved to 'updated_companies_with_location.xlsx'")
