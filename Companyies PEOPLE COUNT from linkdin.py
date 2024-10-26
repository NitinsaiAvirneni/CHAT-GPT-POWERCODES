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

# List to store the number of associated people
associated_people_counts = []

# Iterate through each URL in the Excel file
for index, row in df.iterrows():
    base_url = row['LinkedIn URL']

    # Skip if URL is missing or not a string
    if not isinstance(base_url, str) or pd.isna(base_url):
        associated_people_counts.append(None)
        continue

    # Clean and navigate to the LinkedIn people page
    people_url = f"{base_url.strip()}/people"
    driver.get(people_url)
    time.sleep(5)  # Wait for the page to load completely

    try:
        # Wait for the associated members element to be present
        members_count_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '//h2[contains(@class, "text-heading-xlarge")]')
            )
        )
        
        # Extract text and find the number of associated members
        members_count_text = members_count_element.text
        print(f"People count for {people_url}: {members_count_text}")
        associated_people_counts.append(members_count_text)
    except Exception as e:
        print(f"Could not find the associated people count for {people_url}: {e}")
        associated_people_counts.append(None)
    
    time.sleep(2)

# Close the driver after processing all URLs
driver.quit()

# Save results back to Excel
df['Associated People Count'] = associated_people_counts
df.to_excel('updated_companies.xlsx', index=False)
print("Data has been saved to 'updated_companies.xlsx'")
