from googlesearch import search
import pandas as pd
import requests
from bs4 import BeautifulSoup

# Load your Excel file
df = pd.read_excel('Book.xlsx')  # Replace 'your_file.xlsx' with your actual file path

# Initialize lists to store results
company_names = df['Company']
linkedin_urls = []
company_websites = []

for company in company_names:
    # Google search for LinkedIn URL
    linkedin_query = f"{company} LinkedIn"
    linkedin_url = ""
    for result in search(linkedin_query, num_results=10):
        if 'linkedin.com/company' in result:
            linkedin_url = result
            break
    linkedin_urls.append(linkedin_url)
    
    # Google search for official company website
    website_query = f"{company} official website"
    website_url = ""
    for result in search(website_query, num_results=10):
        if 'http' in result:
            website_url = result
            break
    company_websites.append(website_url)

# Add results to DataFrame
df['LinkedIn URL'] = linkedin_urls
df['Company Website'] = company_websites

# Save updated DataFrame to a new Excel file
df.to_excel('updated_companies.xlsx', index=False)
