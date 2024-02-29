import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

# Function to extract the NumFirm parameter from the info button
def extract_numfirm(info_button):
    if info_button is not None:
        href = info_button.get('href')
        match = re.search(r"InfosPlus\('(\w+)'\)", href)
        if match:
            return match.group(1)
    return None

# Function to extract the address from the redirected page
def get_address_from_url(numfirm):
    url = f"https://www.kerix.net/Fiche.asp?NumFirm={numfirm}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Find the image element
    img_element = soup.find('img', alt="", src="/assets/img/country/morocco.png")
    if img_element:
        # Find the next sibling <p> element
        address_element = img_element.find_next_sibling('p', class_='card-text')
        if address_element:
            address = address_element.get_text(strip=True)
        else:
            address = 'No address found'
    else:
        address = 'No image found'
    
    return address

# Function to extract the activity from the redirected page
def get_activity_from_url(numfirm):
    url = f"https://www.kerix.net/Fiche.asp?NumFirm={numfirm}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Find the 'Activités' section
    activites_section = soup.find('h5', class_='card-title', string=lambda text: text and 'ACTIVITES' in text)
    if activites_section:
        # Find the next sibling <p> element
        activity_element = activites_section.find_next_sibling('p')
        if activity_element:
            activity = activity_element.get_text(strip=True)
        else:
            activity = 'No activity found'
    else:
        activity = 'No activity section found'
    
    return activity

# Make a request to the website
url = "http://maroc1000.net/Ordre-chiffre-d'affaires-2021/Page6"
response = requests.get(url)

# Parse the content of the request with BeautifulSoup
soup = BeautifulSoup(response.content, 'html.parser')

# Find the table
table = soup.find('table', {'id': 'les1000'})

# Extract the data from the table, excluding the first row
data = []
rows = table.find_all('tr')
for row in rows[1:]:  # Skip the first row
    cols = row.find_all('td')
    cols = [col.text.strip() for col in cols]
    info_button = row.find('a', href=lambda value: value and value.startswith('javascript:InfosPlus'))
    numfirm = extract_numfirm(info_button)
    address = get_address_from_url(numfirm) if numfirm else 'No address'
    activity = get_activity_from_url(numfirm) if numfirm else 'No activity'
    # Ensure that the row has exactly four elements
    if len(cols) >= 4:  # Assuming 'id' is the first column
        data.append([cols[0], cols[1], address, activity])  # Get rid of empty values
    else:
        print(f"Skipping row with less than 4 columns: {cols}")

# Create a DataFrame from the extracted data
df = pd.DataFrame(data, columns=['id', 'Raison Sociale', 'adresse', 'Activité'])

# Save the DataFrame to an Excel file
df.to_excel('output_6.xlsx', index=False)