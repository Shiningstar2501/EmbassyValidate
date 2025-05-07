import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
import os
import pandas as pd
import urllib3

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SOCIAL_DOMAINS = ['linkedin.com', 'threads.net', 'facebook.com', 'instagram.com', 'twitter.com']
MAX_PAGES = 30

input_file = 'split_part_1_new (1).xlsx'  # <-- Change this for each app copy
output_file = 'updated_embassy_social_links_rowwise_1.xlsx'  # <-- Change this for each app copy
error_file = 'error_1.txt'  # <-- Change this for each app copy

# Ensure output folder exists
os.makedirs('website_social_links_output', exist_ok=True)

def is_internal_link(base_url, link):
    return urlparse(link).netloc in ['', urlparse(base_url).netloc]

def log_error(embassy_name, website, error_message):
    with open(error_file, "a", encoding="utf-8") as f:
        f.write(f"Embassy Name: {embassy_name}\n")
        f.write(f"Website: {website}\n")
        f.write(f"Error: {error_message}\n\n")

def find_social_links(url, base_url, visited, social_links, pages_crawled, embassy_name):
    if pages_crawled[0] >= MAX_PAGES:
        return
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:120.0) Gecko/20100101 Firefox/120.0"
        }

        response = requests.get(url, headers=headers, timeout=15, verify=False)
        print(f"[INFO] Fetched {url} - Status: {response.status_code}")

        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        pages_crawled[0] += 1

        soup = soup.find_all('a', href=True)
        for a in soup:
            href = a['href']
            full_url = urljoin(base_url, href)
            if not is_internal_link(base_url, full_url):
                if any(domain in href for domain in SOCIAL_DOMAINS):
                    social_links.append(full_url)
            elif full_url not in visited:
                crawl(full_url, base_url, visited, social_links, pages_crawled, embassy_name)

    except requests.RequestException as e:
        log_error(embassy_name, url, f"Request Exception: {str(e)}")
        print(f"[ERROR] {str(e)}")

def crawl(url, base_url, visited, social_links, pages_crawled, embassy_name):
    if url not in visited:
        visited.add(url)
        find_social_links(url, base_url, visited, social_links, pages_crawled, embassy_name)

def extract_social_links(start_url, embassy_name):
    visited = set()
    social_links = []
    pages_crawled = [0]
    crawl(start_url, start_url, visited, social_links, pages_crawled, embassy_name)
    return social_links

def save_links_to_excel(df):
    for idx, row in df.iterrows():
        embassy_name = row['name']
        website2 = row['website2']
        links = []

        if pd.isna(website2):
            continue

        print(f"\n====== Processing: {embassy_name} ======")
        try:
            for link in str(website2).split('||'):
                link = link.strip()
                if link:
                    links.extend(extract_social_links(link, embassy_name))

            # Save the links to respective columns
            for i, domain in enumerate(SOCIAL_DOMAINS):
                df.at[idx, domain] = links[i] if i < len(links) else None

        except Exception as e:
            log_error(embassy_name, website2, str(e))

        # ✅ Save after each row to ensure real-time progress
        df.to_excel(output_file, index=False)
        print(f"[✔] Saved progress to: {output_file}")

# === MAIN BLOCK ===
if __name__ == '__main__':
    df = pd.read_excel(input_file)
    for domain in SOCIAL_DOMAINS:
        if domain not in df.columns:
            df[domain] = None
    save_links_to_excel(df)
    print("✅ Task complete!")
