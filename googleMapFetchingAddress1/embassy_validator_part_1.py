
import pandas as pd
from playwright.sync_api import sync_playwright
import re
from urllib.parse import quote

def clean_address(raw_text):
    text = re.sub(r'<[^>]+>', ' ', str(raw_text))
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'P\.?O\.?\s?Box\s?\d+.*?,?', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\b\d{4,6}\b', '', text)
    return text.strip()

def extract_status_from_locator(page):
    try:
        if page.locator("span.fCEvvc", has_text="Permanently closed").count() > 0:
            return "Permanently Closed"
        elif page.locator("span.fCEvvc", has_text="Temporarily closed").count() > 0:
            return "Temporarily Closed"
        elif "moved" in page.content().lower():
            return "Moved"
        else:
            return "Active / No Info"
    except:
        return "Unknown"

def check_embassy_basic(query):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        encoded_query = quote(query)
        search_url = f"https://www.google.com/maps/search/{encoded_query}"
        print(f"🔗 Visiting: {search_url}")
        page.goto(search_url, timeout=60000)
        page.wait_for_timeout(6000)

        page_url = page.url
        result_type = "Single"
        addresses, phones, websites, hours_list, statuses = [], [], [], [], []

        results = page.locator(".Nv2PK")
        count = results.count()

        if count == 0:
            try:
                print("📍 Single result found")
                result_type = "Single"
                page.wait_for_timeout(4000)
                status = extract_status_from_locator(page)
                statuses.append(status)
                if status != "Permanently Closed":
                    addresses.append(clean_address(page.locator(".Io6YTe, .rogA2c").first.text_content()))
                    phones.append(page.locator('button[data-item-id*="phone"], .UsdlK').first.text_content())
                    websites.append(page.locator('a[aria-label*="Website"], a[href^="http"]').first.get_attribute("href"))
                    hours_list.append(page.locator('div[aria-label*="Hours"], .OqCZI').first.text_content())
            except Exception as e:
                print(f"⚠️ Single result extraction error: {e}")
        else:
            print(f"📋 Multiple results found: {count}")
            result_type = "Multiple"
            for i in range(min(10, count)):
                try:
                    print(f"➡️ Clicking result #{i + 1}")
                    results.nth(i).click()
                    page.wait_for_timeout(5000)
                    status = extract_status_from_locator(page)
                    statuses.append(status)
                    if status == "Permanently Closed":
                        print("❌ Skipping closed listing.")
                        page.go_back()
                        page.wait_for_timeout(3000)
                        continue
                    addresses.append(clean_address(page.locator(".Io6YTe, .rogA2c").first.text_content()))
                    try:
                        phones.append(page.locator('button[data-item-id*="phone"], .UsdlK').first.text_content())
                    except:
                        phones.append("Not Found")
                    try:
                        websites.append(page.locator('a[aria-label*="Website"], a[href^="http"]').first.get_attribute("href"))
                    except:
                        websites.append("Not Found")
                    try:
                        hours_list.append(page.locator('div[aria-label*="Hours"], .OqCZI').first.text_content())
                    except:
                        hours_list.append("Not Found")
                    page.go_back()
                    page.wait_for_timeout(3000)
                except Exception as e:
                    print(f"⚠️ Failed result {i + 1}: {e}")
                    statuses.append("Failed")

        browser.close()

        return {
            "Status Combined": " || ".join(statuses) if statuses else "Not Found",
            "Google Maps Link": page_url,
            "Result Type": result_type,
            "Matched Address": " || ".join(addresses) if addresses else "Not Found",
            "Phone": " || ".join(phones) if phones else "Not Found",
            "Website": " || ".join(websites) if websites else "Not Found",
            "Hours": " || ".join(hours_list) if hours_list else "Not Found"
        }

df = pd.read_excel("split_part_1.xlsx")
df["Status Combined"] = ""
df["Maps Link"] = ""
df["Result Type"] = ""
df["Matched Address"] = ""
df["Phone"] = ""
df["Website"] = ""
df["Hours"] = ""

for idx, row in df.iterrows():
    raw_query = clean_address(f"{row['name']} {row['Address']} embassy")
    print(f"🔍 Processing row {idx + 1}: {raw_query}")
    try:
        result = check_embassy_basic(raw_query)
        for key in result:
            df.at[idx, key] = result[key]
    except Exception as e:
        print(f"❌ Error at row {idx + 1}: {str(e)}")
        df.at[idx, "Status Combined"] = "Error"
        df.at[idx, "Maps Link"] = "N/A"
        df.at[idx, "Matched Address"] = "N/A"
        df.at[idx, "Phone"] = "N/A"
        df.at[idx, "Website"] = "N/A"
        df.at[idx, "Hours"] = "N/A"
        df.at[idx, "Result Type"] = "Error"

df.to_excel("embassy_status_taglevel_closed_check_part_1.xlsx", index=False)
print("✅ Done! File saved as 'embassy_status_taglevel_closed_check_part_1.xlsx'")
