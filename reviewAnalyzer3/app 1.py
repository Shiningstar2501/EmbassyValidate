
from playwright.sync_api import sync_playwright
import pandas as pd
from urllib.parse import quote

def get_newest_review_date(query):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        encoded_query = quote(query)
        search_url = f"https://www.google.com/maps/search/{encoded_query}"
        print(f"🔍 Searching: {search_url}")

        try:
            page.goto(search_url, timeout=60000)
            page.wait_for_timeout(5000)

            buttons = page.locator("button")
            review_button = None
            for i in range(buttons.count()):
                try:
                    label = buttons.nth(i).get_attribute("aria-label")
                    if label and "Reviews for" in label:
                        review_button = buttons.nth(i)
                        break
                except:
                    continue

            if review_button is None:
                browser.close()
                return "Review button not found"

            review_button.click()
            page.wait_for_timeout(3000)
            page.locator('button[aria-label="Sort reviews"]').click()
            page.wait_for_timeout(1500)
            page.locator('div[role="menuitemradio"][data-index="1"]').click()
            page.wait_for_timeout(2000)
            review_date = page.locator('span.rsqaWe').first.inner_text()
            browser.close()
            return review_date
        except Exception as e:
            browser.close()
            return f"Error: {str(e)}"

# Load your final input Excel
input_file = "review_data_part_1.xlsx"
df = pd.read_excel(input_file)

address_cols = [col for col in df.columns if col.startswith("Matched Address")]
for col in address_cols:
    review_col = col.replace("Matched Address", "Review Date")
    if review_col not in df.columns:
        df[review_col] = ""

for idx, row in df.iterrows():
    name = row.get("name", "").strip()
    if not name:
        continue
    print(f"\n📄 Processing row {idx + 1}: {name}")
    for addr_col in address_cols:
        address = row.get(addr_col, "")
        if pd.isna(address) or not str(address).strip():
            continue
        search_query = f"{name} {address}"
        review_date = get_newest_review_date(search_query)
        review_col = addr_col.replace("Matched Address", "Review Date")
        df.at[idx, review_col] = review_date
        print(f"✅ {addr_col} → {review_date}")

df.to_excel("final_review_filled1.xlsx", index=False)
print("✅ All done! Saved to 'final_review_filled.xlsx'")
