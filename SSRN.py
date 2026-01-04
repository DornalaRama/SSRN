from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import csv
import random
import re
import pandas as pd  # for Excel

# ANTI-DETECTION OPTIONS
chrome_options = Options()
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
wait = WebDriverWait(driver, 20)

print("🚀 Starting SSRN scraper...")

# LOGIN
driver.get("https://hq.ssrn.com/login/pubSignInJoin.cfm?rectype=edit&perinf=y&partid=4821630")
time.sleep(random.uniform(5, 8))
input("Press Enter AFTER login successful...")

# COLLECT ARTICLE LINKS
journal_url = "https://papers.ssrn.com/sol3/JELJOUR_Results.cfm?form_name=journalBrowse&journal_id=948092&page=3&sort=0"
driver.get(journal_url)
time.sleep(random.uniform(5, 7))

all_article_links = []
current_page = 1
articles_needed = 50

print("Collecting article links...")
while len(all_article_links) < articles_needed and current_page <= 50:
    article_elements = driver.find_elements(By.CSS_SELECTOR, "div.title a[href*='papers.cfm?abstract_id=']")
    page_links = [el.get_attribute("href") for el in article_elements if el.get_attribute("href")]
    new_links = [link for link in page_links if link not in all_article_links]
    all_article_links.extend(new_links)

    print(f"Page {current_page}: {len(new_links)} new links. Total: {len(all_article_links)}")

    if len(all_article_links) >= articles_needed:
        break

    try:
        next_button = driver.find_element(By.XPATH, "//a[contains(@href, 'per_page=') or contains(text(), 'Next')]")
        driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
        time.sleep(random.uniform(2, 4))
        driver.execute_script("arguments[0].click();", next_button)
        time.sleep(random.uniform(8, 12))
        current_page += 1
    except:
        print("No more pages")
        break

article_links = all_article_links[:50]
print(f"✅ Collected {len(article_links)} articles")

# EMAIL FROM POPUP (first email link only)
def get_all_article_emails(driver, author_divs):
    all_emails = set()
    for div in author_divs:
        try:
            email_anchor = div.find_element(By.XPATH, ".//a[contains(@onclick, 'GetAuthorEmail')]")
            onclick = email_anchor.get_attribute("onclick")
            if onclick and 'GetAuthorEmail' in onclick:
                start = onclick.find("GetAuthorEmail.cfm?abid=")
                if start != -1:
                    end = onclick.find("'", start)
                    email_url_suffix = onclick[start:end] if end != -1 else onclick[start:]
                    email_url = "https://papers.ssrn.com/sol3/" + email_url_suffix.replace("&amp;", "&")

                    main_window = driver.current_window_handle
                    driver.execute_script("window.open(arguments[0]);", email_url)
                    driver.switch_to.window(driver.window_handles[-1])
                    time.sleep(random.uniform(4, 6))

                    email_text = driver.find_element(By.TAG_NAME, "body").text
                    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
                    found_emails = re.findall(email_pattern, email_text)
                    all_emails.update([email.lower() for email in found_emails if len(email) > 5])

                    driver.close()
                    driver.switch_to.window(main_window)
                    return list(all_emails)  # one popup only
        except:
            continue
    return list(all_emails)

authors_emails_data = []
processed_articles = 0

# MAIN PROCESSING
batch_size = 10
for batch_start in range(0, len(article_links), batch_size):
    batch = article_links[batch_start:batch_start + batch_size]
    print(f"\nBatch {batch_start//batch_size + 1}: {len(batch)} articles")

    for i, article_url in enumerate(batch, batch_start + 1):
        try:
            print(f"Article {i}/{len(article_links)}")
            driver.get(article_url)
            time.sleep(random.uniform(3, 5))

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
            time.sleep(random.uniform(2, 4))
            driver.execute_script("window.scrollTo(0, 0);")

            # Title
            try:
                title_element = driver.find_element(By.XPATH, "//div[contains(@class, 'box-abstract-main')]//h1")
                article_title = title_element.text.strip()
            except:
                article_title = "Title not found"

            # Authors
            author_divs = driver.find_elements(By.CSS_SELECTOR, "div.contact-information div.author")
            all_authors = []
            processed_authors_set = set()

            for div in author_divs:
                try:
                    name_element = div.find_element(By.TAG_NAME, "h3")
                    author_name = name_element.text.replace("(Contact Author)", "").strip()
                    if author_name and author_name not in processed_authors_set:
                        all_authors.append(author_name)
                        processed_authors_set.add(author_name)
                except:
                    continue

            if not all_authors:
                b_names = driver.find_elements(By.XPATH, "//td[@colspan='4']/b")
                for b_tag in b_names:
                    author_name = b_tag.text.strip()
                    if author_name and author_name not in processed_authors_set:
                        all_authors.append(author_name)
                        processed_authors_set.add(author_name)

            # Emails
            unique_emails_list = get_all_article_emails(driver, author_divs)
            article_email_str = ", ".join(unique_emails_list) if unique_emails_list else "No emails"

            print(f"  {len(all_authors)} authors, {len(unique_emails_list)} emails")

            # SAVE:
            if all_authors:
                for author_name in all_authors:
                    authors_emails_data.append({
                        "article_url": article_url,
                        "article_title": article_title,
                        "author_name": author_name,
                        "all_article_emails": article_email_str
                    })
            else:
                # Fallback: save one row per article when only emails are available
                if unique_emails_list:
                    authors_emails_data.append({
                        "article_url": article_url,
                        "article_title": article_title,
                        "author_name": "Unknown",
                        "all_article_emails": article_email_str
                    })

            processed_articles += 1

        except Exception as e:
            print(f"  Error: {str(e)[:80]}")
            continue

    if batch_start + batch_size < len(article_links):
        print("Pause 30s between batches...")
        time.sleep(30)

print(f"\nProcessed {processed_articles}/{len(article_links)} articles")
print(f"Saving {len(authors_emails_data)} records...")

# SAVE CSV + EXCEL
csv_filename = "ssrn_50_articles_final.csv"
excel_filename = "ssrn_50_articles_final.xlsx"

with open(csv_filename, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=["article_url", "article_title", "author_name", "all_article_emails"])
    writer.writeheader()
    writer.writerows(authors_emails_data)

df = pd.DataFrame(authors_emails_data)
df.to_excel(excel_filename, index=False)

print(f"CSV saved: {csv_filename}")
print(f"Excel saved: {excel_filename}")
print(f"Total records: {len(authors_emails_data)}")

driver.quit()
