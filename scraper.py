"""
Docstring for scraper
"""

import time
import random
from playwright.sync_api import sync_playwright


class CDEScraper:
    def __init__(self):
        print("Initializing Scraper...")
        self.playwright = sync_playwright().start()

        self.browser = self.playwright.chromium.launch(
            headless=True, args=["--disable-blink-features=AutomationControlled"]
        )

        self.context = self.browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080},
            locale="en-US",
        )

        self.main_page = self.context.new_page()

        print("Navigating to CDE Homepage...")
        self.main_page.goto("https://www.cde.ca.gov/", wait_until="networkidle")
        # Ensure the page is actually ready
        self.main_page.wait_for_selector("input#txtSearchTermSite", timeout=15000)
        self.ignored_extensions = (
            ".pdf",
            ".doc",
            ".docx",
            ".xls",
            ".xlsx",
            ".csv",
            ".zip",
            ".ppt",
            ".pptx",
            ".xml",
        )

    def _human_typing(self, selector, text):
        """Types text naturally."""
        self.main_page.click(selector)  # Focus the element first
        self.main_page.fill(selector, "")  # Clear it
        for char in text:
            self.main_page.type(selector, char)
            time.sleep(random.uniform(0.05, 0.1))

    def _perform_search(self, query):
        """
        Reliably performs the search by pressing ENTER instead of clicking buttons.
        """
        # 1. Try Result Page Input first (if we are already deep in searches)
        if self.main_page.locator("input#searchquery").is_visible():
            print(f"   [Action] using Result Page Search for: {query}")
            self._human_typing("input#searchquery", query)
            self.main_page.press("input#searchquery", "Enter")

        # 2. Try Homepage Input
        elif self.main_page.locator("input#txtSearchTermSite").is_visible():
            print(f"   [Action] using Homepage Search for: {query}")
            self._human_typing("input#txtSearchTermSite", query)
            self.main_page.press("input#txtSearchTermSite", "Enter")

        # 3. Fallback: Reload Homepage
        else:
            print("   [!] Inputs lost. Reloading Homepage...")
            self.main_page.goto("https://www.cde.ca.gov/", wait_until="networkidle")
            self.main_page.wait_for_selector("input#txtSearchTermSite")
            self._human_typing("input#txtSearchTermSite", query)
            self.main_page.press("input#txtSearchTermSite", "Enter")

    def _extract_emails_from_page(self, page):
        """Scrapes visible emails from the current tab."""
        emails = set()
        try:
            # Quick check for mailto links
            email_links = page.locator("a[href^='mailto:']").all()
            for link in email_links:
                text = link.inner_text().strip()
                # If text is empty/generic, try to grab the href content
                if "@" not in text:
                    href = link.get_attribute("href")
                    if href:
                        text = href.replace("mailto:", "").split("?")[0].strip()

                if text and "@" in text:
                    emails.add(text)
        except Exception:
            pass
        return emails

    def find_emails(self, search_query, max_results=3):
        print(f"\n--- Processing: {search_query} ---")
        all_emails = set()

        try:
            # 1. Run Search
            self._perform_search(search_query)

            # 2. Wait for Results
            try:
                self.main_page.wait_for_selector(
                    "a.gs-title", state="visible", timeout=5000
                )
            except:
                print("   [!] Search timed out or no results found.")
                return []

            # 3. Extract Valid URLs
            print("   [i] Extracting target URLs...")
            raw_links = self.main_page.locator("a.gs-title").all()
            target_urls = []

            for link in raw_links:
                if len(target_urls) >= max_results:
                    break

                url = link.get_attribute("href")

                # --- DEFENSE LAYER 1: The Filter ---
                if not url or "http" not in url:
                    continue

                # Check if URL ends with a file extension we don't want
                clean_url = url.lower().strip()
                if clean_url.endswith(self.ignored_extensions):
                    print(f"      [Skipping Download Link]: {url}")
                    continue

                if url not in target_urls:
                    target_urls.append(url)

            if not target_urls:
                print(
                    "   [-] No valid web pages found (only downloads or empty links)."
                )
                return []

            # 4. Visit URLs Loop
            for i, url in enumerate(target_urls):
                print(f"   [{i+1}/{len(target_urls)}] Visiting: {url}")

                result_tab = self.context.new_page()

                # --- DEFENSE LAYER 2: The Canceler ---
                # If the page tries to download something, cancel it immediately.
                result_tab.on("download", lambda download: download.cancel())

                try:
                    result_tab.goto(url, wait_until="domcontentloaded", timeout=15000)

                    # Wait for "reading"
                    time.sleep(random.uniform(1.5, 3.0))

                    # Grab emails
                    found = self._extract_emails_from_page(result_tab)
                    if found:
                        print(f"      [+] Found: {list(found)}")
                        all_emails.update(found)
                    else:
                        print("      [-] No emails.")

                except Exception as e:
                    # Often "net::ERR_ABORTED" happens when we cancel a download. This is good!
                    if "ERR_ABORTED" in str(e):
                        print("      [i] Blocked a forced download.")
                    else:
                        print(f"      [!] Failed to load page: {e}")

                finally:
                    result_tab.close()

            unique = list(all_emails)
            if unique:
                print(f"   [SUCCESS] Collected: {unique}")
            else:
                print("   [RESULT] No emails found.")

            return unique

        except Exception as e:
            print(f"   [CRITICAL] Script error: {e}")
            return []

    def close(self):
        print("Closing browser...")
        self.browser.close()
        self.playwright.stop()


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    targets = ["Lincoln High", "Washington Elementary", "Roosevelt Middle"]
    scraper = CDEScraper()

    try:
        for target in targets:
            scraper.find_emails(target, max_results=3)
            # Tiny pause between schools so we don't look insane
            time.sleep(1)

    except KeyboardInterrupt:
        print("\nStopping early...")
    finally:
        scraper.close()
