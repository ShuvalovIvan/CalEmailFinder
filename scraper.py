import time
import random
import re
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError


class NetworkError(Exception):
    """Custom exception raised when internet connection is lost."""

    pass


class CDEScraper:
    # Files to avoid downloading
    IGNORED_EXTENSIONS = (
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

        # Disable webdriver flags
        self.context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

        self.main_page = self.context.new_page()

        print("Navigating to CDE Homepage...")
        self.main_page.goto("https://www.cde.ca.gov/", wait_until="networkidle")
        self.main_page.wait_for_selector("input#txtSearchTermSite", timeout=15000)

    def _human_typing(self, selector, text):
        """Types text naturally."""
        try:
            self.main_page.click(selector)
            self.main_page.fill(selector, "")
            for char in text:
                self.main_page.type(selector, char)
                time.sleep(random.uniform(0.05, 0.1))
        except:
            # Fallback if click fails
            self.main_page.fill(selector, text)

    def _perform_search(self, query):
        """Executes search using the correct visible input."""
        if self.main_page.locator("input#searchquery").is_visible():
            self._human_typing("input#searchquery", query)
            self.main_page.press("input#searchquery", "Enter")
        elif self.main_page.locator("input#txtSearchTermSite").is_visible():
            self._human_typing("input#txtSearchTermSite", query)
            self.main_page.press("input#txtSearchTermSite", "Enter")
        else:
            print("   [!] Reloading Homepage...")
            self.main_page.goto("https://www.cde.ca.gov/", wait_until="networkidle")
            self.main_page.wait_for_selector("input#txtSearchTermSite")
            self._human_typing("input#txtSearchTermSite", query)
            self.main_page.press("input#txtSearchTermSite", "Enter")

    def _clean_phone(self, text):
        """Formats phone number to 1 (XXX) XXX-XXXX"""
        digits = re.sub(r"\D", "", text)

        if len(digits) == 10:
            return f"1 ({digits[0:3]}) {digits[3:6]}-{digits[6:]}"
        elif len(digits) == 11 and digits.startswith("1"):
            return f"1 ({digits[1:4]}) {digits[4:7]}-{digits[7:]}"

        return text

    def _split_name(self, full_name):
        """
        Splits 'Dr. Ivan Smith' into 'Ivan' and 'Smith'.
        Automatically removes common honorifics.
        """
        # Define honorifics to ignore (case-insensitive)
        honorifics = {
            "mr",
            "mr.",
            "mrs",
            "mrs.",
            "ms",
            "ms.",
            "miss",
            "dr",
            "dr.",
            "prof",
            "prof.",
            "rev",
            "rev.",
            "hon",
            "hon.",
        }

        # Split by spaces
        parts = full_name.strip().split()

        # Check if the first part is an honorific
        if parts and parts[0].lower() in honorifics:
            parts.pop(0)  # Remove it

        if not parts:
            return "Unknown", "Unknown"

        first = parts[0]
        # Join the rest as last name (handles multi-word last names)
        last = " ".join(parts[1:]) if len(parts) > 1 else ""

        return first, last

    def _extract_principal_data_from_page(self, page):
        """
        Scans for a box containing 'Principal', 'Administrator', 'Director', 'Head of School', or 'Superintendent'.
        Returns a dictionary or None.
        """
        try:
            # Regex for job titles
            job_titles_regex = re.compile(
                r"\b(Principal|Administrator|Director|Head|Superintendent)\b",
                re.IGNORECASE,
            )

            candidates = page.get_by_text(job_titles_regex).all()

            for element in candidates:
                try:
                    box = element.locator("xpath=..")
                    text_content = box.inner_text()

                    if not job_titles_regex.search(text_content):
                        continue

                    lines = [
                        line.strip()
                        for line in text_content.split("\n")
                        if line.strip()
                    ]

                    found_name = None
                    found_email = None
                    found_phone = None
                    found_title_text = "N/A"  # Variable to store the specific title

                    # Find the Title Line index
                    title_index = -1
                    for i, line in enumerate(lines):
                        if re.search(
                            r"(Principal|Administrator|Director|Head|Superintendent)",
                            line,
                            re.IGNORECASE,
                        ):
                            title_index = i
                            found_title_text = (
                                line.strip()
                            )  # Capture the actual title found
                            break

                    if title_index > 0:
                        found_name = lines[title_index - 1]

                    # Find Email
                    mailto = box.locator("a[href^='mailto:']").first
                    if mailto.count() > 0:
                        found_email = mailto.inner_text()
                    else:
                        for line in lines:
                            if "@" in line and "." in line:
                                found_email = line
                                break

                    # Find Phone
                    phone_pattern = re.compile(r"\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}")
                    for line in lines:
                        match = phone_pattern.search(line)
                        if match:
                            found_phone = self._clean_phone(match.group())
                            break

                    # --- VALIDATION ---
                    if found_name and (found_email or found_phone):
                        first, last = self._split_name(found_name)

                        return {
                            "First Name": first,
                            "Last Name": last,
                            "Job Title": found_title_text,  # <--- NEW FIELD
                            "Email": found_email if found_email else "N/A",
                            "Phone": found_phone if found_phone else "N/A",
                        }

                except Exception:
                    continue

        except Exception as e:
            pass

        return None

    def find_principal_data(self, query_string):
        print(f"\n--- Processing: {query_string} ---")

        # Default now includes Job Title
        result_data = {
            "First Name": "",
            "Last Name": "",
            "Job Title": "",
            "Email": "",
            "Phone": "",
        }

        try:
            self._perform_search(query_string)

            try:
                self.main_page.wait_for_selector(
                    "a.gs-title", state="visible", timeout=5000
                )
            except Exception as e:
                # Re-raise timeouts for the menu to catch
                if "Timeout" in str(e) or "TargetClosed" in str(e):
                    # Save the URL (or search page) that timed out so Menu can see it
                    self.current_url = self.main_page.url
                    raise e
                print("   [!] No results found.")
                return result_data

            raw_links = self.main_page.locator("a.gs-title").all()
            target_urls = []
            for link in raw_links:
                url = link.get_attribute("href")
                if not url or "http" not in url:
                    continue
                if url.lower().strip().endswith(self.IGNORED_EXTENSIONS):
                    continue
                if url not in target_urls:
                    target_urls.append(url)
                if len(target_urls) >= 3:
                    break

            if not target_urls:
                return result_data

            for i, url in enumerate(target_urls):
                print(f"   [{i+1}/{len(target_urls)}] Visiting: {url}")

                # --- STEP 1: SAVE URL FOR MENU ---
                self.current_url = url
                # ---------------------------------

                result_tab = self.context.new_page()
                result_tab.on("download", lambda download: download.cancel())

                try:
                    result_tab.goto(url, wait_until="domcontentloaded", timeout=15000)
                    time.sleep(random.uniform(1.5, 3.0))

                    data = self._extract_principal_data_from_page(result_tab)

                    if data:
                        print(f"      [SUCCESS] Found Data: {data}")
                        result_data = data
                        result_tab.close()
                        return result_data

                except Exception as e:
                    if (
                        "Timeout" in str(e)
                        or "TargetClosed" in str(e)
                        or "Connection refused" in str(e)
                    ):
                        result_tab.close()
                        # --- STEP 2: RAISE ORIGINAL ERROR ---
                        # We raise 'e' so the outer block knows it's a Timeout.
                        # The menu will read 'self.current_url' separately.
                        raise e
                    pass

                finally:
                    result_tab.close()

            return result_data

        except Exception as e:
            if (
                "Timeout" in str(e)
                or "TargetClosed" in str(e)
                or "Connection refused" in str(e)
            ):
                raise e
            print(f"   [CRITICAL] Script error: {e}")
            return result_data

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
            data = scraper.find_principal_data(target)
            print(f"   -> Final Output: {data}")
            time.sleep(1)

    except KeyboardInterrupt:
        print("\nStopping early...")
    finally:
        scraper.close()
