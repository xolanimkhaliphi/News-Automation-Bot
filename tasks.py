import logging
from RPA.Browser import Browser
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import openpyxl
import datetime
import re

class APNewsBot:
    def __init__(self):
        # Initialize a web browser instance for automation
        self.browser = Browser()
        # Configure logging
        logging.basicConfig(filename='ap_news_bot.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def extract_news(self, search_phrase, news_category, num_months):
        # Open the Associated Press (AP) news website
        self.browser.open_available_browser("https://apnews.com/")
        
        try:
            # Input the search phrase into the search field
            search_button = WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'searchOverlay-search-button')]")))
            search_button.click()
            search_input = WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'searchOverlay-search-input')]")))
            search_input.send_keys(search_phrase)

            # Wait for the search results to load
            WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'CardHeadline')]")))
        except TimeoutException as e:
            # Log timeout error if search results fail to load
            logging.error("Timeout error: Failed to load search results.")
            self.close_browser()
            return None

        # Refine the search by selecting a news category if provided
        if news_category:
            try:
                category_link = WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, news_category)))
                category_link.click()
                WebDriverWait(self.browser.driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'CardHeadline')]")))
            except TimeoutException as e:
                # Log timeout error if news category selection fails
                logging.error(f"Timeout error: Failed to load news category: {news_category}.")
                self.close_browser()
                return None

        # Extract data from the search results
        headlines = WebDriverWait(self.browser.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//h3[contains(@class, 'PagePromo-title')]//span")))
        dates = WebDriverWait(self.browser.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//span[contains(@class, 'Timestamp')]")))
        descriptions = WebDriverWait(self.browser.driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class, 'PagePromo-description')]//span")))

        # Store the extracted data in an Excel file
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Title", "Date", "Description", "Picture Filename", "Search Phrase Count", "Contains Money"])

        for headline, date, description in zip(headlines, dates, descriptions):
            # Calculate the count of occurrences of the search phrase in the title and description
            search_phrase_count = self.count_search_phrase_occurrences(search_phrase, headline.text, description.text)
            # Check if the description contains any mention of money
            contains_money = self.contains_money(description.text)
            # Generate a picture filename based on the title of the news article
            picture_filename = f"{headline.text[:20].strip().replace(' ', '_')}.png"
            # Append the extracted data to the Excel file
            ws.append([headline.text, date.text, description.text, picture_filename, search_phrase_count, contains_money])

        # Generate a filename for the Excel file based on the current timestamp
        filename = f"news_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        # Save the Excel file
        wb.save(filename)

        # Log successful extraction
        logging.info(f"News data has been extracted and saved to: {filename}")

        return filename

    def count_search_phrase_occurrences(self, search_phrase, *texts):
        # Count the occurrences of the search phrase in a list of texts
        count = 0
        for text in texts:
            count += text.lower().count(search_phrase.lower())
        return count

    def contains_money(self, text):
        # Check if a text contains any mention of money using a regular expression
        money_regex = r'\$[\d,]+(\.\d+)?|\d+\s?(dollars|USD)'
        return bool(re.search(money_regex, text))

    def close_browser(self):
        # Close the web browser instance
        self.browser.close_all_browsers()


if __name__ == "__main__":
    bot = APNewsBot()
    # Prompt the user to input the search phrase, news category, and number of months for news
    search_phrase = input("Enter a search phrase: ")
    news_category = input("Enter a news category (leave blank for all): ")
    num_months = int(input("Enter the number of months for news retrieval (0 for current month, 1 for current and previous month, and so on): "))

    # Extract news based on the provided inputs
    news_file = bot.extract_news(search_phrase, news_category, num_months)
    if news_file:
        print(f"News data has been extracted and saved to: {news_file}")
    bot.close_browser()
