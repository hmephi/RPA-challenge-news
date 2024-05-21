import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenPyXLImage
from PIL import Image as PILImage
from io import BytesIO
import os
import logging
from RPA.Robocloud.Items import ItemNotFoundError, WorkItems
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class NewsScraper:
    def __init__(self, search_phrase, news_category, months_to_fetch):
        self.search_phrase = search_phrase
        self.news_category = news_category
        self.months_to_fetch = months_to_fetch
        self.base_url = 'https://www.reuters.com/'  # Replace with the chosen news site URL
        self.search_url = f'{self.base_url}/search/news?blob={self.search_phrase}'
        self.date_limit = self.calculate_date_limit()
        self.driver = self.init_webdriver()
        self.articles = []

    def init_webdriver(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        driver = webdriver.Chrome(options=options)
        return driver

    def calculate_date_limit(self):
        current_date = datetime.now()
        return current_date - timedelta(days=self.months_to_fetch * 30)

    def contains_amount(self, text):
        pattern = re.compile(r'\$\d+(?:,\d{3})*(?:\.\d{1,2})?|\d+ (dollars|USD)')
        return bool(pattern.search(text))

    def count_search_phrase(self, text):
        return text.lower().count(self.search_phrase.lower())

    def fetch_search_results(self):
        self.driver.get(self.search_url)
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'article.story'))  # Adjust selector as needed
            )
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            self.articles = soup.find_all('article', class_='story')  # Adjust the tag and class name as needed
        except Exception as e:
            logger.error(f"Error fetching search results: {e}")
        finally:
            self.driver.quit()

    def extract_article_data(self):
        data = []
        for article in self.articles:
            title = article.find('h3').text.strip()
            date_str = article.find('time')['datetime'].strip()  # Adjust attribute as needed
            date = datetime.fromisoformat(date_str)

            if date < self.date_limit:
                continue

            description = article.find('p').text.strip() if article.find('p') else ''
            image_url = article.find('img')['src']

            # Download the image
            image_response = requests.get(image_url)
            image = PILImage.open(BytesIO(image_response.content))
            image_filename = f"/output/image_{len(data) + 1}.jpg"
            image.save(image_filename)

            phrase_count = self.count_search_phrase(title) + self.count_search_phrase(description)
            contains_money = self.contains_amount(title) or self.contains_amount(description)

            data.append({
                'Title': title,
                'Date': date.strftime('%Y-%m-%d'),
                'Description': description,
                'Image Filename': image_filename,
                'Search Phrase Count': phrase_count,
                'Contains Money': contains_money
            })
        return data

    def save_to_excel(self, data):
        df = pd.DataFrame(data)
        excel_filename = '/output/news_data.xlsx'
        df.to_excel(excel_filename, index=False)

        wb = Workbook()
        ws = wb.active

        header = list(df.columns)
        ws.append(header)

        for row_num, row in df.iterrows():
            row_data = row.tolist()
            ws.append(row_data)
            img = OpenPyXLImage(row['Image Filename'])
            img.anchor = f'E{row_num + 2}'
            ws.add_image(img)

        wb.save(excel_filename)

        for article in data:
            os.remove(article['Image Filename'])

        logger.info('Data extracted and saved to news_data.xlsx')
        return excel_filename


def main():
    work_items = WorkItems()

    try:
        work_items.get_input_work_item()
        search_phrase = work_items.get_work_item_variable('search_phrase')
        news_category = work_items.get_work_item_variable('news_category')
        months_to_fetch = int(work_items.get_work_item_variable('months_to_fetch'))
    except ItemNotFoundError as e:
        logger.error(f"Work item not found: {e}")
        return

    scraper = NewsScraper(search_phrase, news_category, months_to_fetch)
    scraper.fetch_search_results()
    data = scraper.extract_article_data()
    excel_filename = scraper.save_to_excel(data)

    work_items.create_output_work_item()
    work_items.add_work_item_file(excel_filename)
    work_items.save_work_item()


if __name__ == "__main__":
    main()
