import time

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from openpyxl import Workbook


class Parser:
    def __init__(self) -> None:
        self.website = "https://portal.eaeunion.org/sites/commonprocesses/ru-ru/Pages/DrugRegistrationDetails.aspx"
        self.driver = webdriver.Edge()

        # Count IDs
        self.count_db_id = "ComboBox1-input"
        self.count_option_id = "ComboBox1-list4"

        # Pages classes
        self.current_page_class = "ecc-page-number-input"
        self.last_page_class = "eec-page-count"
        self.next_bt_class = "arrow-right"

        # Page
        self.start_page_num = 1
        self.current_page_num = 1
        self.__last_page_num = None

        # Wait time
        self.min_wait_time = 50
        self.avg_wait_time = 150
        self.max_wait_time = 300

        # Countries codes
        self.id2country = {"am": "Армения",
                           "by": "Беларусь",
                           "ru": "Россия",
                           "kg": "Кыргызстан",
                           "kz": "Казахстан"}

        # Xlsx
        self.xlsx_path = "Реестр ЛС ЕАЭС.xlsx"
        self.xlsx_page_title = "Реестр"
        self.headers = ['trade_name', 'int_name', 'rel_form',
                        'manufacturer', 'properties', 'certificate', 'update']

        self.all_country_spans = set()

    @property
    def last_page_num(self) -> int:
        if self.__last_page_num is None:
            last_page_present = EC.presence_of_element_located(
                (By.CLASS_NAME, self.last_page_class))
            WebDriverWait(self.driver, self.avg_wait_time).until(
                last_page_present)
            last_page = self.driver.find_element(
                By.CLASS_NAME, self.last_page_class)
            self.__last_page_num = int(last_page.text)

        return self.__last_page_num

    def change_num_entries(self) -> None:
        # Dropdown
        count_db_present = EC.presence_of_element_located(
            (By.ID, self.count_db_id))
        WebDriverWait(self.driver, self.avg_wait_time).until(count_db_present)
        count_db = self.driver.find_element(By.ID, self.count_db_id)
        count_db.click()

        # Option
        count_option_present = EC.element_to_be_clickable(
            (By.ID, self.count_option_id))
        count_option = WebDriverWait(
            self.driver, self.min_wait_time).until(count_option_present)
        count_option.click()

    def change_page_num(self) -> None:
        page_num_present = EC.presence_of_element_located(
            (By.CLASS_NAME, self.current_page_class))
        WebDriverWait(self.driver, self.avg_wait_time).until(page_num_present)
        page_num = self.driver.find_element(
            By.CLASS_NAME, self.current_page_class)
        page_num.send_keys(str(self.start_page_num))
        page_num.send_keys(Keys.RETURN)

        self.current_page_num = self.start_page_num

        time.sleep(self.min_wait_time)

    def paginate(self) -> None:
        if self.current_page_num < self.last_page_num:
            self.driver.implicitly_wait(self.avg_wait_time)
            class_bt = self.driver.find_element(
                By.CLASS_NAME, self.next_bt_class)
            class_bt.click()

            time.sleep(self.min_wait_time)

        self.current_page_num += 1

    def get_cell_text(self, cell):
        self.driver.implicitly_wait(0)
        div = cell.find_element(By.XPATH, ".//div")
        # country_spans = div.find_elements(
        #     By.XPATH, ".//spa n[contains(@class, 'i-country--')]")

        country_spans = div.find_elements(By.TAG_NAME, "span")

        if country_spans:
            results = []
            for span in country_spans:
                # Extract country code from class
                country_code = span.get_attribute(
                    "class").split("--")[-1].split()[0]
                country_name = self.id2country.get(country_code, country_code)
                # TODO delete
                self.all_country_spans.add(country_code)

                # Get immediately following text using JavaScript
                text = cell.parent.execute_script(
                    "return arguments[0].nextSibling.textContent.trim();", span)

                results.append(f"{country_name} - {text}")
            return ", ".join(results)

        # Simple case with no country codes
        return div.text.strip()

    def get_data(self) -> None:
        # Init dict with empty lists
        data = {k: [] for k in self.headers}

        # Get table and rows
        table = self.driver.find_element(By.XPATH, "//tbody")
        rows = table.find_elements(By.TAG_NAME, "tr")

        # Iterate through rows
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            values = [cell.text for cell in cells]
            values[0] = self.get_cell_text(cells[0])
            for key, value in zip(data.keys(), values):
                data[key].append(value)

        return data

    def create_xlsx(self) -> None:
        # Create a new workbook and select the active worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = self.xlsx_page_title

        # Write headers
        ws.append(self.headers)

        # Save the workbook
        wb.save(self.xlsx_path)

        return wb

    def write_to_xlsx(self,
                      wb: Workbook,
                      data: dict[str, list[str]]) -> None:
        ws = wb.active

        # Write new entries
        rows = zip(*data.values())
        for row in rows:
            ws.append(row)

        # Save the workbook
        wb.save(self.xlsx_path)

    def parse(self) -> None:
        # Set implicitly wait
        self.driver.implicitly_wait(self.avg_wait_time)

        self.driver.get(self.website)

        # Change number of entries per page
        self.change_num_entries()

        # Change current page num
        if self.start_page_num != self.current_page_num:
            self.change_page_num()

        # Create empty xlsx
        wb = self.create_xlsx()

        while self.current_page_num <= self.last_page_num:
            print(f"Parsing {self.current_page_num} page...")
            data = self.get_data()
            self.write_to_xlsx(wb, data)

            self.paginate()

        print("Scraping has finished.")

        # TODO delete
        print(self.all_country_spans)
