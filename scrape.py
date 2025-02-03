import time

from selenium import webdriver

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from openpyxl import Workbook, styles
from openpyxl.utils import get_column_letter


class Parser:
    def __init__(self) -> None:
        self.driver = webdriver.Edge()

        # Count IDs
        self.count_db_id = "ComboBox1-input"
        self.count_option_id = "ComboBox1-list4"

        # Pages classes
        self.current_page_class = "ecc-page-number-input"
        self.last_page_class = "eec-page-count"
        self.next_bt_class = "arrow-right"

        # Page
        self.start_page = 1
        self.current_page_num = self.start_page
        self.__last_page_num = None

        # Wait time
        self.min_wait_time = 20
        self.avg_wait_time = 150
        self.max_wait_time = 500

        # Countries codes
        self.id2country = {"am": "Армения",
                           "by": "Беларусь",
                           "ru": "Россия",
                           "kg": "Кыргызстан",
                           "kz": "Казахстан"}

        # Xlsx
        self.xlsx_path = "Реестр ЛС ЕАЭС.xlsx"
        self.xlsx_page_title = "Реестр"

        self.all_country_spans = set()

    @ property
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

    def paginate(self) -> None:
        self.driver.implicitly_wait(self.avg_wait_time)
        class_bt = self.driver.find_element(By.CLASS_NAME, self.next_bt_class)
        class_bt.click()
        self.current_page_num += 1
        time.sleep(self.min_wait_time)

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

    def get_data(self, data: dict[str, list[str]]):
        table = self.driver.find_element(By.XPATH, "//tbody")
        rows = table.find_elements(By.TAG_NAME, "tr")

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            values = [cell.text for cell in cells]
            values[0] = self.get_cell_text(cells[0])
            for key, value in zip(data.keys(), values):
                data[key].append(value)

    def write_to_xlsx(self, data: dict[str, list[str]]) -> None:
        # Create a new workbook and select the active worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = self.xlsx_page_title

        # Write headers
        ws.append(list(data.keys()))

        # Write entries
        rows = zip(*data.values())
        for row in rows:
            ws.append(row)

        # Save the workbook
        wb.save(self.xlsx_path)

    def parse(self, website: str) -> None:
        # Set implicitly wait
        self.driver.implicitly_wait(self.avg_wait_time)

        self.driver.get(website)

        self.change_num_entries()
        self.change_page_num()

        data = {"trade_name": [],
                "int_name": [],
                "rel_form": [],
                "manufacturer": [],
                "properties": [],
                "certificate": [],
                "update": []}

        while self.current_page_num < self.last_page_num:
            self.get_data(data)
            self.paginate()

        self.write_to_xlsx(data)
        print(self.all_country_spans)
        # time.sleep(self.max_wait_time)


def main() -> None:
    website = "https://portal.eaeunion.org/sites/commonprocesses/ru-ru/Pages/DrugRegistrationDetails.aspx"

    parser = Parser()
    parser.parse(website)


if __name__ == "__main__":
    main()
