from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By

import time


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
        self.current_page_num = 1
        self.__last_page_num = None

        # Wait time
        self.min_wait_time = 10
        self.avg_wait_time = 120
        self.max_wait_time = 500

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
        self.driver.implicitly_wait(self.min_wait_time)

        # Option
        count_option_present = EC.element_to_be_clickable(
            (By.ID, self.count_option_id))
        count_option = WebDriverWait(
            self.driver, self.min_wait_time).until(count_option_present)
        count_option.click()

        # Wait until load page with new num entries
        self.driver.implicitly_wait(self.avg_wait_time)

    def paginate(self) -> None:
        class_bt = self.driver.find_element(By.CLASS_NAME, self.next_bt_class)
        class_bt.click()
        self.current_page_num += 1
        self.driver.implicitly_wait(self.min_wait_time)

    def parse(self, website: str) -> None:
        self.driver.get(website)

        self.change_num_entries()

        while self.current_page_num < self.last_page_num:
            self.paginate()

        time.sleep(self.max_wait_time)


def main() -> None:
    website = "https://portal.eaeunion.org/sites/commonprocesses/ru-ru/Pages/DrugRegistrationDetails.aspx"

    parser = Parser()
    parser.parse(website)


if __name__ == "__main__":
    main()
