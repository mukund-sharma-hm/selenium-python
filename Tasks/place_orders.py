import openpyxl
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By

from Utilities.config import FILEPATH, PASSWORD

from Utilities import drivers

class PlaceOrders():
    def __init__(self):
        #using hints since webdriver will be initialised during runtime
        self.driver : WebDriver
        self.excel = None
        self.wb = None


    def initialize_driver(self):
        self.driver = drivers.initialize_driver("chrome")

    def load_excel(self, excel_file):
        self.excel = excel_file
        self.wb = openpyxl.load_workbook(FILEPATH)

    def login_user(self, username, password):
        self.driver.get("https://www.saucedemo.com/v1/")
        self.driver.maximize_window()
        self.driver.find_element(By.ID, "user-name").send_keys(username)
        self.driver.find_element(By.ID, "password").send_keys(password)
        self.driver.find_element(By.ID, "login-button").click()

    def place_orders(self):
        order_details = self.wb['Order Details']
        for row in order_details.iter_rows(min_row=2, values_only=True):
            #returns tuple of row data
            order_id, user_id, product_id, product_name, quantity, total_price = row
            self.login_user(user_id, PASSWORD)
            element = self.driver.find_element(By.XPATH, f"//*[text()='{product_name}']")
            print(element)

if __name__ == "__main__":

    orders = PlaceOrders()
    orders.initialize_driver()
    orders.load_excel(FILEPATH)
    # orders.login_user()
    orders.place_orders()