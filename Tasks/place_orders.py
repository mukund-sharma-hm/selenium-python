import time
import tests
import sys
import openpyxl
from Utilities.config import FILEPATH, PASSWORD
from Utilities import drivers
from openpyxl import Workbook
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import NoSuchElementException, TimeoutException
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

sys.path.append('../Utilities')
sys.path.append('../Drivers')


class PlaceOrders():
    def __init__(self):
        #using hints since webdriver will be initialised during runtime
        self.wait = None
        self.driver : WebDriver
        self.excel = None
        self.wb: Workbook


    def initialize_driver(self):
        self.driver = drivers.initialize_driver("chrome")
        self.wait = WebDriverWait(self.driver, 2)

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
        sheet = self.wb["Order Details"]

        for row in order_details.iter_rows(min_row=2, values_only=True):
            #returns tuple of row data
            if len(row) == 6:
                order_id, user_id, product_id, product_name, quantity, total_price, = row
            else:
                order_id, user_id, product_id, product_name, quantity, total_price, order_status = row

            self.login_user(user_id, PASSWORD)
            print({user_id:product_name})
            product_name_element = self.driver.find_element(By.XPATH, f"//*[text()='{product_name}']")
            product_name_element.click()
            self.wait()
            add_to_cart_element = self.driver.find_element(By.XPATH, "//*[text()='ADD TO CART']")
            add_to_cart_element.click()
            self.wait()

            cart_element = self.driver.find_element(By.ID, 'shopping_cart_container')
            cart_element.click()
            self.wait()
            checkout_element = self.driver.find_element(By.XPATH, "//*[@class='cart_footer']/a[2]")
            checkout_element.click()
            time.sleep(1)

            self.driver.find_element(By.ID, "first-name").send_keys("random")
            self.driver.find_element(By.ID, "last-name").send_keys("lastname")
            self.driver.find_element(By.ID, "postal-code").send_keys("175101")
            self.driver.find_element(By.XPATH, "//*[@class='checkout_buttons']/input").click()

            try:
                total_item_price_element = self.driver.find_element(By.XPATH, "//*[@class='summary_subtotal_label']")
                total_item_price_element_text = total_item_price_element.text
                total_item_price = total_item_price_element_text.split("$")[1]
                if total_price == total_item_price:
                    order_status = "success"
                else:
                    order_status = "failure"
                self.driver.find_element(By.XPATH, "//*[text()='FINISH']").click()

            except NoSuchElementException:
                order_status = "failure"

            try:
                success_message = self.wait.until(
                    EC.visibility_of_element_located((By.XPATH, "//h2[text()='THANK YOU FOR YOUR ORDER']")))
                if success_message.is_displayed():
                    order_status = "Success"

            except TimeoutException:
                order_status = "Failure"
            column_index = 7
            sheet.cell(row=1, column=column_index,value="Order Status")
            sheet.cell(row=int(order_id)+1, column=column_index, value=order_status)
            self.wb.save(FILEPATH)
