"""

"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.firefox.service import Service as FirefoxService


def initialize_driver(browser_name):
    """
    :param browser_name: name of the browser you want to run
    :return: the instance of the driver class for chosen browser
    """
    match browser_name:
        case "chrome":
            serv_obj = ChromeService('../Drivers/chromedriver.exe')
            driver = webdriver.Chrome(service=serv_obj)
            return driver

        case "edge":
            serv_obj = EdgeService('../Drivers/msedgedriver.exe')
            driver = webdriver.Chrome(service=serv_obj)
            return driver

        case "firefox":
            serv_obj = FirefoxService('../Drivers/geckodriver.exe')
            driver = webdriver.Chrome(service=serv_obj)
            return driver
