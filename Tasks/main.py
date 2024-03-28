from Tasks import tests
from Tasks.place_orders import PlaceOrders
from Tasks.tests import get_user_credentials, login_and_record_errors, standard_user_product_details, orders, close_driver
from Utilities.config import FILEPATH

if __name__ == "__main__":
    # task1
    get_user_credentials()
    # task2
    login_and_record_errors()
    # task 3
    standard_user_product_details()
    # task4
    orders()
    close_driver()
    # tests.run_tests()
    # task 5
    orders = PlaceOrders()
    orders.initialize_driver()
    orders.load_excel(FILEPATH)
    orders.place_orders()
