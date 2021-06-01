import openpyxl
import data_driven_methods
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get("http://demo.guru99.com/test/newtours/index.php")
driver.maximize_window()
data_file = "Login_Data.xlsx"
number_of_rows = data_driven_methods.row_count(data_file, 'Login')
number_of_column = data_driven_methods.column_count(data_file, 'Login')

for r in range(2, number_of_rows + 1):
    username = data_driven_methods.read_data(data_file, 'Login', r, 1)
    password = data_driven_methods.read_data(data_file, 'Login', r, 2)

    driver.find_element(By.NAME, 'userName').send_keys(username)
    driver.find_element(By.NAME, 'password').send_keys(password)
    driver.find_element(By.NAME, 'submit').click()

    if driver.title == "Login: Mercury Tours":
        data_driven_methods.write_data(data_file, 'Login', r, 3, "Passed")
    else:
        data_driven_methods.write_data(data_file, 'Login', r, 3, "Failed")

    driver.find_element(By.LINK_TEXT, 'Home').click()
