from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from io import StringIO
import pandas as pd


def initialize_driver(headless=True):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    return driver


def login_mrcc(driver, email, password):
    print("test")
    driver.get("https://mrcc.purdue.edu/CLIMATE/")
    print("test1")
    email_field = driver.find_element(By.ID, "textName")
    password_field = driver.find_element(By.ID, "textPwd")
    print("test3")
    email_field.send_keys(email)
    password_field.send_keys(password)
    print("test4")
    login_button = driver.find_element(
        By.CSS_SELECTOR, 'input[type="submit"][name="Submit"][value="Log In"]'
    )
    print("test5")
    login_button.click()

    print("Logged in successfully.")


# Function to scrape data from the MRCC website
def scraper(
    driver, start_month, start_day, start_year, end_month, end_day, end_year, state="oh"
):
    def textGrabber(variable):
        driver.get("https://mrcc.purdue.edu/CLIMATE/nClimDiv/clidiv_daily_data.jsp")

        # Select state and dates
        state_select = Select(driver.find_element(By.ID, "selectState"))
        state_select.select_by_value(state)

        Select(driver.find_element(By.ID, "mo1")).select_by_value(str(int(start_month)))
        Select(driver.find_element(By.ID, "dy1")).select_by_value(str(int(start_day)))
        Select(driver.find_element(By.ID, "yr1")).select_by_value(start_year)
        Select(driver.find_element(By.ID, "mo2")).select_by_value(str(int(end_month)))
        Select(driver.find_element(By.ID, "dy2")).select_by_value(str(int(end_day)))
        Select(driver.find_element(By.ID, "yr2")).select_by_value(str(int(end_year)))

        # Select the degree day variable
        radio_button = driver.find_element(
            By.CSS_SELECTOR, f'input[name="variable"][type="radio"][value="{variable}"]'
        )
        radio_button.click()

        # Submit form
        submit_button = driver.find_element(By.ID, "GetClimateData")
        submit_button.click()

        # Wait and extract data
        driver.implicitly_wait(10)
        element = driver.find_element(By.ID, "result")
        ddText = element.text
        lines = ddText.split("\n")
        lines = lines[6:-6]  # Skip header and footer
        table_text = "\n".join(lines)
        return table_text.strip()

    # Process data
    df1 = pd.read_csv(StringIO(textGrabber(variable="1")), delim_whitespace=True)
    df2 = pd.read_csv(StringIO(textGrabber(variable="2")), delim_whitespace=True)

    driver.quit()  # Close WebDriver

    # Combine DataFrames
    final_df = pd.concat(
        [
            df1[["cd", "DD", "Normal", "Departure", "Percent"]],
            df2[["DD", "Normal", "Departure", "Percent"]],
        ],
        axis=1,
    )
    final_df.columns = [
        "cd",
        "W_DD",
        "W_Normal",
        "W_Departure",
        "W_Percent",
        "C_DD",
        "C_Normal",
        "C_Departure",
        "C_Percent",
    ]

    print(final_df)
    final_df.info()

    return final_df
