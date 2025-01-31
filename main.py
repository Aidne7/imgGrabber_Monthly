# Import necessary libraries and modules
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
import requests
import os
import sharedUtils  # Custom utility module
from docx import Document
import pyautogui
import pygetwindow as gw
from PIL import Image
import pandas as pd
import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image, ImageOps


# Functions for downloading maps and handling dates
def download_mrcc_map(state_value, element_value, datatype_value, filename, start_date, end_date, month, year):
    """Downloads a climate map from MRCC."""
    print(f"Downloading map: {filename}...")
    driver.get("https://mrcc.purdue.edu/CLIMATE/Maps/stnMap_btd.jsp")

    # Set state and date selection
    state_month_button = driver.find_element(By.CSS_SELECTOR, 'input[name="areaType"][value="1"]')
    state_month_button.click()
    state_select = driver.find_element(By.CSS_SELECTOR, f'option[value="{state_value}"]')
    state_select.click()

    # Set custom date range
    set_custom_date(start_date, end_date, month, year)

    # Select map variables
    variable_button = driver.find_element(By.CSS_SELECTOR,
                                          f'input[type="radio"][name="element"][value="{element_value}"]')
    variable_button.click()
    calculation_button = driver.find_element(By.CSS_SELECTOR,
                                             f'input[type="radio"][name="datatype"][value="{datatype_value}"]')
    calculation_button.click()

    # Submit request for map generation
    submit_button = driver.find_element(By.CSS_SELECTOR,
                                        'input[type="submit"][name="GetClimateData"][value="Create Map"]')
    submit_button.click()

    # Wait for the map to load and download it
    time.sleep(10)
    map_image = driver.find_element(By.CSS_SELECTOR, 'div#content_in_main img')
    src = map_image.get_attribute("src")
    download_image(src, filename)
    print(f"Map {filename} downloaded successfully.")


def set_custom_date(start_date, end_date, month, year):
    """Sets a custom date range for map generation."""
    print(f"Setting custom date range: {month}/{start_date}/{year} to {month}/{end_date}/{year}")

    # Start date selection
    start_month_select = Select(driver.find_element(By.ID, 'mo1'))
    start_month_select.select_by_value(month)
    start_day_select = Select(driver.find_element(By.ID, 'dy1'))
    start_day_select.select_by_value(start_date)
    start_year_select = Select(driver.find_element(By.ID, 'yr1'))
    start_year_select.select_by_value(year)

    # End date selection
    end_month_select = Select(driver.find_element(By.ID, 'mo2'))
    end_month_select.select_by_value(month)
    end_day_select = Select(driver.find_element(By.ID, 'dy2'))
    end_day_select.select_by_value(end_date)
    end_year_select = Select(driver.find_element(By.ID, 'yr2'))
    end_year_select.select_by_value(year)


def download_image(url, filepath):
    """Downloads an image from a URL and saves it to a file."""
    response = requests.get(url)
    with open(filepath, 'wb') as f:
        f.write(response.content)
    print(f"Image saved to {filepath}")

# Function to generate and save the table as an image with custom styling
def save_table_as_image(df, filepath):
    # Define the colors
    header_bg_color = "#70AD47"  # Dark green
    header_text_color = "#FFFFFF"  # White
    row_colors = ["#E2EFDA", "#C6E0B4"]  # Alternating light green and lighter green
    cell_edge_color = "#FFFFFF"  # White cell outlines

    # Update the column headers with new text and vertical stacking
    new_column_headers = [
        "Climate\nDivision",
        "Heating\nDegree\nDays",
        "Normal",
        "Departure",
        "Cooling\nDegree\nDays",
        "Normal",
        "Departure"
    ]

    # Create a figure and axis for the table
    fig, ax = plt.subplots(figsize=(8, 4))  # Adjust size as necessary
    ax.axis('tight')
    ax.axis('off')

    # Create a table and add it to the axis
    table = ax.table(cellText=df.values, colLabels=new_column_headers, loc='center', cellLoc='center')

    # Adjust font size for all text
    table.auto_set_font_size(False)
    table.set_fontsize(6)  # Reduced font size for data rows

    # Set the header background, text color, and row height, and adjust cell borders
    for (i, j), cell in table.get_celld().items():
        cell.set_edgecolor(cell_edge_color)  # Set cell border color to white

        if i == 0:  # Header row
            cell.get_text().set_color(header_text_color)
            cell.set_facecolor(header_bg_color)
            cell.get_text().set_fontsize(8)  # Slightly smaller header font size
            cell.get_text().set_fontweight('bold')
            cell.set_height(0.15)  # Adjust header row height
        else:  # Data rows
            cell.set_facecolor(row_colors[i % 2])  # Alternate row colors

    # Save the table as an image
    plt.savefig(filepath, bbox_inches='tight', dpi=300)
    plt.close(fig)
    print(f"Table saved as image: {filepath}")


# Function to crop the image and leave a small white margin
def crop_table_image(image_path, margin=10):
    # Open the saved image
    image = Image.open(image_path)

    # Get the bounding box of the non-white areas
    bbox = ImageOps.invert(image.convert('L')).getbbox()

    if bbox:
        # Add margin to the bounding box (if needed)
        left, top, right, bottom = bbox
        left = max(0, left - margin)
        top = max(0, top - margin)
        right = min(image.width, right + margin)
        bottom = min(image.height, bottom + margin)

        # Crop the image to the new bounding box with margin
        cropped_image = image.crop((left, top, right, bottom))

        # Save the cropped image back to the same path
        cropped_image.save(image_path)
        print(f"Cropped image saved with white space margin: {image_path}")
    else:
        print("No content found to crop.")


# Main script execution
if __name__ == "__main__":
    # Date calculations for the previous month
    today = datetime.today()

    # Calculate the first day of the current month and the last day of the previous month
    first_day_of_current_month = today.replace(day=1)
    last_day_of_last_month = first_day_of_current_month - timedelta(days=1)

    # Define start and end dates
    start_date = '01'
    end_date = str(last_day_of_last_month.day).zfill(2)
    month = str(last_day_of_last_month.month).zfill(2)
    year = str(last_day_of_last_month.year)

    # Display date information
    print(f"Start Date: {start_date}")
    print(f"End Date: {end_date}")
    print(f"Month: {month}")
    print(f"Year: {year}")

    # Initialize WebDriver and log in to MRCC
    driver = sharedUtils.initialize_driver()
    sharedUtils.login_mrcc(driver, 'aidenridgway22@gmail.com', 'cARpeU9Vx!')

    # List of maps to download with corresponding element and datatype values
    maps_to_download = [
        {"element": "1", "datatype": "1", "filename": f"AvgPrecip{month}{year[-2:]}.png"},
        {"element": "1", "datatype": "2", "filename": f"PrecipDep{month}{year[-2:]}.png"},
        {"element": "3", "datatype": "1", "filename": f"AvgTemp{month}{year[-2:]}.png"},
        {"element": "3", "datatype": "2", "filename": f"TempDep{month}{year[-2:]}.png"},
    ]

    # Create a new folder for saving images
    folder_name = f"Images{month}{year[-2:]}"
    folder_path = os.path.join(r"C:\Users\AcademicWeapon\Desktop\Byrd", folder_name)
    os.makedirs(folder_path, exist_ok=True)
    print(f"Created folder: {folder_name}")

    # Download MRCC maps and save them in the folder
    for map_info in maps_to_download:
        full_path = os.path.join(folder_path, map_info["filename"])
        download_mrcc_map("33", map_info["element"], map_info["datatype"], full_path, start_date, end_date, month, year)

    # Download CPC outlook images
    print("Downloading CPC images...")
    cpc_urls = [
        {'url': 'https://www.cpc.ncep.noaa.gov/products/predictions/long_range/lead01/off01_temp.gif',
         'filename': f'TempOutlook{month}{year[-2:]}.gif'},
        {'url': 'https://www.cpc.ncep.noaa.gov/products/predictions/long_range/lead01/off01_prcp.gif',
         'filename': f'PrecipOutlook{month}{year[-2:]}.gif'}
    ]

    for cpc in cpc_urls:
        file_path = os.path.join(folder_path, cpc['filename'])
        download_image(cpc['url'], file_path)

    # Download NASA weather images
    print("Downloading NASA weather images...")
    nasa_urls = [
        {'url': f'https://weather.ndc.nasa.gov/sport/dynamic/lis_IN//vsm0-40percent_{year}{month}{end_date}_00z_in.gif',
         'filename': f'40Soil{year}{end_date}.png'},
        {
            'url': f'https://weather.ndc.nasa.gov/sport/dynamic/lis_IN//vsm0-200percent_{year}{month}{end_date}_00z_in.gif',
            'filename': f'200Soil{year}{end_date}.png'}
    ]

    for nasa in nasa_urls:
        file_path = os.path.join(folder_path, nasa['filename'])
        download_image(nasa['url'], file_path)

    print("All images downloaded successfully.")

    # Data scraping from the DD table and Word document handling
    print("Scraping data from DD table")
    # Example usage (replace the Word table section with this)
    final_df = sharedUtils.scraper(driver, month, start_date, year, month, end_date, year)

    columns_to_drop = [4, 8]  # 0-based index of columns to drop
    wordDF = final_df.drop(final_df.columns[columns_to_drop], axis=1)

    # Calculate the average of each column
    avg_row = wordDF.mean().to_frame().T

    # Add the average row to the DataFrame
    wordDF = pd.concat([wordDF, avg_row], ignore_index=True)

    # Rounds the numbers to the nearest integer, converts them to integer, converts them to strings
    wordDF = wordDF.round(0).astype(int).astype(str)

    # Adds the word "State" to the bottom cell of the first column
    wordDF.iloc[-1, 0] = "State"

    # Define the path for saving the table image
    table_image_path = os.path.join(folder_path, f'DDTable{month}{year}.png')

    # Save the table as an image with custom styling
    save_table_as_image(wordDF, table_image_path)

    # After saving the table as an image, crop it
    save_table_as_image(wordDF, table_image_path)
    crop_table_image(table_image_path, margin=5)

    # Quit the WebDriver
    driver.quit()
    print("WebDriver closed.")
