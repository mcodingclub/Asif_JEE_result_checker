import re
from openpyxl import Workbook, load_workbook
from PIL import Image
import pytesseract
import pandas as pd
import requests
import numpy as np
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys


import os

# Tesseract path
tesseract_path = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# Configure Selenium WebDriver
driver_path = "E:\\chromedriver_win32\\chromedriver.exe"
service = Service(driver_path)
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Configure OCR
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# Function to capture screenshot of the entire browser page
def capture_full_browser_screenshot(driver, path):
    # Get the full height of the page
    full_height = driver.execute_script(
        "return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);"
    )

    # Set the window size to the full height
    driver.set_window_size(driver.get_window_size()["width"], full_height)

    # Capture the screenshot
    driver.save_screenshot(path)

    # Reset the window size to its original value
    driver.set_window_size(driver.get_window_size()["width"], driver.get_window_size()["height"])

    # Open the screenshot image
    image = Image.open(path)
    return image

# Function to extract captcha text from image
def get_captcha(driver, element, path):
    location = element.location
    size = element.size

    driver.save_screenshot(path)

    left = location["x"]
    top = location["y"]
    right = location["x"] + size["width"]
    bottom = location["y"] + size["height"]

    image = Image.open(path)
    image_cropped = image.crop((left, top, right, bottom))
    image_cropped.save(path)

    captcha = pytesseract.image_to_string(image_cropped)
    captcha = captcha.replace(" ", "").strip()

    return captcha

try:
    # Load data from Excel file
    dataframe = pd.read_excel('JEE.xlsx')

    # Create an empty dataframe to store the final results
    final_results = pd.DataFrame()

    # Loop through each row in the dataframe
    for i in range(len(dataframe)):
        try:
            # Retrieve data from the current row
            application_number = str(dataframe.loc[i, 'Application Number'])
            day = str(dataframe.loc[i, 'Day']).zfill(2)
            month = str(dataframe.loc[i, 'Month'])
            year = str(dataframe.loc[i, 'Year'])

            # List of URLs
            urls = [
                'https://ntaresults.nic.in/resultservices/JEEMAINauth23s2p1',
                'https://ntaresults.nic.in/resultservices/JEEMAINauth23s2p1',
                # Add more URLs as needed
            ]

            # Loop through each URL and fill the form
            for j, url in enumerate(urls):
                try:
                    # Navigate to the webpage
                    driver.get(url)

                    # Wait for the captcha image to be present
                    element_locator = (By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_captchaimg"]')
                    wait = WebDriverWait(driver, 30)
                    element = wait.until(EC.presence_of_element_located(element_locator))

                    # Fill the form fields
                    Application_input = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtRegNo")
                    Application_input.clear()
                    Application_input.send_keys(application_number)

                    day_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$ddlday')
                    day_input.send_keys(Keys.CONTROL + "a")
                    day_input.send_keys(Keys.DELETE)
                    day_input.send_keys(day)

                    month_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$ddlmonth')
                    month_input.send_keys(Keys.CONTROL + "a")
                    month_input.send_keys(Keys.DELETE)
                    month_select = Select(month_input)
                    month_select.select_by_visible_text(month)

                    year_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$ddlyear')
                    year_input.send_keys(Keys.CONTROL + "a")
                    year_input.send_keys(Keys.DELETE)
                    year_input.send_keys(year)

                    # Wait for the captcha image to be present
                    element = wait.until(EC.presence_of_element_located(element_locator))

                    # Capture captcha image and extract text for the first attempt
                    img_element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_captchaimg"]')
                    captcha_path = os.path.join(os.getcwd(), f"captcha_{i}_{j}.png")
                    captcha = get_captcha(driver, img_element, captcha_path)

                    captcha_input = driver.find_element(By.NAME, 'ctl00$ContentPlaceHolder1$Secpin')
                    captcha_input.clear()
                    captcha_input.send_keys(Keys.CONTROL + "a")
                    captcha_input.send_keys(Keys.DELETE)
                    captcha_input.send_keys(captcha)

                    submit_btn = driver.find_element(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_Submit1')
                    submit_btn.click()

                    result_screenshot_path1 = f"result1_{i}_{j}.png"
                    result_screenshot1 = capture_full_browser_screenshot(driver, result_screenshot_path1)

                    # Wait for the result page to load and become visible for the first attempt
                    wait.until(EC.visibility_of_element_located((By.ID, 'some_unique_id_of_result_page_element')))

                    # Capture screenshot of the result page for the first attempt
                    result_screenshot_path1 = os.path.join(os.path.dirname(captcha_path), f"result1_{i}_{j}_{os.path.basename(captcha_path)}")
                    result_screenshot1.save(result_screenshot_path1)

                    # Perform OCR on the result page for the first attempt
                    result1 = pytesseract.image_to_string(result_screenshot_path1)
                    print(result1,"*****************8")

                    # Extract relevant information from the result1 text
                    lines = result1.split("\n")
                    Application = ""
                    Candidates_Name = ""
                    Total = ""

                    # Find relevant lines and extract information
                    for line in lines:
                        if "Application No.:" in line:
                            Application = line.split(":")[1].strip()
                        elif "Candidates_Name:" in line:
                            Candidates_Name = line.split(":")[1].strip()
                        elif "Total:" in line:
                            Total = line.split(":")[1].strip()

                    # Add the extracted information to the final_results dataframe
                    final_results.loc[i, 'Application_No.'] = Application
                    final_results.loc[i, "Candidate's Name"] = Candidates_Name
                    final_results.loc[i, 'Total'] = Total

                except Exception as e:
                    print(f"Exception in attempt {j+1} for row {i+1}: {str(e)}")

        except Exception as e:
            print(f"Exception in row {i+1}: {str(e)}")

    # Save the final_results dataframe to an Excel file
    final_results.to_excel("result.xlsx", index=False)
    print("Extraction completed successfully.")

except Exception as e:
    print(f"An error occurred: {str(e)}")

# Close the WebDriver
driver.quit()

#     
# image_path = f"result1_0_0.png"
# image = Image.open(image_path)

# # Apply OCR to the image
# ocr_text = pytesseract.image_to_string(image)
# print(ocr_text,"*****************")

# # Extract relevant information from the OCR text
# lines = ocr_text.split("\n")
# Application = ""
# Candidates_Name = ""
# Total = ""

# # Find relevant lines and extract information
# for line in lines:
#     if "Application No.:" in line:
#         Application = line.split(":")[1].strip()
#     elif "Candidates_Name:" in line:
#         Candidates_Name = line.split(":")[1].strip()
#     elif "Total:" in line:
#         Total = line.split(":")[1].strip()

# # Save data in Excel file
# wb = Workbook()
# ws = wb.active
# ws.append(["Application_No.", "Candidate's Name", "Total"])
# ws.append([Application, Candidates_Name, Total])
# # wb.save("results.xlsx")
# excel_file_path = "results.xlsx"
# wb.save(excel_file_path)
# print("Data saved in Excel file:", excel_file_path)




# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>















