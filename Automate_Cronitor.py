from operator import index
import unittest
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.oauth2 import service_account
from googleapiclient.discovery import build
import time
from pynput.keyboard import Controller
import pyautogui
import os
from dotenv import load_dotenv
import pandas as pd
from openpyxl import load_workbook
import random
import shutil
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
import subprocess
import requests
import cronitor

# Load environment variables
load_dotenv()

# Cronitor API Key
cronitor.api_key = os.getenv("CRONITOR_API_KEY")

keyboard = Controller()

class FrontendTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        firefox_options = Options()
        firefox_options.add_argument('--no-sandbox')  # Prevent sandbox error
        
        # Set Firefox preferences for downloads
        download_path = os.getenv("DOWNLOAD_PATH")
        firefox_options.set_preference("browser.download.folderList", 2)
        firefox_options.set_preference("browser.download.dir", download_path)
        firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")

        service = Service(os.getenv("GECKODRIVER_PATH"))
        cls.driver = webdriver.Firefox(service=service, options=firefox_options)
        cls.driver.set_window_size(1400, 900)
        cls.driver.implicitly_wait(10)

        # Save download path for use in download checks
        cls.download_path = download_path
        cls.sheet_service = cls.init_google_sheet()

        # Initialize Cronitor Monitor
        cls.monitor = cronitor.Monitor(os.getenv("CRONITOR_MONITOR_ID"))
        cls.monitor.ping(message="Automate Test setup initialized")

    @classmethod
    def tearDownClass(cls):
        cls.monitor.ping(message="Automate Test teardown completed")
        cls.driver.quit()

    @staticmethod
    def init_google_sheet():
        json_key_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_KEY_PATH")
        credentials = service_account.Credentials.from_service_account_file(
            json_key_path, 
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        service = build('sheets', 'v4', credentials=credentials)
        return service

    def log_result_to_sheet(self, row, status, column):
        spreadsheet_id = os.getenv("SPREADSHEET_ID")
        range_name = f'Automatedtest!{column}{row}'
        values = [[status]]
        body = {'values': values}
        try:
            self.sheet_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption="RAW",
                body=body
            ).execute()
            print(f"Updated cell {range_name} with status: {status}")
        except Exception as e:
            print(f"Failed to log result to sheet: {e}")

    def login(self):
        self.driver.maximize_window()
        self.driver.get(os.getenv("WEBSITE_URL"))
        print("Opened the website.")

        login_button = WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, os.getenv("LOGIN_BUTTON_XPATH")))
        )
        login_button.click()
        print("Login button clicked.")

        time.sleep(1)

        email = WebDriverWait(self.driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, os.getenv("EMAIL_INPUT_XPATH")))
        )
        email.send_keys(os.getenv("TEST_EMAIL"))
        print("Email entered.")

        password = WebDriverWait(self.driver, 30).until(
            EC.visibility_of_element_located((By.XPATH, os.getenv("PASSWORD_INPUT_XPATH")))
        )
        password.send_keys(os.getenv("TEST_PASSWORD"))
        print("Password entered.")

        sign_in_button = WebDriverWait(self.driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, os.getenv("SIGN_IN_BUTTON_XPATH")))
        )
        sign_in_button.click()
        print("Logged in successfully.")

        time.sleep(2)

    def test_studio_misreading(self):
        print("Starting studio test...")
        
        time.sleep(3)
        try:
            self.driver.get(os.getenv("WEBSITE_URL"))
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            print("Website is accessible.")
        except Exception as e:
            self.monitor.ping(
                message="Website is down",
                metrics={'count': 1, 'error_count': 1}
            )
            print(f"Website is down: {e}")
            self.log_result_to_sheet(2, "Fail", column="A")
            raise 
    
        try:
            self.login()
            self.log_result_to_sheet(2, "Pass", column="A")
            self.monitor.ping(message="Login Test Passed", metrics={'count': 1, 'error_count': 0})
            print("Login - Passed")
        except Exception as e:
            self.log_result_to_sheet(2, "Fail", column="A")
            self.monitor.ping(message=f"Login Test Failed: {e}", metrics={'count': 1, 'error_count': 1})
            print(f"Login - Failed: {e}")

if __name__ == "__main__":
    unittest.main(verbosity=2)
