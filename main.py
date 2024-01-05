from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import openpyxl
import time
import sys

# Specify the path to the ChromeDriver executable
chrome_driver_path =r'C:\Users\Inderkiran Singh\Downloads\chromedriver_win32\chromedriver.exe'

# Load Excel workbook
filex = openpyxl.load_workbook(r'C:\Users\Inderkiran Singh\Downloads\Mobile Numbers.xlsx')
sh = filex["Sheet1"]

print("*************************************************************")
print("SMS sending program...")
print("*************************************************************")
print("\n")

print("Please enter the limit from the excel file:")
st = int(input("1. Start point"))
ed = int(input("2. End point"))

# Initialize Chrome browser
driver = webdriver.Chrome(executable_path=chrome_driver_path)
driver.get(r'https://messages.google.com/web/conversations')

# Wait for the page to load
wait = WebDriverWait(driver, 10)
wait.until(EC.title_contains("Messages"))

print("Successfully scanned.....")

# Click on the new conversation button
new_conversation_button_xpath = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/div/mw-fab-link/a'
wait.until(EC.presence_of_element_located((By.XPATH, new_conversation_button_xpath))).click()

# Start a new group conversation
new_conversation_xpath = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-new-conversation-start-group-button/button'
wait.until(EC.presence_of_element_located((By.XPATH, new_conversation_xpath))).click()

count = 0
while st <= ed:
    try:
        cl = sh.cell(st, 1)
        element = driver.find_element_by_xpath('//*[@id="mat-chip-list-0"]/div/input')
        element.send_keys(cl.value)
        time.sleep(1)
        driver.find_element_by_xpath(
            '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-contact-selector-button/button').click()
        st += 1
        count += 1
    except NoSuchElementException as e:
        print(f"Element not found: {e}")

# Add message
message = "Type message here....."
driver.find_element_by_xpath(
    '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/button').click()

# Wait for the message input field to be present
message_input_xpath = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div/mws-message-compose/div/div[2]/div/mws-autosize-textarea/textarea'
wait.until(EC.presence_of_element_located((By.XPATH, message_input_xpath))).send_keys(message)

# Send the message
send_button_xpath = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div/mws-message-compose/div/mws-message-send-button/button'
wait.until(EC.presence_of_element_located((By.XPATH, send_button_xpath))).click()

print(f"{count} messages sent already..")

# Close the browser
driver.quit()
sys.exit()
