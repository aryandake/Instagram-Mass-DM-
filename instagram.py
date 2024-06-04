from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from selenium.common.exceptions import NoSuchElementException
from dotenv import load_dotenv
import time
import pandas as pd
df = pd.read_excel('Format.xlsx')

load_dotenv()
count = 0

# Get username and password from environment variables
INSTAGRAM_USERNAME = os.getenv("INSTAGRAM_USERNAME")
INSTAGRAM_PASSWORD = os.getenv("INSTAGRAM_PASSWORD")

service = Service("chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.maximize_window()

insta_url = "https://instagram.com"
driver.get(insta_url)

time.sleep(5)

username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))
username_field.send_keys(INSTAGRAM_USERNAME + Keys.TAB)

time.sleep(5)

password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "password")))
password_field.send_keys(INSTAGRAM_PASSWORD + Keys.ENTER)


# Handle "Not Now" for saving login info
notnow_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/section/main/div/div/div/div/div")))
notnow_field.click()

# Handle "Turn on Notifications" pop-up
nonotification_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/div[3]/button[2]")))
nonotification_field.click()




def send_dm(naam, messages):
    try:
        # Click on the search icon
        search_icon = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[1]/div/div/div/div/div[2]/div[2]/span/div/a/div")))
        search_icon.click()
        time.sleep(2)
        
        # Find the search input field
        search_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[1]/div/div/div[2]/div/div/div[2]/div/div/div[1]/div/div/input")))
        search_box.click()
        search_box.clear()
        search_box.send_keys(naam)
        time.sleep(2)  # Wait for the search results to load
        search_box.send_keys(Keys.RETURN)
        time.sleep(1)
        search_box.send_keys(Keys.RETURN)
        time.sleep(2)

        # Click on the user's profile
        user_profile = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f"//a[@href='/{naam}/']")))
        user_profile.click()
        time.sleep(2)

        # Click on the Message button
        try:
            message_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[2]/div/div[2]/section/main/div/header/section[2]/div/div/div[2]/div/div[2]/div")))
            message_button.click()
            time.sleep(5)
        except NoSuchElementException:
            print(f"Message button not found for user: {username_field}")
            driver.get("https://www.instagram.com")  # Navigate back to the main page
            print("This is the First Error")
            time.sleep(3)
            return  # Skip to the next user if no message button is found


        # Send each message
        for message in messages:
            try:
                message_box = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/div[2]/div/div/div[1]/div[1]/div[2]/section/div/div/div/div[1]/div/div[2]/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div[1]/p")))
                message_box.send_keys(message)
                message_box.send_keys(Keys.RETURN)
                time.sleep(2)
            except Exception as e:
                print(f"Couldn't send message to {naam}:{e}")
                print("This is case 2")
                driver.get("https://instagram.com")
                time.sleep(3)
                return
                

    except Exception as e:
            print(f"Error processing user {naam}: {e}")
            print("This is case 3, the outermost case")
            driver.get("https://instagram.com")
            time.sleep(3)
            return
        
        

# Iterate through each row in the DataFrame and send DM
for index, row in df.iterrows():
    naam = row['naam']  # Get the username from the 'naam' column
    messages = [row[f'message{i}'] for i in range(1, 4) if pd.notna(row[f'message{i}'])]  # Get the messages from 'message1', 'message2', 'message3' columns
    
    # Call the send_dm function with the username and messages
    send_dm(naam, messages)
    count = count + 1
    print(f"The current count is {count}")

# Close the WebDriver
driver.quit() 


