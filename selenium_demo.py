from selenium import webdriver
from tkinter import *
import tkinter.messagebox
from flask import Flask
import time
import win32com.client as win32

app = Flask(__name__)


@app.route('/')
def home():
    driver = webdriver.Chrome('./chromedriver')
    driver.maximize_window()
    # Part 1: Opening google chrome nad searching in google
    try:
        driver.get("https://www.google.com")
        search_field = driver.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input")
        search_field.send_keys("Current temperature in Mumbai")
        search_button = driver.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[3]/center/input[1]")
        search_button.click()
        getTemp = str(driver.find_element_by_xpath('//*[@id="wob_tm"]').text).strip()
        search_field = driver.find_element_by_xpath('//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[2]/input')
        search_field.clear()
        search_field.send_keys("Current time in Mumbai")
        search_button = driver.find_element_by_xpath('//*[@id="tsf"]/div[1]/div[1]/div[2]/button')
        search_button.click()
        getTime = str(driver.find_element_by_xpath('//*[@id="rso"]/div[1]/div/div/div[1]/div[1]').text).strip()
        driver.close()
        if str(getTime[-2:]).lower() == 'am':
            time_now = int(getTime[:-6])
        else:
            time_now = int(getTime[:-6])+12
        if int(getTime[:-6]) == 12 and str(getTime[-2:]).lower() == 'am':
            time_now = 0
        greeting = ''
        if 3 <= time_now < 12:
            greeting = 'Good Morning!'
        elif 12 <= time_now < 16:
            greeting = 'Good Afternoon!'
        elif 16 <= time_now < 20:
            greeting = 'Good Evening!'
        else:
            greeting = 'Good Night!'
        print("Hi {}! The current temperature in Mumbai is {}°C and time is {}.".format(greeting, str(getTemp), str(getTime)))
    except Exception as error:
        error = str(error)
        print('Error in part 1:' + error)

    # Part 2: Service now access
    # try:
    #     driver.get('https://signon.service-now.com/ssologin.do?RelayState=%252Fapp%252Ftemplate_saml_2_0%252Fk317zlfESMUHAFZFXMVB%252Fsso%252Fsaml%253FRelayState%253Dhttps%25253A%25252F%25252Fdeveloper.servicenow.com%25252Fdev.do&redirectUri=&email=')
    #     time.sleep(120)
    #     driver.close()
    # except Exception as error:
    #     error = str(error)
    #     print('Error in part 2:' + error)

    # Part 3: Read mail from outlook
    try:
        outlook = win32.Dispatch('outlook.application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(6)
        mail = inbox.Items
        lastMail = mail.GetLast()
        body = lastMail.Body
        print(body)
    except Exception as error:
        error = str(error)
        print('Error in part 3:' + error)

    # root = Tk()
    # root.withdraw()
    # tkinter.messagebox.showinfo(greeting, "Hi! The current temperature in Mumbai is {}°C and time is {}.".format(str(getTemp), str(getTime)))
    # root.destroy()
    # root.mainloop()

    return "Thank you! The selenium PoC is completed."


if __name__ == '__main__':
    app.run(host='localhost', port=8080, debug=True)
