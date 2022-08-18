from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from webdriver_manager.chrome import ChromeDriverManager

import time
import random
from datetime import datetime

from send_email_outlook import send_email

SIGN_IN_URL = "https://ais.usvisa-info.com/en-ca/niv/users/sign_in"

LOGIN_PAGE = "Sign in or Create an Account | Official U.S. Department of State Visa Appointment Service | " + \
                "Canada | English"
SUMMARY_PAGE = "Groups | Official U.S. Department of State Visa Appointment Service | Canada | English"
HOME_PAGE = "Official U.S. Department of State Visa Appointment Service | Canada | English"
SCHEDULE_PAGE = "Schedule Appointments | Official U.S. Department of State Visa Appointment Service | Canada | English"
MAINTAINANCE_PAGE = "Under construction (503)"

PAGES = [LOGIN_PAGE, SUMMARY_PAGE, HOME_PAGE, SCHEDULE_PAGE, MAINTAINANCE_PAGE]

EMAIL = "email_address"
USER_EMAIL = "user_email"   # email used to login
PASSWARD = "password"   # passward used to login
# the script will schedule date that is in or before the desired month, 18 means (18 - 12 = June) of next year.
DESIRED_MONTH = 20   


def openLink():
    service = Service(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.get(SIGN_IN_URL)
    assert driver.title == LOGIN_PAGE
    return driver


def logIn(driver, log_file):

    time.sleep(random.random() * 1)
    driver.implicitly_wait(5)

    for i in range(3):
        try:
            if driver.title != LOGIN_PAGE:
                # log(log_file, "ERROR: it's not on login page")
                return
            driver.find_element(By.NAME,
                                "user[email]").send_keys(USER_EMAIL)
            time.sleep(random.random() * 1)
            driver.implicitly_wait(5)
            driver.find_element(By.NAME,
                                "user[password]").send_keys(PASSWARD)
            time.sleep(random.random() * 1)
            driver.implicitly_wait(5)
            driver.find_element(
                By.XPATH, "//label[@for='policy_confirmed']/div").click()
            time.sleep(random.random() * 1)
            driver.implicitly_wait(5)
            driver.find_element(By.NAME, "commit").click()

            driver.implicitly_wait(5)
            if driver.title == SUMMARY_PAGE:
                break
        except ElementNotInteractableException:
            log(log_file, "from login %d" % i)
            log(log_file, str(driver.title))
            driver.get("https://ais.usvisa-info.com/en-ca/niv/users/sign_in")

    driver.implicitly_wait(5)
    assert driver.title == SUMMARY_PAGE


def home(driver, log_file):

    time.sleep(random.random() * 1)
    driver.implicitly_wait(5)

    if driver.title != SUMMARY_PAGE:
        log(log_file, "ERROR: it's not no summary page")
        return

    driver.find_element(By.XPATH, "//a[@class='button primary small']").click()

    driver.implicitly_wait(5)
    assert driver.title == HOME_PAGE


def reschedule(driver, log_file):

    time.sleep(random.random() * 1)
    driver.implicitly_wait(5)

    # click reschedule
    for i in range(60):
        if driver.title != HOME_PAGE:
            log(log_file, "ERROR: it's not on home page")
            return
        driver.find_element(By.XPATH,
                            "//section[@id='forms']/ul/li[4]").click()
        time.sleep(random.random() * 1)
        driver.implicitly_wait(5)
        driver.find_element(
            By.XPATH,
            "//section[@id='forms']/ul/li[4]/div/div/div[2]/p[2]/a").click()

        driver.implicitly_wait(5)

        if driver.title == "429 Too Many Requests":
            driver.back()
            time.sleep(0.1)
            driver.implicitly_wait(5)
        elif driver.title == SCHEDULE_PAGE:
            log(log_file, str(i))
            break
    assert driver.title == SCHEDULE_PAGE


def schedule(driver, max_i, log_file):

    time.sleep(random.random() * 1)
    driver.implicitly_wait(5)

    if driver.title != SCHEDULE_PAGE:
        log(log_file, "ERROR: it's not on schedule page")
        return -1

    # find and click Date of Appointment, to show the calendar
    try:
        driver.find_element(
            By.ID, "appointments_consulate_appointment_date_input").click()
    except ElementNotInteractableException:
        close(driver)
        return -1

    # find if any date is available, if not click next month.
    i = 0
    for _ in range(max_i):
        if i >= max_i:
            break

        if driver.title != SCHEDULE_PAGE:
            log(log_file, "ERROR: it's not on schedule page")
            return -1

        try:
            # driver.implicitly_wait(5)
            driver.find_element(By.XPATH, "//td[@class=' undefined']").click()
            break
        except NoSuchElementException:
            driver.find_element(By.XPATH, "//a[@data-handler='next']").click()
            i += 1

    if i < max_i and i > 0:  # found available date
        # select the first available date
        # time.sleep(random.random() * 1)
        # driver.implicitly_wait(5)
        # driver.find_element(By.XPATH, "//td[@class=' undefined']").click()

        # find and click Time of Appointment
        time.sleep(random.random() * 1)
        driver.implicitly_wait(5)
        driver.find_element(
            By.XPATH,
            "//select[@name='appointments[consulate_appointment][time]']/option[2]"
        ).click()
        send_email(
            EMAIL,
            "Found desired date in month %d" % (i + datetime.now().month + 1))

        submit(driver)

        driver.implicitly_wait(5)
        driver.find_element(By.XPATH, "//a[@class='button alert']").click()

        # close(driver)

        return i

    close(driver)
    return -2


def sign_out(driver):
    driver.implicitly_wait(5)
    driver.find_element(By.XPATH,
                        "//div[@class='top-bar-right']/div/ul/li[3]").click()
    driver.implicitly_wait(5)
    driver.find_element(
        By.XPATH,
        "//div[@class='top-bar-right']/div/ul/li[3]/ul/li[4]").click()
    driver.get(SIGN_IN_URL)


def submit(driver):
    time.sleep(random.random() * 1)
    driver.implicitly_wait(5)
    driver.find_element(By.XPATH, "//input[@name='commit']").click()


def close(driver):
    time.sleep(random.random() * 1)
    driver.implicitly_wait(5)
    driver.find_element(By.XPATH, "//a[text()='Close']").click()


def log(log_file, msg):
    with open(log_file, 'a') as f:
        f.write(msg + '\n')


def under_construction(driver):
    if driver.title == "Under construction (503)":
        return True


def too_many_request(driver):
    if driver.title == "429 Too Many Requests":
        driver.back()


def main():
    log_file = "log_" + str(datetime.now()).replace(' ', '_').replace(
        '.', "_").replace(':', "-")

    with open(log_file, 'w') as f:
        f.write("Current before month: %d \n" % DESIRED_MONTH)

    max_i = DESIRED_MONTH - datetime.now().month - 1
    driver = openLink()

    start_time = time.time()
    log(log_file, "start_time:" + str(start_time))
    while (1):
        if time.time() - start_time >= 3600 * 3.5:
            log(log_file, "Signed out")
            start_time = time.time()
            sign_out(driver)

        logIn(driver, log_file)
        home(driver, log_file)
        reschedule(driver, log_file)
        scheduled_i = schedule(driver, max_i, log_file)

        if under_construction(driver):
            driver.get(SIGN_IN_URL)
            continue

        too_many_request(driver)

        if scheduled_i == -1:
            log(log_file, "No available date. %s" % datetime.now())
        elif scheduled_i == -2:
            log(log_file, "No desired date. %s" % datetime.now())
        else:
            max_i = scheduled_i - 1
            scheduled_month = (scheduled_i + datetime.now().month + 1)

            send_email(
                EMAIL, "[Appointment] FOUND NEW APPOINTMENT IN %d!!!!!!!!!!" %
                scheduled_month)
            input("Scheduled month: %d. Continue?" % scheduled_month)

    time.sleep(5)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        raise
    except:
        send_email(EMAIL, "[Appointment] Stopped unexpected ")
        input("Stopped unexpected1: %d. Continue?")
        raise
