import os
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.keys import Keys

from config import CHROME_DRIVER_PATH, AUTH_DATA, XPATH, USER_AGENT, URL, HOME_DIR, FOLDER_FILE_NAMES, SEARCH_PARAMS


def initialization(folder_name):
    options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': os.path.join(HOME_DIR, f'Desktop\\{folder_name}\\')}
    options.add_experimental_option("prefs", prefs)
    options.add_argument(f'user_agent={USER_AGENT}')
    service = Service(executable_path=CHROME_DRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    return driver


def get_report(driver, url, dir_name):
    try:
        driver.get(url)

        # skip alert
        WebDriverWait(driver, 10).until(ec.alert_is_present())
        driver.switch_to.alert.dismiss()

        # auth
        auth_button = driver.find_element(By.ID, 'loginWithoutEds')
        auth_button.click()

        login_input = driver.find_element(By.XPATH, XPATH.get('login'))
        login_input.clear()
        login_input.send_keys(AUTH_DATA.get('login'))

        password_input = driver.find_element(By.XPATH, XPATH.get('password'))
        password_input.clear()
        password_input.send_keys(AUTH_DATA.get('password'))
        time.sleep(0.5)
        password_input.send_keys(Keys.ENTER)

        # find report
        driver.find_element(By.LINK_TEXT, 'На главную').click()
        driver.find_element(By.LINK_TEXT, 'Личный кабинет').click()
        driver.find_element(By.XPATH, XPATH.get('reports')).click()
        driver.find_element(By.XPATH, XPATH.get('report on the execution of the plan')).click()

        change_year = driver.find_element(By.XPATH, XPATH.get('change_year'))
        change_year.clear()
        change_year.send_keys(SEARCH_PARAMS.get('year'))
        time.sleep(0.5)

        if dir_name == 'ДПЗ':
            selector = driver.find_element(By.XPATH, SEARCH_PARAMS.get('plane_kind').get('long-time'))
            selector.click()
        time.sleep(0.5)

        find_button = driver.find_element(By.XPATH, XPATH.get('find_button'))
        find_button.click()

        # WebDriverWait(driver, 60).until(
        #     ec.presence_of_element_located((By.XPATH, XPATH.get('pagination')))
        # )
        # time.sleep(0.5)

        WebDriverWait(driver, 60).until(
            ec.invisibility_of_element_located((By.XPATH, XPATH.get('progress_bar')))
        )
        time.sleep(0.5)

        # download report
        download_button = driver.find_element(By.XPATH, XPATH.get('download_button'))
        download_button.click()
        time.sleep(0.5)

        WebDriverWait(driver, 60).until(
            ec.element_to_be_clickable((By.XPATH, XPATH.get('download_button')))
        )
        time.sleep(5)
    except Exception as ex:
        print(ex)
    finally:
        driver.close()
        driver.quit()


def run():
    for i in FOLDER_FILE_NAMES.items():
        new_download_dir = os.path.join(HOME_DIR + '\\Desktop\\', i[0])
        old_filename = new_download_dir + '\\download.xls'
        new_filename = new_download_dir + f'\\{i[1]}_{datetime.now().strftime("%m.%d.%Y %H.%M.%S")}.xlsx'

        try:
            os.mkdir(new_download_dir)
        except FileExistsError:
            ...

        get_report(driver=initialization(i[0]), url=URL, dir_name=i[0])

        try:
            os.rename(old_filename, new_filename)
        except FileNotFoundError:
            print('Не удалось найти отчёт по заданным параметрам')


if __name__ == '__main__':
    run()
