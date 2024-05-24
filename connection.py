from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from bs4 import BeautifulSoup
import time
from captcha_handler import CaptchaHandler


class Connection:

    def check_internet(self):
        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.unhandled_prompt_behavior = 'ignore'
        options.add_argument('--headless=new')
        self.driver = webdriver.Chrome(options=options)
        self.driver.get('https://www.feynmanlectures.caltech.edu/III_toc.html')
    
    def connect(self, url, mode):
        options = Options()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.unhandled_prompt_behavior = 'ignore'
        if not mode:
            options.add_argument('--headless=new')
        self.driver = webdriver.Chrome(options=options)
        self.driver.get(url)

    def enter_usn(self, usn):
        self.driver.find_element(By.NAME, 'lns').send_keys(usn)

    def get_captcha(self):
        captcha_image = self.driver.find_element(By.XPATH, '//*[@id="raj"]/div[2]/div[2]/img').screenshot_as_png

        text = None
        error = False

        try:
            text = CaptchaHandler().get_captcha_from_image(captcha_image)
        except:
            error = True

        return text, error
    
    def captcha_submit(self, captcha):
        self.driver.find_element(By.NAME, 'captchacode').send_keys(captcha)
        self.driver.find_element(By.ID, 'submit').click()

    def get_info(self, soup_dict):
        student_name = self.driver.find_element(By.XPATH,'//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td[2]').text.split(':')[1].strip()
        student_usn = self.driver.find_element(By.XPATH,'//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]').text.split(':')[1].strip()

        soup = BeautifulSoup(self.driver.page_source, 'lxml')

        soup_dict[f'{student_usn}+{student_name}'] = soup

        return soup_dict

    def sleep(self, secs):
        time.sleep(secs)

    def check_alert(self):
        return WebDriverWait(self.driver, 1).until(EC.alert_is_present())
    
    def stuck_page(self):
        soup = BeautifulSoup(self.driver.page_source, 'lxml')
        return soup.find_all('b', string='University Seat Number')