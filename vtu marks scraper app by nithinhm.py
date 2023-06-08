from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import *
from PIL import Image
from io import BytesIO
import pytesseract
import pandas as pd
from bs4 import BeautifulSoup
import time


pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pixel_range = [(i, i, i) for i in range(102, 130)]

def get_captcha_from_image(target_image):
    image_data = BytesIO(target_image)

    image = Image.open(image_data)
    width, height = image.size

    image = image.convert("RGB")

    white_image = Image.new("RGB", (width, height), "white")

    for x in range(width):
        for y in range(height):
            pixel = image.getpixel((x, y))

            if pixel in pixel_range:
                white_image.putpixel((x, y), pixel)

    return pytesseract.image_to_string(white_image).strip()


skipped_usns = []
data_dict = {}


service = ChromeService(executable_path='/chromedriver.exe')
options = Options()
options.unhandled_prompt_behavior = 'ignore'


art = '''
 _    __________  __   __  ___           __           _____                                    ___               
| |  / /_  __/ / / /  /  |/  /___ ______/ /_______   / ___/______________ _____  ___  _____   /   |  ____  ____  
| | / / / / / / / /  / /|_/ / __ `/ ___/ //_/ ___/   \__ \/ ___/ ___/ __ `/ __ \/ _ \/ ___/  / /| | / __ \/ __ \\
| |/ / / / / /_/ /  / /  / / /_/ / /  / ,< (__  )   ___/ / /__/ /  / /_/ / /_/ /  __/ /     / ___ |/ /_/ / /_/ / 
|___/ /_/  \____/  /_/  /_/\__,_/_/  /_/|_/____/   /____/\___/_/   \__,_/ .___/\___/_/     /_/  |_/ .___/ .___/  
                                                                       /_/                       /_/   /_/       
'''

print(art)
print("\nWelcome to the VTU marks scraping app.\n\nApp developed by\nProf. Nithin H M\nAssistant Professor\nDepartment of Physics\nAMC Engineering College\nBangalore - 560083.")
print("\nProcedure:\n1. Fill the details as mentioned below and press enter.\n2. Sit back and relax.\n3. You can keep an eye on this console to check for status updates or errors that occur.")

print('\nExample USN: 1AM22CS010')

while True:
    try:
        coll_code = input('\nEnter your college code as in USN (Ex: 1AM): ').upper().strip()
    except:
        print('Error! Enter a valid 3-character code.')
    else:
        if len(coll_code) != 3 or (not isinstance(int(coll_code[0]), int)) or (not isinstance(coll_code[1:], str)):
            print('Error! Enter a valid 3-character code.')
        else:
            break

while True:
    try:
        batch = int(input('\nEnter batch year as in USN (Ex: 22 or 21 or 19): ').strip())
    except:
        print('Error! Enter a valid 2-digit number.')
    else:
        batch = str(batch)
        if len(batch) != 2:
            print('Error! Enter a valid 2-digit number.')
        else:
            break

while True:
    try:
        branch = input('\nEnter branch as in USN (Ex: CS or MT or CI): ').upper().strip()
    except:
        print('Error! Enter a valid 2-character branch.')
    else:
        if len(branch) != 2:
            print('Error! Enter a valid 2-character branch.')
        else:
            break

while True:
    try:
        first_number = int(input("\nEnter the first number of this branch's USN (Ex: 1 or 25 or 140): ").strip())
    except:
        print('Error! Enter a valid 3-digit number.')
    else:
        if first_number > 999:
            print('Error! Enter a valid 3-digit number.')
        else:
            break

while True:
    try:
        final_number = int(input("\nEnter the last number of this branch's USN (Ex: 5 or 23 or 145): ").strip())
    except:
        print('Error! Enter a valid 3-digit number.')
    else:
        if final_number > 999:
            print('Error! Enter a valid 3-digit number.')
        else:
            break

while True:
    try:
        retry_delay = int(input('\nIf there are network issues, how many seconds do you wish to wait before retrying again?: ').strip())
    except:
        print('Error! Enter a valid number.')
    else:
        break

while True:
    try:
        max_retries = int(input('\nHow many times do you wish to retry before ending the session?: ').strip())
    except:
        print('Error! Enter a valid number.')
    else:
        break

while True:
    try:
        url = input('\nEnter the URL of the results login page (which contains fields for entering USN and CAPTCHA; it should start with https://): ').strip()
    except:
        print('Error! Enter a valid url.')
    else:
        if 'https://' not in url:
            print("Error! Please include https:// in the url.")
        else:
            break

driver = webdriver.Chrome(service=service, options=options)

driver.get(url)

k = first_number - 1

window_is_present = True

print("\n")

while k < final_number and window_is_present:
    
    k += 1
    this_retry = 0
    while this_retry < max_retries:
        try:
            usn = f'{coll_code}{batch}{branch}{k:03d}'

            driver.find_element(By.NAME, 'lns').send_keys(usn)

            captcha_image = driver.find_element(By.XPATH, '//*[@id="raj"]/div[2]/div[2]/img').screenshot_as_png
            captcha_text = get_captcha_from_image(captcha_image)

            if len(captcha_text) != 6:
                print("\nError! The length of CAPTCHA is invalid. Trying again.")
                driver.refresh()
                continue

            driver.find_element(By.NAME, 'captchacode').send_keys(captcha_text)

            driver.find_element(By.ID, 'submit').click()

            student_name = driver.find_element(By.XPATH,'//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td[2]').text.split(':')[1].strip()
            student_usn = driver.find_element(By.XPATH,'//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]').text.split(':')[1].strip()

            soup = BeautifulSoup(driver.page_source, 'lxml')

            marks_data = soup.find('div', class_='divTableBody')

            data_dict[f'{student_usn}+{student_name}'] = marks_data

            print(f'\nData successsfully collected for {usn}')

            time.sleep(1)

            driver.back()

            break

        except UnexpectedAlertPresentException:
            alert = WebDriverWait(driver, 1).until(EC.alert_is_present())
            alert_text = alert.text

            print(f'\nError for {usn} because {alert_text}.')

            if alert_text == 'University Seat Number is not available or Invalid..!':
                print('Moving to the next USN.')
                k += 1
                skipped_usns.append(usn)
                alert.accept()
            else:
                print('Trying again.')
                alert.accept()

        except NoSuchWindowException:
            print('\nError! Window closed prematurely. Data collected so far will be saved.')
            window_is_present = False
            break

        except:
            soup2 = BeautifulSoup(driver.page_source, 'lxml')
            occur = soup2.find_all('b', string='University Seat Number')
            if len(occur) > 0:
                print(f'\nThere was an error collecting data for {usn}. Trying again.')
                driver.back()
            else:
                print(f'\nError! Retrying after {retry_delay} seconds. Retry {this_retry+1} of {max_retries}')
                this_retry += 1
                time.sleep(retry_delay)
                driver.refresh()

    else:
        try:
            print(f'\nMaximum number of retries reached ({max_retries}). Data collected so far will be saved.')
        except NoSuchWindowException:
            print('\nError! Window closed prematurely. Data collected so far will be saved.')
            window_is_present = False
        finally:
            break

driver.quit()

if len(skipped_usns) > 0:
    print(f'\nThese USNs were skipped {skipped_usns}.')
else:
    print('\nNo USNs were skipped.')

list_of_student_dfs = []

for id, marks_data in data_dict.items():

    this_usn, this_name = tuple(id.split('+'))
    rows = marks_data.find_all('div', class_='divTableRow')

    data = []
    for row in rows:
        cells = row.find_all('div', class_='divTableCell')
        data.append([cell.text.strip() for cell in cells])
    
    df_temp = pd.DataFrame(data[1:], columns=data[0])

    subjects = [f'{name} ({code})' for name, code in zip(df_temp['Subject Name'], df_temp['Subject Code'])]
    headers = df_temp.columns[2:-1]

    ready_columns = [(name, header) for name in subjects for header in headers]

    student_df = pd.DataFrame([this_usn, this_name] + list(df_temp.iloc[:,2:-1].to_numpy().flatten()), index= [('USN',''), ('Student Name','')] + ready_columns).T
    student_df.columns = pd.MultiIndex.from_tuples(student_df.columns, names=['', ''])

    list_of_student_dfs.append(student_df)

final_df = pd.concat(list_of_student_dfs).reset_index(drop=True)

cols = list(final_df.columns)[2:]
cols.sort(key = lambda x: x[0].split('(')[-1][5:-1])
final_df = final_df[[('USN',''), ('Student Name','')] + cols]

final_df.index += 1

df2 = final_df.apply(pd.to_numeric, errors='ignore')

collected_usns = list(df2['USN'])

first_USN, last_USN = collected_usns[0], collected_usns[-1]

df2.to_excel(f'20{batch} {branch} {first_USN} to {last_USN} VTU results.xlsx')

print(f'\nData collected for USNs {branch} {first_USN} to {last_USN} and saved in an excel file.\n')

input("Press enter to close this app.")