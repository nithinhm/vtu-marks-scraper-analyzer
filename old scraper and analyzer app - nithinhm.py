from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import scrolledtext
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
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import os
import threading
import webbrowser

matplotlib.use('agg')
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

art = r'''
 _    __________  __   __  ___           __           _____                                    ___               
| |  / /_  __/ / / /  /  |/  /___ ______/ /_______   / ___/______________ _____  ___  _____   /   |  ____  ____  
| | / / / / / / / /  / /|_/ / __ `/ ___/ //_/ ___/   \__ \/ ___/ ___/ __ `/ __ \/ _ \/ ___/  / /| | / __ \/ __ \ 
| |/ / / / / /_/ /  / /  / / /_/ / /  / ,< (__  )   ___/ / /__/ /  / /_/ / /_/ /  __/ /     / ___ |/ /_/ / /_/ / 
|___/ /_/  \____/  /_/  /_/\__,_/_/  /_/|_/____/   /____/\___/_/   \__,_/ .___/\___/_/     /_/  |_/ .___/ .___/  
                                                                       /_/                       /_/   /_/       
'''

service = ChromeService(executable_path='/chromedriver.exe')

default_values = ['1AM', '22', 'CS', '1', '100', '1', '5', '5', 'https://results.vtu.ac.in/JFEcbcs23/index.php']
to_abort = False

def check_connection_thread():
    toggle_entries('disabled')
    for button in buttons.winfo_children():
        button.config(state='disabled')

    threading.Thread(target=check_connection).start()

def check_connection():
    wait = Toplevel()
    wait.title('Just a moment')
    wait_label = Label(wait, text='Checking your internet connection...', width=30)
    wait_label.pack(padx=20, pady=5)

    try:
        options = Options()
        options.add_argument("--headless=new")
        driver = webdriver.Chrome(service=service, options=options)
        driver.get('https://www.feynmanlectures.caltech.edu/III_toc.html')
        wait.destroy()
        driver.quit()
        messagebox.showinfo(title='Internet Status', message='Internet is connected.\n\nPresss OK to continue.')
        reset_entries()

    except:
        check_again = messagebox.askretrycancel(title='Internet Status', message='Internet is not connected.\n\nWould you like to retry?\n\nPresss Cancel to quit the app.')
        if not check_again:
            driver.quit()
            window.quit()
        else:
            wait.destroy()
            driver.quit()
            check_connection()

def get_entry_widgets():
    return [e for e in form.winfo_children() if isinstance(e, ttk.Entry)]

def set_default_values():
    for i, entry_widget in enumerate(get_entry_widgets()):
        entry_widget.delete(0, 'end')
        entry_widget.insert(0, default_values[i])

def get_values():
    entry_widgets = get_entry_widgets()
    values = [e.get().upper().strip() for e in entry_widgets[:-1]] + [entry_widgets[-1].get().strip()]
    return tuple(values)

def toggle_entries(state):
    for entry_widget in get_entry_widgets():
        entry_widget.config(state=state)

def check_error():
    error_list = []

    code_value, batch_value, branch_value, firstnum_value, lastnum_value, sem_value, delay_value, retries_value, url_value = get_values()

    if len(code_value) != 3 or (not isinstance(int(code_value[0]), int)) or (not isinstance(code_value[1:], str)):
        error_list.append('Code Error! Enter a valid 3-character college code.')
    if len(batch_value) != 2 or (not isinstance(int(batch_value), int)):
        error_list.append('Batch Error! Enter a valid 2-digit batch number.')
    if len(branch_value) != 2:
        error_list.append('Branch Error! Enter a valid 2-character branch.')
    if int(firstnum_value) < 1 or int(firstnum_value) > 999:
        error_list.append('First USN Error! Enter a valid 3-digit first USN number.')
    if int(lastnum_value) < 1 or int(lastnum_value) > 999:
        error_list.append('Last USN Error! Enter a valid 3-digit last USN number.')
    if int(sem_value) < 1 or int(sem_value) > 8:
        error_list.append('Semester Error! Enter a valid semester number.')
    if not isinstance(int(delay_value), int):
        error_list.append('Delay Error! Enter a valid number for the delay.')
    if not isinstance(int(retries_value), int):
        error_list.append('Retries Error! Enter a valid number for the number of retries.')
    if 'https://' not in url_value:
        error_list.append('URL Error! Please include https:// in the URL.')

    if len(error_list) == 0:
        message = 'Here are the values that you entered:\n\n'+'\n'.join([f'Code: {code_value}', f'Batch: {batch_value}', f'Branch: {branch_value}', f'First USN: {firstnum_value}', f'Last USN: {lastnum_value}', f'Semester: {sem_value}', f'Delay: {delay_value}', f'Retries: {retries_value}', f'URL: {url_value}'])+'\n\nWould you like to proceed?\n\nIf you wish to make some changes, press No.'
        answer = messagebox.askyesno(title='Confirmation', message=message)

        if answer:
            toggle_entries('disabled')
            verify_button.config(state='disabled')
            reset_button.config(state='normal')
            collect_button.config(state='normal')
            status_progress.grid_remove()
    
    else:
        message = 'Kindly correct the following error(s) before proceeding:\n\n'+'\n'.join(error_list)
        answer = messagebox.showerror(title='ERROR', message=message)

def reset_entries():
    toggle_entries('normal')
    set_default_values()
    verify_button.config(state='normal')
    reset_button.config(state='disabled')
    collect_button.config(state='disabled')

def try_again():
    toggle_entries('normal')
    verify_button.config(state='normal')
    abort_button.config(state='disabled')
    progress.config(value=0)

def status_update(statement):
    statusbox.config(state='normal')
    statusbox.insert(END, statement+'\n')
    statusbox.see(END)
    statusbox.config(state='disabled')

def abort_app():
    global to_abort
    to_abort = messagebox.askyesno(title='ABORT', message='Are you sure you want to abort the data collection process?\n\nData collected so far (if any) will be saved.')
    if to_abort:
        status_update('Aborting...\n')


def start_app():
    global to_abort
    to_abort = False
    loading_label.grid(column=0, row=11, columnspan=2, pady=10)
    skipped_usns = []
    data_dict = {}

    code_value, batch_value, branch_value, firstnum_value, lastnum_value, sem_value, delay_value, retries_value, url_value = get_values()

    try:
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(url_value)
    except:
        messagebox.showerror(title='Connection Error', message='There was an unknown error.\n\nPlease try again after some time.')
        loading_label.grid_remove()
        try_again()
        return
    else:
        loading_label.grid_remove()
        status_progress.grid(column=0, row=11, pady=5)
        progress.grid(column=0, row=0, columnspan=2)
        statusbox.grid(column=0, row=1, columnspan=2, pady=20)
        statusbox.config(state='normal')
        statusbox.delete('1.0', END)
        statusbox.config(state='disabled')

    k = int(firstnum_value) - 1

    while k < int(lastnum_value):
        k += 1
        this_retry = 0

        try:
            abort_button.config(state='normal')

            while this_retry < int(retries_value) and not to_abort:
                try:
                    usn = f'{code_value}{batch_value}{branch_value}{k:03d}'

                    driver.find_element(By.NAME, 'lns').send_keys(usn)
                    this_retry = 0

                    captcha_image = driver.find_element(By.XPATH, '//*[@id="raj"]/div[2]/div[2]/img').screenshot_as_png
                    captcha_text = get_captcha_from_image(captcha_image)

                    if len(captcha_text) != 6:
                        status_update('Tried reading CAPTCHA. The length was invalid. Trying again.\n')
                        driver.refresh()
                        continue

                    driver.find_element(By.NAME, 'captchacode').send_keys(captcha_text)

                    driver.find_element(By.ID, 'submit').click()

                    student_name = driver.find_element(By.XPATH,'//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td[2]').text.split(':')[1].strip()
                    student_usn = driver.find_element(By.XPATH,'//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]').text.split(':')[1].strip()

                    soup = BeautifulSoup(driver.page_source, 'lxml')

                    marks_data = soup.find('div', class_='divTableBody')

                    data_dict[f'{student_usn}+{student_name}'] = marks_data

                    status_update(f'Data successsfully collected for {usn}\n')

                    progress.config(value=(k - int(firstnum_value) + 1)/(int(lastnum_value) - int(firstnum_value) + 1)*100)

                    time.sleep(2)

                    driver.back()

                    break

                except UnexpectedAlertPresentException:
                    alert = WebDriverWait(driver, 1).until(EC.alert_is_present())
                    alert_text = alert.text

                    status_update(f'Error for {usn} because {alert_text}')

                    if alert_text == 'University Seat Number is not available or Invalid..!':
                        status_update('Moving to the next USN.\n')
                        progress.config(value=(k - int(firstnum_value) + 1)/(int(lastnum_value) - int(firstnum_value) + 1)*100)
                        skipped_usns.append(usn)
                        alert.accept()
                        break
                    else:
                        status_update('Trying again.\n')
                        alert.accept()

                except:
                    soup2 = BeautifulSoup(driver.page_source, 'lxml')
                    occur = soup2.find_all('b', string='University Seat Number')
                    if len(occur) > 0:
                        status_update(f'There was an error collecting data for {usn}. Trying again.\n')
                        driver.back()
                    else:
                        status_update(f'Error! Retrying after {delay_value} seconds. Retry {this_retry+1} of {retries_value}\n')
                        this_retry += 1
                        time.sleep(int(delay_value))
                        driver.refresh()

            else:
                if to_abort:
                    status_update('ABORTED\n')
                    break
                else:
                    messagebox.showerror(title='Connection Error', message=f'Maximum number of retries reached ({retries_value}).\nData collected so far (if any) will be saved.\n\nPlease try again after some time.')
                    break
                    
        except:
            messagebox.showerror(title='Unknown Error', message='There was an unknown error.\nData collected so far (if any) will be saved.\n\nPlease try again after some time.')
            break
    
    driver.quit()
    try_again()

    if bool(data_dict):
        if len(skipped_usns) > 0:
            status_update(f'These USNs were skipped {skipped_usns}.\n')
        else:
            status_update('No USNs were skipped.\n')

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

        overall_column = df2[df2.iloc[:,4::4].columns].replace('-', 0).fillna(0).astype(int).sum(axis=1)
        temp_df = df2.iloc[:,5::4].apply(lambda x: x.value_counts(), axis=1).fillna(0).astype(int)

        result_cases = ['A', 'P', 'F', 'W', 'X']
        not_present1 = list(set(result_cases) - set(list(temp_df.columns)))

        if len(not_present1) > 0:
            for i in not_present1:
                temp_df[i] = 0

        students_passed_all = sum(temp_df['F'] == 0)
        students_failed = sum(temp_df['F'] > 0)
        students_failed_one = sum(temp_df['F'] == 1)
        students_eligible = len(temp_df[temp_df['A'] != len(cols)]['F'])
        overall_pass_percentage = round(students_passed_all/students_eligible*100, 2)

        stats_df = pd.Series({'Number of students passed in all subjects':students_passed_all, 'Number of students failed atleast 1 subject':students_failed, 'Number of students failed in only 1 subject':students_failed_one, 'Number of eligible students':students_eligible, 'Overall pass percentage':overall_pass_percentage}, name='')

        result_df = df2.iloc[:,5::4].apply(lambda x: x.value_counts(), axis=0)
        result_df.columns = [x[0] for x in result_df.columns]
        result_df = result_df.T
        not_present2 = list(set(result_cases) - set(list(result_df.columns)))

        if len(not_present2) > 0:
            for i in not_present2:
                result_df[i] = 0
        
        result_df = result_df.rename(columns={'A': 'Absent', 'P':'Passed', 'F':'Failed', 'X':'Not Eligible', 'W':'Withheld'})

        pass_percentage_column = result_df.fillna(0).apply(lambda x: round(x['Passed']/(x['Passed'] + x['Failed'] + x['Not Eligible'])*100, 2), axis=1)
        result_df['Subject Pass Percentage'] = pass_percentage_column

        labels = [x.split('(')[-1][:-1] for x in result_df.index]

        x = np.arange(len(labels))

        fig, ax = plt.subplots(figsize=(15,7))
        ax.bar(x, pass_percentage_column)

        ax.set_xlabel('Subject Code', fontsize='x-large')
        ax.set_ylabel('Pass Percentage', fontsize='x-large')
        ax.set_title('Subject-wise Pass Percentages', fontsize='xx-large')
        ax.set_xticks(x, labels, fontsize='x-large')
        ax.set_yticks(ax.get_yticks(), fontsize='large')

        for i,v in enumerate(result_df.index):
            ax.text(i, pass_percentage_column[i]+2, result_df.loc[v, 'Subject Pass Percentage'], ha='center', fontsize='x-large')

        fig.tight_layout()

        folder_path = f'20{batch_value} {branch_value} semester {sem_value} VTU results'
        os.makedirs(folder_path, exist_ok=True)
        file_path_image = os.path.join(folder_path, f'Subject-wise Pass Percentages of {branch_value} branch students.jpg')

        fig.savefig(file_path_image)

        df2['Overall_Total'] = overall_column

        file_path_excel = os.path.join(folder_path, f'20{batch_value} {branch_value} semester {sem_value} {first_USN} to {last_USN} VTU results.xlsx')

        with pd.ExcelWriter(file_path_excel) as writer:
            df2.to_excel(writer, sheet_name='Student-wise results')
            stats_df.to_excel(writer, sheet_name='Stats of students')
            result_df.fillna(0).to_excel(writer, sheet_name='Subject-wise results')

        continue_app = messagebox.askyesno(title='Continue?', message=f'Data collected for USNs {first_USN} to {last_USN} and saved in an excel file.\nResult analysis also saved.\n\nPress YES to continue to collect data of other students.\n\nPress No to quit the app.')
        if not continue_app:
            window.quit()
        else:
            status_update('Start Again.\n')
            try_again()

    else:
        messagebox.showinfo(title='No Data', message='No data was collected.')
        status_update('Start Again.\n')
        try_again()

def start_thread():
    global options
    options = Options()
    options.unhandled_prompt_behavior = 'ignore'
    if not messagebox.askyesno(title='Look at Process', message='Do you wish to look at the automated data collection process?'):
        options.add_argument("--headless=new")

    for button in buttons.winfo_children():
        button.config(state='disabled')

    threading.Thread(target=start_app).start()


window = Tk()
window.title('VTU Marks Scraper App')
window.geometry('650x800')
window.config(padx=40, pady=10)

form = Frame(window)
form.grid()

Label(form, text=art, font=('Courier', 7)).grid(column=0, row=0, columnspan=2)

Label(form, font=('Segoe UI',10), text='Region+College code (Ex: 1AM or 2KE etc): ').grid(column=0, row=1, sticky='w')
code_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
code_entry.grid(column=1, row=1)

Label(form, font=('Segoe UI',10), text='Batch (Ex: 22 or 21 etc): ').grid(column=0, row=2, sticky='w')
batch_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
batch_entry.grid(column=1, row=2)

Label(form, font=('Segoe UI',10), text='Branch (Ex: CS or MT or CI etc): ').grid(column=0, row=3, sticky='w')
branch_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
branch_entry.grid(column=1, row=3)

Label(form, font=('Segoe UI',10), text='First Number in USN (Ex: 1 or 25 etc): ').grid(column=0, row=4, sticky='w')
firstnum_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
firstnum_entry.grid(column=1, row=4)

Label(form, font=('Segoe UI',10), text='Last Number in USN (Ex: 5 or 50 etc): ').grid(column=0, row=5, sticky='w')
lastnum_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
lastnum_entry.grid(column=1, row=5)

Label(form, font=('Segoe UI',10), text='Semester (Ex: 1 or 2 etc): ').grid(column=0, row=6, sticky='w')
sem_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
sem_entry.grid(column=1, row=6)

Label(form, font=('Segoe UI',10), text='Retry time delay (seconds): ').grid(column=0, row=7, sticky='w')
delay_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
delay_entry.grid(column=1, row=7)

Label(form, font=('Segoe UI',10), text='Number of retries: ').grid(column=0, row=8, sticky='w')
retries_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
retries_entry.grid(column=1, row=8)

Label(form, font=('Segoe UI',10), text='Result page URL (starts with https://): ').grid(column=0, row=9, sticky='w')
url_entry = ttk.Entry(form, width=40, font=('Segoe UI',10))
url_entry.grid(column=1, row=9)

buttons = Frame(window)
buttons.grid(column=0, row=10, pady=10)

verify_button = ttk.Button(buttons, text='Verify', command=check_error, width=15)
verify_button.grid(column=0, row=0, padx=10, pady=10, columnspan=3)

collect_button = ttk.Button(buttons, text='Collect', command=start_thread, width=15, state='disabled')
collect_button.grid(column=1, row=1, padx=10, pady=10)

reset_button = ttk.Button(buttons, text='Reset', command=reset_entries, width=15, state='disabled')
reset_button.grid(column=0, row=1, padx=10, pady=10)

abort_button = ttk.Button(buttons, text='Abort', width=15, command=abort_app, state='disabled')
abort_button.grid(column=2, row=1, padx=10, pady=10)

loading_label = Label(window, font=('Segoe UI',11), text='Loading...')

status_progress = Frame(window)

progress = ttk.Progressbar(status_progress, orient='horizontal', length=550, mode='determinate')

statusbox = scrolledtext.ScrolledText(status_progress, state='disabled', wrap=WORD, width=80, height=8, font=('Segoe UI', 10))

my_credit = Frame(window)
my_credit.grid(column=0, row=12, columnspan=2)

Label(my_credit, font=('Segoe UI', 8), text='App developed by\nProf. Nithin H M\nAssistant Professor\nDepartment of Physics\nAMC Engineering College\nBangalore - 560083').pack()

gitlink = Label(my_credit, text="My GitHub", fg="blue", cursor="hand2")
gitlink.pack()
gitlink.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/nithinhm"))

dinlink = Label(my_credit, text="My LinkedIn", fg="blue", cursor="hand2")
dinlink.pack()
dinlink.bind("<Button-1>", lambda e: webbrowser.open_new("https://linkedin.com/in/nithinhm13"))

check_connection_thread()

window.mainloop()