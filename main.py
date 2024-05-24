from tkinter import ttk, scrolledtext, messagebox, Toplevel, BooleanVar
from tkinter import filedialog as fd
from selenium.common.exceptions import *
import re
import threading
import os
from gui import TemplateWindow
from connection import Connection
from data_processor import DataProcessor

mess = \
'''

    Welcome.

    DO NOT CLOSE THIS TERMINAL WHILE THE APP IS RUNNING!

    If you see a "DevTools listening on..." message, you can ignore it.

    If you see any "Traceback" exception/error messages in this terminal while using the app, please close everything and restart again.

    For assistance, contact Nithin H M through email: nithinmanju111@gmail.com
    or go through the GitHub repo https://github.com/nithinhm/vtu-marks-scraper-analyzer

'''

current_dir = os.path.dirname(os.path.abspath(__file__))

conn_support = Connection()

class MainFrame(TemplateWindow):
    def create_frame(self):
        self.title("Welcome | Select app")
        self.config(pady=20)

        self.frame = ttk.Frame(self)
        self.frame.pack()

        self.frame.welcome_label = ttk.Label(self.frame, font=('Segoe UI', 10), text='Welcome to the VTU Marks Scraper and Analyzer App')
        self.frame.welcome_label.grid(columnspan=2)

        self.frame.select_app_label = ttk.Label(self.frame, font=('Segoe UI', 10), text='Select the app to proceed:')
        self.frame.select_app_label.grid(row=1, columnspan=2, sticky='w', padx=20, pady=20)

        self.frame.main_buttons = ttk.Frame(self.frame)
        self.frame.main_buttons.grid(padx=20)

        self.frame.s_button = ttk.Button(self.frame.main_buttons, text='\n\n\nMarks Scraper\n\n\n', command=self.open_scraper, width=30)
        self.frame.s_button.grid(row=2, column=0)

        self.frame.a_button = ttk.Button(self.frame.main_buttons, text='\n\n\nMarks Analyzer\n\n\n', command=self.open_analyzer, width=30)
        self.frame.a_button.grid(row=2, column=1)


    def open_scraper(self):
        self.withdraw()
        self.scraper_window = ScraperFrame()
        self.scraper_window.grab_set()
        self.scraper_window.protocol("WM_DELETE_WINDOW", self.close_scraper)
        self.scraper_window.mainloop()

    def close_scraper(self):
        self.scraper_window.destroy()
        self.deiconify()


    def open_analyzer(self):
        self.withdraw()
        self.analyzer_window = AnalyzerFrame()
        self.analyzer_window.grab_set()
        self.analyzer_window.protocol("WM_DELETE_WINDOW", self.close_analyzer)
        self.analyzer_window.mainloop()

    def close_analyzer(self):
        self.analyzer_window.destroy()
        self.deiconify()        


class ScraperFrame(TemplateWindow):
    scraper_art = r'''
     _    __________  __   __  ___           __           _____                                    ___               
    | |  / /_  __/ / / /  /  |/  /___ ______/ /_______   / ___/______________ _____  ___  _____   /   |  ____  ____  
    | | / / / / / / / /  / /|_/ / __ `/ ___/ //_/ ___/   \__ \/ ___/ ___/ __ `/ __ \/ _ \/ ___/  / /| | / __ \/ __ \ 
    | |/ / / / / /_/ /  / /  / / /_/ / /  / ,< (__  )   ___/ / /__/ /  / /_/ / /_/ /  __/ /     / ___ |/ /_/ / /_/ / 
    |___/ /_/  \____/  /_/  /_/\__,_/_/  /_/|_/____/   /____/\___/_/   \__,_/ .___/\___/_/     /_/  |_/ .___/ .___/  
                                                                           /_/                       /_/   /_/       
    '''

    default_entry_values = ['1AM23CS001', '1AM23CS100', '1', '5', '5', 'https://results.vtu.ac.in/DJcbcs24/index.php']

    to_abort = None

    def create_frame(self):
        self.update_idletasks()
        self.title('VTU Marks Scraper App')
        self.config(padx=40, pady=10)

        self.frame = ttk.Frame(self)
        self.frame.pack()

        self.frame.form = ttk.Frame(self.frame)
        self.frame.form.grid()

        ttk.Label(self.frame.form, text=self.scraper_art, font=('Courier', 7)).grid(column=0, row=0, columnspan=2)

        texts = ['First USN: ', 'Last USN: ', 'Current semester: ', 'Retry time delay (seconds): ', 'Number of retries: ', 'Result page URL: ']

        for i, text in enumerate(texts):
            ttk.Label(self.frame.form, font=('Segoe UI',10), text=text).grid(column=0, row=i+1, sticky='w')
        
        for i in range(len(texts)):
            ttk.Entry(self.frame.form, width=40, font=('Segoe UI',10)).grid(column=1, row=1+i)
        
        self.frame.form.entries = [e for e in self.frame.form.winfo_children() if isinstance(e, ttk.Entry)]
        
        self.frame.form.is_reval = BooleanVar(self.frame.form)
        self.frame.form.reval = ttk.Label(self.frame.form, font=('Segoe UI',10), text='Will collect revaluation marks')

        self.frame.buttons = ttk.Frame(self.frame)
        self.frame.buttons.grid(pady=10)

        self.frame.buttons.verify = ttk.Button(self.frame.buttons, text='Verify', command=self.verify_for_error, width=15)
        self.frame.buttons.verify.grid(column=0, row=0, padx=10, pady=10, columnspan=3)

        self.frame.buttons.reset = ttk.Button(self.frame.buttons, text='Reset', command=self.reset_default_entry_values, width=15, state='disabled')
        self.frame.buttons.reset.grid(column=0, row=1, padx=10, pady=10)

        self.frame.buttons.collect = ttk.Button(self.frame.buttons, text='Collect', command=self.on_collect_click, width=15, state='disabled')
        self.frame.buttons.collect.grid(column=1, row=1, padx=10, pady=10)

        self.frame.buttons.abort = ttk.Button(self.frame.buttons, text='Abort', command=self.abort_app, width=15, state='disabled')
        self.frame.buttons.abort.grid(column=2, row=1, padx=10, pady=10)

        self.frame.buttons.buttons = {name:b for name, b in zip(['v', 'r', 'c', 'a'], self.frame.buttons.winfo_children())}

        self.frame.loading = ttk.Frame(self.frame)

        self.frame.loading.label = ttk.Label(self.frame.loading, font=('Segoe UI',11), text='Loading...')
        self.frame.loading.label.grid(columnspan=2)

        self.frame.status = ttk.Frame(self.frame)

        self.frame.status.progress = ttk.Progressbar(self.frame.status, orient='horizontal', length=550, mode='determinate')
        self.frame.status.progress.grid(columnspan=2)

        self.frame.status.textbox = scrolledtext.ScrolledText(self.frame.status, state='disabled', wrap='word', width=80, height=8, font=('Segoe UI', 10))
        self.frame.status.textbox.grid(columnspan=2, pady=20)

        self.toggle_entries('disabled')
        self.toggle_buttons(v=0, r=0, c=0, a=0)

        threading.Thread(target=self.start_check_connection).start()


    def toggle_entries(self, state):
        for e in self.frame.form.entries:
            e.config(state=state)
    
    def toggle_buttons(self, **kwargs):
        '''
        **kwargs will look like v=1, r=0, c=1, a=1

        where 0 means to disable and 1 means to enable

        and the letters represent the different buttons

        v - verify, r - reset, c - collect, a - abort
        '''
        for name, i in kwargs.items():
            if i==1:
                self.frame.buttons.buttons[name].config(state='normal')
            elif i==0:
                self.frame.buttons.buttons[name].config(state='disabled')
    
    def reset_default_entry_values(self):
        self.frame.form.reval.grid_remove()
        self.toggle_entries('normal')
        for i, e in enumerate(self.frame.form.entries):
            e.delete(0, 'end')
            e.insert(0, self.default_entry_values[i])
        self.toggle_buttons(v=1, r=0, c=0)

    def open_top_wait(self):
        self.top_wait = Toplevel(self)
        self.top_wait.update_idletasks()
        self.top_wait.title('Just a moment')
        ttk.Label(self.top_wait, text='Checking your internet connection...').pack(padx=20, pady=5)
        self.top_wait.lift()
        x = self.winfo_x() + self.winfo_width()//2 - self.top_wait.winfo_width()//2
        y = self.winfo_y() + self.winfo_height()//2 - self.top_wait.winfo_height()//2
        self.top_wait.geometry(f"+{x}+{y}")
        self.top_wait.protocol("WM_DELETE_WINDOW", self.disable_event)

    def disable_event(self):
        pass

    def start_check_connection(self):
        self.open_top_wait()

        try:
            conn_support.check_internet()

        except:
            conn_support.driver.quit()
            self.top_wait.destroy()
            if messagebox.askretrycancel(title='Internet Status', message='Internet is not connected.\n\nWould you like to retry?\n\nPresss "Cancel" to quit the app.'):
                self.start_check_connection()

        else:
            conn_support.driver.quit()
            self.top_wait.destroy()
            messagebox.showinfo(title='Internet Status', message='Internet is connected.\n\nPresss OK to continue.')
            self.reset_default_entry_values()

    def verify_for_error(self):
        error_list = []

        pattern_usn = r'^\d{1}[A-Za-z]{2}\d{2}[A-Za-z]{2}\d{3}$'

        first_usn, last_usn, main_sem, delay_value, retries_value, url_value = tuple([e.get().strip() for e in self.frame.form.entries])

        if not (re.match(pattern_usn, first_usn) and re.match(pattern_usn, last_usn)):
            error_list.append('USN Error! Enter valid USN(s)')
        elif first_usn[:-3] != last_usn[:-3]:
            error_list.append('USN Error! First and last USNs need to match (except for last three characters)')
        elif int(first_usn[-3:]) > int(last_usn[-3:]):
            error_list.append('USN Error! First USN has to be less than last USN')

        first_usn = first_usn.upper()
        last_usn = last_usn.upper()
        
        if not main_sem.isnumeric():
            error_list.append('Semester Error! Enter a number for the semester.')
        elif int(main_sem) not in range(1,9):
            error_list.append('Semester Error! Enter a valid number for the semester.')

        if not delay_value.isnumeric():
            error_list.append('Delay Error! Enter a valid number for the delay.')
        if not retries_value.isnumeric():
            error_list.append('Retries Error! Enter a valid number for the number of retries.')
        if not re.match('https://results.vtu.ac.in/[a-zA-Z0-9]+/index.php', url_value):
            error_list.append('URL Error! Enter correct URL. (starts with https://results.vtu.ac.in/ and ends with index.php)')

        if not error_list:
            message = 'Here are the values that you entered:\n\n'+'\n'.join([f'First USN: {first_usn}', f'Last USN: {last_usn}', f'Current semester: {main_sem}', f'Delay: {delay_value}', f'Retries: {retries_value}', f'URL: {url_value}'])+'\n\nWould you like to proceed?\n\nIf you wish to make some changes, press No.'
            answer = messagebox.askyesno(title='Confirmation', message=message)

            if answer:
                self.usn_begin = first_usn[:-3]
                self.first_num = int(first_usn[-3:])
                self.last_num = int(last_usn[-3:])
                self.main_sem = main_sem
                self.delay_value = int(delay_value)
                self.retries_value = int(retries_value)
                self.url_value = url_value

                if 'RV' in self.url_value:
                    self.frame.form.is_reval.set(True)
                    self.frame.form.reval.grid(columnspan=2, pady=10)
                else:
                    self.frame.form.is_reval.set(False)
                    self.frame.form.reval.grid_remove()

                self.toggle_entries('disabled')
                self.toggle_buttons(v=0, r=1, c=1)
                self.frame.status.grid_remove()
        
        else:
            message = 'Kindly correct the following error(s) before proceeding:\n\n'+'\n'.join(error_list)
            answer = messagebox.showerror(title='ERROR', message=message)
    
    def try_again(self):
        self.toggle_entries('normal')
        self.toggle_buttons(v=1, a=0)
        self.frame.status.progress.config(value=0)
        
    def status_update(self, statement):
        self.frame.status.textbox.config(state='normal')
        self.frame.status.textbox.insert('end', statement+'\n')
        self.frame.status.textbox.see('end')
        self.frame.status.textbox.config(state='disabled')

    def abort_app(self):
        to_abort = messagebox.askyesno(title='ABORT', message='Are you sure you want to abort the data collection process?\n\nData collected so far (if any) will be saved.')
        if to_abort:
            self.to_abort = to_abort
            self.status_update('Aborting...\n')
        
    def on_collect_click(self):
        self.see_process = messagebox.askyesno(title='Look at Process', message='Do you wish to look at the automated data collection process?')

        self.toggle_buttons(v=0, r=0, c=0, a=0)

        threading.Thread(target=self.start_scraping).start()

    def start_scraping(self):
        self.frame.loading.grid(pady=10)

        self.skipped_usns = []

        self.soup_dict = {}

        try:
            conn_support.connect(self.url_value, self.see_process)
        except:
            messagebox.showerror(title='Connection Error', message='There was an unknown error.\n\nPlease try again after some time.')
            self.frame.loading.grid_remove()
            self.try_again()
            return
        else:
            self.frame.loading.grid_remove()
            self.frame.status.grid(pady=5)
            self.frame.status.textbox.config(state='normal')
            self.frame.status.textbox.delete('1.0', 'end')
            self.frame.status.textbox.config(state='disabled')

        k = self.first_num - 1

        cool = 0

        while k < self.last_num:
            k += 1
            this_retry = 0

            try:
                self.toggle_buttons(a=1)

                while this_retry < self.retries_value and not self.to_abort:
                    try:
                        usn = self.usn_begin+f'{k:03d}'
                        
                        conn_support.enter_usn(usn)

                        this_retry = 0

                        captcha_text, error = conn_support.get_captcha()

                        if error:
                            self.status_update('Error! Tesseract not configured. Retry after configuring.')
                            self.to_abort = True
                            break

                        if len(captcha_text) != 6:
                            self.status_update('Invalid captcha. Trying again.\n')
                            conn_support.driver.refresh()
                            continue
                        else:
                            conn_support.captcha_submit(captcha_text)
                        
                        self.soup_dict = conn_support.get_info(self.soup_dict)

                        self.status_update(f'Data successsfully collected for {usn}\n')
                        cool = 0

                        self.frame.status.progress.config(value=(k - self.first_num + 1)/(self.last_num - self.first_num + 1)*100)

                        conn_support.sleep(2)
                        conn_support.driver.back()

                        break    
                                    
                    except UnexpectedAlertPresentException:
                        alert = conn_support.check_alert()
                        alert_text = alert.text

                        self.status_update(f'Error for {usn} because {alert_text}')

                        if alert_text == 'University Seat Number is not available or Invalid..!' or alert_text == 'You have not applied for reval or reval results are awaited !!!':
                            cool += 1
                            self.status_update('Moving to the next USN.\n')
                            self.frame.status.progress.config(value=(k - self.first_num + 1)/(self.last_num - self.first_num + 1)*100)
                            self.skipped_usns.append(usn)
                            alert.accept()
                            break
                        elif alert_text == '':
                            self.status_update('\nUnfortunately, this IP address has been blocked due to excessive requests in a short period of time.\nYou will not be able to access this particular link using this IP address anymore.\nUse an internet connection with a different IP address or set a proxy server in your system settings.\n')
                            self.to_abort = True
                            alert.accept()
                            break
                        else:
                            cool += 1
                            self.status_update('Trying again.\n')
                            alert.accept()                    

                    except:
                        occur = conn_support.stuck_page()
                        
                        if len(occur) > 0:
                            self.status_update(f'There was an error collecting data for {usn}. Trying again.\n')
                            conn_support.driver.back()
                        else:
                            self.status_update(f'Error! Retrying after {self.delay_value} seconds. Retry {this_retry+1} of {self.retries_value}\n')
                            this_retry += 1
                            conn_support.sleep(self.delay_value)
                            conn_support.driver.refresh()

                    if cool > 5:
                        self.status_update('Waiting for a bit to avoid IP block.\n')
                        cool = 0
                        conn_support.sleep(2)

                else:
                    if self.to_abort:
                        self.status_update('ABORTED\n')
                        self.to_abort = False
                        break
                    else:
                        messagebox.showerror(title='Connection Error', message=f'Maximum number of retries reached ({self.retries_value}).\nData collected so far (if any) will be saved.\n\nPlease try again after some time.')
                        break
                        
            except:
                messagebox.showerror(title='Unknown Error', message='There was an unknown error.\nData collected so far (if any) will be saved.\n\nPlease try again after some time.')
                break
        
        conn_support.driver.quit()

        if bool(self.soup_dict):
            if self.skipped_usns:
                self.status_update(f'These USNs were skipped {self.skipped_usns}.\n')
            else:
                self.status_update('No USNs were skipped.\n')

            messagebox.showinfo(title='Select folder', message='Select a folder in which to save the downloaded files.')
            fp_given = False
            while not fp_given:
                root_folder = fd.askdirectory(title='Select folder')
                if root_folder:
                    fp_given = True

            dataproc = DataProcessor()
            dataproc.preprocess(self.soup_dict, self.frame.form.is_reval.get(), self.main_sem, self.usn_begin, root_folder)

            messagebox.showinfo(title='Success', message=f'Data collected for USNs {dataproc.main_first_USN} to {dataproc.main_last_USN} and saved in a csv file.\n\nYou can continue to collect data of other students.\n\nOr you can close the scraper window to return to the main interface.')

        else:
            messagebox.showinfo(title='No Data', message='No data was collected.')
        
        self.status_update('Start again or close this window.\n')
        self.try_again()


class AnalyzerFrame(TemplateWindow):
    analyzer_art = r'''
     __   _______ _   _   __  __          _           _             _                     _             
     \ \ / /_   _| | | | |  \/  |__ _ _ _| |__ ___   /_\  _ _  __ _| |_  _ ______ _ _    /_\  _ __ _ __ 
      \ V /  | | | |_| | | |\/| / _` | '_| / /(_-<  / _ \| ' \/ _` | | || |_ / -_) '_|  / _ \| '_ \ '_ \
       \_/   |_|  \___/  |_|  |_\__,_|_| |_\_\/__/ /_/ \_\_||_\__,_|_|\_, /__\___|_|   /_/ \_\ .__/ .__/
                                                                      |__/                   |_|  |_|   
    '''

    filepaths = []

    def create_frame(self):
        self.title('VTU Marks Analyzer App')
        self.config(padx=10, pady=20)

        self.frame = ttk.Frame(self)
        self.frame.pack()

        self.frame.selection = ttk.Frame(self.frame)
        self.frame.selection.grid()

        ttk.Label(self.frame.selection, text=self.analyzer_art, font=('Courier', 7)).grid(columnspan=2, padx=10)
        ttk.Label(self.frame.selection, font=('Segoe UI', 10), text='Select the csv file(s) from the appropriate folder:').grid(column=0, row=1, padx=10, pady=10)

        browse_button = ttk.Button(self.frame.selection, text='Browse', command=self.browse_files)
        browse_button.grid(column=1, row=1, padx=10, sticky='w')

        self.frame.buttons = ttk.Frame(self.frame)
        self.frame.buttons.grid(pady=10)

        self.frame.buttons.clear = ttk.Button(self.frame.buttons, text='Clear', command=self.clear_box, state='disabled')
        self.frame.buttons.clear.grid(column=1, row=0, padx=10)

        self.frame.buttons.analyze = ttk.Button(self.frame.buttons, text='Analyze', command=self.analyze_and_save, state='disabled')
        self.frame.buttons.analyze.grid(column=3, row=0, padx=10)

        self.frame.output = ttk.Frame(self.frame)

        self.frame.output.filebox = scrolledtext.ScrolledText(self.frame.output, state='disabled', wrap='word', width=80, height=8, font=('Segoe UI', 10))
        self.frame.output.filebox.grid(padx=20, pady=20)

    def browse_files(self):
        if self.filepaths:
            file_list = list(fd.askopenfilenames(title='Select file(s)', filetypes=[("csv files", "*.csv")]))
        else:
            file_list = list(fd.askopenfilenames(title='Select file(s)', initialdir=current_dir, filetypes=[("csv files", "*.csv")]))
        
        if file_list:
            self.frame.output.grid()
            self.frame.output.filebox.config(state='normal')
            for path in file_list:
                if path not in self.filepaths:
                    self.frame.output.filebox.insert('end', path+'\n\n')
                    self.filepaths.append(path)
            self.frame.output.filebox.config(state='disabled')

            self.frame.buttons.clear.config(state='normal')
            self.frame.buttons.analyze.config(state='normal')

    def clear_box(self):
        self.frame.output.filebox.config(state='normal')
        self.frame.output.filebox.delete('1.0', 'end')
        self.frame.output.filebox.config(state='disabled')
        self.frame.output.grid_remove()
        self.frame.buttons.clear.config(state='disabled')
        self.frame.buttons.analyze.config(state='disabled')
        self.filepaths.clear()

    def analyze_and_save(self):
        processor = DataProcessor()

        processor.analyze_data(self.filepaths)

        if processor.no_data:
            messagebox.showerror(title='Select file', message='Select atleast one result file other than Reval file.')
            return

        messagebox.showinfo(title='Select folder', message='Select a folder in which to save the excel file.')
        folder_path = fd.askdirectory(title='Select folder')

        if not folder_path:
            return

        processor.save_data(folder_path)

        messagebox.showinfo(title='Analysis Complete.', message=f'The data has been analyzed.\n\nYou can continue to analyze data of other students.\n\nOr you can close the analyzer window to return to the main interface.')
        self.clear_box()


if __name__=='__main__':
    print(mess)

    window = MainFrame()
    window.mainloop()