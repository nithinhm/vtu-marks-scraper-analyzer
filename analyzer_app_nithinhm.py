from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter import ttk
from tkinter import scrolledtext
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import os
import webbrowser

matplotlib.use('agg')
current_dir = os.path.dirname(os.path.abspath(__file__))

full_data = None
filepaths = []

art = r'''
 __   _______ _   _   __  __          _           _             _                     _             
 \ \ / /_   _| | | | |  \/  |__ _ _ _| |__ ___   /_\  _ _  __ _| |_  _ ______ _ _    /_\  _ __ _ __ 
  \ V /  | | | |_| | | |\/| / _` | '_| / /(_-<  / _ \| ' \/ _` | | || |_ / -_) '_|  / _ \| '_ \ '_ \
   \_/   |_|  \___/  |_|  |_\__,_|_| |_\_\/__/ /_/ \_\_||_\__,_|_|\_, /__\___|_|   /_/ \_\ .__/ .__/
                                                                  |__/                   |_|  |_|   
'''

def clear_box():
    global filepaths
    filebox.config(state='normal')
    filebox.delete('1.0', END)
    filebox.config(state='disabled')
    output.grid_remove()
    clear_button.config(state='disabled')
    analyze_button.config(state='disabled')
    filepaths.clear()

def browse_files():
    global filepaths

    if filepaths:
        file_list = list(fd.askopenfilenames(title='Select file(s)', filetypes=[("csv files", "*.csv")]))
    else:
        file_list = list(fd.askopenfilenames(title='Select file(s)', initialdir=current_dir, filetypes=[("csv files", "*.csv")]))
    
    if file_list:
        output.grid(row=2, column=0)
        filebox.grid(padx=20, pady=20)
        filebox.config(state='normal')

        for path in file_list:
            if path not in filepaths:
                filebox.insert(END, path+'\n\n')
                filepaths.append(path)
        filebox.config(state='disabled')

        clear_button.config(state='normal')
        analyze_button.config(state='normal')

def analyze_data():
    global filepaths

    filepaths.sort(key = lambda x: x.split('/')[-1])

    list_of_data = [pd.read_csv(filepath, header=[0,1]) for filepath in filepaths]

    full_data = pd.concat(list_of_data).reset_index(drop=True)
    full_data = full_data.drop_duplicates(subset=full_data.columns[1]).reset_index(drop=True)
    full_data.drop(full_data.columns[0], axis=1, inplace=True)
    full_data.rename(columns={name:'' for name in full_data.columns.levels[1] if 'level' in name}, inplace=True)

    cols = list(full_data.columns)[2:]
    cols.sort(key = lambda x: x[0].split('(')[-1][5:-1])
    full_data = full_data[[('USN',''), ('Student Name','')] + cols]

    full_data.index += 1

    full_data = full_data.apply(pd.to_numeric, errors='ignore')

    cols = list(full_data.columns)[2:]

    USNs = list(full_data['USN'])
    first_USN, last_USN = USNs[0], USNs[-1]
    branch_value = first_USN[5:7]
    batch_value = first_USN[3:5]

    overall_column = full_data[full_data.iloc[:,4::4].columns].replace('-', 0).fillna(0).astype(int).sum(axis=1)
    temp_df = full_data.iloc[:,5::4].apply(lambda x: x.value_counts(), axis=1).fillna(0).astype(int)

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

    result_df = full_data.iloc[:,5::4].apply(lambda x: x.value_counts(), axis=0)
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
        ax.text(i, pass_percentage_column[i]+2, f"{result_df.loc[v, 'Subject Pass Percentage']}%", ha='center', fontsize='x-large')

    fig.tight_layout()

    folder_path = None

    while not folder_path:
        messagebox.showinfo(title='Select a folder', message='Select a folder in which to save the excel file.')
        folder_path = fd.askdirectory(title='Select folder')

    file_path_image = os.path.join(folder_path, f'Subject-wise Pass Percentages of {branch_value} branch students.jpg')

    fig.savefig(file_path_image)

    full_data['Overall_Total'] = overall_column

    file_path_excel = os.path.join(folder_path, f'20{batch_value} {branch_value} {first_USN} to {last_USN} VTU results.xlsx')

    with pd.ExcelWriter(file_path_excel) as writer:
        full_data.to_excel(writer, sheet_name='Student-wise results')
        stats_df.to_excel(writer, sheet_name='Stats of students')
        result_df.fillna(0).to_excel(writer, sheet_name='Subject-wise results')

    if messagebox.askyesno(title='Analysis Complete.', message=f'The data has been analyzed.\n\nYou can find the excel file and the image file in the selected folder:\n\n{folder_path}\n\nPress "Yes" to analyze data for other branch students.\n\nPress "No" to quit the app.'):
        clear_box()
    else:
        window.quit()

window = Tk()
window.title('VTU Marks Analyzer App')
window.config(padx=10, pady=20)

selection = Frame(window)
selection.grid()

Label(selection, text=art, font=('Courier', 7)).grid(column=0, row=0, columnspan=2)
Label(selection, font=('Segoe UI', 10), text='Select the csv file(s) from the appropriate folder:').grid(column=0, row=1, padx=10, pady=10)

browse_button = ttk.Button(selection, text='Browse', command=browse_files)
browse_button.grid(column=1, row=1, padx=10, sticky='w')

buttons = Frame(window)
buttons.grid(pady=10)

clear_button = ttk.Button(buttons, text='Clear', command=clear_box, state='disabled')
clear_button.grid(column=1, row=0, padx=10)

analyze_button = ttk.Button(buttons, text='Analyze', command=analyze_data, state='disabled')
analyze_button.grid(column=3, row=0, padx=10)

output = Frame(window)

filebox = scrolledtext.ScrolledText(output, state='disabled', wrap=WORD, width=80, height=8, font=('Segoe UI', 10))

my_credit = Frame(window)
my_credit.grid(row=3, column=0)

Label(my_credit, font=('Segoe UI', 8), text='App developed by\nProf. Nithin H M\nAssistant Professor\nDepartment of Physics\nAMC Engineering College\nBangalore - 560083').pack()

gitlink = Label(my_credit, text="My GitHub", fg="blue", cursor="hand2")
gitlink.pack()
gitlink.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/nithinhm"))

dinlink = Label(my_credit, text="My LinkedIn", fg="blue", cursor="hand2")
dinlink.pack()
dinlink.bind("<Button-1>", lambda e: webbrowser.open_new("https://linkedin.com/in/nithinhm13"))

window.mainloop()