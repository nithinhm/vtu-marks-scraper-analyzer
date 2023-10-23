import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import os


matplotlib.use('agg')

class DataProcessor():
    def preprocess(self, data_dict):
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

            student_df = pd.DataFrame([this_usn, this_name] + list(df_temp.iloc[:,2:-1].to_numpy().flatten()), index= [('USN',''), ('Student Name','')]+ready_columns).T
            
            student_df.columns = pd.MultiIndex.from_tuples(student_df.columns, names=['', ''])

            list_of_student_dfs.append(student_df)

        final_df = pd.concat(list_of_student_dfs).reset_index(drop=True)

        df2 = final_df.apply(pd.to_numeric, errors='ignore')

        collected_usns = list(df2[('USN','')])

        self.first_USN, self.last_USN = collected_usns[0], collected_usns[-1]

        batch_value = self.first_USN[3:5]
        branch_value = self.first_USN[5:7]
        
        folder_path = f'20{batch_value} {branch_value} VTU results'
        os.makedirs(folder_path, exist_ok=True)
        file_path_csv = os.path.join(folder_path, f'{branch_value} {self.first_USN} to {self.last_USN}.csv')

        df2.to_csv(file_path_csv)
    
    def analyze_data(self, filepaths):
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
        self.first_USN, self.last_USN = USNs[0], USNs[-1]
        self.branch_value = self.first_USN[5:7]
        self.batch_value = self.first_USN[3:5]

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

        self.stats_df = pd.Series({'Number of students passed in all subjects':students_passed_all, 'Number of students failed atleast 1 subject':students_failed, 'Number of students failed in only 1 subject':students_failed_one, 'Number of eligible students':students_eligible, 'Overall pass percentage':overall_pass_percentage}, name='')

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

        self.result_df = result_df

        labels = [x.split('(')[-1][:-1] for x in result_df.index]

        x = np.arange(len(labels))

        self.fig, ax = plt.subplots(figsize=(15,7))
        ax.bar(x, pass_percentage_column)

        ax.set_xlabel('Subject Code', fontsize='x-large')
        ax.set_ylabel('Pass Percentage', fontsize='x-large')
        ax.set_title('Subject-wise Pass Percentages', fontsize='xx-large')
        ax.set_xticks(x, labels, fontsize='x-large')
        ax.set_yticks(ax.get_yticks(), fontsize='large')

        for i,v in enumerate(result_df.index):
            ax.text(i, pass_percentage_column[i]+2, f"{result_df.loc[v, 'Subject Pass Percentage']}%", ha='center', fontsize='x-large')

        self.fig.tight_layout()

        full_data['Overall_Total'] = overall_column

        self.full_data = full_data


    def save_data(self, folder_path):
        file_path_image = os.path.join(folder_path, f'Subject-wise Pass Percentages of {self.branch_value} branch students.jpg')

        self.fig.savefig(file_path_image)

        file_path_excel = os.path.join(folder_path, f'20{self.batch_value} {self.branch_value} {self.first_USN} to {self.last_USN} VTU results.xlsx')

        with pd.ExcelWriter(file_path_excel) as writer:
            self.full_data.to_excel(writer, sheet_name='Student-wise results')
            self.stats_df.to_excel(writer, sheet_name='Stats of students')
            self.result_df.fillna(0).to_excel(writer, sheet_name='Subject-wise results')

