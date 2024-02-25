import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import os
import re


matplotlib.use('agg')

class DataProcessor():
    def preprocess(self, soup_dict, is_reval, main_sem):

        dict_of_sems_dfs = {f'{sem}':[] for sem in range(8,0,-1)}

        for id, soup in soup_dict.items():

            this_usn, this_name = tuple(id.split('+'))
            sems_divs = soup.find_all('div', style="text-align:center;padding:5px;")
            sems_num = [x.text.split(':')[-1].strip() for x in sems_divs]
            sems_data = [sem_div.find_next_sibling('div') for sem_div in sems_divs]

            for sem, marks_data in zip(sems_num, sems_data):
                rows = marks_data.find_all('div', class_='divTableRow')

                data = []
                for row in rows:
                    cells = row.find_all('div', class_='divTableCell')
                    data.append([cell.text.strip() for cell in cells])
                
                df_temp = pd.DataFrame(data[1:], columns=data[0])

                subjects = [f'{name} ({code})' for name, code in zip(df_temp['Subject Name'], df_temp['Subject Code'])]

                if not is_reval:
                    headers = df_temp.columns[2:-1]
                else:
                    headers = df_temp.columns[2:]

                ready_columns = [(name, header) for name in subjects for header in headers]

                if not is_reval:
                    student_sem_df = pd.DataFrame([this_usn, this_name] + list(df_temp.iloc[:,2:-1].to_numpy().flatten()), index= [('USN',''), ('Student Name','')]+ready_columns).T
                else:
                    student_sem_df = pd.DataFrame([this_usn, this_name] + list(df_temp.iloc[:,2:].to_numpy().flatten()), index= [('USN',''), ('Student Name','')]+ready_columns).T

                student_sem_df.columns = pd.MultiIndex.from_tuples(student_sem_df.columns, names=['', ''])

                dict_of_sems_dfs[sem].append(student_sem_df)

        dict_of_sems_dfs = {key:value for key, value in dict_of_sems_dfs.items() if value}

        for sem in dict_of_sems_dfs:
            df2 = pd.concat(dict_of_sems_dfs[sem]).reset_index(drop=True)

            # df2 = final_df.apply(pd.to_numeric, errors='ignore')

            collected_usns = list(df2[('USN','')])

            self.first_USN, self.last_USN = collected_usns[0], collected_usns[-1]

            batch_value = self.first_USN[3:5]
            branch_value = self.first_USN[5:7]
            
            if sem == main_sem:
                self.main_first_USN = self.first_USN
                self.main_last_USN = self.last_USN

                if not is_reval:
                    folder_path = f'Regular 20{batch_value} {branch_value} VTU results'
                    os.makedirs(folder_path, exist_ok=True)
                    file_path_csv = os.path.join(folder_path, f'Raw Regular {branch_value} semester {sem} {self.first_USN} to {self.last_USN}.csv')
                else:
                    folder_path = f'Reval Regular 20{batch_value} {branch_value} VTU results'
                    os.makedirs(folder_path, exist_ok=True)
                    file_path_csv = os.path.join(folder_path, f'Raw Reval Regular {branch_value} semester {sem} {self.first_USN} to {self.last_USN}.csv')
            
            else:
                if not is_reval:
                    folder_path = f'Arrear 20{batch_value} {branch_value} VTU results'
                    os.makedirs(folder_path, exist_ok=True)
                    file_path_csv = os.path.join(folder_path, f'Arrear {branch_value} semester {sem} {self.first_USN} to {self.last_USN}.csv')
                else:
                    folder_path = f'Reval Arrear 20{batch_value} {branch_value} VTU results'
                    os.makedirs(folder_path, exist_ok=True)
                    file_path_csv = os.path.join(folder_path, f'Reval Arrear {branch_value} semester {sem} {self.first_USN} to {self.last_USN}.csv')

            df2.to_csv(file_path_csv)
    
    def analyze_data(self, filepaths):
        self.no_data = None
        self.reval = None

        filepaths.sort(key = lambda x: x.split('/')[-1])

        list_of_data = []
        reval_data = []

        for filepath in filepaths:
            data = pd.read_csv(filepath, header=[0,1])
            if 'Reval' in filepath:
                reval_data.append(data)
            else:
                list_of_data.append(data)

        if not list_of_data:
            self.no_data = True
            return

        full_data = pd.concat(list_of_data).reset_index(drop=True)
        full_data = full_data.drop_duplicates(subset=full_data.columns[1]).reset_index(drop=True)
        full_data.drop(full_data.columns[0], axis=1, inplace=True)
        full_data.rename(columns={name:'' for name in full_data.columns.levels[1] if 'level' in name}, inplace=True)

        cols = list(full_data.columns)[2:]
        cols.sort(key = lambda x: re.findall(r'\d+', x[0])[-1])
        full_data = full_data[[('USN',''), ('Student Name','')] + cols]

        USNs = list(full_data['USN'])
        self.first_USN, self.last_USN = USNs[0], USNs[-1]
        self.branch_value = self.first_USN[5:7]
        self.batch_value = self.first_USN[3:5]

        # full_data = full_data.apply(pd.to_numeric, errors='ignore')

        if reval_data:
            full_data.set_index(('USN',''), inplace=True)

            full_reval_data = pd.concat(reval_data).reset_index(drop=True)
            full_reval_data = full_reval_data.drop_duplicates(subset=full_reval_data.columns[1]).reset_index(drop=True)
            full_reval_data.drop(full_reval_data.columns[0], axis=1, inplace=True)
            full_reval_data.rename(columns={name:'' for name in full_reval_data.columns.levels[1] if 'level' in name}, inplace=True)

            reval_cols = list(full_reval_data.columns)[2:]
            reval_cols.sort(key = lambda x:re.findall(r'\d+', x[0])[-1])
            full_reval_data = full_reval_data[[('USN',''), ('Student Name','')] + reval_cols]

            # full_reval_data = full_reval_data.apply(pd.to_numeric, errors='ignore')
            full_reval_data.set_index(('USN',''), inplace=True)

            sub_cols = list(set([x[0] for x in full_reval_data.columns if '(' in x[0]]).intersection(set([x[0] for x in full_data.columns if '(' in x[0]])))

            # Iterate over the rows of the second dataframe
            for index, row in full_reval_data.iterrows():
                # Iterate over the columns (subjects) of the second dataframe
                for subject in sub_cols:
                    # Check if the student has applied for revaluation in this subject
                    if not pd.isna(row[(subject, 'Final Marks')]):
                        # Update the marks in the first dataframe
                        full_data.loc[index, (subject, 'External Marks')] = row[(subject, 'Final Marks')]
                        # Update the result in the first dataframe
                        full_data.loc[index, (subject, 'Result')] = row[(subject, 'Final Result')]
                        # update the total marks in the first dataframe
                        full_data.loc[index, (subject, 'Total')] = int(full_data.loc[index, (subject, 'Internal Marks')]) + int(full_data.loc[index, (subject, 'External Marks')])

            full_data.reset_index(inplace=True)

            self.reval = True

        full_data.index += 1

        if (full_data == 'P *').any().any():
            full_data.replace('P *', 'P', inplace=True)

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
        students_eligible = len(temp_df[temp_df['A'] != len(cols)]['F']) #students who have not written any exams are ineligible here
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

        result_df['Eligible and Appeared'] = result_df.fillna(0).apply(lambda x: x['Passed'] + x['Failed'], axis=1)

        pass_percentage_column = result_df.fillna(0).apply(lambda x: round(x['Passed']/x['Eligible and Appeared']*100, 2), axis=1)
        result_df['Subject Pass Percentage'] = pass_percentage_column

        self.result_df = result_df

        labels = [x.split('(')[-1][:-1] for x in result_df.index]

        x = np.arange(len(labels))

        self.fig, ax = plt.subplots(figsize=(25,10))
        ax.bar(x, pass_percentage_column)

        ax.set_xlabel('Subject Code', fontsize='x-large')
        ax.set_ylabel('Pass Percentage', fontsize='x-large')
        if not self.reval:
            ax.set_title('Subject-wise Pass Percentages', fontsize='xx-large')
        else:
            ax.set_title('After Reval Subject-wise Pass Percentages', fontsize='xx-large')
        ax.set_xticks(x, labels, fontsize='x-large')
        # ax.set_yticks(ax.get_yticks())

        for i,v in enumerate(result_df.index):
            ax.text(i, pass_percentage_column.iloc[i]+2, f"{result_df.loc[v, 'Subject Pass Percentage']}%", ha='center', fontsize='x-large')

        self.fig.tight_layout()

        full_data['Overall Total'] = overall_column

        self.full_data = full_data


    def save_data(self, folder_path):
        if not self.reval:
            file_path_image = os.path.join(folder_path, f'Subject-wise Pass Percentages of 20{self.batch_value} {self.branch_value} branch students.jpg')
            file_path_excel = os.path.join(folder_path, f'20{self.batch_value} {self.branch_value} {self.first_USN} to {self.last_USN} VTU results and analysis.xlsx')
        else:
            file_path_image = os.path.join(folder_path, f'After Reval Subject-wise Pass Percentages of 20{self.batch_value} {self.branch_value} branch students.jpg')
            file_path_excel = os.path.join(folder_path, f'After Reval 20{self.batch_value} {self.branch_value} {self.first_USN} to {self.last_USN} VTU results and analysis.xlsx')

        self.fig.savefig(file_path_image)

        # to avoid getting blank row after column names
        def save_double_column_df(df, xl_writer, startrow=0, **kwargs):
            df.drop(df.index).to_excel(xl_writer, startrow=startrow, **kwargs)
            df.to_excel(xl_writer, startrow=startrow + 1, header=False, **kwargs)

        sheet_data = {'Student-wise results':self.full_data, 'Stats of students':self.stats_df, 'Subject-wise results':self.result_df.fillna(0)}

        with pd.ExcelWriter(file_path_excel) as writer:
            for name, data in sheet_data.items():
                save_double_column_df(data, writer, sheet_name=name)