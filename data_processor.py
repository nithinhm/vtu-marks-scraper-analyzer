import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import os
import re


matplotlib.use('agg')

class DataProcessor():
    def preprocess(self, soup_dict, is_reval, main_sem, usn_begin, root_folder):

        dict_of_sems_dfs = {f'{sem}':[] for sem in range(8,0,-1)}

        for id, soup in soup_dict.items():

            this_usn, this_name = id.split('+')
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

        batch_value = usn_begin[3:5]
        branch_value = usn_begin[-2:]
        
        folder_path = os.path.join(root_folder, f'{batch_value} batch semester {main_sem} {branch_value} VTU results')
        os.makedirs(folder_path, exist_ok=True)

        subs_for_creds = []

        for sem in dict_of_sems_dfs:
            df2 = pd.concat(dict_of_sems_dfs[sem]).reset_index(drop=True)

            # df2 = final_df.apply(pd.to_numeric, errors='ignore')
            s = [x for x in df2.columns.levels[0] if '(' in x]
            s.sort(key = lambda x: re.findall(r'\d+', x)[-1])
            subs_for_creds.extend(s)

            collected_usns = df2[('USN','')].to_list()

            first_USN, last_USN = collected_usns[0], collected_usns[-1]

            if sem == main_sem:
                self.main_first_USN = first_USN
                self.main_last_USN = last_USN

                if not is_reval:
                    file_path_csv = os.path.join(folder_path, f'Raw Regular semester {sem} {first_USN} to {last_USN}.csv')
                else:
                    file_path_csv = os.path.join(folder_path, f'Raw Reval Regular semester {sem} {first_USN} to {last_USN}.csv')
            else:
                if not is_reval:
                    file_path_csv = os.path.join(folder_path, f'Raw Arrear semester {sem} {first_USN} to {last_USN}.csv')
                else:
                    file_path_csv = os.path.join(folder_path, f'Raw Reval Arrear semester {sem} {first_USN} to {last_USN}.csv')
            
            df2.to_csv(file_path_csv)

        if not is_reval:
            credf = pd.Series(index=subs_for_creds, name='credits')
            file_path_credit_csv = os.path.join(folder_path, f'Credit info {branch_value}.csv')
            credf.to_csv(file_path_credit_csv)
    
    def analyze_data(self, filepaths):
        self.no_data = None
        self.sgpa = None
        self.reval = None

        filepaths.sort(key = lambda x: x.split('/')[-1])

        regular_data = []
        arrear_data = []
        reval_regular_data = []
        reval_arrear_data = []
        credits_data = []

        for filepath in filepaths:
            if 'Credit' not in filepath:
                data = pd.read_csv(filepath, header=[0,1])
                if 'Raw Regular' in filepath:
                    regular_data.append(data)
                elif 'Raw Arrear' in filepath:
                    arrear_data.append(data)
                elif 'Raw Reval Regular' in filepath:
                    reval_regular_data.append(data)
                elif 'Raw Reval Arrear' in filepath:
                    reval_arrear_data.append(data)
            else:
                data = pd.read_csv(filepath)
                credits_data.append(data)
        
        if not regular_data:
            self.no_data = True
            return

        full_data = pd.concat(regular_data).reset_index(drop=True)
        full_data = full_data.drop_duplicates(subset=full_data.columns[1]).reset_index(drop=True)
        full_data.drop(full_data.columns[0], axis=1, inplace=True)
        full_data.rename(columns={name:'' for name in full_data.columns.levels[1] if 'level' in name}, inplace=True)

        cols = full_data.columns.to_list()[2:]
        cols.sort(key = lambda x: re.findall(r'\d+', x[0])[-1])
        full_data = full_data[[('USN',''), ('Student Name','')] + cols]

        if reval_regular_data:
            full_data.set_index(('USN',''), inplace=True)

            full_reval_regular_data = pd.concat(reval_regular_data).reset_index(drop=True)
            full_reval_regular_data = full_reval_regular_data.drop_duplicates(subset=full_reval_regular_data.columns[1]).reset_index(drop=True)
            full_reval_regular_data.drop(full_reval_regular_data.columns[0], axis=1, inplace=True)
            full_reval_regular_data.rename(columns={name:'' for name in full_reval_regular_data.columns.levels[1] if 'level' in name}, inplace=True)

            reval_cols = full_reval_regular_data.columns.to_list()[2:]
            reval_cols.sort(key = lambda x:re.findall(r'\d+', x[0])[-1])
            full_reval_regular_data = full_reval_regular_data[[('USN',''), ('Student Name','')] + reval_cols]

            full_reval_regular_data.set_index(('USN',''), inplace=True)

            sub_cols = list(set([x[0] for x in full_reval_regular_data.columns if '(' in x[0]]).intersection(set([x[0] for x in full_data.columns if '(' in x[0]])))

            for index, row in full_reval_regular_data.iterrows():
                for subject in sub_cols:
                    if not pd.isna(row[(subject, 'Final Marks')]):
                        full_data.loc[index, (subject, 'External Marks')] = row[(subject, 'Final Marks')]
                        full_data.loc[index, (subject, 'Result')] = row[(subject, 'Final Result')]
                        full_data.loc[index, (subject, 'Total')] = int(full_data.loc[index, (subject, 'Internal Marks')]) + int(full_data.loc[index, (subject, 'External Marks')])

            full_data.reset_index(inplace=True)

            subs = [x for x in full_reval_regular_data.columns.levels[0] if '(' in x]
            subs.sort(key = lambda x: re.findall(r'\d+', x)[-1])
            want = ['Old Result', 'Final Result']

            rev_df = full_reval_regular_data[[(sub, wan) for sub in subs for wan in want]]

            rev_columns = ['Failed and Applied for Reval', 'Changed to Pass After Reval']
            rev_report = pd.DataFrame(index=subs, columns=rev_columns, dtype=int).fillna(0)

            for subject in subs:
                for index, row in rev_df[[subject]].iterrows():
                    if row[(subject, 'Old Result')] == 'F':
                        rev_report.loc[subject, rev_columns[0]] += 1
                        if row[(subject, 'Final Result')] == 'P':
                            rev_report.loc[subject, rev_columns[1]] += 1

            rev_report['Conversion %'] = rev_report.apply(lambda x: round(x.iloc[1]/x.iloc[0]*100, 2), axis=1)
            self.rev_report = rev_report
            self.reval = True

        full_data.index += 1

        if (full_data == 'P *').any().any():
            full_data.replace('P *', 'P', inplace=True)

        if arrear_data:
            # full_arrear_data = pd.concat(arrear_data).reset_index(drop=True)
            for x in arrear_data:
                x.drop(x.columns[0], axis=1, inplace=True)
                x.rename(columns={name:'' for name in x.columns.levels[1] if 'level' in name}, inplace=True)
                s = x.columns.to_list()[2:]
                s.sort(key = lambda x: re.findall(r'\d+', x[0])[-1])
                x = x[[('USN',''), ('Student Name','')] + s]

            if len(arrear_data) > 1:
                full_arrear_data = pd.merge(arrear_data[-1], arrear_data[-2], how='outer')
                for i in range(2, len(arrear_data)):
                    full_arrear_data = pd.merge(full_arrear_data, arrear_data[-i-1], how='outer')
            else:
                full_arrear_data = arrear_data[0]

            full_arrear_data = full_arrear_data.drop_duplicates().reset_index(drop=True)

            # arrear_cols = full_arrear_data.columns.to_list()[2:]
            # arrear_cols.sort(key = lambda x: re.findall(r'\d+', x[0])[-1])
            # full_arrear_data = full_arrear_data[[('USN',''), ('Student Name','')] + arrear_cols]

            if reval_arrear_data:
                full_arrear_data.set_index(('USN',''), inplace=True)

                # full_reval_arrear_data = pd.concat(reval_arrear_data).reset_index(drop=True)
                for x in reval_arrear_data:
                    x.drop(x.columns[0], axis=1, inplace=True)
                    x.rename(columns={name:'' for name in x.columns.levels[1] if 'level' in name}, inplace=True)
                    s = x.columns.to_list()[2:]
                    s.sort(key = lambda x: re.findall(r'\d+', x[0])[-1])
                    x = x[[('USN',''), ('Student Name','')] + s]

                # full_reval_arrear_data = pd.concat(reval_arrear_data).reset_index(drop=True)
                if len(reval_arrear_data) > 1:
                    full_reval_arrear_data = pd.merge(reval_arrear_data[-1], reval_arrear_data[-2], how='outer')
                    for i in range(2, len(reval_arrear_data)):
                        full_reval_arrear_data = pd.merge(full_reval_arrear_data, reval_arrear_data[-i-1], how='outer')
                else:
                    full_reval_arrear_data = reval_arrear_data[0]

                full_reval_arrear_data = full_reval_arrear_data.drop_duplicates().reset_index(drop=True)

                # reval_arrear_cols = full_reval_arrear_data.columns.to_list()[2:]
                # reval_arrear_cols.sort(key = lambda x:re.findall(r'\d+', x[0])[-1])
                # full_reval_arrear_data = full_reval_arrear_data[[('USN',''), ('Student Name','')] + reval_arrear_cols]

                full_reval_arrear_data.set_index(('USN',''), inplace=True)

                sub_cols = list(set([x[0] for x in full_reval_arrear_data.columns if '(' in x[0]]).intersection(set([x[0] for x in full_arrear_data.columns if '(' in x[0]])))

                for index, row in full_reval_arrear_data.iterrows():
                    for subject in sub_cols:
                        if not pd.isna(row[(subject, 'Final Marks')]):
                            full_arrear_data.loc[index, (subject, 'External Marks')] = row[(subject, 'Final Marks')]
                            full_arrear_data.loc[index, (subject, 'Result')] = row[(subject, 'Final Result')]
                            full_arrear_data.loc[index, (subject, 'Total')] = int(full_arrear_data.loc[index, (subject, 'Internal Marks')]) + int(full_arrear_data.loc[index, (subject, 'External Marks')])

                full_arrear_data.reset_index(inplace=True)

        if credits_data:
            full_credits_data = pd.concat(credits_data).reset_index(drop=True)
            full_credits_data = full_credits_data.drop_duplicates(subset=full_credits_data.columns[0]).set_index(full_credits_data.columns[0])

            if (not full_credits_data.isna().any().iloc[0]) and (full_credits_data.dtypes.iloc[0] == 'int64'):

                def to_grade(score):
                    grade = None

                    if isinstance(score, str):
                        return score

                    if score >= 90:
                        grade = 10
                    elif score >= 80:
                        grade = 9
                    elif score >= 70:
                        grade = 8
                    elif score >= 60:
                        grade = 7
                    elif score >= 55:
                        grade = 6
                    elif score >= 50:
                        grade = 5
                    elif score >= 40:
                        grade = 4
                    else:
                        grade = 0
                    
                    return grade
                
                
                full_credits_data = full_credits_data.T

                new = pd.merge(full_data, full_arrear_data, how='outer') if arrear_data else full_data.copy()

                subjs = [x[0] for i, x in enumerate(new.columns) if '(' in x[0] and i%4==0]

                for sub in subjs:
                    new[(sub, 'Grade Point')] = pd.Series(dtype=int)
                    new[(sub, 'Max Credits')] = pd.Series(dtype=int)

                wants = ['Result', 'Total', 'Grade Point', 'Max Credits']

                colms = [(sub, wan) for sub in subjs for wan in wants]

                sgpa_report = new[[('USN',''), ('Student Name','')] + colms]
                
                sgpa_report[('SGPA', '')] = pd.Series(dtype=float)

                for i, row in sgpa_report.iterrows():
                    num = 0
                    den = 0
                    for sub in subjs:
                        if not pd.isna(row[(sub, 'Result')]):
                            c = full_credits_data.loc['credits', sub]
                            sgpa_report.loc[i, (sub, 'Max Credits')] = c                  
                            
                            if row[(sub, 'Result')] == 'P':
                                gp = to_grade(row[(sub, 'Total')])
                            else:
                                gp = 0
                            
                            sgpa_report.loc[i, (sub, 'Grade Point')] = gp

                            num += gp*c
                            den += c

                    sgpa_report.loc[i, ('SGPA', '')] = round(num/den, 2)
                
                sgpa_report = sgpa_report.convert_dtypes()
                if arrear_data:
                    sgpa_report.index += 1
                self.sgpa = True
                self.sgpa_report = sgpa_report

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

        self.f_usn, self.l_usn = full_data[('USN', '')].iloc[0], full_data[('USN', '')].iloc[-1]

        self.batch_value = self.f_usn[3:5]
        self.branch_value = self.l_usn[5:7]
       
        self.full_data = full_data


    def save_data(self, folder_path):
        if not self.reval:
            file_path_image = os.path.join(folder_path, f'Subject-wise Pass Percentages of {self.batch_value} batch {self.branch_value} branch students.jpg')
            file_path_excel = os.path.join(folder_path, f'{self.f_usn} to {self.l_usn} VTU results and analysis.xlsx')
        else:
            file_path_image = os.path.join(folder_path, f'After Reval Subject-wise Pass Percentages of {self.batch_value} batch {self.branch_value} branch students.jpg')
            file_path_excel = os.path.join(folder_path, f'After Reval {self.f_usn} to {self.l_usn} VTU results and analysis.xlsx')

        self.fig.savefig(file_path_image)

        # to avoid getting blank row after column names
        def save_double_column_df(df, xl_writer, startrow=0, **kwargs):
            df.drop(df.index).to_excel(xl_writer, startrow=startrow, **kwargs)
            df.to_excel(xl_writer, startrow=startrow + 1, header=False, **kwargs)

        sheet_data = {'Student-wise results':self.full_data, 'Stats of students':self.stats_df, 'Subject-wise results':self.result_df.fillna(0)}

        if self.reval:
            sheet_data['Revaluation Report'] = self.rev_report
        
        if self.sgpa:
            sheet_data['SGPA Report'] = self.sgpa_report

        with pd.ExcelWriter(file_path_excel) as writer:
            for name, data in sheet_data.items():
                save_double_column_df(data, writer, sheet_name=name)