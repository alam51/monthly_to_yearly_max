import glob
import os

import pandas as pd


class Summarizer:
    def __init__(self, folder_path: str):
        self.files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith('.xlsx')]
        self.df_list = []

        for file in self.files:
            try:
                df_raw = pd.read_excel(file, sheet_name='QF-LDC-33', skiprows=6,
                                       nrows=160, usecols='B:F', index_col=None, header=None)
                df = df_raw.dropna(how='any', axis=0)
                # df = df_raw
                df.columns = ['ss_name', 'MW', 'day', 'month', 'time']
                for i in df.index:
                    try:
                        dt_year_month = df.loc[i, "month"]
                        df.loc[i, 'dt'] = pd.to_datetime(
                            f'{dt_year_month.year}-{dt_year_month.month}-{int(df.loc[i, "day"])} '
                            f'{df.loc[i, "time"]}')
                    except:
                        print(f'failed on i = "{i}", ss_name = {df.loc[i, "ss_name"]}, month = {df.loc[i, "month"]}'
                              f'\nfile_name = {file}\n')
                # dt = pd.to_datetime(f'{df.month[0].year}-{df.month[0].month}-{int(df.day[0])} {df.time[0]}')
                df1 = df.loc[:, ['ss_name', 'MW', 'dt']]
                self.df_list.append(df1)
            except:
                print(f'problem in reading file {file}')

        self.combined_df = pd.concat(self.df_list, axis=0, ignore_index=True)

        # type casting
        self.combined_df['MW'] = self.combined_df['MW'].apply(lambda x: float(str(x)) if str(x).replace('.', '1').
                                                              isdigit() else 0.0)
        self.combined_df['dt'] = self.combined_df['dt'].apply(lambda x: str(x)[:19])
        # self.max_df = self.combined_df.groupby(['ss_name']).max()
        self.max_df = pd.DataFrame()
        max_mw_df_list = []
        ss_name_list = []
        for ss_name in self.combined_df['ss_name']:
            if ss_name not in ss_name_list:
                ss_name_list += [ss_name]
                ss_df = self.combined_df[self.combined_df['ss_name'] == ss_name]
                max_mw = ss_df['MW'].max()
                # ss_max_mw_df = ss_df[ss_df['MW'] == max_mw].iloc[0, :]  # choose first row when multiple max MW
                # present
                ss_max_mw_df = ss_df[ss_df['MW'] == max_mw]
                max_mw_df_list.append(ss_max_mw_df)

        self.max_df = pd.concat(max_mw_df_list).drop_duplicates(subset=['ss_name'], keep='first', ignore_index=True)
        b = 5
        # self.max_df.loc[len(ss_name_list), 'ss_name'] = self.combined_df


folder_path = r'G:/Other computers/My Computer/Data [F]/MOD/2021'
# folder_path = r'G:\Other computers\My Computer\Data [F]\MOD\test'
a = Summarizer(folder_path)
with pd.ExcelWriter(os.path.join(folder_path, 'collective_rep.xlsx')) as writer:
    a.max_df.to_excel(writer, sheet_name='summary')
    a.combined_df.to_excel(writer, sheet_name='combined')
b = 5
