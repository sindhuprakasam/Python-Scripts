"""
This module is to generate MBSP daily report to track requirements progress.

Author : Sindhu Prakasam
"""
import os
import datetime
import numpy as np
from win32com.client import Dispatch
import textwrap
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd
from datetime import timedelta
import warnings
warnings.filterwarnings("ignore")


class GenerateReport():
    """
    Main class to generate report
    """
    def __init__(self):
        """
        Set initial values
        """
        self.day1_flag = 0
        self.sep_flag = 0
        self.flag2 = 1
        self.ids_day1 = 0
        self.up_report = 0
        self.proj_spe = 0
        self.data = pd.read_csv(r"path_to_input_csv_file.csv")
        d = self.data['Report Run Date'][0]
        self.date = datetime.datetime.strptime(d[:10], '%d.%m.%Y')
        if self.day1_flag != 1:
            self.prev_man_vw = pd.read_pickle('prev_man_vw.pkl')

            if self.up_report != 1:
                self.delta_week_vw = pd.read_pickle('delta_week_vw.pkl')
                self.all_data_vw = pd.read_pickle('all_data_vw.pkl')
            else:
                self.delta_week = pd.read_pickle('delta_week_vw.pkl')
                self.all_data = pd.read_pickle('all_data_vw.pkl')
                self.prev_man = pd.read_pickle('prev_man_vw.pkl')
        else:
            self.delta_week_vw = pd.DataFrame(columns=['date', 'id', 'colnm', 'total', 'delta', 'week'])
            for fn in ['prev_grp1.pkl', 'prev_grp2.pkl', 'prev_grp3.pkl', 'prev_man.pkl', 'all_data.pkl']:
                os.remove(fn) if os.path.exists(fn) else None
            self.prev_man_vw = pd.DataFrame(columns=['date', 'id', 'Doc ID', 'Total (functional / nonfunctional)', 'Specified', 'Accepted', 'Trace-up', 'Trace-down'])
            self.all_data_vw = pd.DataFrame()

    def get_type_group(self, grp_name, ids_flag=0):
        """
        Split the data according to the Type.
        """
        if ids_flag == 1:
            data = self.data.loc[self.data['ID'].isin(self.ids.tolist())]
        else:
            data = self.data

        proj_values = ['/MBSP_System', '/MBSP_System/Customer_MAN', '/MBSP_System/Customer_VW/Cobra_1.0',
                       '/MBSP_System/Customer_SCANIA', '/MBSP_SW_Assembly']
        if self.proj_spe == 1:
            data = data[data['Project'].isin(proj_values)]

        if grp_name in data['WABCO Type'].tolist():
            grp = data.groupby(['WABCO Type']).get_group(grp_name)
            return grp
        else:
            return None

    def get_concat_df(self, grp_list):
        no_flag = 0
        for grp in grp_list:
            if grp is not None:
                no_flag = 1
            else:
                continue
        return no_flag

    def find_trace_down(self, grp_data):
        """
        This method is to calculate a output column
        :param grp_data: input data for a particular group
        :return: calculated trace down value
        """
        trc_dwn = []
        trc_dwn_fn = []
        for ind, row in grp_data.iterrows():
            if row['WABCO Type'] in ['RQ L1hw Doc', 'RQ L1sw Doc', 'RQ L1me Doc']:
                trc_dwn.append(row['Acc Req downstream traceability count-Satisfied By" / "Modelled By""'])
                trc_dwn_fn.append(row['Acc Req downstream traceability count-Satisfied By" / "Modelled By"(Functional)"'])
            else:
                trc_dwn.append(row['Acc Req downstream traceability count-Decomposes To""'])
                trc_dwn_fn.append(row['Acc Req downstream traceability count-Decomposes To" (Functional)"'])
        return trc_dwn, trc_dwn_fn

    def gen_excel(self, ids_flag=0, ids_file=0):
        """
        Calculate all the output column values for the report.
        """
        if ids_flag == 1:
            print("for ids report")
            ids_flag_list = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1]
            if ids_file == 'vw':
                self.writer = pd.ExcelWriter('mbsp_report_vw.xlsx', engine='xlsxwriter')
                ids_excel = pd.read_excel(r"path_to_ID_list.xlsx")
                ids_excel = ids_excel.loc[ids_excel['Relevance'].isin(['in VW scope'])]
                self.ids = ids_excel['ID']
                self.ids.dropna(inplace=True)
                self.ids = self.ids.astype(int)
            else:
                self.writer = pd.ExcelWriter('mbsp_report_man.xlsx', engine='xlsxwriter')
                self.ids = pd.read_excel(r"path_to_doc_IDs.xls")
        else:
            ids_flag_list = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
            self.writer = pd.ExcelWriter('mbsp_report.xlsx', engine='xlsxwriter')

        grp11, grp12, grp13, grp14, grp21, grp22, grp31, grp32, grp33, grp4 = map(self.get_type_group, ['IN Customer Doc', 'IN WABCO Doc', 'IN Edit Doc', 'IN Safety Doc', 'RQ L3 Doc', 'RQ L2 Doc', 'RQ L1hw Doc', 'RQ L1sw Doc', 'RQ L1me Doc', 'SP Doc'], ids_flag_list)
        grp1 = pd.concat([grp11, grp12, grp13, grp14]) if self.get_concat_df([grp11, grp12, grp13, grp14]) == 1 else None
        grp2 = pd.concat([grp21, grp22]) if self.get_concat_df([grp21, grp22]) == 1 else None
        grp3 = pd.concat([grp31, grp32, grp33]) if self.get_concat_df([grp31, grp32, grp33]) == 1 else None
        self.grp = pd.concat([grp1, grp2, grp3, grp4])

        dtl_rep = self.grp[['Report Run Date', 'Project', 'ID', 'WABCO Type', 'Summary', 'Total number of requirements', 'Specified Requirements', 'Accepted  Requirements', 'Rejected Requirements', 'Deleted Requirements', 'Req upstream traceability count-Decomposed From" / "Defined By""', 'Req test traceability count', 'Req test pass count', 'Total number of requirements (Functional)', 'Accepted Requirements (Functional)', 'Req upstream traceability count-Decomposed From" / "Defined By" (Functional)"', 'Reused Count', 'Stability']]
        dtl_rep.rename(columns={'Project': 'Project Name', 'ID': 'Document ID', 'Summary': 'Name of the Document', 'Total number of requirements': 'Total relevant [IN - tbd,on hold,agreed,conditionally agreed], [RQ/SP -Created, Specified, Accepted, Completed]', 'Specified Requirements': 'Specified [RQ/SP], on hold [IN]', 'Accepted  Requirements': 'Accepted / Completed [RQ/SP], Ageed/Cond. [IN]', 'Deleted Requirements': 'deleted [RQ,SP] / "n/a" [RQ,SP,IN]', 'Req upstream traceability count-Decomposed From" / "Defined By""': 'Trace-up', 'Req test traceability count': 'Test Trace', 'Req test pass count': 'Test Pass', 'Req upstream traceability count-Decomposed From" / "Defined By" (Functional)"': 'Trace Up (functional)'}, inplace=True)

        created = list(map(lambda a, b, c: a-(b+c), self.grp['Total number of requirements'].tolist(), self.grp['Specified Requirements'].tolist(), self.grp['Accepted  Requirements'].tolist()))
        total = list(map(lambda a, b: a+b, self.grp['Total number of requirements'], self.grp['Rejected Requirements']))

        trace_down, trace_down_func = self.find_trace_down(self.grp)

        dtl_rep.insert(5, 'Total', total)
        dtl_rep.insert(7, 'Created [REQ/SP], tbd [IN]', created)
        dtl_rep.insert(13, 'Trace-Down', trace_down)
        dtl_rep.insert(16, 'Trace-Down (functional)', trace_down_func)

        grp1_man, grp2_man = map(self.fetch_grp3_data, [grp1, grp2], ['Input', 'System'], [ids_flag, ids_flag])
        res1, res2, res3 = map(self.fetch_grp3_data, [grp32, grp31, grp33], ['Software', 'Hardware', 'Mechanical'], [ids_flag, ids_flag, ids_flag])

        empty_ids = []
        for g, i in zip([grp1, grp2, grp32, grp31, grp33], ['Input', 'System', 'Software', 'Hardware', 'Mechanical']):
            if g is None:
                empty_ids.append(i)

        self.write_data_excel(pd.concat([grp1_man, grp2_man, res1, res2, res3]), ids_flag, ids_file, 'grpm')

        tot_nonfunc = list(map(lambda a, b: a - b, self.grp['Total number of requirements'].tolist(), dtl_rep['Total number of requirements (Functional)'].tolist()))
        dtl_rep.insert(17, 'Total number of requirements (Non Functional)', tot_nonfunc)

        self.delta_data = dtl_rep
        self.delta_data.fillna(0, inplace=True)
        print("before calling")
        self.save_del_week = pd.DataFrame()
        val = list(map(self.create_delta_week,
                  [('Created [REQ/SP], tbd [IN]', 'Daily Delta - Created [REQ/SP], tbd [IN]', 'Weekly Delta - Created [REQ/SP], tbd [IN]'),
                   ('Specified [RQ/SP], on hold [IN]', 'Daily Delta - Specified [RQ/SP], on hold [IN]', 'Weekly Delta - Specified [RQ/SP], on hold [IN]'),
                   ('Accepted / Completed [RQ/SP], Ageed/Cond. [IN]', 'Daily Delta - Accepted / Completed [RQ/SP], Ageed/Cond. [IN]', 'Weekly Delta - Accepted / Completed [RQ/SP], Ageed/Cond. [IN]'),
                   ('Rejected Requirements', 'Daily Delta - Rejected', 'Weekly Delta - Rejected'),
                   ('deleted [RQ,SP] / "n/a" [RQ,SP,IN]', 'Daily Delta - deleted [RQ,SP] / "n/a" [RQ,SP,IN]', 'Weekly Delta - deleted [RQ,SP] / "n/a" [RQ,SP,IN]'),
                   ('Trace-up', 'Daily Delta - Trace-up', 'Weekly Delta - Trace-up')]))

        print("after delta")
        if self.up_report == 1:
            dw_s = pd.DataFrame()
            all_data = pd.DataFrame()
            man_s = self.prev_man
        else:
            if ids_flag == 1:
                dw_s = self.delta_week_vw if ids_file == 'vw' else self.delta_week_man
                all_data = self.all_data_vw if ids_file == 'vw' else self.all_data_man
                man_s = self.prev_man_vw if ids_file == 'vw' else self.prev_man_man
            else:
                dw_s = self.delta_week
                all_data = self.all_data
                man_s = self.prev_man

        self.delta_data['date'] = str(self.date)[:10]
        d_cols = list(self.delta_data)
        delta_cols = [col for col in d_cols if col[:11] == 'Daily Delta']
        week_cols = [col for col in d_cols if col[:12] == 'Weekly Delta']
        d_cols = d_cols[:16] + d_cols[21:23] + delta_cols + week_cols + d_cols[18:21] + d_cols[16:18] + d_cols[-1:]
        self.delta_data = self.delta_data[d_cols]
        self.delta_data.rename(columns={'Total number of requirements (Non Functional)': 'Total relevant (Non-functional)', 'Total number of requirements (Functional)': 'Total relevant (Functional)'}, inplace=True)
        acc_nonfunc = list(map(lambda a, b: a - b, self.delta_data['Accepted / Completed [RQ/SP], Ageed/Cond. [IN]'], self.delta_data['Accepted Requirements (Functional)']))
        trup_nonfunc = list(map(lambda a, b: a - b, self.delta_data['Trace-up'], self.delta_data['Trace Up (functional)']))
        trdn_nonfunc = list(map(lambda a, b: a - b, self.delta_data['Trace-Down'], self.delta_data['Trace-Down (functional)']))

        self.delta_data.insert(36, 'Accepted Requirements (Non-functional)', acc_nonfunc)
        self.delta_data.insert(37, 'Trace Up (Non-functional)', trup_nonfunc)
        self.delta_data.insert(38, 'Trace-Down (Non-functional)', trdn_nonfunc)

        if self.date.weekday() == 0 and self.up_report != 1:
            all_data = all_data[all_data['date'] == str(self.date-timedelta(days=3))[:10]]

        dr_data = pd.concat([self.delta_data, all_data])
        all_dates = dr_data['date'].unique().tolist()
        print(all_dates)

        for date in all_dates:
            day_data = dr_data[dr_data['date'] == date]
            d = datetime.datetime.strptime(date[:10], '%Y-%M-%d').date()
            sht_nm = 'DR-Day' + str(d.weekday()+1)
            sht_nm = 'DR-Day0' if (date[:10] == str(self.date-timedelta(days=(self.date.weekday()+3)))[:10]) else sht_nm

            day_data.drop(['date'], axis=1, inplace=True)
            day_data = day_data[['Report Run Date', 'Project Name', 'Document ID', 'WABCO Type', 'Name of the Document', 'Total', 'Total relevant [IN - tbd,on hold,agreed,conditionally agreed], [RQ/SP -Created, Specified, Accepted, Completed]','Created [REQ/SP], tbd [IN]' ,
                                 'Specified [RQ/SP], on hold [IN]', 'Accepted / Completed [RQ/SP], Ageed/Cond. [IN]', 'Rejected Requirements','deleted [RQ,SP] / "n/a" [RQ,SP,IN]', 'Trace-up', 'Trace-Down', 'Test Trace', 'Test Pass', 'Reused Count','Stability',
                                 'Daily Delta - Created [REQ/SP], tbd [IN]','Daily Delta - Specified [RQ/SP], on hold [IN]','Daily Delta - Accepted / Completed [RQ/SP], Ageed/Cond. [IN]','Daily Delta - Rejected','Daily Delta - deleted [RQ,SP] / "n/a" [RQ,SP,IN]',
                                 'Daily Delta - Trace-up','Weekly Delta - Created [REQ/SP], tbd [IN]','Weekly Delta - Specified [RQ/SP], on hold [IN]','Weekly Delta - Accepted / Completed [RQ/SP], Ageed/Cond. [IN]',
                                 'Weekly Delta - Rejected','Weekly Delta - deleted [RQ,SP] / "n/a" [RQ,SP,IN]','Weekly Delta - Trace-up','Total relevant (Functional)','Accepted Requirements (Functional)','Trace Up (functional)',
                                 'Trace-Down (functional)','Total relevant (Non-functional)','Accepted Requirements (Non-functional)','Trace Up (Non-functional)','Trace-Down (Non-functional)']]
            day_data.to_excel(self.writer, sht_nm, index=False)

        self.workbook.add_worksheet()
        self.writer.save()
        print("saved excel")

        xl = Dispatch("Excel.Application")
        #xl.Visible = True # You can remove this line if you don't want the Excel application to be visible

        out_file = 'mbsp_report.xlsx'
        if ids_flag == 1:
            out_file = 'mbsp_report_vw.xlsx' if ids_file == 'vw' else 'mbsp_report_man.xlsx'
        wb1 = xl.Workbooks.Open(Filename='mbsp_help.xlsx')
        wb2 = xl.Workbooks.Open(Filename='\\' + out_file)

        sheet_num = len(all_dates) + 2
        ws1 = wb1.Worksheets(1)
        print(sheet_num)
        ws1.Copy(Before=wb2.Worksheets(sheet_num))
        desheet = wb2.Sheets('Sheet'+str(sheet_num))
        desheet.Delete()
        wb2.Close(SaveChanges=True)
        wb1.Close()
        xl.Quit()

        if grp1 is not None:
            self.save_df(pd.concat([grp1_man, grp2_man, res1, res2, res3]), 'prev_man.pkl', ids_file, man_s)
        else:
            self.save_df(pd.concat([grp1_man, grp2_man, res1, res2, res3]), 'prev_man.pkl', ids_file, man_s)

        self.save_df(self.delta_data, 'all_data.pkl', ids_file, all_data)
        self.save_df(self.save_del_week, 'delta_week.pkl', ids_file, dw_s)

    def save_df(self, curr_data, file_nm, ids_file, prev_data=None):
        """
        Saving the dataframe for next day/future
        """
        if ids_file != 0:
            file_nm = file_nm[:-4] + "_" + ids_file + file_nm[-4:]
        print(file_nm)
        curr_data.reset_index(drop=True, inplace=True)
        if (self.day1_flag == 1) or (self.day1_flag != 1 and self.date.weekday() == 0) or (self.ids_day1 == 1) or (self.up_report == 1):
            if self.date.weekday() == 0 and self.up_report != 1:
                prev_data = prev_data[prev_data['date'] == str(self.date-timedelta(days=3))[:10]]
                fnl_data = pd.concat([prev_data, curr_data])
                fnl_data.to_pickle(file_nm)
            else:
                curr_data.to_pickle(file_nm)
        else:
            fnl_data = pd.concat([prev_data, curr_data])
            fnl_data.to_pickle(file_nm)

    def get_prev_delta_week(self, id_num, date, p_data):
        """
        Reading previous week's data from pickle files.
        """
        if date.weekday() == 0:
            yes_data = p_data[(p_data['id'] == id_num)]
            yes_data = yes_data[(yes_data['date'] == str(date-timedelta(days=3))[:10])]
        else:
            yes_data = p_data[(p_data['id'] == id_num) & (p_data['date'] == str(date-timedelta(days=1))[:10])]
        int_cols = ['id', 'total', 'delta', 'week']
        yes_data[int_cols] = yes_data[int_cols].astype(int)
        return yes_data['total'].tolist()[0], yes_data['week'].tolist()[0]

    def create_delta_week(self, col_nm):
        """
        Creating delta value with current and previous day's numbers.
        """
        print(col_nm[0])
        grp_data = self.delta_data
        if self.up_report == 1:
            p_data = self.delta_week[self.delta_week['colnm'] == col_nm[0]]
        else:
            p_data = self.delta_week_vw[self.delta_week_vw['colnm'] == col_nm[0]]
        p_data.reset_index(drop=True, inplace=True)
        #print(p_data)
        res = []
        for ind, row in grp_data.iterrows():
            #print(row['Document ID'])
            if self.up_report == 1:
                res.append([row['Document ID'], col_nm[0], row[col_nm[0]], 0, 0])
            else:
                if str(row['Document ID']) in ['5755570', '5754619', '5759961', '5760027']:
                    #print("here")
                    res.append([row['Document ID'], col_nm[0], row[col_nm[0]], 0, 0])
                else:
                    #print("getting prev data")
                    if row['Document ID'] in p_data['id'].tolist():
                        y_tot, y_week = self.get_prev_delta_week(row['Document ID'], self.date, p_data)
                        res.append([row['Document ID'], col_nm[0], row[col_nm[0]], (row[col_nm[0]] - y_tot), y_week + (row[col_nm[0]] - y_tot)])
                    else:
                        res.append([row['Document ID'], col_nm[0], row[col_nm[0]], 0, 0])
        grp_res = pd.DataFrame(res, columns=['id', 'colnm', 'total', 'delta', 'week'])
        grp_res['date'] = str(self.date)[:10]
        int_cols = ['id', 'total', 'delta', 'week']
        grp_res[int_cols] = grp_res[int_cols].astype(int)
        grp_res = grp_res[['date', 'id', 'colnm', 'total', 'delta', 'week']]
        self.delta_data[col_nm[1]] = grp_res['delta'].values
        self.delta_data[col_nm[2]] = grp_res['week'].values
        self.save_del_week = pd.concat([self.save_del_week, grp_res])
        return 1

    def fetch_grp3_data(self, data, id, ids_flag):
        """
        To calculate output data values for a different group. 
        """
        if data is None:
            res = pd.DataFrame(columns=['date', 'id', 'Doc ID', 'Total (functional / nonfunctional)', 'Req Evaluated', 'in_Trace-down', 'type'])
            return res
        res = []
        grp_flag = 'grpm'

        if (id != 'Input' and self.sep_flag == 1) or ((id != 'Input' and self.flag2 == 1)):
            stb_totreq = list(map(lambda a, b: a*b, data['Stability'], data['Total number of requirements']))
            if id in ['Software', 'Hardware', 'Mechanical']:
                trc_dwn = data['Acc Req downstream traceability count-Satisfied By" / "Modelled By""'].sum()
            else:
                trc_dwn = data['Acc Req downstream traceability count-Decomposes To""'].sum()

            res.append(['Total', (data['Total number of requirements'].sum() + data['Rejected Requirements'].sum() + data['Deleted Requirements'].sum()),
                       data['Total number of requirements'].sum(),
                        (data['Total number of requirements'].sum() - (data['Specified Requirements'].sum() + data['Accepted  Requirements'].sum())),
                        data['Specified Requirements'].sum(),
                        data['Accepted  Requirements'].sum(),
                        data['Rejected Requirements'].sum(),
                        data['Deleted Requirements'].sum(),
                        data['Req upstream traceability count-Decomposed From" / "Defined By""'].sum(),
                        trc_dwn,
                        data['Reused Count'].sum(),
                        (sum(stb_totreq)/(data['Total number of requirements'].sum()))])
            grp_res = pd.DataFrame(res, columns=['Doc ID', 'Total Requirements', 'Total relevant [Created, Specified, Accepted, Completed]', 'Created', 'Specified', 'Accepted/ Completed', 'Rejected',  'deleted / "n/a"', 'Trace-up', 'Trace-down (Only for accepted)', 'Reused', 'Stability Index'])
        else:
            res.append(['Total', (data['Total number of requirements'].sum() + data['Rejected Requirements'].sum() + data['Deleted Requirements'].sum()),
                       data['Total number of requirements'].sum(),
                        (data['Total number of requirements'].sum() - (data['Specified Requirements'].sum() + data['Accepted  Requirements'].sum())),
                        data['Specified Requirements'].sum(),
                        data['Accepted  Requirements'].sum(),
                        data['Rejected Requirements'].sum(),
                        data['Deleted Requirements'].sum(),
                        (data['Acc Req downstream traceability count-Decomposes To""'].sum())])
            grp_res = pd.DataFrame(res, columns=['Doc ID', 'Total Requirements', 'Total relevant [tbd,on hold,agreed,conditionally agreed]', 'tbd', 'on hold', 'agreed/ conditionally agreed', 'rejected', 'n/a', 'Trace-down (Only for agreed)'])

        grp_res['date'] = str(self.date)[:10]
        grp_res['type'] = grp_flag
        grp_res['id'] = id
        if (id != 'Input' and self.sep_flag == 1) or (id != 'Input' and self.flag2 == 1):
            grp_res = grp_res[['date', 'id', 'Doc ID', 'Total Requirements', 'Total relevant [Created, Specified, Accepted, Completed]', 'Created', 'Specified', 'Accepted/ Completed', 'Rejected',  'deleted / "n/a"', 'Trace-up', 'Trace-down (Only for accepted)', 'Reused', 'Stability Index', 'type']]
        else:
            grp_res = grp_res[['date', 'id', 'Doc ID', 'Total Requirements', 'Total relevant [tbd,on hold,agreed,conditionally agreed]', 'tbd', 'on hold', 'agreed/ conditionally agreed', 'rejected', 'n/a', 'Trace-down (Only for agreed)', 'type']]
        return grp_res

    def get_summary(self, all_ids, look_data=None):
        """
        Getting the summary data for all the metrics
        """
        sumry = []
        for id in all_ids:
            sumry.append(look_data[look_data['id'] == id]['summary'].tolist()[0])
        return sumry

    def write_data_excel(self, data, ids_flag, ids_file, gnm):
        """
        Writing all the required output values to excel sheets.
        """
        if self.day1_flag == 1 or (self.up_report == 1):
            all_df = data
        else:
            if ids_flag == 1:
                save_data = self.prev_man_vw if ids_file == 'vw' else self.prev_man_man
            else:
                save_data = self.prev_man
            save_data.reset_index(drop=True, inplace=True)
            if save_data[save_data['type'] == gnm].empty:
                all_df = data
            elif self.day1_flag != 1 and self.date.weekday() == 0:
                save_data = save_data[save_data['date'] == str(self.date - timedelta(days=3))[:10]]
                all_df = pd.concat([save_data[save_data['type'] == gnm], data])
            else:
                all_df = pd.concat([save_data[save_data['type'] == gnm], data])

        self.workbook = self.writer.book
        worksheet = self.workbook.add_worksheet('Requirements overview')

        all_dates = all_df['date'].unique().tolist()
        worksheet.write(1, 0, 'Date')
        srow = 2
        for date in all_dates:
            worksheet.write(srow, 0, date)
            srow = srow + 1

        all_ids = all_df['id'].unique().tolist()
        ids_colr = [('#FFD39B', '#FFE4C4'), ('#EEA9B8', '#FFE4E1'), ('#CD96CD', '#EED2EE'), ('#6E8B3D', '#CAFF70'), ('#33A1C9', '#B2DFEE') ]

        r1 = r2 = 0
        c1 = 1
        c2 = 0
        scol = 1
        srow = 1

        isrow = len(all_dates)+5
        for id in all_ids:
            iscol = 2
            wdf = all_df[all_df['id'] == id]

            all_dates = wdf['date'].tolist()

            wdf = wdf[wdf['Doc ID'] == 'Total']
            tot_req = wdf['Total Requirements']

            if id in ['System', 'Domain', 'Hardware', 'Software', 'Mechanical']:
                wdf.drop(['Total relevant [tbd,on hold,agreed,conditionally agreed]', 'tbd', 'on hold', 'agreed/ conditionally agreed', 'rejected', 'n/a', 'Trace-down (Only for agreed)'], axis=1, inplace=True)
                wdf = wdf[['Total Requirements', 'Total relevant [Created, Specified, Accepted, Completed]', 'Created', 'Specified', 'Accepted/ Completed', 'Rejected', 'deleted / "n/a"', 'Trace-up', 'Trace-down (Only for accepted)', 'Reused', 'Stability Index']]
                tot_del = wdf['deleted / "n/a"']
            elif id == 'Input':
                wdf.drop(['Total relevant [Created, Specified, Accepted, Completed]', 'Created', 'Specified', 'Accepted/ Completed', 'Rejected', 'deleted / "n/a"', 'Trace-up', 'Trace-down (Only for accepted)', 'Reused', 'Stability Index'], axis=1, inplace=True)
                wdf = wdf[['Total Requirements', 'Total relevant [tbd,on hold,agreed,conditionally agreed]', 'tbd', 'on hold', 'agreed/ conditionally agreed', 'rejected', 'n/a', 'Trace-down (Only for agreed)']]
                tot_del = wdf['n/a']

            #print(tot_req)
            #print(tot_del)

            id_cols = list(wdf)
            id_data = wdf
            c2 = c2 + len(id_cols)
            cell_format = self.workbook.add_format({'align': 'center', 'bg_color': ids_colr[all_ids.index(id)][0]})
            if id in ['Hardware', 'Software', 'Mechanical']:
                id_name = "Main Assembly - " + id.upper()
            id_name = id_name if id in ['Hardware', 'Software', 'Mechanical'] else id
            worksheet.merge_range(r1, c1, r2, c2, id_name, cell_format)
            c1 = c2+1

            width = 0.25
            fig, ax = plt.subplots(figsize=(6, 4.5))
            x_dates = [datetime.datetime.strptime(d, '%Y-%m-%d').date() for d in all_dates]
            if id == 'Input':
                bar_cols = [id_data['tbd'], id_data['on hold'], id_data['agreed/ conditionally agreed'], id_data['rejected'], id_data['n/a']]
                series_labels = ['tbd', 'on hold', 'agreed/ conditionally agreed', 'rejected', 'n/a']
            else:
                bar_cols = [id_data['Created'], id_data['Specified'], id_data['Accepted/ Completed'], id_data['Rejected'], id_data['deleted / "n/a"']]
                series_labels = ['Created', 'Specified', 'Accepted/ Completed', 'Rejected', 'deleted / "n/a"']

            def stacked_bar(data, series_labels, bar_colr, tot_req, tot_del, category_labels=None,
                            show_values=False, value_format="{}", y_label=None,
                            grid=True, reverse=False):

                ny = len(data[0])
                ind = list(range(ny))

                axes = []
                cum_size = np.zeros(ny)

                data = np.array(data)

                if reverse:
                    data = np.flip(data, axis=1)
                    category_labels = reversed(category_labels)

                for i, row_data in enumerate(data):
                    plt.rcParams["figure.figsize"] = [6, 4.5]
                    plt.rcParams.update({'font.size':8})
                    axes.append(plt.bar(ind, row_data, width, bottom=cum_size, color=bar_colr[i],
                                        label=series_labels[i]))
                    cum_size += row_data

                if category_labels:
                    plt.xticks(ind, category_labels)
                    tot_del = 0 if str(tot_del) in ['nan', 'NaN'] else tot_del
                    tot = tot_req + tot_del
                    max_y = tot+(round(int(tot/4)))
                    if tot / 2 < 2000:
                        step_val = 200
                        max_y = tot + (round(int(tot / 3)))
                    elif tot/2 < 5000 and tot > 2000:
                        step_val = 500
                    elif tot/2 > 5000 and tot < 10000:
                        step_val = 2000
                    else:
                        step_val = 2500

                    y_values = range(step_val, int(max_y), step_val)
                    plt.yticks(y_values)

                if y_label:
                    plt.ylabel(y_label)

                plt.legend(loc='upper center', bbox_to_anchor=(0.5, 1.15), fancybox=True, shadow=True, ncol=3) #, fontsize='x-small'

                if show_values:
                    for axis in axes:
                        for bar in axis:
                            w, h = bar.get_width(), bar.get_height()
                            h = 0 if str(h) in ['nan', 'NaN'] else h
                            plt.text(bar.get_x() + w / 2, bar.get_y() + h / 2,
                                     value_format.format(int(h)), ha="center",
                                     va="center")

            stacked_bar(
                bar_cols,
                series_labels,
                ['#4682B4', '#CD5555', '#6E8B3D', '#9A32CD', '#33A1C9'],
                max(tot_req), max(tot_del),
                category_labels=x_dates,
                show_values=True,
                value_format="{:d}",
                y_label=""
            )

            plt.savefig(id + '_.png')
            plt.close()
            worksheet.insert_image(isrow, iscol, id + '_.png')

            if id in ['Hardware', 'Software', 'Mechanical']:
                id_name = "Main Assembly - " + id.upper()
            id_name = id_name if id in ['Hardware', 'Software', 'Mechanical'] else id
            worksheet.merge_range(isrow, 0, isrow, 1, id_name, cell_format)
            for idc in id_cols:
                srow = 1
                format1 = self.workbook.add_format({'bg_color': ids_colr[all_ids.index(id)][1]})
                col_data = id_data[idc]
                worksheet.write(srow, scol, idc, format1)
                for d in col_data:
                    d = '' if str(d) == 'nan' else d
                    srow = srow + 1
                    format1.set_align('left')
                    format1.set_text_wrap()
                    worksheet.write(srow, scol, d, format1)
                scol = scol + 1

                if idc in ['Total relevant [tbd,on hold,agreed,conditionally agreed]', 'Total relevant [Created, Specified, Accepted, Completed]', 'Trace-up', 'Specified', 'agreed/ conditionally agreed', 'Accepted/ Completed', 'Trace-down (Only for agreed)', 'Trace-down (Only for accepted)']:
                    total = wdf[idc].tolist()
                    fig, ax = plt.subplots(figsize=(4.5, 4.5))
                    ax.tick_params(labelsize=8)
                    if idc in ['Total relevant [tbd,on hold,agreed,conditionally agreed]', 'Total relevant [Created, Specified, Accepted, Completed]']:
                        line_col = '#4682B4'
                        iscol = 12
                    elif idc in ['Trace-up']:
                        line_col = '#FFB90F'
                        iscol = 19
                    elif idc in ['Specified']:
                        line_col = '#CD5555'
                        iscol = 26
                    elif idc in ['agreed/ conditionally agreed', 'Accepted/ Completed']:
                        line_col = '#6E8B3D'
                        iscol = 33
                    elif idc in ['Trace-down (Only for agreed)', 'Trace-down (Only for accepted)']:
                        line_col = '#9A32CD'
                        iscol = 40

                    #print(x_dates)
                    #print(total)
                    #print(idc)
                    ax.plot(x_dates, total, color=line_col, label=idc)
                    ax.grid(axis='y')
                    ax.set_xticks(x_dates)

                    box = ax.get_position()
                    ax.set_position([box.x0, box.y0 + box.height * 0.1, box.width, box.height * 0.9])

                    # Put a legend below current axis
                    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), fancybox=True, shadow=True, ncol=4, fontsize='x-small')
                    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))

                    fig.autofmt_xdate()

                    ax.set_title("\n".join(textwrap.wrap(idc, 35)), fontsize=12)
                    #fig.suptitle("\n".join(textwrap.wrap(idc, 20)), fontsize=12)
                    img_flnm = id+idc+".png"
                    for ch in ["/", "\""]:
                        img_flnm = img_flnm.replace(ch, '')

                    #fig.tight_layout()
                    fig.savefig(img_flnm, dpi=150)
                    #iscol = iscol + 10
                    #print("at location", isrow, iscol, id+idc+".png")
                    worksheet.insert_image(isrow, iscol, img_flnm)
                    plt.gcf().clear()
            scol = scol
            isrow = isrow + 25

    def close_report(self, ids_flag):
        pass


class WriteExcel():

    def __init__(self, gr):
        self.gr = gr

    def write_cell_dates(self, ids=None, ids_file=0):
        if ids == 1:
            gm = self.gr.gen_excel(1, ids_file)
            #out_file = "ids_report.xlsx"
            self.gr.close_report(ids)
        else:
            gm = self.gr.gen_excel()
            #out_file = "overall_report.xlsx"
            self.gr.close_report(ids)


if __name__ == "__main__":
    report = GenerateReport()
    we = WriteExcel(report)
    we.write_cell_dates(1, 'vw')
    #we.write_cell_dates(1, 'man')
    #we.write_cell_dates()

