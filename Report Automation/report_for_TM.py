#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
This module generates an excel report for a given input Test Plan.

Author : Sindhu Prakasam
"""

import os
import sys
import logging
import time
from datetime import datetime
import pandas as pd
from tkinter import *
import subprocess
from bs4 import BeautifulSoup
import codecs
import warnings
warnings.filterwarnings("ignore")

hostname = os.environ.get("MKSSI_HOST")
username = os.environ.get("MKSSI_USER")

#hostname = 'wdecspdpmks'
#username = 'vikramm'


class Interface():
    """GUI"""
    def __init__(self, master):
        self.master = master
        master.title(" Enter Test Plan/Sample/Phase ID and Submit ")

        frame1 = Frame(self.master, borderwidth=2, bg='old lace', relief=RAISED)
        frame1.pack(side=TOP, padx=20, pady=20)

        self.dev_var = StringVar()
        label1 = Label(frame1, text='Test Plan ID', bg='old lace').grid(row=1, sticky=E)
        entry1 = Entry(frame1, textvariable=self.dev_var, width=60, exportselection=1).grid(row=1, column=2)

        # Adding instructions for input
        explanation = """
        1. Please enter a valid test plan ID or Test Sample ID or Test Phase ID\n
        2. We advise you to run this script for items which has number of linked relationship items less than 700\n
        3. User should have access to that item\n
        4. No need to open/select the test plan to run this script"""

        w2 = Label(root, justify=LEFT, text='Instructions for Input:', font='Helvetica 12 bold').pack(side="top")
        w3 = Label(root, justify=LEFT, text=explanation, font='Helvetica 10').pack(side="top")

        bouton_recup = Button(master, text="Submit", relief=RAISED, command=self.recupere)# Call the function 'recupere'
        bouton_recup.pack(side=RIGHT, padx=5, pady=5)

    def recupere(self):
        "Function to get the variable"
        self.dev = self.dev_var.get()
        self.master.destroy()
        return self.dev_var.set(self.dev)

    def get_value(self):
        return self.dev


class InterfaceCount():
    """GUI"""
    def __init__(self, master):
        self.master = master

        master.title("Cannot proceed script execution")

        frame1 = Frame(self.master, borderwidth=2, bg='old lace', relief=RAISED)
        frame1.pack(side=TOP, padx=20, pady=20)

        # Adding instructions for input
        explanation1 = """
        Problem:\n
        Given Test Plan may contain/linked to many relationship items (more than 700).\n
        Why stopped:\n
        Running the script for the given test plan might affect PTC performance.\n
        
        Solution:\n
        - Try to divide the large Test Plan and run the script for 
        Test Plan\Test Sample\Test Phase with less number of items linked to it.\n
        - If not, Contact PTC Support team to run this script further."""

        w2 = Label(root, justify=LEFT, text="Script execution cannot be started", font='Helvetica 14 bold').pack(side="top")
        w3 = Label(root, justify=LEFT, text=explanation1, font='Helvetica 10').pack(side="top")

        bouton_recup = Button(master, text="Close", relief=RAISED, command=self.recupere)# Call the function 'recupere'
        bouton_recup.pack(side=RIGHT, padx=5, pady=5)

    def recupere(self):
        "Function to get the variable"
        self.master.destroy()


class InterfaceHalfdone():
    """GUI"""

    def __init__(self, master):
        self.master = master

        master.title("Cannot proceed script execution")

        frame1 = Frame(self.master, borderwidth=2, bg='old lace', relief=RAISED)
        frame1.pack(side=TOP, padx=20, pady=20)

        explanation2 = """
        Script took more than 30 minutes to run. This will affect PTC performance
        if we continue running the script.

        Script hasn't processed all the items in the given test plan.
        Report has been generated for the items processed till now.
        """

        heading_text = "Cannot proceed (Run time exceeded)"
        explanation = explanation2

        w2 = Label(root, justify=LEFT, text=heading_text,
                   font='Helvetica 14 bold').pack(side="top")
        w3 = Label(root, justify=LEFT, text=explanation,
                   font='Helvetica 10').pack(side="top")

        bouton_recup = Button(master, text="Close", relief=RAISED,
                              command=self.recupere)  # Call the function 'recupere'
        bouton_recup.pack(side=RIGHT, padx=5, pady=5)

    def recupere(self):
        "Function to get the variable"
        self.master.destroy()


def execommand(command, commandfile, id):
    """execute binaries"""
    error_flag = 0
    exe_cmd = u'' + command
    try:
        logging.info("Executing Command:" + command)
        #text_mode = ('' is None)
        p = subprocess.Popen(exe_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE) #universal_newlines=True,
        p.wait()
        output, err = p.communicate()

        try:
            #output = str(output, "utf-8")
            #str = unicode(str, errors='ignore')
            output = output.decode("utf-8", errors='replace')
        except Exception as e:
            print(str(e))
            print("Error in decoding german text")
            logging.error("Error in decoding german text")
            time.sleep(5)

        if p.returncode == 0:
            out = output.splitlines()
        else:
            out = err.splitlines()

    except Exception as e:
        print(str(e))
        print("Host not available!")
        print("Exit...")
        logging.error("Error while running/reading the PTC Command.")
        print("Error running PTC Command")
        time.sleep(5)
        sys.exit(1)

    for line in out:
        if line.startswith('***') or ('does not exist' in line) or ('you may' in line) or (
                'was unexpected' in line) or \
                ('not recognized as an internal or external command' in line) or (
                'operable program or batch file' in line) or \
                ('supervisor or administrator' in line) or ('Could not save' in line) or \
                ('system cannot find' in line) or ('You may not' in line) or \
                ('command requires operands:' in line) or ('The command line is too long' in line) or ('not a valid Item' in line):
            logging.error("****************** Error Executing Command ***************************")
            logging.error("Output of Command -" + ' '.join(out))
            error_flag = 1
            err_file.write("\nGetting Error for ID-" + id + "\n")
            err_file.write("Error Excuting the following command\n" + exe_cmd + "\n")
            err_file.write("Output of Command -" + ' '.join(out) + "\n\n")
    return out, error_flag


def run_report(test_plan_id):
    logging.info("Started fetching data from PTC ...")
    print("Started fetching data from PTC ...")
    #test_plan_id = '6016617'
    report_out_flnm = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ptcreport_' + str(test_plan_id) + ".html")
    run_report_cmd = "im runreport --yes --substituteParams --user='" + \
                     username + "' --hostname='" + hostname + \
                     "'  --issues=" + str(test_plan_id) + ' --outputfile="' + report_out_flnm + '"  "Input for Test Report Overview -- USED FOR CUSTOM SCRIPT"'

    run_report_out, run_err = execommand(run_report_cmd, '', '')
    if run_err == 1:
        logging.error("Error executing PTC report. Reason may be related to PTC server slowness.")
        logging.info("Try after sometime or Contact Script Developer")
        logging.info("Command output:" + ",".join(run_report_out))
        sys.exit("Error executing run report, Please try after some time or Contact developer.")

    logging.info("Data fetch from PTC completed.")
    return report_out_flnm


def read_html(test_plan_id, start_time, runtime):
    report_out_flnm = run_report(test_plan_id)
    #report_out_flnm = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ptcreport_6655148.html')

    print("Generating report ...")
    f = codecs.open(report_out_flnm, 'r', 'utf-8')
    soup = BeautifulSoup(f.read(), 'lxml')
    all_tables = soup.find_all('table')
    res_df = pd.DataFrame()
    last_type = ''
    add_test_flag = 0
    handld_ids = []
    handld_sesstc = []
    for table in all_tables:
        for row in table.find_all("tr"):

            if len(row.text.split('\n')) > 2 and row.text.split('\n')[2] == 'ALM_Test Plan':
                text = row.text.split('\n')[1:-1]
                if text[0] in handld_ids:
                    continue

                print(text[0])
                tb = pd.DataFrame([text + ['level0']],
                                   columns=['id', 'type', 'summary', 'state', 'user', 'plan_cnt',
                                            'pass_cnt', 'fail_cnt', 'othr_cnt', 'vm_level',
                                            'plan_enddt', 'rel_name', 'env_picklist', 'hard_vname',
                                            'soft_vname', 'level'])
                tb['text'] = ''
                tb['result'] = ''
                tb['state'] = ''

                # adding on july 23
                tb['prcr_link'] = ''
                tb['prcr_state'] = ''
                handld_ids.append(tb['id'].tolist()[0])
                res_df = pd.concat([res_df, tb], ignore_index=True)
                nowtime = time.time()

                if (nowtime - start_time) > runtime:
                    return res_df, 1

            for row1 in row.find_all("table"):
                for row2 in row1.find_all("tr"):
                    sess_tc = []
                    if row2.attrs and 'class' in row2.attrs.keys():
                        text = row2.text.split('\n')
                        text = text[1:-1]
                        if row2.attrs['class'][0] != 'level1':
                            print(text[0])
                            tb1 = pd.DataFrame([text + row2.attrs['class']],
                                           columns=['id', 'type', 'summary', 'text',
                                                    'state', 'user', 'plan_cnt',
                                                    'pass_cnt', 'fail_cnt', 'othr_cnt', 'vm_level',
                                                    'plan_enddt', 'rel_name', 'env_picklist',
                                                    'hard_vname', 'soft_vname',
                                                    's_plan', 's_pass', 's_fail', 's_other',
                                                    'level'])
                        else:
                            print(text[0])
                            tb1 = pd.DataFrame([text + row2.attrs['class']],
                                               columns=['id', 'type', 'summary', 'state', 'user',
                                                        'plan_cnt', 'pass_cnt', 'fail_cnt', 'othr_cnt',
                                                        'vm_level', 'plan_enddt', 'rel_name', 'env_picklist',
                                                        'hard_vname', 'soft_vname', 'level'])

                            tb1['text'] = ''
                        handld_ids.append(tb1['id'].tolist()[0])
                        '''
                        tb1 = [['id', 'type', 'summary', 'text', 'state', 'user', 'plan_cnt', 'pass_cnt', 'fail_cnt',
                                'othr_cnt', 'vm_level', 'plan_enddt', 'rel_name', 'env_picklist', 'hard_vname', 'soft_vname', 'level']]
                        tb1.reset_index()
                        '''
                        res_df = pd.concat([res_df, tb1], ignore_index=True)
                        nowtime = time.time()
                        if (nowtime - start_time) > runtime:
                            return res_df, 1

                    for row3 in row2.find_all("tr"):

                        if row3.find_all("td"):
                            row3_attrs = row3.find_all("td")[0].attrs
                        else:
                            continue

                        if row3_attrs == {}:
                            if 'NO_TEST_RESULTS_VALUE' in row3.text.split('\n'):
                                if last_type == 'ALM_Test Session':
                                    skip_flag = 1
                                else:
                                    continue
                            else:
                                tc_dtls = row3.text.split('\n')

                                if len(tc_dtls) != 5 or tc_dtls[0] != '' or tc_dtls[4] != '':
                                    continue
                                elif skip_flag == 1 or tc_dtls[2] in tc_dtl_list:
                                    continue

                                if add_test_flag == 1 and tc_dtls[1] + '_' + tc_dtls[2] not in sess_tc:
                                    tc_df_now = pd.DataFrame([tc_dtls[1:-1]], columns=['sess', 'tc', 'verdict'])
                                    tc_df = pd.concat([tc_df, tc_df_now], ignore_index=True)
                                    sess_tc.append(tc_dtls[1] + '_' + tc_dtls[2])

                        elif 'class' not in row3_attrs.keys():
                            if len(row3.text.split('\n')) > 2:
                                last_type = row3.text.split('\n')[2]
                            continue
                        elif row3_attrs and row3_attrs['class'][0].startswith('level'):
                            text = row3.text.split('\n')
                            text = text[1:-1]
                            if (text[1] != 'ALM_Test Case' and text[0] in handld_ids) or \
                                    (text[1] == 'ALM_Test Case' and last_sess_id + '_' + text[0] in handld_ids):
                                continue

                            if text[1] == 'ALM_Test Case' and skip_flag == 1:
                                continue

                            if text[1] == 'ALM_Test Case' and last_sess_id + '_' + text[0] in handld_sesstc:
                                continue
                            if text[1] == 'ALM_Test Case' and text[0] not in tc_df['tc'].tolist():
                                last_type = 'ALM_Test Case'
                                continue

                            #adding this on July 24
                            if text[1] == 'ALM_Test Case' and last_sess_id + '_' + text[0] not in sess_tc:
                                continue
                            print(text[0])
                            tb2 = pd.DataFrame([text + row3_attrs['class']],
                                               columns=['id', 'type', 'summary',
                                                        'text', 'state', 'user',
                                                        'plan_cnt', 'pass_cnt', 'fail_cnt',
                                                        'othr_cnt', 'vm_level',
                                                        'plan_enddt', 'rel_name', 'env_picklist',
                                                        'hard_vname', 'soft_vname',
                                                        's_plan', 's_pass', 's_fail', 's_other',
                                                        'level'])

                            last_type = tb2['type'].tolist()[0]
                            if last_type == 'ALM_Test Session':
                                skip_flag = 0
                                tc_df = pd.DataFrame([], columns=['sess', 'tc', 'verdict'])
                                tc_dtl_list = []
                                handld_sesstc = []
                                sess_tc = [] #July 24
                                last_sess_id = text[0]
                                add_test_flag = 1

                                # Copy Test Session's count values
                                tb2['plan_cnt'] = tb2['s_plan']
                                tb2['pass_cnt'] = tb2['s_pass']
                                tb2['fail_cnt'] = tb2['s_fail']
                                tb2['othr_cnt'] = tb2['s_other']
                                tb2.drop(['s_plan', 's_pass', 's_fail', 's_other'], axis=1, inplace=True)

                            if last_type == 'ALM_Test Case':
                                tb2['result'] = tc_df[tc_df['tc'] == tb2['id'].tolist()[0]]['verdict'].tolist()[0]
                                tc_dtl_list.append(text[0])
                                handld_sesstc.append(last_sess_id + '_' + text[0])
                                tb2['summary'] = tb2['text']

                                #print("session: ", last_sess_id.strip(), ", tc df:", tc_df['tc'].tolist())

                                # Finding PR/CR Links
                                cmd = "tm viewresult " + "--hostname='" + \
                                      hostname + "' --user='" + username + "'" + \
                                      " " + last_sess_id.strip() + ":" + text[0].strip()
                                prcr_out, prcr_err = execommand(cmd, '', '')

                                if "Related Items: " in prcr_out:
                                    st_index = prcr_out.index('Related Items: ')
                                    end_index = prcr_out.index('Attachments: ')
                                    Relitem_list = prcr_out[st_index + 1:end_index]
                                    if len(Relitem_list) == 0:
                                        tb2['prcr_link'] = ''
                                        tb2['prcr_state'] = ''
                                    else:
                                        for each_PRCR in Relitem_list:
                                            cmd = "im issues " + "--hostname='" + \
                                                  hostname + "' --user='" + username + \
                                                  "' --fields=State " + each_PRCR
                                            prcr_state, state_err = execommand(cmd, '', '')
                                            tb2['prcr_link'] = each_PRCR.strip()
                                            tb2['prcr_state'] = prcr_state[0].strip()

                                    handld_ids.append(last_sess_id.strip() + "_" + text[0].strip())

                            else:
                                if last_type != 'ALM_Test Session':
                                    add_test_flag = 0
                                tc_df = pd.DataFrame([], columns=['sess', 'tc', 'verdict'])
                                handld_ids.append(tb2['id'].tolist()[0])
                                tb2['result'] = ''
                                tb2['state'] = ''
                                tb2['prcr_link'] = ''
                                tb2['prcr_state'] = ''

                            res_df = pd.concat([res_df, tb2], ignore_index=True)
                            nowtime = time.time()
                            if (nowtime - start_time) > runtime:
                                return res_df, 1

    return res_df, 0


def write_to_excel(test_plan_id, start_time, runtime):

    res_df, half_flag = read_html(test_plan_id, start_time, runtime)

    #res_df.to_excel("temp.xlsx", index=False)
    res_df['summary'] = res_df['summary'].astype('str')
    res_df['summary'] = res_df['summary'].apply(lambda x: x[:25])
    res_df['text'] = res_df['text'].apply(lambda x: x[:25])

    all_levels = list(res_df["level"].unique())

    for level in all_levels:
        if len(level) > 5 and level[5] != '0':
            tab_text = "    " * int(level[5])
            res_df.loc[res_df["level"] == level, ["summary"]] = tab_text + \
                                                                res_df.loc[res_df["level"] == level, ["summary"]]["summary"]

    res_df = res_df[['id', 'summary', 'state', 'result', 'plan_cnt', 'pass_cnt',
                     'fail_cnt', 'othr_cnt', 'prcr_link', 'prcr_state',
                     'vm_level', 'plan_enddt', 'user', 'rel_name',
                     'env_picklist', 'hard_vname', 'soft_vname']]

    rename_dict = {'id': 'ID (PTC)', 'summary': 'Summary / Text',
                   'state': 'State',
                   'plan_cnt': 'Total Planned Count', 'pass_cnt': 'Total Pass Count',
                   'fail_cnt': 'Total Fail Count', 'othr_cnt': 'Total Other Count',
                   'prcr_link': 'PR/CR    ID --> link', 'prcr_state': 'PR/CR    State',
                   'user': 'Assigned User', 'vm_level': 'VM Level',
                   'plan_enddt': 'planned end date', 'rel_name': 'Release Name',
                   'env_picklist': 'test environment pick list',
                   'hard_vname': 'HW Version Name', 'soft_vname': 'SW Version Name'}
    res_df.rename(columns=rename_dict, inplace=True)

    out_flnm = os.path.join(os.path.dirname(os.path.abspath(__file__)),'test_report_overview' + '_' + test_plan_id + '.xlsx')
    res_df.to_excel(out_flnm, index=False)

    if half_flag == 1:
        return 1
    else:
        return 0


def find_count(test_plan_id):
    cnt_flnm = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   'count_' + str(test_plan_id) + ".csv")
    count_report_cmd = "im runreport --yes --substituteParams --user='" + \
                     username + "' --hostname='" + hostname + \
                     "'  --issues=" + str(test_plan_id) + ' --outputfile="' + \
                       cnt_flnm + '"  "Test Objective Count -- Input for Test Report Overview -- USED FOR CUSTOM SCRIPT"'

    run_report_out, run_err = execommand(count_report_cmd, '', '')
    if run_err == 1:
        logging.error("Error executing PTC report. Reason may be related to PTC server slowness.")
        logging.info("Try after sometime or Contact Script Developer")
        logging.info("Command output:" + ",".join(run_report_out))
        sys.exit("Error executing run report, Please try after some time or Contact developer.")

    cnt_data = pd.read_csv(cnt_flnm)
    test_obj_cnt = cnt_data['Test Objective Count'].tolist()[0]
    os.remove(cnt_flnm)

    tc_cnt_cmd = "im viewissue " + "--user='" + username + "' --hostname='" + \
                 hostname + "' " + test_plan_id

    tccnt_out, tccnt_err = execommand(tc_cnt_cmd, '', '')
    if tccnt_err == 1:
        logging.error("Error while fetching 'Total Planned Count' from Test Plan.")
        logging.info("Contact developer to fix the issue.")
        sys.exit("script execution not started.")

    for out in tccnt_out:
        if out.startswith('Total Planned Count'):
            tot_plan_tc_cnt = out[out.find(':')+1:]
            break

    logging.info("Test Objective count - " + str(test_obj_cnt))
    logging.info("Test Case Planned Count - " + tot_plan_tc_cnt)

    if int(test_obj_cnt) > 50 or ((int(test_obj_cnt)+int(tot_plan_tc_cnt)) > 700):
        count_status = "not okay"
    else:
        count_status = "okay"

    logging.info("Count check of the given test plan Completed, Status - " + count_status)
    return count_status


if __name__ == "__main__":
    start_time = datetime.now()

    root = Tk()
    my_interface = Interface(root)
    root.mainloop()
    root_item = my_interface.get_value()
    #root_item = '6655148'
    print("Test Plan Given")
    print(root_item)

    logging.basicConfig(filename=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              'tm_testreport_' + root_item + '_logfile' + '.log'),
        filemode='a', format='%(asctime)s %(levelname)s: %(message)s',
        datefmt='%m/%d/%Y %H:%M:%S', level=logging.DEBUG)
    logging.info("\n\n\nStarting Python Script .....")

    err_flnm = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "tm_testreport_" + root_item + "_errorfile" + '.txt')
    if os.path.exists(err_flnm):
        os.remove(err_flnm)
    err_file = open(err_flnm, "w")
    logging.info("Script started ...")
    logging.info("Test Plan/Sample/Phase ID Given :" + str(root_item))
    logging.info("Hostname: "+hostname+" , Username: "+username)

    count_status = find_count(root_item)
    #print(count_status)
    if count_status != "okay":
        root = Tk()
        my_interface = InterfaceCount(root)
        root.mainloop()
        sys.exit("script execution cannot be started.")
    print("Execution Started ...")

    runtime_in_secs = 2400
    half_flag = write_to_excel(root_item, time.time(), runtime_in_secs)

    logging.info("")
    logging.info("")
    logging.info("************************************************************")
    logging.info("************************************************************")
    #Only half of the data processed
    if half_flag == 1:
        root = Tk()
        half_done = InterfaceHalfdone(root)
        root.mainloop()
        logging.info("****** Processed only few items ******")
        logging.info("Run time limit exceeded")
    else:
        logging.info("Script execution completed successfully.")

    if username not in ['vikramm', 'adm_vikramm', 'sindhu']:
        os.remove(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               'ptcreport_'+root_item+'.html'))

    print("Process completed...")
    time.sleep(5)
    end_time = datetime.now()
    logging.info("Total Run Duration -"+str(end_time - start_time))
    print('Run Duration: {}'.format(end_time - start_time))

    subprocess.Popen([os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   'test_report_overview_' + root_item + '.xlsx')], shell=True)


