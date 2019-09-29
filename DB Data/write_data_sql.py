import os
import urllib2
import pandas as pd
import pyodbc
from datetime import datetime
import sqlalchemy
from sqlalchemy import create_engine
import urllib
import mysql.connector
import collections
import sys
import numpy as np
import warnings

warnings.filterwarnings("ignore")

def read_url_df(url):
    response = urllib2.urlopen(url)
    with open('data_text.txt', 'w') as f:
        f.write(response.read())

    if os.stat('data_text.txt').st_size != 0:
        data_df = pd.read_csv('data_text.txt', sep="\t", header=None)
        return data_df
    else:
        return None


def read_both_url(url, all_data_flag=None):
    #all_data_url = "https://winssdcapp02.vcs.privad.net/imdsi?FILTER=AND(IN(%27project%27,%27/ABS_System/E8%27),IN(%27type%27,%27Release%27,%27Problem%20Report%27,%27Change_Request%27,%20%27Task%27))&OUTPUT=project,id,Type,alm_spawns,covers,effort_planned__h_,effort_remaining__h_,effort_spent__h_"
    all_data = read_url_df(url)

    if all_data is None:
        all_data = pd.DataFrame(columns=['project', 'id', 'type', 'spawns', 'covers', 'e_planned', 'e_remaining', 'e_spent'])
        return all_data

    all_data.columns = ['project', 'id', 'type', 'spawns', 'covers', 'e_planned', 'e_remaining', 'e_spent']

    if all_data_flag == 1:
        all_data.to_excel('all_data.xlsx', index=False)

    all_data['id'].apply(str)
    all_data['spawns'].apply(str)
    return all_data


def frame_filter(in_list):
    proj_filter = ""
    for id in in_list:
        proj_filter = proj_filter + "%27" + id + "%27,"
    proj_filter = proj_filter[:-1]

    return proj_filter


def traverse_release_data(in_data, all_data=None, first_time_flag=None,
                          e_planned_sum=None, e_remaining_sum=None,
                          e_spent_sum=None, rel_dtl_dict={}, handld_tasks=[], task_df=pd.DataFrame()):

    if first_time_flag is not None:
        all_release = in_data.loc[in_data['type'] == 'Release']
        #all_release = all_release.loc[all_release['id'] == 1098677]

        all_data = in_data

        e_planned_sum = 0
        e_remaining_sum = 0
        e_spent_sum = 0
        #task_df = pd.DataFrame()
    else:
        all_release = in_data

    dup_flag = 0
    tmp_res = {}
    tmp_res_dict = {}
    all_release = all_release.reset_index(drop=True)
    for index, row in all_release.iterrows():
        dup_flag = 0
        if first_time_flag is not None:
            print "Current Release ID-", row['id']
            e_planned_sum = 0
            e_remaining_sum = 0
            e_spent_sum = 0
            handld_tasks = []
            tmp_res = {}
        else:
            print "Current just ID-", row['id']
            #print e_planned_sum, e_remaining_sum, e_spent_sum
            #print handld_tasks

        r_covers = [] if str(row['covers']) == 'nan' else row['covers'].split(',')

        row['spawns'] = str(row['spawns'])
        try:
            r_spawns = [] if str(row['spawns']) == 'nan' else row['spawns'].split(',')
        except Exception as e:
            print "Error Here"
            print str(e)
            print "id-", row['id']
            print "covers-", row['covers']
            print "spawns-", row['spawns']
            print type(row['spawns'])
            sys.exit("check")

        r_relationships = r_spawns + r_covers
        print "Relationships for release ID-", row['id']
        print r_relationships

        other_proj_ids = np.setdiff1d(r_relationships, all_data['id'].tolist())
        if len(other_proj_ids) > 0:
            print "Other Project items"
            print other_proj_ids
            #sys.exit("check")

            proj_filter = frame_filter(other_proj_ids)

            other_ids_url = "https://winssdcapp02.vcs.privad.net/imdsi?FILTER=AND(IN(%27id%27," + \
                             proj_filter + \
                             "))&OUTPUT=project,id,Type,alm_spawns,covers,effort_planned__h_,effort_remaining__h_,effort_spent__h_"

            other_ids_data = read_both_url(other_ids_url)
            #print other_ids_data

            other_projs = other_ids_data['project'].unique().tolist()
            print "Other project names"
            print other_projs

            if other_projs:
                other_proj_filter = frame_filter(other_projs)
                other_proj_url = "https://winssdcapp02.vcs.privad.net/imdsi?FILTER=AND(IN(%27project%27," + \
                            other_proj_filter + \
                             "))&OUTPUT=project,id,Type,alm_spawns,covers,effort_planned__h_,effort_remaining__h_,effort_spent__h_"
                other_proj_data = read_both_url(other_proj_url)

                if not other_proj_data.empty:
                    all_data = pd.concat([all_data, other_proj_data], ignore_index=True)
                    all_data['spawns'].apply(str)


        rr_dtls = all_data[all_data['id'].isin(r_relationships)]
        rr_task_dtls = rr_dtls.loc[rr_dtls['type'] == 'Task']
        rr_nottask_dtls = rr_dtls.loc[rr_dtls['type'] != 'Task']

        if not rr_task_dtls.empty:

            if not handld_tasks:
                #print "Current values, if-", [rr_task_dtls['e_planned'].sum(), rr_task_dtls['e_remaining'].sum(), rr_task_dtls['e_spent'].sum()]
                rr_task_dtls = rr_task_dtls.drop_duplicates(subset='id', keep='first')
                handld_tasks = handld_tasks + rr_task_dtls['id'].tolist()
            else:
                #print handld_tasks
                #print "sum, before-", e_planned_sum, e_remaining_sum, e_spent_sum
                #print "Current values, else-", [rr_task_dtls['e_planned'].sum(), rr_task_dtls['e_remaining'].sum(), rr_task_dtls['e_spent'].sum()]
                rr_task_dtls = rr_task_dtls.drop_duplicates(subset='id', keep='first')
                #print "Current values, after drop-", [rr_task_dtls['e_planned'].sum(), rr_task_dtls['e_remaining'].sum(), rr_task_dtls['e_spent'].sum()]
                #print handld_tasks
                #print rr_task_dtls['id'].tolist()
                if handld_tasks:
                    rr_task_dtls = rr_task_dtls.loc[~rr_task_dtls['id'].isin(handld_tasks)]
                #print rr_task_dtls['id'].tolist()
                #print "Current values, after tasks drop-", [rr_task_dtls['e_planned'].sum(), rr_task_dtls['e_remaining'].sum(), rr_task_dtls['e_spent'].sum()]
                handld_tasks = handld_tasks + rr_task_dtls['id'].tolist()

            rr_task_dtls = rr_task_dtls.fillna(0)
            task_df = pd.concat([task_df, rr_task_dtls], ignore_index=True)

            #print "Current values-", [rr_task_dtls['e_planned'].sum(), rr_task_dtls['e_remaining'].sum(), rr_task_dtls['e_spent'].sum()]

            e_planned_sum = e_planned_sum + rr_task_dtls['e_planned'].sum()
            e_remaining_sum = e_remaining_sum + rr_task_dtls['e_remaining'].sum()
            e_spent_sum = e_spent_sum + rr_task_dtls['e_spent'].sum()

            #print "temp sum, after-", e_planned_sum, e_remaining_sum, e_spent_sum

        if not rr_nottask_dtls.empty:
            #print "before calling next"
            #print e_planned_sum, e_remaining_sum, e_spent_sum
            tmp_res_dict, temp_task_df = traverse_release_data(rr_nottask_dtls, all_data, None, e_planned_sum,
                                      e_remaining_sum, e_spent_sum, rel_dtl_dict, handld_tasks, task_df)

            task_df = temp_task_df

            #print "sum before-", e_planned_sum, e_remaining_sum, e_spent_sum

            if tmp_res_dict != {}:
                #print "before calling next"
                #print e_planned_sum, e_remaining_sum, e_spent_sum
                #e_planned_sum = e_planned_sum + tmp_res_dict.values()[0][0]
                #e_remaining_sum = e_remaining_sum + tmp_res_dict.values()[0][1]
                #e_spent_sum = e_spent_sum + tmp_res_dict.values()[0][2]

                e_planned_sum = tmp_res_dict.values()[0][0]
                e_remaining_sum = tmp_res_dict.values()[0][1]
                e_spent_sum = tmp_res_dict.values()[0][2]

            #e_planned_sum = e_planned_sum + tmp_res_dict.values()[0][0]
            #e_remaining_sum = e_remaining_sum + tmp_res_dict.values()[0][1]
            #e_spent_sum = e_spent_sum + tmp_res_dict.values()[0][2]

            #print "sum after-", e_planned_sum, e_remaining_sum, e_spent_sum

        #print "Type-", row['type']
        #if row['type'] == 'Release':
        if first_time_flag is not None:

            '''
            if tmp_res_dict != {}:
                e_planned_sum = e_planned_sum + tmp_res_dict.values()[0][0]
                e_remaining_sum = e_remaining_sum + tmp_res_dict.values()[0][1]
                e_spent_sum = e_spent_sum + tmp_res_dict.values()[0][2]
            '''

            #print "adding values for Release ID-", row['id'], " : ", [e_planned_sum, e_remaining_sum, e_spent_sum]
            rel_dtl_dict[str(row['id'])] = [e_planned_sum, e_remaining_sum, e_spent_sum]

        else:
            tmp_res[1] = [e_planned_sum, e_remaining_sum, e_spent_sum]
            #print "temp res"
            #print tmp_res

            if index == len(all_release)-1:
                rel_dtl_dict = tmp_res
                #return e_planned_sum, e_remaining_sum, e_spent_sum
            else:
                continue

    print "last return"
    print rel_dtl_dict
    return rel_dtl_dict, task_df


def convert_dict_to_df(in_dict):
    dates_url = "https://winssdcapp02.vcs.privad.net/imdsi?OUTPUT=dates"
    dates_df = read_url_df(dates_url)
    #print dates_df
    all_dates = dates_df[0].tolist()
    latest_date = str(all_dates[-1])
    print latest_date
    ptc_data_rundate = latest_date[:4] + '-' + latest_date[4:6] + '-' + latest_date[6:]
    print ptc_data_rundate

    rel_eff_list = [[key]+value for key, value in in_dict.iteritems()]
    rel_eff_df = pd.DataFrame(rel_eff_list, columns=['release_id',
                'effort_planned', 'effort_rem', 'effort_spent'])
    rel_eff_df['effort_planned'] = rel_eff_df['effort_planned'].astype('int')
    rel_eff_df['effort_rem'] = rel_eff_df['effort_rem'].astype('int')
    rel_eff_df['effort_spent'] = rel_eff_df['effort_spent'].astype('int')
    rel_eff_df['effort_ideal'] = rel_eff_df['effort_rem'] + rel_eff_df['effort_spent']

    rel_eff_df['ptc_data_rundate'] = ptc_data_rundate
    rel_eff_df['our_rundate'] = datetime.today().strftime('%Y-%m-%d')
    rel_eff_df['project_name'] = '/ABS_System'
    rel_eff_df = rel_eff_df[['ptc_data_rundate', 'our_rundate', 'project_name',
                             'release_id', 'effort_planned', 'effort_spent',
                             'effort_rem', 'effort_ideal']]
    return rel_eff_df


def write_to_table(rel_eff_df):
    quoted = urllib.pathname2url("DRIVER={SQL Server Native Client 11.0};SERVER=WDEGSSQL3;DATABASE=MP_PE_REL;uid=mission_pe;pwd=mission@sql3")
    engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))

    coltype = {'ptc_data_rundate': sqlalchemy.DATE, 'our_rundate': sqlalchemy.DATE, 'project_name': sqlalchemy.NVARCHAR, 'release_id': sqlalchemy.NVARCHAR}
    rel_eff_df.to_sql('ptc_release_efforts', con=engine, if_exists='append', index=False, dtype=coltype)


if __name__ == "__main__":
    start_time = datetime.now()

    all_data_url = "https://winssdcapp02.vcs.privad.net/imdsi?FILTER=AND(IN(%27project%27,%27/ABS_System/E8%27),IN(%27type%27,%27Release%27,%27Problem%20Report%27,%27Change_Request%27,%20%27Task%27))&OUTPUT=project,id,Type,alm_spawns,covers,effort_planned__h_,effort_remaining__h_,effort_spent__h_"
    all_data = read_both_url(all_data_url)
    rel_dtl_dict, task_df = traverse_release_data(all_data, all_data, 1)
    task_df.to_excel('task_df.xlsx', index=False)

    print "*************************************************"
    print len(rel_dtl_dict)
    print rel_dtl_dict

    #rel_dtl_dict = {'1477844': [757.0, 0.0, 475.0], '1161763': [910.0, 0.0, 910.0], '1829284': [656.0, 324.0, 168.0], '2626221': [220.0, 1.0, 198.0], '1710067': [3055.0, 834.0, 1176.0], '3006885': [13.0, 0.0, 17.0], '1872781': [178.0, 0.0, 183.0], '1468017': [2291.0, 6.0, 2226.0], '686190': [15488.0, 579.0, 12188.0], '2264035': [0, 0, 0], '1820233': [1001.0, 130.0, 634.0], '1071883': [1028.0, 0.0, 564.0], '1820230': [5335.0, 1574.0, 3254.0], '2119531': [0, 0, 0], '513652': [9239.0, 20.0, 6428.0], '1334044': [1757.0, 431.0, 852.0], '2102492': [176.0, 4.0, 171.0], '1532716': [271.0, 0.0, 201.0], '2417781': [14.0, 0.0, 20.0], '1334043': [2605.0, 217.0, 1439.0], '1412245': [36.0, 3.0, 25.0], '979449': [0, 0, 0], '1098675': [3179.0, 0.0, 2765.0], '1098677': [2303.0, 17.0, 2025.0], '3505993': [151.0, 0.0, 111.0], '3848071': [2466.0, 47.0, 2788.0], '603298': [4863.0, 2.0, 5372.0], '3644271': [318.0, 0.0, 331.0], '3164572': [65.0, 0.0, 58.0], '1281444': [1766.0, 78.0, 1587.0], '3507421': [1595.0, 11.0, 1468.0], '4672472': [747.0, 7.0, 864.0], '2845579': [48.0, 2.0, 53.0], '3359863': [135.0, 107.0, 61.0], '1461284': [26.0, 0.0, 44.0], '1652347': [1001.0, 130.0, 634.0], '1652348': [1141.0, 635.0, 490.0], '1468022': [2781.0, 148.0, 2261.0], '3191660': [39.0, 0.0, 75.0], '644021': [2513.0, 0.0, 1160.0], '3250472': [245.0, 5.0, 304.0], '3206137': [1640.0, 44.0, 1559.0], '4499269': [198.0, 20.0, 175.0], '1038927': [0, 0, 0], '1709732': [1376.0, 84.0, 1435.0], '1547396': [0, 0, 0], '1547397': [0, 0, 0], '1827222': [1458.0, 26.0, 1070.0], '1826513': [1458.0, 26.0, 1070.0], '2319728': [5.0, 3.0, 2.0], '2721678': [801.0, 24.0, 717.0], '3290700': [38.0, 0.0, 40.0], '605134': [11811.0, 0.0, 9402.0], '1609172': [3178.0, 794.0, 2130.0], '1313038': [1977.0, 0.0, 1673.0], '1313039': [2524.0, 0.0, 2411.0], '2267902': [32.0, 0.0, 32.0], '1445334': [4782.0, 219.0, 3227.0], '1445238': [7175.0, 237.0, 5533.0], '2903689': [1211.0, 49.0, 1220.0], '2018305': [201.0, 10.0, 227.0], '2361451': [48.0, 0.0, 55.0], '1477432': [638.0, 0.0, 350.0], '1211901': [360.0, 240.0, 0.0], '2112456': [48.0, 20.0, 29.0], '1673074': [44.0, 0.0, 64.0], '2234602': [0, 0, 0], '1445249': [8899.0, 592.0, 6427.0], '3239791': [0, 0, 0], '2234604': [24.0, 0.0, 26.0], '1820236': [1141.0, 635.0, 490.0], '2590427': [15.0, 0.0, 23.0], '4757910': [8.0, 0.0, 8.0], '3854069': [139.0, 0.0, 149.0], '1322747': [44.0, 0.0, 52.0], '2788677': [104.0, 2.0, 116.0], '1826943': [0, 0, 0], '3021791': [0, 0, 0], '3021793': [0, 0, 0], '2822165': [504.0, 8.0, 594.0], '3021795': [0, 0, 0], '1336522': [9556.0, 2620.0, 5679.0], '2243069': [421.0, 87.0, 195.0], '1071935': [616.0, 21.0, 74.0], '1071932': [2410.0, 2.0, 1192.0], '2337193': [307.0, 6.0, 185.0], '1692070': [0, 0, 0], '3379965': [230.0, 0.0, 244.0], '1445251': [36344.0, 1627.0, 27375.0]}
    rel_eff_df = convert_dict_to_df(rel_dtl_dict)
    write_to_table(rel_eff_df)

    end_time = datetime.now()
    #logging.info("Total Run Duration -"+str(end_time - start_time))
    print('Run Duration: {}'.format(str(end_time - start_time)))

