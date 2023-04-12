# Import necessary libraries
import dash
from dash import Dash, html, Input, Output, ctx, dcc, State
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import time
from time import sleep
import threading
from datetime import datetime, timezone
import xlsxwriter
import openpyxl
import pandas as pd
import numpy as np
import pytz
import pathlib
import shareplum



#Reads Agent file in order to get list of agent names
df = pd.read_excel(r'Agents.xlsx')
agent_list = df['Name']


#Sets up blank lists to be used later
dfs_mbm = []
dfs_uet = []

#Cycles through list of agents and opens each excel files
for i in agent_list:
    agent_list == i
    dict_df = pd.read_excel(i+'.xlsx', sheet_name=[i+'_MBM_Worked', i+'_UET_Worked'])

    dfw_m = dict_df.get(i+'_MBM_Worked')
    dfw_m.rename(columns ={i+'_MBM_Worked': "Date_Time"}, inplace = True)
    dfw_m['Name'] = i
    dfs_mbm.append(dfw_m)

    dfw_u = dict_df.get(i+'_UET_Worked')
    dfw_u.rename(columns ={i+'_UET_Worked': "Date_Time"}, inplace = True)
    dfw_u['Name'] = i
    dfs_uet.append(dfw_u)


#Raw combined files with each spreadsheet (MBM/UET, with both columns)
dft_mbm = pd.concat(dfs_mbm, ignore_index=True)
dft_uet = pd.concat(dfs_uet, ignore_index=True)

# MBM sort and organize by date
dft_mbm['Date_Time'] = pd.to_datetime(dft_mbm['Date_Time'])
dft_mbm.sort_values(by='Date_Time', ascending=True, inplace=True)
dft_mbm.reset_index(drop=True, inplace=True)
dft_mbm_master = pd.DataFrame(dft_mbm)

# UET sort and organize by date
dft_uet['Date_Time'] = pd.to_datetime(dft_uet['Date_Time'])
dft_uet.sort_values(by='Date_Time', ascending=True, inplace=True)
dft_uet.reset_index(drop=True, inplace=True)
dft_uet_master = pd.DataFrame(dft_uet)

# Writes data to excel file
writer = pd.ExcelWriter('MASTER.xlsx', engine='xlsxwriter')
with pd.ExcelWriter('MASTER.xlsx') as writer:
    dft_mbm_master.to_excel(writer, sheet_name='MASTER_MBM_Worked', index=False)
    dft_uet_master.to_excel(writer, sheet_name='MASTER_UET_Worked', index=False)

#Counts undo counts and returns value
mbm_undo_count = dft_mbm_master['Action'].str.count('Undone').sum()
uet_undo_count = dft_uet_master['Action'].str.count('Undone').sum()
report_details = 'Finsihed report, there are {} counts of MBM Cases being undone and {} counts of UET Tickets being undone.'.format(mbm_undo_count, uet_undo_count)

#Removes all rows that are not designated as assigning a ticket and resets index
dft_mbm_master = dft_mbm_master[dft_mbm_master["Action"].str.contains('Assigned Ticket')]
dft_mbm_master.reset_index(drop=True, inplace=True)

#Convert MBM timestamps to datetime objects and sets the index as timestamps
dft_mbm_master['Date_Time'] = pd.to_datetime(dft_mbm_master['Date_Time'])
dft_mbm_master.set_index('Date_Time', inplace=True)

# Sets up timestamp increments by day
daily_mbm_counts = dft_mbm_master.resample('D').count()

#Removes all rows that are not designated as assigning a ticket and resets index
dft_uet_master = dft_uet_master[dft_uet_master["Action"].str.contains('Assigned Ticket')]
dft_uet_master.reset_index(drop=True, inplace=True)

#Convert UET timestamps to datetime objects and sets the index as timestamps
dft_uet_master['Date_Time'] = pd.to_datetime(dft_uet_master['Date_Time'])
dft_uet_master.set_index('Date_Time', inplace=True)