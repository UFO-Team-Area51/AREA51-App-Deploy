# Import necessary libraries 
import dash
from dash import Dash, html, Input, Output, ctx, dcc, State
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc

import time
import threading
from datetime import datetime, timezone, date
import xlsxwriter
import openpyxl
import  pandas as pd
import numpy as np
import pytz

import plotly.express as px
import pathlib


##########################
#### Intial report run ###
##########################
PATH = pathlib.Path(__file__).parent
DATA_PATH = PATH.joinpath("../").resolve()


# Import data from excel files
df = pd.read_excel(DATA_PATH.joinpath("Agents.xlsx"))

add_mstr_value = 'MASTER'
df.loc[df.index.max() + 1, 'Name'] = add_mstr_value
name_list = df['Name']
first_name_list = df.loc[0, 'Name']

dt_mstr_mbm = pd.read_excel(DATA_PATH.joinpath("MASTER.xlsx"), sheet_name='MASTER_MBM_Worked')
dt_mstr_uet = pd.read_excel(DATA_PATH.joinpath("MASTER.xlsx"), sheet_name='MASTER_UET_Worked')

#Removes all rows that are not designated as assigning a ticket and resets index
dt_mstr_mbm = dt_mstr_mbm[dt_mstr_mbm["Action"].str.contains('Assigned Ticket')]
dt_mstr_mbm.reset_index(drop=True, inplace=True)

#Convert MBM timestamps to datetime objects and sets the index as timestamps
dt_mstr_mbm['Date_Time'] = pd.to_datetime(dt_mstr_mbm['Date_Time'])
dt_mstr_mbm.set_index('Date_Time', inplace=True)

# Sets up timestamp increments by day
daily_mbm_counts = dt_mstr_mbm.resample('D').count()

#Removes all rows that are not designated as assigning a ticket and resets index
dt_mstr_uet = dt_mstr_uet[dt_mstr_uet["Action"].str.contains('Assigned Ticket')]
dt_mstr_uet.reset_index(drop=True, inplace=True)

#Convert UET timestamps to datetime objects and sets the index as timestamps
dt_mstr_uet['Date_Time'] = pd.to_datetime(dt_mstr_uet['Date_Time'])
dt_mstr_uet.set_index('Date_Time', inplace=True)

# Sets up timestamp increments by day
daily_uet_counts = dt_mstr_uet.resample('D').count()


##########################################


#app = dash.Dash()

layout = dbc.Container([

    html.Div(children=[

    #Run report button
    html.Br(),
    html.Center(dbc.Button("Run Report", style={"width": "30%", 'textAlign': 'center'}, color="primary", id='run-report_btn', n_clicks=0),),
    
    html.Div(id='report-output', style = {'color' : '#00d257', 'textAlign': 'center', 'font-size' : '120%'}),

    html.Br(),
    dbc.Col(html.Hr(style={"borderWidth": "0.5vh", "width": "117%", "borderColor": "#00c3ff", "opacity": "unset",}), width={"size": 10, "offset": 0},), 

    html.Details([

        html.Summary('Click me to expand/collapse MBM Case Report', style = {'color' : '#00d257', 'textAlign': 'center', 'font-size' : '120%'}),
        dcc.DatePickerRange(id='mbm-date-picker-range',
        start_date=daily_mbm_counts.index.min().date(),
        end_date=daily_mbm_counts.index.max().date()),

        dcc.Graph(id='mbm-bar-chart'),
        
        html.Br(),
        html.Br(),

        #MBM Download
        html.Center(dcc.Dropdown(name_list, first_name_list,  id='download_mbm_dropdown', clearable=False, className="text-success", style={
            'textAlign': 'center', 
            'width': '55%', 
            'left': '10%',
            'transform': 'translateX(8%)',
            }
            ),
        ),
        html.Br(),
        html.Center(dbc.Button("Download MBM Excel", style={"width": "30%"}, color="success", outline=True, id='download_mbm_btn', n_clicks=0, ),),
        dcc.Download(id="download-mbm-dataframe-xlsx"),
                
    ]),

    html.Br(),
    dbc.Col(html.Hr(style={"borderWidth": "0.5vh", "width": "117%", "borderColor": "#00c3ff", "opacity": "unset",}), width={"size": 10, "offset": 0},),

    html.Details([
        
        html.Summary('Click me to expand/collapse UET TIcket Report', style = {'color' : '#00d257', 'textAlign': 'center', 'font-size' : '120%'}),
        dcc.DatePickerRange(id='uet-date-picker-range',
        start_date=daily_uet_counts.index.min().date(),
        end_date=daily_uet_counts.index.max().date()),

        dcc.Graph(id='uet-bar-chart'),

        html.Br(),
        html.Br(),

        html.Center(dcc.Dropdown(name_list, first_name_list,  id='download_uet_dropdown', clearable=False, className="text-success", style={
            'textAlign': 'center', 
            'width': '55%', 
            'left': '10%',
            'transform': 'translateX(8%)',
            }
            ),),
        html.Br(),
        html.Center(dbc.Button("Download UET Excel", style={"width": "30%"}, color="success", outline=True, id='download_uet_btn', n_clicks=0),),
        dcc.Download(id="download-uet-dataframe-xlsx"),
    ],
    ),
    html.Br(),
    html.Br(),
    html.Br(),
    ], style={'textAlign': 'center'} 
    ),
])

