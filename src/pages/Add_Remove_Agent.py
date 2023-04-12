# Import necessary libraries 
import dash
from dash import Dash, html, Input, Output, ctx, dcc, State
import dash_bootstrap_components as dbc

import time
from datetime import datetime

import openpyxl
import threading
import  pandas as pd
import numpy as np
import pathlib


PATH = pathlib.Path(__file__).parent
DATA_PATH = PATH.joinpath("../").resolve()


# Import data from excel files
dt = pd.read_excel(DATA_PATH.joinpath("Data.xlsx"))
df = pd.read_excel(DATA_PATH.joinpath("Agents.xlsx"))

# Assign list of names
df['Name'] = df['Name'].str.upper()
name_list = df['Name']
first_name_list = df.loc[0, 'Name']

# Establishes timestamp
now = datetime.now()
timestamp_num = datetime.timestamp(now)
timestamp = datetime.fromtimestamp(timestamp_num)


# Define the page layout
layout = dbc.Container([
    

    dbc.Row([
        html.Center(html.H1("Add/Remove Agent", style = {'color' : '#00c3ff'})),
        html.Br(),
        html.Hr(),
        dbc.Col([
            html.P("Select the name of the Agent you want to Add", style = {'color' : '#00d257'}), 
            
            #Add agent
            html.Div(dcc.Input('', id='new-name-input', type='text')),
            dbc.Button('Submit', color="primary", id='add-agent-btn', n_clicks=0),
            html.Div(id='add-agent-output', children='Enter a Name and press submit', style = {'color' : '#ffffff'}),


        ]), 
        dbc.Col([
            html.P("Select the name of the Agent you want to Remove", style = {'color' : '#00d257'}), 


            #Remove agent
            html.Div(dcc.Dropdown(options=[{'label': name, 'value': index} for index, name in df['Name'].to_dict().items()], value=df.index[0], id='demo-dropdown', clearable=False)),
            html.Div(id='remove-agent-output', style = {'color' : '#ffffff'}),

            dcc.ConfirmDialogProvider(children=dbc.Button('Remove',), id='remove-agent-btn', message='Are you sure you want to delete this agent?'),

            html.Hr(style={"margin-bottom": "700px"}),

        ])
    ])
])



