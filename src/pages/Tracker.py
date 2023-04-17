from dash import html, dcc
import dash_bootstrap_components as dbc
import pandas as pd
# noinspection PyCompatibility
import pathlib

PATH = pathlib.Path(__file__).parent
DATA_PATH = PATH.joinpath("../").resolve()

# Import data from excel files
dt = pd.read_excel(DATA_PATH.joinpath("Data.xlsx"))
df = pd.read_excel(DATA_PATH.joinpath("Agents.xlsx"))

df['Name'] = df['Name'].str.upper()
name_list = df['Name']

#working_list = df.loc[df['Working'] == True, 'Name'].to_string()

first_name_list = df.loc[0, 'Name']

# Define the Tracker page layout
# noinspection LongLine
layout = dbc.Container([

    dbc.Row([
        html.Center(
            html.H1("Ticket Tracker", style={'color': '#00c3ff'})
        ),
        html.Br(),
        dbc.Row(
            dbc.Col(
                html.Hr(
                    style={"borderWidth": "0.5vh", "width": "121%", "borderColor": "#00d257", "opacity": "unset", }),
                width={"size": 10, "offset": 0},
            ),
        ),

        dbc.Col([
            # Ticket assigner
            html.P(" - Ticket Assigner - ",
                   style={'color': '#00c3ff', "font-weight": "bold", 'textAlign': 'center', 'font-size': '120%'}),
            dbc.Row([
                dcc.Dropdown(name_list, first_name_list, id='assigner_dropdown', clearable=False,
                             className="text-success", style={'textAlign': 'center', 'width': '60%'}),
                dbc.Button("Set Assigner", style={"width": "30%"}, color="primary", id='assigner-btn', n_clicks=0),
            ], justify="center"),
            html.Div(id='assigner_output', style={'color': '#00d257', 'textAlign': 'center', 'font-size': '120%'}),
            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.5vh", "width": "120%", "borderColor": "#00d257",
                                   "opacity": "unset", }), width={"size": 10, "offset": 0},
                ),
            ),

            # Auto-Assign MBM Ticket Button
            html.Div('-MBM Assign-',
                     style={'color': '#00c3ff', "font-weight": "bold", 'textAlign': 'center', 'font-size': '120%'}),
            html.Div(id='mbm-output-container', style={'textAlign': 'center', 'color': '#ffffff'}),
            html.Br(),
            dbc.Row(
                dbc.Button("Auto-Assign MBM Case", style={"width": "30%"}, color="primary", id='ambm-btn', n_clicks=0),
                justify="center"),

            # Manually assign MBM
            html.Br(),
            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.2vh", "width": "120%", "borderColor": "#00c3ff",
                                   "opacity": "unset", }), width={"size": 10, "offset": 0},
                ),
            ),
            dbc.Row([
                dbc.Button("Manually Assign MBM case", style={"width": "40%"}, color="success", id='mmbm-btn',
                           n_clicks=0),
                dcc.Dropdown(name_list, first_name_list, id='mbm_dropdown', clearable=False, className="text-success",
                             style={'textAlign': 'center', 'width': '60%'}),
            ], align="center", justify="center"),

            # Auto-Assign UET Ticket Button
            html.Br(),
            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.5vh", "width": "120%", "borderColor": "#00d257",
                                   "opacity": "unset", }), width={"size": 10, "offset": 0},
                ),
            ),
            html.Div('-UET Assign-',
                     style={'color': '#00c3ff', "font-weight": "bold", 'textAlign': 'center', 'font-size': '120%'}),
            html.Div(id='uet-output-container', style={'textAlign': 'center', 'color': '#ffffff'}),
            html.Br(),
            dbc.Row(
                dbc.Button("Auto-Assign UET Ticket", style={"width": "30%"}, color="primary", id='auet-btn',
                           n_clicks=0), justify="center"),

            # Manually assign UET
            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.2vh", "width": "120%", "borderColor": "#00c3ff",
                                   "opacity": "unset", }), width={"size": 10, "offset": 0}, ), ),
            dbc.Row([
                dbc.Button("Manually Assign UET ticket", style={"width": "40%"}, color="success", id='muet-btn',
                           n_clicks=0),
                dcc.Dropdown(name_list, first_name_list, id='uet_dropdown', clearable=False, className="text-success",
                             style={'textAlign': 'center', 'width': '60%'}),
            ], align="center", justify="center"),

        ]),

        dbc.Col([

            # Working Agents List
            html.P(" - Agents Working Today - ",
                   style={'color': '#00c3ff', "font-weight": "bold", 'textAlign': 'center', 'font-size': '120%'}),

            dbc.Row([
                dcc.Dropdown(options=name_list, id='work-list', multi=True, persistence=True,
                             persistence_type='session'),
                dbc.Button("Set Working List", style={"width": "30%"}, color="primary", id='working-btn', n_clicks=0),
            ], justify="center"),

            html.Br(),
            html.Div(id='mdd-output-container', style={'textAlign': 'center', 'color': '#ffffff'}),

            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.5vh", "width": "117%", "borderColor": "#00d257",
                                   "opacity": "unset", }), width={"size": 10, "offset": 0}, ), ),

            # MBM cases display
            html.P("- MBM Information -",
                   style={'color': '#00c3ff', 'textAlign': 'center', "font-weight": "bold", 'font-size': '120%'}),
            html.Tr([
                html.Td(children='Previous MBM Case was assigned to: ', style={'color': '#ffffff'}),
                html.Td(id='mbm-prev-assignee-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('Current MBM Case is assigned to: ', style={'color': '#ffffff'}),
                html.Td(id='mbm-assignee-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('Next MBM Case will be assigned to: ', style={'color': '#ffffff'}),
                html.Td(id='mbm-next-assignee-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('MBM Cases assigned today = ', style={'color': '#ffffff'}),
                html.Td(id='mbm-day-count-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('Total MBM Cases worked = ', style={'color': '#ffffff'}),
                html.Td(id='mbm-count-output', style={'color': '#ffffff'})
            ]),
            html.Br(),

            # MBM Undo button
            dcc.ConfirmDialogProvider(children=dbc.Button('Undo Last MBM Case Assignment', outline=True, color="danger",
                                                          style={"width": "40%"}), id='undo-mbm-btn',
                                      message='Are you sure you want to undo the last assigment?'
                                      ),
            html.Br(),
            dbc.Row(
                dbc.Col(
                    html.Hr(style={"borderWidth": "0.5vh", "width": "117%", "borderColor": "#00d257",
                                   "opacity": "unset", }), width={"size": 10, "offset": 0},
                ),
            ),

            # UET tickets display
            html.P("- UET Information -",
                   style={'color': '#00c3ff', 'textAlign': 'center', "font-weight": "bold", 'font-size': '120%'}
                   ),
            html.Tr([
                html.Td('Previous UET Ticket was assigned to: ', style={'color': '#ffffff'}),
                html.Td(id='uet-prev-assignee-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('Current UET Ticket is assigned to: ', style={'color': '#ffffff'}),
                html.Td(id='uet-assignee-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('Next UET Ticket will be assigned to: ', style={'color': '#ffffff'}),
                html.Td(id='uet-next-assignee-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('UET Tickets assigned today = ', style={'color': '#ffffff'}),
                html.Td(id='uet-day-count-output', style={'color': '#ffffff'})
            ]),
            html.Tr([
                html.Td('Total UET Tickets worked = ', style={'color': '#ffffff'}),
                html.Td(id='uet-count-output', style={'color': '#ffffff'})
            ]),
            html.Br(),

            # UET Undo button
            dcc.ConfirmDialogProvider(
                children=dbc.Button('Undo Last UET Ticket Assignment', outline=True, color="danger",
                                    style={"margin-bottom": "225px", "width": "40%"}),
                id='undo-uet-btn', message='Are you sure you want to undo the last assigment?'
            ),
        ])
    ])
])
