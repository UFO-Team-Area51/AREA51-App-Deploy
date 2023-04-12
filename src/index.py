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

# Connect to main app.py file
from app import app
from app import server

# Connect to your app pages
from pages import Tracker, Add_Remove_Agent, Analytics

# Connect the navbar to the index
from components import navbar

# define the navbar
nav = navbar.Navbar()


# Define the index page layout
app.layout = dbc.Container(html.Div(
    children=[

    # Core Homepage
    dcc.Location(id='url', refresh=True),
    nav, 
    html.Div(id='page-content', children=[], style={'backgroundColor':'#4c4c4c', 'opacity' : '0.95'}), 
    
    # dcc.Store inside the user's current browser session
    dcc.Store(id='store-data-mbm', data=[], storage_type='memory'), # 'local' or 'session'
    dcc.Store(id='store-data-uet', data=[], storage_type='memory'), # 'local' or 'session'

]),
    style={
    'background-image': 'url("/assets/UFOtruePNG3.png")',
    'background-repeat': 'no-repeat',
    'background-position': 'center center',
    'background-size': '850px 800px'
    },
    fluid=True,
    className="stylesheets"
)


@app.callback(Output('page-content', 'children'),
              [Input('url', 'pathname')])
def display_page(pathname):
    if pathname == '/Tracker':
        return Tracker.layout
    if pathname == '/Add_Remove_Agent':
        return Add_Remove_Agent.layout
    if pathname == '/Analytics':
        return Analytics.layout

    else:
        return html.Div("Please choose a link", style = {'color' : '#ffffff'}), html.Div(style={"margin-bottom": "875px"})



##########################
#### Add/Remove Agent ####
##########################


@app.callback(
    Output('add-agent-output', 'children'),
    Output('remove-agent-output', 'children'),
    Output('demo-dropdown', 'options'),

    
    
    Input('add-agent-btn', 'n_clicks'),
    Input('remove-agent-btn', 'submit_n_clicks'),
    Input('demo-dropdown', 'value'),
    State('demo-dropdown', 'options'),
    State('new-name-input', 'value'),
    prevent_initial_call=False
)

#######################
#### Add New Agent ####
#######################

def update_add_remove_agent(n_clicks1, n_clicks2, selected_name_value, options, name_input):
    
    df = pd.read_excel(r'Agents.xlsx')
    
    options = df['Name']
    
    RmvAgentMsg = 'Select name of the agent who you want to remove from the list of agents'
    AddAgentMsg = 'Select name of the agent who you want to add to the list of agents'

    ctx = dash.callback_context
    triggered_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if triggered_id == "add-agent-btn":

        #checks if value is blank
        if name_input.isalpha() == False:
            AddAgentMsg = 'No spaces, numbers or special characters are allowed'

        else:
            value_name = name_input
            selected_name = value_name.upper()

            # Creates seperate list to compare if duplicates exist.
            df = pd.read_excel(r'Agents.xlsx')
            dup_check = True
            name_list = df['Name']
            name_list.loc[len(name_list)] = selected_name
        
            #If duplicates exist on seperate list, error message generated. Else, add name to main list.     
            if len(name_list) != len(set(name_list)):
                AddAgentMsg = 'Duplicate name, please enter a different name!'
                dup_check = True
            else:
                dup_check = False


            #Checks if duplicates status is false, then adds name to DF and DT dataframes.    
            if dup_check == False:

                AddAgentMsg = 'Adding {} to the list of agents!'.format(selected_name)
                
                add_name = [selected_name, 0, 1, 0, 1, False]
                df.loc[len(df)] = add_name

                now = datetime.now()
                timestamp_num = datetime.timestamp(now)
                time_now = datetime.fromtimestamp(timestamp_num)
                timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')
        
                writer = pd.ExcelWriter(selected_name+'.xlsx', engine='xlsxwriter')
                with pd.ExcelWriter(selected_name+'.xlsx') as writer:

                    dict = {selected_name+'_MBM_Worked': [timestamp], "Action": ['Agent Created']}
                    da_m = pd.DataFrame(dict)
                    dict = {selected_name+'_UET_Worked': [timestamp], "Action": ['Agent Created']}
                    da_u = pd.DataFrame(dict)
                        
                    da_m.to_excel(writer, sheet_name=selected_name+'_MBM_Worked', index=False)
                    da_u.to_excel(writer, sheet_name=selected_name+'_UET_Worked', index=False)

                df['Name'] = df['Name'].str.upper()
                name_list = df['Name']

                df.to_excel('Agents.xlsx', index = False)
                
                options = df['Name']
                    
            else:
                pass

######################
#### Remove Agent ####                                                         
######################

    if triggered_id == 'remove-agent-btn':
        
        if not selected_name_value:
            
            RmvAgentMsg = 'No agent selected'
        else:
            
            df = pd.read_excel(r'Agents.xlsx')
            agent_count = df.shape[0]
            
            if agent_count == 1:
                RmvAgentMsg = 'You need at least one agent on the list'
        
            else:

                RmvAgentMsg = 'You have removed {} from the list of agents'.format(selected_name_value)
            
                df = df[df["Name"].str.contains(selected_name_value) == False]

                df.to_excel('Agents.xlsx', index = False) # Directory needs to be updated
            
    options = df['Name']

    return AddAgentMsg, RmvAgentMsg, options


#################################
##### Set Ticket Assigner  ######
#################################


@app.callback(
    Output('assigner_output', 'children'),
    Output('assigner_dropdown', 'options'),
    
    Input('assigner-btn', 'n_clicks'),
    State('assigner_dropdown', 'value'),
    State('assigner_dropdown', 'options'),
    prevent_initial_call=False
)

def set_assigner(button, value, options):

    sleep(0.3)

    dt = pd.read_excel(r'Data.xlsx')
    df = pd.read_excel(r'Agents.xlsx')
    assigner = dt.at[0, 'Assigner']

    triggered_id = ctx.triggered_id

    if triggered_id == 'assigner-btn':
        if value == 0:
            message = 'You need to select and angent first'
        
        else:
            dt['Assigner'] = value
            assigner = dt.at[0, 'Assigner']

            dt.to_excel('Data.xlsx', index = False)
    else:
        pass

    message = '{} is set as the ticket assigner today'.format(assigner)
    
    options = df['Name']

    return message, options


#################################
##### Change Working State ######
#################################

@app.callback(
    Output('store-data-mbm', 'data'),
    Output('store-data-uet', 'data'),
    Output('mdd-output-container', 'children'),
    Output('work-list', 'options'),

    Input('working-btn', 'n_clicks'),
    Input('work-list', 'value'),
    State('work-list', 'options'),
    prevent_initial_call=False   
    )

def update_working(button, value1, options):

    
    df = pd.read_excel(r'Agents.xlsx')
    dt = pd.read_excel(r'Data.xlsx')

    now = datetime.now()
    timestamp_num = datetime.timestamp(now)
    time_now = datetime.fromtimestamp(timestamp_num)

    # Sets up today's date and yesterdays data
    day = time_now.strftime("%d")
    today = day
    yesterday = dt.loc[0,'Date']
    dt.loc[0, 'Date'] = today

    mbm_next_assignee = 'push Set Working List button to get next agent'
    uet_next_assignee = 'push Set Working List button to get next agent'

    df4 = pd.DataFrame(df.loc[df['Working'] == True, 'Name'])
    worklist_name = df4['Name'].str.split()
    msg2 = 'List of agents working: ' + worklist_name.to_string(index=False)

    # checks if time is different day, if true, resets tickets worked to zero
    if int(today) != yesterday:
        
        df['MBM_Worked'] = 0
        df['UET_Worked'] = 0
        dt['Assigner'] = 'No One'

    else:
        pass


    triggered_id = ctx.triggered_id

    if triggered_id == 'working-btn':
        
        #Checks if there are any workers selected
        if not value1:
        
            df['Working'] = False
        
            msg2 = "There aren't any agents selected, please select an Agent first!"
        
        else:

            # Updates selected agent values and updates working list. 
            list_value = value1
            df_value1 = pd.Series(list_value)
            working_value = df['Name'].isin(df_value1)
            df['Working'] = working_value
                
            #Finds agent with highest selection value and is working
            work_list = df[df['Working'] == True]
            
            df2 = pd.DataFrame(work_list)
            max_work_index = df2['MBM_Selected'].idxmax()
            max_agent_name_mbm = df.at[max_work_index, 'Name']
            dt.at[0, 'Next_Agent_MBM'] = max_agent_name_mbm  

            mbm_next_assignee = max_agent_name_mbm
        

            df3 = pd.DataFrame(work_list)
            max_work_index = df3['UET_Selected'].idxmax()
            max_agent_name_uet = df.at[max_work_index, 'Name']
            dt.at[0, 'Next_Agent_UET'] = max_agent_name_uet

            uet_next_assignee = max_agent_name_uet

            df4 = pd.DataFrame(df.loc[df['Working'] == True, 'Name'])
            worklist_name = df4['Name'].str.split()
            msg2 = 'List of agents working: ' + worklist_name.to_string(index=False)

    else:
        pass

    options = df['Name']

    df.to_excel('Agents.xlsx', index = False) # Directory needs to be updated
    dt.to_excel('Data.xlsx', index = False) # Directory needs to be updated

    return mbm_next_assignee, uet_next_assignee, msg2, options



##################################################
### Manual/Automatic Assign & Undo MBM Cases  ####
##################################################


@app.callback(
    #Auto/Mnl MBM assign
    Output('mbm-output-container', 'children'),
    Output('mbm-count-output', 'children'),
    Output('mbm-assignee-output', 'children'),
    Output('mbm-day-count-output', 'children'),
    Output('mbm-prev-assignee-output', 'children'),
    Output('mbm-next-assignee-output', 'children'),
    Output('mbm_dropdown', 'options'),
    
    Input('undo-mbm-btn', 'submit_n_clicks'),
    Input('ambm-btn', 'n_clicks'),
    Input('mmbm-btn', 'n_clicks'),
    State('mbm_dropdown', 'options'),
    State("mbm_dropdown", "value"),
    Input('store-data-mbm', 'data'),


    prevent_initial_call=False
)

#Manual assign MBM ticket and setups button call for Automatic assign MBM case 
def update_mbm(submit_n_clicks, button2, button3, options, value, data):

    sleep(0.5)

    dt = pd.read_excel(r'Data.xlsx')
    df = pd.read_excel(r'Agents.xlsx')

    message1 = "Either manually or automatically assign agent a ticket"
    
    min_work_index = df['MBM_Selected'].idxmin()
    min_agent_name_mbm = df.at[min_work_index, 'Name']
    assignee = '{} at {}'.format(min_agent_name_mbm, dt.at[0, 'MBM_Time'])

    #Gets sum of total tickets worked for the day
    mbm_day_output = df['MBM_Worked'].sum()

    mbm_total_output = dt['Total_MBM_Cases']
    
    prev_assignee = dt.at[0, 'Prev_Agent_MBM']
    next_assignee = data

    #Sets trigger ID
    triggered_id = ctx.triggered_id
    
    if triggered_id == 'mmbm-btn':
         
        #Reads/creates dataframes from Excel files
        df = pd.read_excel(r'Agents.xlsx')
        dt = pd.read_excel(r'Data.xlsx')
        
        #Gets agent name from dropdown and finds index value
        selected_name = value                                                      # df['Name'].str.capitalize() )
        mbm_selected_index = df[df['Name']==selected_name].index.values

        #Gets sum of total tickets worked for the day
        mbm_day_output = df['MBM_Worked'].sum()

        #Checks if selected agent is working today then returns True/False value
        working_true = df.at[int(mbm_selected_index), 'Working']

        #Checks if agent is working today value is False
        if working_true == False:

            message1 = "Are you sure you selected the right agent? This agent isn't set to work today!"
            assignee = "Agent not assigned"

            #Finds agent who has lowest selection value (last agent who worked case or a new agent)
            prev_assignee = dt.at[0, 'Prev_Agent_MBM']
            next_assignee = dt.at[0, 'Next_Agent_MBM']

    
        else:

            #Gets the time
            now = datetime.now()
            timestamp_num = datetime.timestamp(now)
            time_now = datetime.fromtimestamp(timestamp_num)
            timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')

            time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')
            
            dt['MBM_Time'] = time_ez
            time_atm = dt.at[0, 'MBM_Time']

            #Finds agent who has lowest selection value (last agent who worked case or a new agent)
            min_work_index = df['MBM_Selected'].idxmin()
            min_agent_name_mbm = df.at[min_work_index, 'Name']
            prev_assignee = min_agent_name_mbm
            dt.at[0, 'Prev_Agent_MBM'] = prev_assignee

            #Adds +1 ticket worked to selected agent
            df.at[int(mbm_selected_index), 'MBM_Worked'] += 1
            
            #Add +1 to selection count value to all agents
            df['MBM_Selected'] += 1

            #Resets selection value count of selected agent to zero
            df.at[int(mbm_selected_index), 'MBM_Selected'] = 0

            #Add +1 to total worked tickets
            dt['Total_MBM_Cases'] += 1

            #Reads personal agent Excel files            
            da_m = pd.read_excel(selected_name+'.xlsx', sheet_name=selected_name+'_MBM_Worked') 
            da_u = pd.read_excel(selected_name+'.xlsx', sheet_name=selected_name+'_UET_Worked')  

            #Makes copy of agent data and adds new data
            da_m1 = pd.DataFrame(da_m[[selected_name+'_MBM_Worked', 'Action']])
            da_m2 = pd.DataFrame({selected_name+'_MBM_Worked': [timestamp], "Action": ['Manually Assigned Ticket']})
            da_m = pd.concat([da_m1, da_m2])

            #writes data to excel file
            writer = pd.ExcelWriter(selected_name+'.xlsx', engine='xlsxwriter')  
            with pd.ExcelWriter(selected_name+'.xlsx') as writer:
                da_m.to_excel(writer, sheet_name=selected_name+'_MBM_Worked', index=False)
                da_u.to_excel(writer, sheet_name=selected_name+'_UET_Worked', index=False)
            
            #Assigns value of total tickets worked
            mbm_total_output = dt['Total_MBM_Cases']
            
            #Gets sum of total tickets worked for the day
            mbm_day_output = df['MBM_Worked'].sum()

            #Returns message of changes made
            message1 = '{} was manually assigned the next MBM Case at {}'.format(selected_name, time_ez)
            assignee = '{} at {}'.format(selected_name, time_atm)
            
            #Finds agent with highest selection value and is working
            work_list = df[df['Working'] == True]
            df2 = pd.DataFrame(work_list)
            max_work_index = df2['MBM_Selected'].idxmax()
            max_agent_name_mbm = df.at[max_work_index, 'Name']
            next_assignee = max_agent_name_mbm
            dt.at[0, 'Next_Agent_MBM'] = next_assignee                                          

            #Writes data to Excel files
            df.to_excel('Agents.xlsx', index = False)
            dt.to_excel('Data.xlsx', index = False)

    #calls on auto_mbm assign button function if pressed
    if triggered_id == 'ambm-btn':
        return Auto_MBM()


    elif triggered_id == 'undo-mbm-btn':
        return Undo_MBM()
    else:
        pass
    options = df['Name']

    return message1, mbm_total_output, assignee, mbm_day_output, prev_assignee, next_assignee, options


#Auto Assign ticket based on a selection value counter. The agent with highest value is selected for next ticket. Then selection value is reset for selected agent
#And all other agents gain +1 to selection value.
def Auto_MBM():

    df = pd.read_excel(r'Agents.xlsx')            # Directory needs to be updated
    dt = pd.read_excel(r'Data.xlsx')              # Directory needs to be updated

    #Checks if any agents are working
    if df['Working'].values.sum() != 0:

        #Gets the time
        now = datetime.now()
        timestamp_num = datetime.timestamp(now)
        time_now = datetime.fromtimestamp(timestamp_num)
        timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')
        
        time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

        dt['MBM_Time'] = time_ez
        time_atm = dt.at[0, 'MBM_Time']

        #Creates list of agents who are working today
        work_list = df[df['Working'] == True]

        #Finds agent who has lowest selection value (last agent who worked case or a new agent)
        min_work_index = df['MBM_Selected'].idxmin()
        min_agent_name_mbm = df.at[min_work_index, 'Name']
        prev_assignee = min_agent_name_mbm
        dt.at[0, 'Prev_Agent_MBM'] = prev_assignee
        
        #Creates Data Frame of agents who are working today, then finds the agent with highest slection count value
        df2 = pd.DataFrame(work_list)
        max_work_index = df2['MBM_Selected'].idxmax()
        max_agent_name_mbm = df.at[max_work_index, 'Name']
        
        #Reads and creates Data Frame of personal agent Excel file
        da_m = pd.read_excel(max_agent_name_mbm+'.xlsx', sheet_name=max_agent_name_mbm+'_MBM_Worked')  
        da_u = pd.read_excel(max_agent_name_mbm+'.xlsx', sheet_name=max_agent_name_mbm+'_UET_Worked') 

        #Adds +1 count to all agent's selection count value
        df['MBM_Selected'] += 1

        #Adds +1 total number of worked tickets
        df.at[max_work_index, 'MBM_Worked'] += 1

        #Gets name of agent with highest selection count value then returns it in message
        assignee_name = df.at[max_work_index, 'Name']
        assignee = '{} at {}'.format(assignee_name, time_ez)
        message1 = "{} was automatically assigned the next MBM case at {}".format(assignee_name, time_atm)

        #Resets value of highest selection count agent to 0
        df.at[max_work_index, 'MBM_Selected'] = 0

        #Adds +1 to total ever worked tickets
        dt['Total_MBM_Cases'] += 1

        #Gets sum of total tickets worked for the day
        mbm_day_output = df['MBM_Worked'].sum()
        
        #Assigns value of total tickets worked
        mbm_total_output = dt['Total_MBM_Cases']


        #Makes copy of selected agent data frame, writes timestamp to copy of data frame and then merges them together
        da_m1 = pd.DataFrame(da_m[[max_agent_name_mbm+'_MBM_Worked', 'Action']])
        da_m2 = pd.DataFrame({max_agent_name_mbm+'_MBM_Worked': [timestamp], "Action": ['Automatically Assigned Ticket']})
        da_m = pd.concat([da_m1, da_m2])

        #writes data to agent's personal excel file
        writer = pd.ExcelWriter(max_agent_name_mbm+'.xlsx', engine='xlsxwriter')  
        with pd.ExcelWriter(max_agent_name_mbm+'.xlsx') as writer:
            da_m.to_excel(writer, sheet_name=max_agent_name_mbm+'_MBM_Worked', index=False)
            da_u.to_excel(writer, sheet_name=max_agent_name_mbm+'_UET_Worked', index=False)

        #Finds agent with highest selection value and is working
        work_list = df[df['Working'] == True]
        df2 = pd.DataFrame(work_list)
        max_work_index = df2['MBM_Selected'].idxmax()
        max_agent_name_mbm = df.at[max_work_index, 'Name']
        next_assignee = max_agent_name_mbm
        dt['Next_Agent_MBM'] = next_assignee   

        #Saves changes to Excel files
        df.to_excel('Agents.xlsx', index = False)              # Directory needs to be updated
        dt.to_excel('Data.xlsx', index = False)              # Directory needs to be updated

    else:
        
        #Returns error messages
        message1 = 'No agents selected to work today!'
        assignee = 'Agent not assigned'
        mbm_day_output = df['MBM_Worked'].sum()
        
        mbm_total_output = dt['Total_MBM_Cases']

        prev_assignee = dt.at[0, 'Prev_Agent_MBM']
        next_assignee = dt.at[0, 'Next_Agent_MBM']

    options = df['Name']

    return message1, mbm_total_output, assignee, mbm_day_output, prev_assignee, next_assignee, options



##########################
#### MBM Undo Button  ####
##########################


#Undo last MBM case assignment
def Undo_MBM():

    df = pd.read_excel(r'Agents.xlsx')
    dt = pd.read_excel(r'Data.xlsx')

    min_work_index = df['MBM_Selected'].idxmin()

    #Checks if any agents are working, if true: search agent with lowest count then -= 1 to worked tickets. Adds 9001 to count for assignee agent, then -= 1 count to all other agents. 
    if df.at[min_work_index, 'MBM_Worked'] == 0:
        
        time_atm = dt.at[0, 'MBM_Time']

        message1 = 'Unable to undo further'
        
        min_work_index = df['MBM_Selected'].idxmin()
        min_agent_name_mbm = df.at[min_work_index, 'Name']

        assignee = 'Unable to undo further'

        mbm_total_output = dt['Total_MBM_Cases']
        
        mbm_day_output = df['MBM_Worked'].sum()

        prev_assignee = dt.at[0, 'Prev_Agent_MBM']
        next_assignee = dt.at[0, 'Next_Agent_MBM']
    else:
        
        if df.at[min_work_index, 'Working'] == True:

            #Gets the time
            now = datetime.now()
            timestamp_num = datetime.timestamp(now)
            time_now = datetime.fromtimestamp(timestamp_num)
            timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')

            time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

            dt['MBM_Time'] = time_ez
            time_atm = dt.at[0, 'MBM_Time']

            message1 = 'Click this button to undo last MBM assignment'

            min_work_index = df['MBM_Selected'].idxmin()
            min_agent_name_mbm = df.at[min_work_index, 'Name']

            mbm_day_output = df['MBM_Worked'].sum()

            da_u = pd.read_excel(min_agent_name_mbm+'.xlsx', sheet_name=min_agent_name_mbm+'_UET_Worked')   
            da_m = pd.read_excel(min_agent_name_mbm+'.xlsx', sheet_name=min_agent_name_mbm+'_MBM_Worked') 

            df['MBM_Selected'] -= 1
            df.at[min_work_index, 'MBM_Worked'] -= 1
            df.at[min_work_index, 'MBM_Selected'] = 9001
            dt['Total_MBM_Cases'] -= 1
              
            mbm_total_output = dt['Total_MBM_Cases']

            #Gets last agent who was assigned ticket and creates a dataframe
            da_mbm_last_agent = pd.DataFrame(da_m[[min_agent_name_mbm+'_MBM_Worked', 'Action']])
    
            #Gets the index value of the last row which contains 'Assigned Ticket' value
            last_assign_ticket_index = da_mbm_last_agent.loc[da_mbm_last_agent['Action'].str.contains('Assigned Ticket')].index[-1]
    
            #Replaces value of last "Assigned Ticket" with "Undone" statement and timestamp
            da_mbm_last_agent.iloc[last_assign_ticket_index, da_mbm_last_agent.columns.get_loc('Action')] = 'Case Assignment Undone @ '+timestamp
            da_m = da_mbm_last_agent


            #writes data to excel file
            writer = pd.ExcelWriter(min_agent_name_mbm+'.xlsx', engine='xlsxwriter')  
            with pd.ExcelWriter(min_agent_name_mbm+'.xlsx') as writer:
                da_m.to_excel(writer, sheet_name=min_agent_name_mbm+'_MBM_Worked', index=False)
                da_u.to_excel(writer, sheet_name=min_agent_name_mbm+'_UET_Worked', index=False)
            
            #Collects information numbers
            mbm_day_output = df['MBM_Worked'].sum()

            #Finds agent with highest selection value and is working
            work_list = df[df['Working'] == True]
            df2 = pd.DataFrame(work_list)
            max_work_index = df2['MBM_Selected'].idxmax()
            max_agent_name_mbm = df.at[max_work_index, 'Name']
            next_assignee = max_agent_name_mbm

            min_work_index = df['MBM_Selected'].idxmin()
            min_agent_name_mbm1 = df.at[min_work_index, 'Name']
            dt['Prev_Agent_MBM'] = min_agent_name_mbm1 
            prev_assignee = min_agent_name_mbm1
        
            assignee = '-Undo- {}'.format(time_atm)
        
            message1 = "{} was unassigned a MBM Ticket at {}.".format(min_agent_name_mbm, time_atm)

            #Saves changes to Excel files
            df.to_excel('Agents.xlsx', index = False)              # Directory needs to be updated
            dt.to_excel('Data.xlsx', index = False)              # Directory needs to be updated

        else:

            min_work_index = df['MBM_Selected'].idxmin()
            min_agent_name_mbm1 = df.at[min_work_index, 'Name']
            dt['Prev_Agent_MBM'] = min_agent_name_mbm1 
            prev_assignee = min_agent_name_mbm1

            message1 = 'Please ensure {} is set to working today and try again!'.format(prev_assignee)
        
            mbm_total_output = dt['Total_MBM_Cases']
        
            assignee = 'Error - Unable to Undo'

            mbm_day_output = df['MBM_Worked'].sum()

            prev_assignee = dt.at[0, 'Prev_Agent_MBM']
            next_assignee = dt.at[0, 'Next_Agent_MBM']

    options = df['Name']

    return message1, mbm_total_output, assignee, mbm_day_output, prev_assignee, next_assignee, options



###########################################
### Manual/Automatic Assign UET Ticket ####
###########################################


@app.callback(
    #Auto/Mnl UET assign
    Output('uet-output-container', 'children'),
    Output('uet-count-output', 'children'),
    Output('uet-assignee-output', 'children'),
    Output('uet-day-count-output', 'children'),
    Output('uet-prev-assignee-output', 'children'),
    Output('uet-next-assignee-output', 'children'),
    Output('uet_dropdown', 'options'),
    
    Input('undo-uet-btn', 'submit_n_clicks'),
    Input('auet-btn', 'n_clicks'),
    Input('muet-btn', 'n_clicks'),
    State('uet_dropdown', 'options'),
    State("uet_dropdown", "value"),
    Input('store-data-uet', 'data'),

    prevent_initial_call=False
)

#Manual assign UET ticket and setups button call for Automatic assign UET case 
def update_uet(button1, button2, button3, options, value, data):

    sleep(0.4)

    dt = pd.read_excel(r'Data.xlsx')
    df = pd.read_excel(r'Agents.xlsx')

    #Assigns value of total tickets worked
    uet_total_output = dt.at[0, 'Total_UET_Tickets']
    message1 = "Either manually or automatically assign agent a ticket"
    
    min_work_index = df['UET_Selected'].idxmin()
    min_agent_name_uet = df.at[min_work_index, 'Name']
    assignee = '{} at {}'.format(min_agent_name_uet, dt.at[0, 'UET_Time'])

    #Gets sum of total tickets worked for the day
    uet_day_output = df['UET_Worked'].sum()

    #Waits for button click to be triggered
    triggered_id = ctx.triggered_id

    #Finds the agents with highest and lowest selection value
    prev_assignee = dt.at[0, 'Prev_Agent_UET']
    next_assignee = data

    
    if triggered_id == 'muet-btn':
         
        #Reads/creates dataframes from Excel files
        df = pd.read_excel(r'Agents.xlsx')
        dt = pd.read_excel(r'Data.xlsx')
        
        #Gets agent name from dropdown and finds index value
        selected_name = value                                                      # df['Name'].str.capitalize() )
        uet_selected_index = df[df['Name']==selected_name].index.values

        #Assigns value of total tickets worked
        uet_total_output = dt.at[0, 'Total_UET_Tickets']

        #Gets sum of total tickets worked for the day
        uet_day_output = df['UET_Worked'].sum()

        #Checks if selected agent is working today then returns True/False value
        working_true = df.at[int(uet_selected_index), 'Working']

        #Checks if agent is working today value is False
        if working_true == False:
            
            message1 = "Are you sure you selected the right agent? This agent isn't set to work today!"
            assignee = "Agent not assigned"

            #Finds agent who has lowest selection value (last agent who worked case or a new agent)
            prev_assignee = dt.at[0, 'Prev_Agent_UET']
            next_assignee = dt.at[0, 'Next_Agent_UET']


        else:

            #Gets the time
            now = datetime.now()
            timestamp_num = datetime.timestamp(now)
            time_now = datetime.fromtimestamp(timestamp_num)
            timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')

            time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

            dt['UET_Time'] = time_ez
            time_atm = dt.at[0, 'UET_Time']

            #Finds agent who has lowest selection value (last agent who worked case or a new agent)
            min_work_index = df['UET_Selected'].idxmin()
            min_agent_name_uet = df.at[min_work_index, 'Name']
            prev_assignee = min_agent_name_uet
            dt.at[0, 'Prev_Agent_UET'] = prev_assignee

            #Adds +1 ticket worked to selected agent
            df.at[int(uet_selected_index), 'UET_Worked'] += 1
            
            #Add +1 to selection count value to all agents
            df['UET_Selected'] += 1

            #Resets selection value count of selected agent to zero
            df.at[int(uet_selected_index), 'UET_Selected'] = 0

            #Add +1 to total worked tickets
            dt['Total_UET_Tickets'] += 1

            #Reads personal agent Excel files            
            da_m = pd.read_excel(selected_name+'.xlsx', sheet_name=selected_name+'_MBM_Worked') 
            da_u = pd.read_excel(selected_name+'.xlsx', sheet_name=selected_name+'_UET_Worked')  

            #Makes copy of agent data and adds new data
            da_u1 = pd.DataFrame(da_u[[selected_name+'_UET_Worked', 'Action']])
            da_u2 = pd.DataFrame({selected_name+'_UET_Worked': [timestamp], "Action": ['Manually Assigned Ticket']})
            da_u = pd.concat([da_u1, da_u2])

            #writes data to excel file
            writer = pd.ExcelWriter(selected_name+'.xlsx', engine='xlsxwriter')  
            with pd.ExcelWriter(selected_name+'.xlsx') as writer:
                da_m.to_excel(writer, sheet_name=selected_name+'_MBM_Worked', index=False)
                da_u.to_excel(writer, sheet_name=selected_name+'_UET_Worked', index=False)
            
            #Assigns value of total tickets worked
            uet_total_output = dt.at[0, 'Total_UET_Tickets']
            
            #Gets sum of total tickets worked for the day
            uet_day_output = df['UET_Worked'].sum()

            #Returns message of changes made
            message1 = '{} was manually assigned the next UET ticket at {}'.format(selected_name, time_ez)
            assignee = '{} at {}'.format(selected_name, time_atm)
            
            #Finds agent with highest selection value and is working
            work_list = df[df['Working'] == True]
            df2 = pd.DataFrame(work_list)
            max_work_index = df2['UET_Selected'].idxmax()
            max_agent_name_uet = df.at[max_work_index, 'Name']
            next_assignee = max_agent_name_uet
            dt.at[0, 'Next_Agent_UET'] = next_assignee                                          

            #Writes data to Excel files
            df.to_excel('Agents.xlsx', index = False)
            dt.to_excel('Data.xlsx', index = False)

        
    #calls on auto_uet assign button function if pressed
    if triggered_id == 'auet-btn':
         return Auto_UET()

    elif triggered_id == 'undo-uet-btn':
        return Undo_UET()

    options = df['Name']

    return message1, uet_total_output, assignee, uet_day_output, prev_assignee, next_assignee, options


#Auto Assign ticket based on a selection value counter. The agent with highest value is selected for next ticket. Then selection value is reset for selected agent
#And all other agents gain +1 to selection value.
def Auto_UET():

    df = pd.read_excel(r'Agents.xlsx')            # Directory needs to be updated
    dt = pd.read_excel(r'Data.xlsx')              # Directory needs to be updated

    #Checks if any agents are working
    if df['Working'].values.sum() == 0:
        
        #Returns error messages
        message1 = 'No agents selected to work today!'
        assignee = 'Agent not assigned'
        uet_day_output = df['UET_Worked'].sum()
        uet_total_output = dt.at[0, 'Total_UET_Tickets']

        #Finds agent who has lowest selection value (last agent who worked case or a new agent)
        prev_assignee = dt.at[0, 'Prev_Agent_UET']
        next_assignee = dt.at[0, 'Next_Agent_UET']

        
    else:
        #Gets the time
        now = datetime.now()
        timestamp_num = datetime.timestamp(now)
        time_now = datetime.fromtimestamp(timestamp_num)
        timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')

        time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

        dt['UET_Time'] = time_ez
        time_atm = dt.at[0, 'UET_Time']

        #Creates list of agents who are working today
        work_list = df[df['Working'] == True]

        #Finds agent who has lowest selection value (last agent who worked case or a new agent)
        min_work_index = df['UET_Selected'].idxmin()
        min_agent_name_uet = df.at[min_work_index, 'Name']
        prev_assignee = min_agent_name_uet
        dt.at[0, 'Prev_Agent_UET'] = prev_assignee
        
        #Creates Data Frame of agents who are working today, then finds the agent with highest slection count value
        df2 = pd.DataFrame(work_list)
        max_work_index = df2['UET_Selected'].idxmax()
        max_agent_name_uet = df.at[max_work_index, 'Name']
        
        #Reads and creates Data Frame of personal agent Excel file
        da_m = pd.read_excel(max_agent_name_uet+'.xlsx', sheet_name=max_agent_name_uet+'_MBM_Worked')  
        da_u = pd.read_excel(max_agent_name_uet+'.xlsx', sheet_name=max_agent_name_uet+'_UET_Worked') 

        #Adds +1 count to all agent's selection count value
        df['UET_Selected'] += 1

        #Adds +1 total number of worked tickets
        df.at[max_work_index, 'UET_Worked'] += 1

        #Gets name of agent with highest selection count value then returns it in message
        assignee_name = df.at[max_work_index, 'Name']
        assignee = '{} at {}'.format(assignee_name, time_ez)
        message1 = "{} was automatically assigned the next UET ticket at {}".format(assignee_name, time_atm)

        #Resets value of highest selection count agent to 0
        df.at[max_work_index, 'UET_Selected'] = 0

        #Adds +1 to total ever worked tickets
        dt['Total_UET_Tickets'] += 1

        #Gets sum of total tickets worked for the day
        uet_day_output = df['UET_Worked'].sum()
        
        #Assigns value of total tickets worked
        uet_total_output = dt.at[0, 'Total_UET_Tickets']

        #Makes copy of selected agent data frame, writes timestamp to copy of data frame and then merges them together
        da_u1 = pd.DataFrame(da_u[[max_agent_name_uet+'_UET_Worked', 'Action']])
        da_u2 = pd.DataFrame({max_agent_name_uet+'_UET_Worked': [timestamp], "Action": ['Automatically Assigned Ticket']})
        da_u = pd.concat([da_u1, da_u2])

        #writes data to agent's personal excel file
        writer = pd.ExcelWriter(max_agent_name_uet+'.xlsx', engine='xlsxwriter')  
        with pd.ExcelWriter(max_agent_name_uet+'.xlsx') as writer:
            da_m.to_excel(writer, sheet_name=max_agent_name_uet+'_MBM_Worked', index=False)
            da_u.to_excel(writer, sheet_name=max_agent_name_uet+'_UET_Worked', index=False)

        #Finds agent with highest selection value and is working
        work_list = df[df['Working'] == True]
        df2 = pd.DataFrame(work_list)
        max_work_index = df2['UET_Selected'].idxmax()
        max_agent_name_uet = df.at[max_work_index, 'Name']
        next_assignee = max_agent_name_uet
        dt['Next_Agent_UET'] = next_assignee   

        #Saves changes to Excel files
        df.to_excel('Agents.xlsx', index = False)              # Directory needs to be updated
        dt.to_excel('Data.xlsx', index = False)              # Directory needs to be updated

    options = df['Name']

    return message1, uet_total_output, assignee, uet_day_output, prev_assignee, next_assignee, options



##########################
#### UET Undo Button  ####
##########################


#Undo last UET case assignment
def Undo_UET():

    df = pd.read_excel(r'Agents.xlsx')
    dt = pd.read_excel(r'Data.xlsx')

    min_work_index = df['UET_Selected'].idxmin()

    #Checks if any agents are working, if true: search agent with lowest count then -= 1 to worked tickets. Adds 9001 to count for assignee agent, then -= 1 count to all other agents. 
    if df.at[min_work_index, 'UET_Worked'] == 0:
        
        time_atm = dt.at[0, 'UET_Time']

        message1 = 'Unable to undo further'
        
        min_work_index = df['UET_Selected'].idxmin()
        min_agent_name_uet = df.at[min_work_index, 'Name']

        assignee = 'Unable to undo further'

        uet_total_output = dt['Total_UET_Tickets']
        
        uet_day_output = df['UET_Worked'].sum()

        prev_assignee = dt.at[0, 'Prev_Agent_UET']

        next_assignee = dt.at[0, 'Next_Agent_UET']

    else:

        if df.at[min_work_index, 'Working'] == True:

            #Gets the time
            now = datetime.now()
            timestamp_num = datetime.timestamp(now)
            time_now = datetime.fromtimestamp(timestamp_num)
            timestamp = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('%H:%M:%S %m-%d-%Y')

            time_ez = time_now.astimezone(pytz.timezone('US/Eastern')).strftime('|| %H:%M-%Z | %m/%d/%Y ||')

            dt['UET_Time'] = time_ez
            time_atm = dt.at[0, 'UET_Time']

            min_work_index = df['UET_Selected'].idxmin()
            min_agent_name_uet = df.at[min_work_index, 'Name']

            message1 = 'Click this button to undo last UET assignment'

            uet_day_output = df['UET_Worked'].sum()
            uet_total_output = dt.at[0, 'Total_UET_Tickets']

            da_u = pd.read_excel(min_agent_name_uet+'.xlsx', sheet_name=min_agent_name_uet+'_UET_Worked')   
            da_m = pd.read_excel(min_agent_name_uet+'.xlsx', sheet_name=min_agent_name_uet+'_MBM_Worked') 

            df['UET_Selected'] -= 1
            df.at[min_work_index, 'UET_Worked'] -= 1
            df.at[min_work_index, 'UET_Selected'] = 9001
            dt['Total_UET_Tickets'] -= 1
                
            uet_total_output = dt['Total_UET_Tickets']


            #Gets last agent who was assigned ticket and creates a dataframe
            da_uet_last_agent = pd.DataFrame(da_u[[min_agent_name_uet+'_UET_Worked', 'Action']])
    
            #Gets the index value of the last row which contains 'Assigned Ticket' value
            last_assign_ticket_index = da_uet_last_agent.loc[da_uet_last_agent['Action'].str.contains('Assigned Ticket')].index[-1]
    
            #Replaces value of last "Assigned Ticket" with "Undone" statement and timestamp
            da_uet_last_agent.iloc[last_assign_ticket_index, da_uet_last_agent.columns.get_loc('Action')] = 'Ticket Assignment Undone @ '+timestamp
            da_u = da_uet_last_agent

            #writes data to excel file
            writer = pd.ExcelWriter(min_agent_name_uet+'.xlsx', engine='xlsxwriter')  
            with pd.ExcelWriter(min_agent_name_uet+'.xlsx') as writer:
                da_m.to_excel(writer, sheet_name=min_agent_name_uet+'_MBM_Worked', index=False)
                da_u.to_excel(writer, sheet_name=min_agent_name_uet+'_UET_Worked', index=False)
                
            #Collectss information numbers
            uet_day_output = df['UET_Worked'].sum()

            #Finds agent with highest selection value and is working
            work_list = df[df['Working'] == True]
            df2 = pd.DataFrame(work_list)
            max_work_index = df2['UET_Selected'].idxmax()
            max_agent_name_uet = df.at[max_work_index, 'Name']
            dt['Next_Agent_UET'] = max_agent_name_uet  
            next_assignee = max_agent_name_uet

            min_work_index = df['UET_Selected'].idxmin()
            min_agent_name_uet1 = df.at[min_work_index, 'Name']
            prev_assignee = min_agent_name_uet1
            dt['Prev_Agent_UET'] = min_agent_name_uet1
            
            assignee = '-Undo- {}'.format(time_atm)
        
            message1 = "{} was unassigned a UET Ticket at {}.".format(min_agent_name_uet, time_atm)

            #Saves changes to Excel files
            df.to_excel('Agents.xlsx', index = False)              # Directory needs to be updated
            dt.to_excel('Data.xlsx', index = False)              # Directory needs to be updated

        else:
        
            min_work_index = df['UET_Selected'].idxmin()
            min_agent_name_uet1 = df.at[min_work_index, 'Name']
            dt['Prev_Agent_UET'] = min_agent_name_uet1 
            prev_assignee = min_agent_name_uet1
            
            message1 = 'Please ensure {} is set to working today and try again!'.format(prev_assignee)
        
            uet_total_output = dt['Total_UET_Tickets']
        
            assignee = 'Error - Unable to Undo'

            uet_day_output = df['UET_Worked'].sum()

            prev_assignee = dt.at[0, 'Prev_Agent_UET']
            next_assignee = dt.at[0, 'Next_Agent_UET']

    options = df['Name']

    return message1, uet_total_output, assignee, uet_day_output, prev_assignee, next_assignee, options


############################################################################
### Runs Report - Reads Excel Files and Combines into Master Dataframe #####
############################################################################

@app.callback(
    Output('report-output', 'children'),
    Output('mbm-bar-chart', 'figure'),
    Output('uet-bar-chart', 'figure'),
    
    Input("run-report_btn", "n_clicks"),
    Input('mbm-date-picker-range', 'start_date'),
    Input('mbm-date-picker-range', 'end_date'),
    Input('uet-date-picker-range', 'start_date'),
    Input('uet-date-picker-range', 'end_date'),

    prevent_initial_call=True,
)
def run_report(n_clicks, start_date_mbm, end_date_mbm, start_date_uet, end_date_uet):

    # Sets button trigger ID
    triggered_id = ctx.triggered_id

    if triggered_id == 'run-report_btn':

        # Reads Agent file in order to get list of agent names
        df = pd.read_excel(r'Agents.xlsx')
        agent_list = df['Name']

        # Sets up blank lists to be used later
        dfs_mbm = []
        dfs_uet = []

        # Cycles through list of agents and opens each excel files
        for i in agent_list:
            agent_list == i
            dict_df = pd.read_excel(i + '.xlsx', sheet_name=[i + '_MBM_Worked', i + '_UET_Worked'])

            dfw_m = dict_df.get(i + '_MBM_Worked')
            dfw_m.rename(columns={i + '_MBM_Worked': "Date_Time"}, inplace=True)
            dfw_m['Name'] = i
            dfs_mbm.append(dfw_m)

            dfw_u = dict_df.get(i + '_UET_Worked')
            dfw_u.rename(columns={i + '_UET_Worked': "Date_Time"}, inplace=True)
            dfw_u['Name'] = i
            dfs_uet.append(dfw_u)

        # Raw combined files with each spreadsheet (MBM/UET, with both columns)
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

    else:
        pass

    dft_master = pd.read_excel(r'Master.xlsx')

    #Counts undo counts and returns value
    mbm_undo_count = dft_master['Action'].str.count('Undone').sum()
    uet_undo_count = dft_master['Action'].str.count('Undone').sum()
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

    # Sets up timestamp increments by day
    daily_uet_counts = dft_uet_master.resample('D').count()
    dft_uet_master
    mask_mbm = (daily_mbm_counts.index >= start_date_mbm) & (daily_mbm_counts.index <= end_date_mbm)
    filtered_counts_mbm = daily_mbm_counts.loc[mask_mbm]
    figure_mbm = {
        'data': [{

            'x': filtered_counts_mbm.index,
            'y': filtered_counts_mbm['Action'],

            'type': 'bar'
        }],
            'layout': {
            'title': 'MBM Case Assignment History',
            'xaxis': {'title': 'Date'},
            'yaxis': {'title': 'Case Count'}
        }
    }
        
    mask_uet = (daily_uet_counts.index >= start_date_uet) & (daily_uet_counts.index <= end_date_uet)
    filtered_counts_uet = daily_uet_counts.loc[mask_uet]
    figure_uet = {
        'data': [{

            'x': filtered_counts_uet.index,
            'y': filtered_counts_uet['Action'],

            'type': 'bar'
        }],
            'layout': {
            'title': 'UET Ticket Assignment History',
            'xaxis': {'title': 'Date'},
            'yaxis': {'title': 'Ticket Count'}
        }
    }

    return report_details, figure_mbm, figure_uet



#MBM Download callback
@app.callback(
    Output("download-mbm-dataframe-xlsx", "data"),
    Input("download_mbm_btn", "n_clicks"),
    Input("download_mbm_dropdown", "value"),
    prevent_initial_call=True,
)
def download_mbm_xlsx(n_clicks, value):

    #Sets trigger ID
    triggered_id = ctx.triggered_id
    
    if triggered_id == 'download_mbm_btn':

        df_mbm_selected_dwnld = pd.read_excel(value+'.xlsx', sheet_name=value+'_MBM_Worked') 
        df_mbm_dwnld = pd.DataFrame(df_mbm_selected_dwnld)
        return dcc.send_data_frame(df_mbm_dwnld.to_excel, value+'_MBM_Worked.xlsx', sheet_name="MBM_Cases_Worked")
    else:
        pass


#UET Download callback
@app.callback(
    Output("download-uet-dataframe-xlsx", "data"),
    Input("download_uet_btn", "n_clicks"),
    Input("download_uet_dropdown", "value"),
    prevent_initial_call=True,
)
def download_uet_xlsx(n_clicks, value):

    #Sets trigger ID
    triggered_id = ctx.triggered_id
    
    if triggered_id == 'download_uet_btn':

        df_uet_selected_dwnld = pd.read_excel(value+'.xlsx', sheet_name=value+'_UET_Worked') 
        df_uet_dwnld = pd.DataFrame(df_uet_selected_dwnld)
        return dcc.send_data_frame(df_uet_dwnld.to_excel, value+'_UET_Worked.xlsx', sheet_name="UET_Tickets_Worked")
    else:
        pass




# Run the app on localhost:8050
if __name__ == '__main__':
    app.run_server(debug=True)



    