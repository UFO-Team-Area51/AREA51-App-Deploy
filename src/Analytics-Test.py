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

dft_master = pd.read_excel(r'Master.xlsx')

print(dft_master)