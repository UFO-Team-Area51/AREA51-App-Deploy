# Define the Dash App and it's properties here 
import dash
import dash_bootstrap_components as dbc

# noinspection IncorrectFormatting
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.COSMO],
                meta_tags=[{"name": "viewport", "content": "width=device-width",}
                           ],
                suppress_callback_exceptions=True)

server = app.server

app.config.suppress_callback_exceptions = True
