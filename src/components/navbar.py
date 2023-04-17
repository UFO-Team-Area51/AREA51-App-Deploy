from dash import html
import dash_bootstrap_components as dbc

image_path = "/assets/Annie.png"


# Define the navbar structure
def navbar():
    layout = html.Div([
        dbc.NavbarSimple(
            children=[
                dbc.NavItem(dbc.NavLink("Tracker", href="/Tracker")),
                dbc.NavItem(dbc.NavLink("Add/Remove Agent", href="/Add_Remove_Agent")),
                dbc.NavItem(dbc.NavLink("Analytics", href="/Analytics")),
                html.Img(src=image_path),
            ],

            brand="Agent Rotation Errand Assiger v5.1 (AREA 51)",
            brand_href="/Tracker",
            color="#00c3ff",  # Was 00d257
            dark=True,
        ),
    ])

    return layout
