import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import pandas as pd
import datetime
import random

# Step 1: Set up the Dash app
app = dash.Dash(__name__)

# Step 2: Define the layout
app.layout = html.Div(children=[
    html.H1(children='Dynamic Bar Graph Example'),
    dcc.Graph(id='bar-graph'),
    dcc.Interval(
        id='interval-component',
        interval=2000,  # Update interval in milliseconds (2 seconds)
        n_intervals=0
    )
])

# Step 3: Define the callback to update the bar graph
@app.callback(
    Output('bar-graph', 'figure'),
    [Input('interval-component', 'n_intervals')]
)
def update_graph(n_intervals):
    # Simulate changing data in a DataFrame
    categories = ['A', 'B', 'C', 'D']
    values = [random.randint(1, 10) for _ in range(len(categories))]
    df = pd.DataFrame({'Category': categories, 'Value': values})

    # Update the bar graph
    fig = go.Figure(
        go.Bar(
            x=df['Category'],
            y=df['Value'],
            marker_color='blue'
        )
    )
    fig.update_layout(title='Dynamic Bar Graph Example (Updating Data)')
    return fig

# Step 4: Run the app
if __name__ == '__main__':
    app.run_server(debug=True)