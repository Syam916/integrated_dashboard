import pandas as pd
from dash import Dash, dcc, html, Input, Output
import plotly.express as px
import plotly.graph_objects as go

# Load data from Excel sheet
data_file = 'DBdata.xlsx'  # Replace with your Excel file path
data = pd.read_excel(data_file)

# Initialize the Dash app
app = Dash(__name__)

def create_donut_chart(value, total, color):
    fig = go.Figure(go.Pie(
        values=[0, total],  # Start with an empty pie (0 for the value part)
        hole=0.7,  # Thin donut
        marker=dict(colors=[color, '#E0E0E0']),  # Color the value and make the rest gray
        textinfo='none',  # Hide text in pie chart
        direction='clockwise',  # Fill starting from the right
        showlegend=False  # No legends
    ))

    # Animation properties to make the donut fill automatically
    fig.update_traces(
        values=[value, total - value],  # Actual values to fill
        selector=dict(type='pie')
    )

    fig.update_layout(
        annotations=[dict(text=str(value), x=0.5, y=0.5, font_size=20, font_family='Arial', showarrow=False)],  # Center number
        margin=dict(l=0, r=20, t=0, b=0),  # Remove margins
        height=85,  # Height of the chart
        width=105,  # Width of the chart
        showlegend=False,  # No legend
        paper_bgcolor="rgba(0,0,0,0)",  # Transparent background
        plot_bgcolor="rgba(0,0,0,0)",  # Transparent plot background
    )

    return fig

# App layout
app.layout = html.Div([
    # 1st Section: Dropdowns
    html.Div([
        html.Div([
            html.Label('Select Topic Id:', className='font-weight-bold'),
            dcc.Dropdown(
                id='topic-dropdown',
                options=[{'label': str(topic), 'value': topic} for topic in data['Topic Id'].unique()],
                value=data['Topic Id'].unique()[0],  # Default to the first Topic Id
                className='mb-3'
            )
        ], style={'width': '48%', 'display': 'inline-block', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#f9f9f9', 'padding': '10px'}),

        html.Div([
            html.Label('Select Source:', className='font-weight-bold'),
            dcc.Dropdown(
                id='source-dropdown',
                options=[{'label': 'All', 'value': 'All'}] + [
                    {'label': str(source), 'value': source} for source in data['Source'].unique()
                ],
                value='All',  # Default to "All"
                className='mb-3'
            )
        ], style={'width': '38%', 'display': 'inline-block', 'marginLeft': '4%', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#f9f9f9', 'padding': '10px'})
    ], style={'marginBottom': '20px'}),

    # Metrics Section with Donut Charts
    html.Div([
        html.Div([
            html.H5("Total", className='text-center font-weight-bold', style={'fontSize': '18px', 'fontFamily': 'Arial'}),
            dcc.Graph(id='total-metric', config={'displayModeBar': False}, style={'margin-left': '115px'})
        ], style={'width': '20%', 'display': 'inline-block', 'margin': '10px', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#f5f5f5', 'padding': '15px', 'textAlign': 'center'}),

        html.Div([
            html.H5("Yes", className='text-center font-weight-bold', style={'fontSize': '18px', 'fontFamily': 'Arial'}),
            dcc.Graph(id='yes-metric', config={'displayModeBar': False}, style={'margin-left': '115px'})
        ], style={'width': '20%', 'display': 'inline-block', 'margin': '10px', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#f5f5f5', 'padding': '15px', 'textAlign': 'center'}),

        html.Div([
            html.H5("No", className='text-center font-weight-bold', style={'fontSize': '18px', 'fontFamily': 'Arial'}),
            dcc.Graph(id='no-metric', config={'displayModeBar': False}, style={'margin-left': '115px'})
        ], style={'width': '20%', 'display': 'inline-block', 'margin': '10px', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#f5f5f5', 'padding': '15px', 'textAlign': 'center'}),

        html.Div([
            html.H5("Maybe", className='text-center font-weight-bold', style={'fontSize': '18px', 'fontFamily': 'Arial'}),
            dcc.Graph(id='maybe-metric', config={'displayModeBar': False}, style={'margin-left': '115px'})
        ], style={'width': '20%', 'display': 'inline-block', 'margin': '10px', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#f5f5f5', 'padding': '15px', 'textAlign': 'center'})
    ], style={'marginBottom': '20px', 'display': 'flex', 'justifyContent': 'center'}),

    # 2nd Section: Charts
    html.Div([
        html.Div([
            dcc.Graph(id='pie-chart')
        ], style={'width': '48%', 'display': 'inline-block', 'margin': '10px', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#ffffff', 'padding': '20px'}),

        html.Div([
            dcc.Graph(id='bar-chart')
        ], style={'width': '48%', 'display': 'inline-block', 'margin': '10px', 'border': '2px solid #ccc', 'borderRadius': '8px', 'boxShadow': '2px 2px 8px rgba(0, 0, 0, 0.1)', 'backgroundColor': '#ffffff', 'padding': '20px'})
    ], style={'display': 'flex', 'justifyContent': 'center'})
], style={'scroll': 'hidden', 'textAlign': 'center'})

# Callback for updating charts and metrics
@app.callback(
    [Output('pie-chart', 'figure'), Output('bar-chart', 'figure'),
     Output('total-metric', 'figure'), Output('yes-metric', 'figure'),
     Output('no-metric', 'figure'), Output('maybe-metric', 'figure')],
    [Input('topic-dropdown', 'value'), Input('source-dropdown', 'value')]
)
def update_charts_and_metrics(selected_topic, selected_source):
    # Ensure data types match
    data['Topic Id'] = data['Topic Id'].astype(str)
    selected_topic = str(selected_topic)

    # Filter data based on dropdowns
    filtered_data = data[data['Topic Id'] == selected_topic]
    if selected_source != 'All':
        filtered_data = filtered_data[filtered_data['Source'] == selected_source]

    # Pie Chart
    pie_fig = px.pie(
        filtered_data,
        names='Response',
        title='Response Percentages',
        hole=0.7,
        color='Response',
        color_discrete_map={'Y': '#32CD32', 'N': '#FF4500', 'M': '#8A2BE2'}
    )

    # Bar Chart
    bar_data = data[data['Topic Id'] == selected_topic]
    bar_data_grouped = bar_data.groupby(['Source', 'Response']).size().reset_index(name='Count')
    bar_fig = px.bar(
        bar_data_grouped,
        x='Source',
        y='Count',
        color='Response',
        barmode='group',
        title='Response Counts by Source',
        color_discrete_map={'Y': '#32CD32', 'N': '#FF4500', 'M': '#8A2BE2'}
    )

    # Metrics filtered by both dropdowns
    total_count = filtered_data.shape[0]
    yes_count = filtered_data[filtered_data['Response'] == 'Y'].shape[0]
    no_count = filtered_data[filtered_data['Response'] == 'N'].shape[0]
    maybe_count = filtered_data[filtered_data['Response'] == 'M'].shape[0]

    # Create Donut Charts for Metrics
    total_metric = create_donut_chart(total_count, total_count, '#6c757d')
    yes_metric = create_donut_chart(yes_count, total_count, '#32CD32')
    no_metric = create_donut_chart(no_count, total_count, '#FF4500')
    maybe_metric = create_donut_chart(maybe_count, total_count, '#8A2BE2')

    return pie_fig, bar_fig, total_metric, yes_metric, no_metric, maybe_metric

# Run the Dash app
if __name__ == '__main__':
    app.run_server(debug=True)
