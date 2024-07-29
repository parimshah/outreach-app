import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
from dash_table import DataTable
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Initialize the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.config.suppress_callback_exceptions = True

# Set up Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'C://Users//rohin//Documents//gen//Edvance AI//Python Scripts//edvance-ai-outreach-app-838a8051589e.json', scope)
client = gspread.authorize(creds)


# Function to fetch data from a specific sheet
def fetch_sheet_data(sheet_name):
    try:
        sheet = client.open("Marketing Contacts").worksheet(sheet_name)
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        df.reset_index(inplace=True)
        df.rename(columns={"index": "Row Index"}, inplace=True)
        return df
    except gspread.SpreadsheetNotFound:
        return pd.DataFrame()  # Return an empty DataFrame if the spreadsheet is not found
    except Exception as e:
        print(f"An error occurred: {e}")
        return pd.DataFrame()


# Define the app layout
app.layout = html.Div([
    dbc.Container([
        html.H1("Edvance AI Outreach App", className='text-center mb-4', style={
            'font-family': 'Montserrat, sans-serif', 'color': '#6ca1cd', 'font-weight': 'bold', 'margin-top': '30px',
            'text-shadow': '2px 2px 4px rgba(0, 0, 0, 0.2)'}),

        dbc.Row([
            dbc.Col([
                html.Div([
                    html.H3("Automated Email", style={'font-family': 'Montserrat, sans-serif', 'color': '#6ca1cd', 'text-align': 'center'}),
                    html.Label("Select Contacts", style={'font-weight': 'bold', 'font-size': '18px', 'margin-bottom': '10px', 'font-family': 'Arial, sans-serif'}),
                    html.Br(),
                    html.Div([
                        dbc.Button("Row Index                              ➔", id='row-button', style={'background-color': 'transparent', 'border': 'none', "color": "Black", 'text-align': 'left', 'width': '100%', 'box-shadow': '0px 4px 6px rgba(0, 0, 0, 0.1)', 'position': 'relative'}),
                        html.Div(id='row-div', style={'display': 'none', 'margin-top': '20px'}, children=[
                            dcc.Input(id='row-range-start', type='number', placeholder='Start row', style={'width': '20%', 'display': 'inline-block', 'margin-right': '2%', 'margin-top': '10px', 'padding': '5px'}),
                            dcc.Input(id='row-range-end', type='number', placeholder='End row', style={'width': '20%', 'display': 'inline-block', 'padding': '5px', 'margin-top': '10px'}),
                        ]),
                    ], style={'margin-bottom': '10px'}),

                    html.Div([
                        dbc.Button("Filter  ➔", id='filter-button', style={'background-color': 'transparent', 'border': 'none', "color": "Black", 'text-align': 'left', 'width': '100%', 'box-shadow': '0px 4px 6px rgba(0, 0, 0, 0.1)', 'position': 'relative'}),
                        html.Div(id='filter-div', style={'display': 'none', 'margin-top': '20px'}, children=[
                            html.Label("Coming soon.", style={'font-size': '12px', 'margin-bottom': '10px', 'font-family': 'Arial, sans-serif', 'margin-top': '10px'}),
                        ]),
                    ], style={'margin-bottom': '10px'}),

                    html.Div([
                        dbc.Button("Manual  ➔", id='manual-button', style={'background-color': 'transparent', 'border': 'none', "color": "Black", 'text-align': 'left', 'width': '100%', 'box-shadow': '0px 4px 6px rgba(0, 0, 0, 0.1)', 'position': 'relative'}),
                        html.Div(id='manual-div', style={'display': 'none', 'margin-top': '20px'}, children=[
                            html.Label("Coming soon.", style={'font-size': '12px', 'margin-bottom': '10px', 'font-family': 'Arial, sans-serif', 'margin-top': '10px'}),
                        ]),
                    ], style={'margin-bottom': '10px'}),

                    html.Label("Subject", style={'font-weight': 'bold', 'font-size': '18px', 'margin-bottom': '10px', 'font-family': 'Arial, sans-serif', 'margin-top': '10px'}),
                    dcc.Input(id='email-subject', type='text', placeholder='Enter email subject', style={'width': '100%', 'border-radius': '15px', 'padding': '10px', "margin-bottom": '10px'}),
                    html.Label("Body", style={'font-weight': 'bold', 'font-size': '18px', 'margin-bottom': '10px', 'font-family': 'Arial, sans-serif'}),
                    html.Label("Use {School} and {Name} for dynamic text.", style={'font-size': '14px', 'margin-bottom': '10px', 'font-family': 'Arial, sans-serif', 'margin-top': '10px', 'color': '#6ca1cd', 'margin-left': '5px'}),
                    dcc.Textarea(id='email-body', placeholder='Enter email body', style={'width': '100%', 'height': '150px', 'border-radius': '15px', 'padding': '10px', "margin-bottom": '10px'}),

                    dbc.Button("Send Emails", id='send-button', style={'background-color': '#6ca1cd', 'border-color': '#6ca1cd'}, className='w-100'),
                ], style={'border-radius': '15px', 'padding': '20px', 'background': '#f9f9f9', 'box-shadow': '0 4px 8px rgba(0, 0, 0, 0.1)'}),
            ], width=6),

            dbc.Col([
                html.Div([
                    html.H3("Google Sheets", style={'font-family': 'Montserrat, sans-serif', 'color': '#6ca1cd', 'text-align': 'center'}),
                    dcc.Dropdown(
                        id='data-table-dropdown',
                        options=[
                            {'label': 'US CS Teachers', 'value': 'US CS Teachers'},
                            {'label': 'US Counselors', 'value': 'US Counselors'},
                            {'label': 'Int. CS Teachers', 'value': 'Int. CS Teachers'},
                        ],
                        value='US CS Teachers',
                        style={'border-radius': '10px', 'margin-bottom': '20px'}
                    ),
                    html.Div(id='data-table-container')
                ], style={'border-radius': '15px', 'padding': '20px', 'background': '#f9f9f9', 'box-shadow': '0 4px 8px rgba(0, 0, 0, 0.1)'})
            ], width=6),

            dbc.Col([
                html.Div([
                    html.H3("Email Log", style={'font-family': 'Montserrat, sans-serif', 'color': '#6ca1cd', 'text-align': 'center'}),
                    html.Div(id='email-log', style={'border-radius': '15px', 'padding': '20px'})
                ], style={'border-radius': '15px', 'padding': '20px', 'background': '#f9f9f9', 'box-shadow': '0 4px 8px rgba(0, 0, 0, 0.1)', "width": "49%", "margin-top":"10px"})
            ], width=12)
        ], className='mb-3'),
    ], fluid=True),
])


# Callback to update the data table based on dropdown selection
@app.callback(
    Output('data-table-container', 'children'),
    [Input('data-table-dropdown', 'value'),
     Input('manual-button', 'n_clicks')]
)
def update_data_table(selected_table, manual_clicks):
    df = fetch_sheet_data(selected_table)
    if not df.empty:
        df.reset_index(inplace=True)
        df.rename(columns={"index": "Row Index"}, inplace=True)
        columns = [{"name": i, "id": i} for i in df.columns]
        data = df.to_dict('records')
        if manual_clicks and manual_clicks % 2 == 1:
            columns.insert(0, {"name": "Select", "id": "select", "presentation": "checkbox"})
            data = [{"select": False, **item} for item in data]
        return DataTable(
            id='data-table',
            columns=columns[1:],  # Skip the first column which is 'Row Index'
            data=data,
            row_selectable='multi' if manual_clicks and manual_clicks % 2 == 1 else False,
            style_table={'overflowX': 'auto'},
            style_header={'backgroundColor': '#6ca1cd', 'fontWeight': 'bold', 'color': 'white'},
            style_cell={'textAlign': 'left', 'font-family': 'Arial, sans-serif', 'fontSize': '14px', 'padding': '10px'},
            page_size=10,
            style_data_conditional=[
                {
                    'if': {'column_id': 'Row Index'},
                    'backgroundColor': '#f2f2f2',
                    'fontWeight': 'bold'
                }
            ]
        )
    return html.Div("No data available")


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Define the function to send an email
def send_email(recipient, subject, body):
    try:
        sender_email = 'shahparim3@gmail.com'
        sender_password = 'popj xnos zmuh ovvv'

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        return "Sent"
    except Exception as e:
        return f"Error: {str(e)}"


# Define callback to handle sending email
@app.callback(
    [Output('send-button', 'children'),
     Output('email-log', 'children')],
    [Input('send-button', 'n_clicks')],
    [State('data-table', 'data'),
     State('data-table', 'selected_rows'),
     State('row-range-start', 'value'),
     State('row-range-end', 'value'),
     State('email-subject', 'value'),
     State('email-body', 'value'),
     State('email-log', 'children')]
)
def handle_send_email(n_clicks, data, selected_rows, row_start, row_end, subject, body, log):
    df = pd.DataFrame(data)
    rows_to_email = selected_rows if selected_rows else list(range(row_start, row_end + 1))
    new_log_entries = []

    for idx in rows_to_email:
        if idx < len(df):
            recipient = df.iloc[idx]["Email"]
            name = df.iloc[idx]["Name"]
            body_with_placeholder = body.format(Name=name, School=df.iloc[idx]["School"])
            send_email(recipient, subject, body_with_placeholder)
            new_log_entries.append(f"Email sent to {name} at {recipient}")

    log_entries = log.split('\n') if log else []
    new_log_entries = [entry for entry in new_log_entries if entry]  # Remove empty entries if any
    log_entries.extend(new_log_entries)
    updated_log = '\n'.join(log_entries)
    return "Emails Sent", updated_log


# Callback to show/hide divs based on button clicks
@app.callback(
    [Output('row-div', 'style'),
     Output('filter-div', 'style'),
     Output('manual-div', 'style')],
    [Input('row-button', 'n_clicks'),
     Input('filter-button', 'n_clicks'),
     Input('manual-button', 'n_clicks')]
)
def toggle_divs(row_clicks, filter_clicks, manual_clicks):
    row_div_style = {'display': 'block'} if row_clicks and row_clicks % 2 == 1 else {'display': 'none'}
    filter_div_style = {'display': 'block'} if filter_clicks and filter_clicks % 2 == 1 else {'display': 'none'}
    manual_div_style = {'display': 'block'} if manual_clicks and manual_clicks % 2 == 1 else {'display': 'none'}
    return row_div_style, filter_div_style, manual_div_style


# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
