import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px
from dash import dash_table
from PyQt5.QtCore import QLocale
from excel_loader import run_excel_loader

QLocale.setDefault(QLocale(QLocale.English, QLocale.UnitedStates))

# Load and preprocess data
df = pd.read_excel('official_report_data.xlsx')

# Convert 'Date' column to datetime
df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y')

# Define fiscal year function
def fiscal_year(date):
    year = date.year
    if date.month < 9:
        return str(year - 1)
    else:
        return str(year)

# Apply fiscal year function
df['Fiscal Year'] = df['Date'].apply(fiscal_year)

# Extract month and year from 'Date'
df_sorted_fiscal_year = df.sort_values(by="Date")
df['Month-Year'] = df['Date'].dt.to_period('M').astype(str)
df_sorted_item_detial = df.sort_values(by="Item Detail")

df['Date'] = pd.to_datetime(df['Date'])
df['Year'] = df['Date'].dt.year
df['Month'] = df['Date'].dt.strftime('%b')  # Get the month name as string (e.g., 'Jan', 'Feb')

df = df.sort_values(['Site', 'Item', 'Item Detail', 'Fiscal Year'])

# Create a pivot table
df_grid = df.pivot_table(index=['Site', 'Item', 'Item Detail', 'Fiscal Year'],
                         columns='Month',
                         values='Amount',  # Assuming 'Amount' is the column with the numeric data
                         aggfunc='sum').reset_index()

# Reorder the columns to match the desired output
month_order = ['Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug']
df_grid = df_grid[['Site', 'Item', 'Item Detail', 'Fiscal Year'] + month_order]

# Infer proper dtypes first to avoid deprecation warnings
df_grid = df_grid.infer_objects(copy=False)

# Now safely replace NaN values with 0
df_grid.fillna(0, inplace=True)

# Add a "Total" column to sum from "Sep" to "Aug"
df_grid['Total'] = df_grid[month_order].sum(axis=1)

# Ensure numerical columns remain numeric
df_grid[month_order + ['Total']] = df_grid[month_order + ['Total']].apply(pd.to_numeric, errors='coerce')

# Store a separate formatted version for display
df_grid_display = df_grid.copy()
df_grid_display[month_order + ['Total']] = df_grid_display[month_order + ['Total']].apply(lambda col: col.map(lambda x: '{:,}'.format(int(x)) if x != 0 else ''))

# Initialize Dash app
app = dash.Dash(__name__)

# App layout: Display charts and data grid side by side
app.layout = html.Div([
    html.Div([
        dcc.Dropdown(
            id='site-selector1',
            options=[{'label': site, 'value': site} for site in df['Site'].unique()],
            value='SDCT',  # Set "SDCT" as the default value
            multi=True,
            style={'margin-bottom': '10px'}
        ),
        dcc.Dropdown(
            id='item-detail-selector1',
            options=[{'label': item_detail, 'value': item_detail} for item_detail in df_sorted_item_detial['Item Detail'].unique()],
            value=['[1003] Revenue Total', '[1027] Gross Profit'],
            multi=True,
            style={'margin-bottom': '10px'}
        ),
        dcc.Dropdown(
            id='fiscal-year-filter',
            options=[{'label': year, 'value': year} for year in df_sorted_fiscal_year['Fiscal Year'].unique()],
            value='none',
            multi=True,
            style={'margin-bottom': '10px'}
        )        
    ], style={'width': '100%', 'padding': '10px'}),  # Full width dropdown
    
    # Main content area: charts and data grid
    html.Div([
        html.Div([
            dcc.Graph(id='line-chart1', style={'height': '400px'})  # Set height for the line chart
        ], style={'width': '65%', 'display': 'inline-block', 'padding': '10px'}),  # Left side for charts
        
        html.Div([
            dcc.Graph(id='bar-chart1', style={'height': '350px'})  # Set height for the bar chart
        ], style={'width': '35%', 'display': 'inline-block', 'padding': '10px'})  # Right side for bar chart
    ], style={'display': 'flex', 'justify-content': 'space-between', 'padding': '10px'}),
    
    html.Div([
        dash_table.DataTable(
            id='data-table',
            columns=[{"name": i, "id": i} for i in df_grid.columns],
            data=df_grid.to_dict('records'),
            page_size=100,  # Set page size to 100 rows per page
            fixed_rows={'headers': True},  # Fix headers to enable scrolling
            style_table={'overflowy': 'auto'},
            filter_action='none',  # Disable filter actions

            # Style for the header
            style_header={
                'backgroundColor': '#0070C0',  # Set header background color (e.g., blue)
                'color': 'white',  # Set font color (e.g., white)
                'fontWeight': 'bold',  # Make header font bold
                'textAlign': 'center'  # Center align the header text
            },
            
            # Style for the cell
            style_cell_conditional=[
                {'if': {'column_id': 'Site'}, 'textAlign': 'center', 'width': '80px'},
                {'if': {'column_id': 'Item'}, 'textAlign': 'center', 'width': '100px'},
                {'if': {'column_id': 'Item Detail'}, 'textAlign': 'left', 'width': '250px'},
                {'if': {'column_id': 'Fiscal Year'}, 'textAlign': 'center', 'width': '100px'},
                {'if': {'column_id': 'Sep'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Oct'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Nov'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Dec'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Jan'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Feb'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Mar'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Apr'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'May'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Jun'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Jul'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Aug'}, 'textAlign': 'right', 'width': '100px'},
                {'if': {'column_id': 'Total'}, 'textAlign': 'right', 'width': '100px'},

            ]
        )
    ], style={'width': '100%', 'padding': '10px'})
])

# Unified callback for updating charts and data grid
@app.callback(
    [Output('line-chart1', 'figure'),
     Output('bar-chart1', 'figure'),
     Output('data-table', 'data')],
    [Input('site-selector1', 'value'),
     Input('item-detail-selector1', 'value'),
     Input('fiscal-year-filter', 'value')]
)
def update_content(selected_sites, selected_item_details, selected_fiscal_years):
    # Handle blank selections by assigning all data when a filter is left blank
    if not selected_sites:  # Handle empty or None selection
        selected_sites = df['Site'].unique().tolist()
    elif isinstance(selected_sites, str):
        selected_sites = [selected_sites]
    
    if not selected_item_details:
        selected_item_details = df['Item Detail'].unique().tolist()
    elif isinstance(selected_item_details, str):
        selected_item_details = [selected_item_details]
    
    if not selected_fiscal_years or 'none' in selected_fiscal_years:
        selected_fiscal_years = df['Fiscal Year'].unique().tolist()
    elif isinstance(selected_fiscal_years, str):
        selected_fiscal_years = [selected_fiscal_years]
    
    # Filter data for the charts
    filtered_df = df[
        (df['Site'].isin(selected_sites)) & 
        (df['Item Detail'].isin(selected_item_details)) &
        (df['Fiscal Year'].isin(selected_fiscal_years))
    ]
    
    # Create the line chart with value labels in millions
    fig1 = px.line(
        filtered_df,
        x='Month-Year',
        y='Amount',
        color='Item Detail',
        markers=True
    )
    
    fig1.update_traces(
        texttemplate='%{y:.1f}M',  # Display the value divided by 1,000,000 with 'M'
        textposition='top center'
    )
    
    fig1.update_layout(
        title='Monthly Topic',
        xaxis_title='Year',
        yaxis_title='Amount (in M)',
        xaxis_tickformat='%b %Y',
        xaxis_tickangle=-45,
        showlegend=False,  # Hide the legend
        margin=dict(l=40, r=40, t=40, b=40)  # Adjust margins for better visibility
    )
    
    # Filter data for the bar chart (sum of each fiscal year and item detail)
    filtered_df2 = filtered_df.groupby(['Fiscal Year', 'Item Detail'], as_index=False)['Amount'].sum()
    
    # Convert amounts to millions
    filtered_df2['Amount'] = filtered_df2['Amount'] / 1_000_000

    # Create the bar chart
    fig2 = px.bar(
        filtered_df2,
        x='Fiscal Year',
        y='Amount',
        color='Item Detail',  # Add color by 'Item Detail' for series representation
        text='Amount'
    )
    fig2.update_traces(
        texttemplate='%{y:.1f}M',  # Format with thousands separator and 'M' for millions
        textposition='outside'
    )
    
    fig2.update_layout(
        title='Total Fiscal Year',
        xaxis_title='Fiscal Year',
        yaxis_title='Total Amount (in M)',
        margin=dict(l=40, r=40, t=40, b=40),  # Adjust margins for better visibility
        showlegend=False  # Hide the legend
    )
    
    # Update the data table
    df_grid_filtered = df_grid.copy()
    
    # Apply filters to the grid data
    if selected_sites:
        df_grid_filtered = df_grid_filtered[df_grid_filtered['Site'].isin(selected_sites)]
    if selected_item_details:
        df_grid_filtered = df_grid_filtered[df_grid_filtered['Item Detail'].isin(selected_item_details)]
    if selected_fiscal_years:
        df_grid_filtered = df_grid_filtered[df_grid_filtered['Fiscal Year'].isin(selected_fiscal_years)]
    
    df_grid_filtered = df_grid_filtered.sort_values(['Site', 'Item Detail', 'Fiscal Year'], ascending=True)
    

    data=df_grid_display.to_dict('records')
    return fig1, fig2, data

server = app.server  # 👈 ADD THIS LINE at the global level for Render to find the server

if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=10000, debug=False)
