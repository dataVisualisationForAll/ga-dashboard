import pandas as pd
import plotly.express as px
import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output

# spreadsheet directory
workbook = './data/ga_rawdata.xlsx'

# read all excel sheets
all_dfs = pd.read_excel(workbook, sheet_name=None)

# drop unwanted sheet from dictionary
removed = all_dfs.pop('Copy of Report Configuration')

merged_df = pd.DataFrame()

# for iteration to modify all dataframes
for df in all_dfs:

    # drop possibly empty columns 
    all_dfs[df] = all_dfs[df].iloc[:,0:5]

    # rename columns (drop 'ga:')
    for col in all_dfs[df].columns:
        all_dfs[df].rename(columns={col : col[3:]}, inplace=True)
    
    # create month and year columns for easier manipulation & drop yearMonth   
    all_dfs[df]['year'] = all_dfs[df]['yearMonth'].apply(lambda x: str(x)[:4])
    all_dfs[df]['month'] = all_dfs[df]['yearMonth'].apply(lambda x: str(x)[4:])
    all_dfs[df].drop(columns=['yearMonth'], inplace=True)
    

    # merge all dataframes 
    all_dfs[df]['market'] = df

    merged_df = pd.concat(all_dfs, ignore_index=True)

# generate new metrics
merged_df['bounce_rate'] = merged_df['bounces']/merged_df['sessions'] 
merged_df['pageviews_per_session'] = merged_df['pageviews']/merged_df['sessions']
merged_df['sessions_per_user'] = merged_df['sessions']/merged_df['users'] 


# create summary dataframe
summary_df = merged_df.groupby(['market','year']).agg({
                    'users':'mean',
                    'bounces' : 'mean',
                    'pageviews' : 'mean',
                    'sessions' : 'mean',
                    'bounce_rate' : 'mean',
                    'sessions_per_user' : 'mean',
                    'pageviews_per_session' : 'mean'
                })

# rename columns as average
for col in summary_df.columns:
    summary_df.rename(columns={col:str('average_'+col)}, inplace=True)

# sort values based on market number
summary_df.reset_index(inplace=True)
summary_df['labels'] = summary_df.market.str[-2:]
summary_df = summary_df.sort_values('labels').drop('labels', axis=1)
summary_df.dropna(inplace=True)

# create a datetime column
merged_df.dropna(inplace=True)
merged_df['date'] = pd.to_datetime(merged_df[['year','month']].assign(DAY=1))


# declaring functions

# adds slider
def add_slider(fig):
    fig.update_layout(
        xaxis=dict(
            rangeselector=dict(
                buttons=list([
                    dict(count=1,
                        label="1m",
                        step="month",
                        stepmode="backward"),
                    dict(count=6,
                        label="6m",
                        step="month",
                        stepmode="backward"),
                    dict(count=1,
                        label="YTD",
                        step="year",
                        stepmode="todate"),
                    dict(count=1,
                        label="1y",
                        step="year",
                        stepmode="backward"),
                    dict(step="all")
                ])
            ),
            rangeslider=dict(
                visible=True
            ),
            type="date"
        )
    )


# App initialisation
app = dash.Dash(__name__)

server = app.server

app.layout = html.Div(children=[

    html.H1(children=['Interactive Dashboard', html.H6('A Web Application for data visualisation by Federico Denni')]),

    html.Div(
        className='first-chart-wrapper',
        children=[

            html.H4('Select Metric'),

            dcc.RadioItems(
                id='radio-metric-select',
                options=[
                    {'label':'Users', 'value':'users'},
                    {'label':'Sessions', 'value':'sessions'},
                    {'label':'Pageviews', 'value':'pageviews'},
                    {'label':'Bounces', 'value':'bounces'}                
              ],
                value='users'
           ),

            dcc.Graph(
                id='first-chart'
            )

        ]        
    ),

    html.Div(
        className='second-chart-wrapper',
        children=[
            html.H4('Select Metric'),

            dcc.RadioItems(
                id='radio-avg-metric-select',
                options=[
                    {'label':'Average Users', 'value':'average_users'},
                    {'label':'Average Sessions per User', 'value':'average_sessions_per_user'},
                    {'label':'Average Pageviews per Session', 'value':'average_pageviews_per_session'},
                    {'label':'Average Bounce Rate', 'value':'average_bounce_rate'}                
              ],
                value='average_users'
            ),

            dcc.Graph(
                id='second-chart'
            )

        ]
    )

  ]
)

@app.callback(
    Output('first-chart', 'figure'),
    [Input('radio-metric-select', 'value')]
)

def update_first_chart(metric):
    import plotly.express as px
    fig = px.line(merged_df,color='market',x='date',y=metric)
    add_slider(fig)
    return fig

@app.callback(
    Output('second-chart', 'figure'),
    [Input('radio-avg-metric-select', 'value')]
)

def update_second_chart(metric):
    import plotly.express as px
    fig = px.bar(summary_df, x='market', color='year', y=metric, barmode='group')
    return fig

if __name__ == '__main__':
    app.run_server(debug=False)