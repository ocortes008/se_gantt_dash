import pandas as pd
import plotly.express as px
import os
from dash import Dash, dcc, html, Input, Output, State
from flask import request
import dash_bootstrap_components as dbc
import dash_auth
import datetime
# Load Excel
#df = pd.read_excel(r"C:\Users\ocortes\OneDrive - Niagara Bottling, LLC\Files for new PC\SE Tracking\SE_Gantt_Chart.xlsx", sheet_name='Sheet3')
df = pd.read_excel(r"SE_Gantt_Chart.xlsx", sheet_name = 'Sheet3')
df.columns = df.columns.str.strip()
df['Project Status'] = df['Project Status'].str.strip()
#print("Columns in Excel:", df.columns.tolist())
# Parse dates
df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')
df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
# Remove rows with missing or invalid dates
df.loc[(df['Project Status'] == 'Pending Approval') & (df['End Date'].isna()), 'End Date'] = pd.Timestamp.today() + pd.Timedelta(days=1)
df.dropna(subset=['Start Date', 'End Date'], inplace=True)
# Filter out tasks with 0 duration (same start/end)
df = df[df['Start Date'] < df['End Date']]
# Confirm dates; this was just to confrim I was reading correctly from the excel sheet
# print(df[['Task', 'Start Date', 'End Date']])
# app setup
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server=app.server #for deployment
# Get unique filter options
plants = sorted(df["Plant"].dropna().unique())
project_names = sorted(df['Project Name'].dropna().unique())
project_types = sorted(df['Project Type'].dropna().unique())
project_status = sorted(df["Project Status"].dropna().unique())
resources = sorted(df['SE'].dropna().unique())
managers = sorted(df['SE Manager'].dropna().unique())



app.layout = dbc.Container([

    html.H2("E80 Gantt Chart", className="mt-3"),
    html.Button("Refresh Chart", id="refresh_button", n_clicks=0, className="btn btn-primary"),
    #Filters
    dbc.Row([
        dbc.Col([
            html.Label("Plant"),
            dcc.Dropdown(
                options=[{"label":d,"value":d} for d in plants],
                id="plant-filter",
                multi=True,
                placeholder="All Plants"
            )
        ]),
        dbc.Col([
               html.Label("Project Name"),
               dcc.Dropdown(
                options=[{"label":n,"value":n} for n in project_names],
                id="name-filter",
                multi=True,
                placeholder="All Projects"
            )  
        ]),
        dbc.Col([
               html.Label("Project Type"),
               dcc.Dropdown(
                options=[{"label":p,"value":p} for p in project_types],
                id="type-filter",
                multi=True,
                placeholder="All Project Types"
            )  
        ]),
        dbc.Col([
               html.Label("Project Status"),
               dcc.Dropdown(
                options=[{"label":s,"value":s} for s in project_status],
                id="status-filter",
                multi=True,
                placeholder="All Project Status"
            )  
        ]),
        dbc.Col([
               html.Label("SE"),
               dcc.Dropdown(
                options=[{"label":r,"value":r} for r in resources],
                id="resource-filter",
                multi=True,
                placeholder="All SE Resources"
            )  
        ]),
        dbc.Col([
               html.Label("SE Manager"),
               dcc.Dropdown(
                options=[{"label":m,"value":m} for m in managers],
                id="manager-filter",
                multi=True,
                placeholder="All SE Managers"
            )  
        ])
    ], className="mb-4"),
    
    dcc.Graph(id="gantt-chart"),


    #html.Hr(),
    #html.H4("Task Count by Resource"),
    #dash_table.DataTable(
    #    id="resource-task-table",
    #    columns=[
    #        {"name":"Resource","id":"Resource"},
    #        {"name":"Task Count", "id": "Task Count"}
    #    ],
    #    style_table={"overflowX":"auto"},
    #    style_cell={"textAlign":"left","padding":"8px"},
    #    style_header={"backgroundColor":"#f2f2f2", "fontweight": "bold"},
    #)
], fluid=True)
# Call back to update Gantt Chart based on filters
from dash import Output, Input
@app.callback(
    Output("gantt-chart", "figure"),
    Input("refresh_button", "n_clicks"),
    #Output("resource-task-table","data"),
    State("plant-filter", "value"),
    State("name-filter", "value"),
    State("type-filter", "value"),
    State("status-filter", "value"),
    State("resource-filter", "value"),
    State("manager-filter", "value"),
)
def update_chart(n_clicks, pnt, pname, ptype, pstatus, resource, manager):
    #df = pd.read_excel(r"C:\Users\ocortes\OneDrive - Niagara Bottling, LLC\Files for new PC\SE Tracking\SE_Gantt_Chart.xlsx", sheet_name = 'Sheet3')
    df = pd.read_excel(r"C:\Users\ocortes\Desktop\Python\SE_Gantt_Chart.xlsx", sheet_name = 'Sheet3')
    df.columns = df.columns.str.strip()
    df['Project Status'] = df['Project Status'].str.strip()
#print("Columns in Excel:", df.columns.tolist())
# Parse dates
    df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')
    df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
# Remove rows with missing or invalid dates
    df.loc[(df['Project Status'] == 'Pending Approval') & (df['End Date'].isna()), 'End Date'] = pd.Timestamp.today() + pd.Timedelta(days=1)
    df.dropna(subset=['Start Date', 'End Date'], inplace=True)
# Filter out tasks with 0 duration (same start/end)
    df = df[df['Start Date'] < df['End Date']]
    
    filtered_df = df.copy()
    if pnt:
        filtered_df = filtered_df[filtered_df["Plant"].isin(pnt)]
    if pname:
        filtered_df = filtered_df[filtered_df["Project Name"].isin(pname)]
    if ptype:
        filtered_df = filtered_df[filtered_df["Project Type"].isin(ptype)]
    if pstatus:
        filtered_df = filtered_df[filtered_df["Project Status"].isin(pstatus)]
    if resource:
        filtered_df = filtered_df[filtered_df["SE"].isin(resource)]
    if manager:
        filtered_df = filtered_df[filtered_df["SE Manager"].isin(manager)]

    if filtered_df.empty:
        return px.timeline(
            pd.DataFrame(dict(Task=[], Start=[], End=[])),
            x_start='Start', x_end='End', y='Project Name'
        ).update_layout(title="No tasks match the selected filters.")
    filtered_df["Label"] = (
        "SE: " + filtered_df["SE"] +
        "<br>Start: " + filtered_df["Start Date"].dt.strftime("%m/%d/%Y") +
        "<br>End: " + filtered_df["End Date"].dt.strftime("%m/%d/%Y")
    )
    filters_active = any([pnt, pname, ptype, pstatus, resource, manager])
# build gantt chart
    if filters_active:
        fig = px.timeline(
            filtered_df,
            x_start="Start Date",
            x_end = "End Date",
            y="Project Name",
            color = "Project Status",
            text='Label',
            color_discrete_map={"Pending Approval":"gray"}
        )

        fig.update_traces(
            textposition="inside",
            insidetextanchor="start",
            textfont=dict(
                size = 13,
                color="white",
                family="Arial Black"
            )
        )
        fig.update_traces(
            hoverlabel=dict(
                bgcolor="White",
                font=dict(
                    color="Black",
                    size = 14,
                    family="Arial Black"
                ),
                bordercolor="Black"
            )
        )
    else:
        fig = px.timeline(
            filtered_df,
            x_start="Start Date",
            x_end="End Date",
            y="Project Name",
            color="Project Status",
            title="E80 Gantt Chart",
            hover_data=['SE'],
            #text='SE',
            color_discrete_map ={
                "Pending Approval": "gray"
            }
        )
        fig.update_traces(
            hoverlabel=dict(
                bgcolor="White",
                font=dict(
                    color="Black",
                    size = 14,
                    family="Arial Black"
                ),
                bordercolor="Black"
            )
        )

    fig.update_yaxes(autorange="reversed")
    fig.update_layout(xaxis_tickformat="%m/%d/%Y")
    fig.update_layout(
        yaxis=dict(tickfont=dict(size=12)
                   )
    )
#build table

#    task_summary= (
#        filtered_df.groupby("Resource")
#        .size()
#        .reset_index(name="Task Count")
#        .sort_values(by="Task Count", ascending=False)
#    )
    fig.update_layout(
        height=max(200,20*len(df)),
        margin=dict(l=20,r=20,t=60,b=40)
    )
    return fig
#, task_summary.to_dict("records")

#Run the app

if __name__ == "__main__":
    app.run()

# Create chart
#fig = px.timeline(df, x_start='Start Date', x_end='End Date', y='Task', title="E80 Gantt Chart")
#fig.update_yaxes(autorange="reversed")
# Format dates and adjust zoom by day, removed this cause  .show() did a better job of zooming so the bars can actually show, when these lines were included I couldnt see bars
#fig.update_layout(
 #  xaxis_tickformat="%m/%d/%Y"
   #,xaxis_range=[
       #df['Start Date'].min() - pd.Timedelta(days=10),
       #df['End Date'].max() + pd.Timedelta(days=10)
   #]
#)
#fig.update_xaxes(tickformat="%m/%d/%Y", ticklabelmode="instant")
# Save image; this is to save the image/pdf on desktop, was having issues with the formatting, getting 1970 dates and no bars
#desktop_folder = os.path.join(os.path.expanduser("~"), "Desktop", "Python")
#os.makedirs(desktop_folder, exist_ok=True)
#fig.write_image(os.path.join(desktop_folder, "E80_Resource_Gantt_Chart.png"))

#fig.show()

#C:\Users\ocortes\Desktop\Python\gantt_env\gantt_chart.py