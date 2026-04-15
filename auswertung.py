import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from dash import Dash, html, dcc, Input, Output
import os

# ================= DATA =================

table = 'table'
excel_files = [f for f in os.listdir(table) if f.endswith(('.xlsx', '.xls'))]

if not excel_files:
    raise FileNotFoundError("Keine Excel-Datei gefunden")

file_path = os.path.join(table, excel_files[0])
df = pd.read_excel(file_path)

if "Date" in df.columns:
    df["Date"] = pd.to_datetime(df["Date"])

# sichere Stress-Spalte
stress_candidates = [c for c in df.columns if 'stress' in c.lower()]
if not stress_candidates:
    raise ValueError("Keine Stress-Spalte gefunden")
stress_col = stress_candidates[0]

participant_col = 'Participant_ID'

numeric_cols = df.select_dtypes(include='number').columns.tolist()
if stress_col in numeric_cols:
    numeric_cols.remove(stress_col)

participants = sorted(df[participant_col].dropna().unique().tolist())
participants.append("Durchschnitt")

# ================= DASH =================

app = Dash(__name__)

graph_style = {
    "backgroundColor": "#111111",
    "padding": "10px",
    "border": "1px solid white",
    "borderRadius": "8px",
    "marginBottom": "15px"
}

app.layout = html.Div([

    html.H1("Stress Analyse Dashboard", style={"color": "white"}),

    # Sidebar
    html.Div([
        html.Label("Proband auswählen", style={"color": "white"}),
        dcc.Dropdown(id="participant",
                     options=[{"label": p, "value": p} for p in participants],
                     value=participants[0]),

        html.Br(),

        html.Label("Variable auswählen", style={"color": "white"}),
        dcc.Dropdown(id="variable",
                     options=[{"label": v, "value": v} for v in numeric_cols],
                     value=numeric_cols[0]),

    ], style={
        "width": "20%",
        "display": "inline-block",
        "verticalAlign": "top",
        "padding": "20px",
        "backgroundColor": "#111111",
        "borderRight": "1px solid #333"
    }),

    # MAIN
    html.Div([

        html.Div([
            html.Div(dcc.Graph(id="scatter"), style=graph_style),
            html.Div(dcc.Graph(id="heatmap"), style=graph_style),
            html.Div(dcc.Graph(id="histogram"), style=graph_style),
        ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"}),

        html.Div([
            html.Div(dcc.Graph(id="boxplot"), style=graph_style),
            html.Div(dcc.Graph(id="time"), style={**graph_style, "height": "300px"}),
        ], style={"width": "48%", "display": "inline-block", "verticalAlign": "top"}),

        html.Div(id="info-box", style={
            "marginTop": "15px",
            "padding": "10px",
            "border": "1px solid #333",
            "backgroundColor": "#111111",
            "color": "white",
            "width": "80%"
        })

    ], style={"width": "75%", "display": "inline-block", "padding": "20px"})

], style={"backgroundColor": "#000000"})

# ================= CALLBACK =================

@app.callback(
    Output("scatter", "figure"),
    Output("boxplot", "figure"),
    Output("heatmap", "figure"),
    Output("time", "figure"),
    Output("histogram", "figure"),
    Output("info-box", "children"),
    Input("participant", "value"),
    Input("variable", "value")
)
def update_dashboard(participant, variable):

    df_p = df.copy() if participant == "Durchschnitt" else df[df[participant_col] == participant]

    df_temp = df_p[[stress_col, variable]].dropna()

    total_before = len(df_temp)
    removed_stress = (df_temp[stress_col] == 0).sum()
    removed_var = (df_temp[variable] == 0).sum()

    df_filtered = df_temp[(df_temp[stress_col] != 0) & (df_temp[variable] != 0)].copy()

    total_after = len(df_filtered)
    total_removed = total_before - total_after

    # Scatter
    scatter_fig = px.scatter(
        df_filtered,
        x=stress_col,
        y=variable,
        trendline="ols",
        template="plotly_dark"
    )

    # Boxplot
    box_fig = go.Figure()
    if not df_filtered.empty:
        df_filtered['stress_bin'] = pd.cut(df_filtered[stress_col], bins=5).astype(str)
        box_fig.add_trace(go.Box(x=df_filtered['stress_bin'], y=df_filtered[variable]))
    box_fig.update_layout(template="plotly_dark")

    # Korrelation
    corr = df_p.select_dtypes(include='number').corr()
    stress_corr = corr[[stress_col]].drop(stress_col)

    heatmap_fig = go.Figure(go.Heatmap(
        z=stress_corr.values,
        x=["Stress"],
        y=stress_corr.index,
        colorscale='RdBu_r',
        zmin=-1,
        zmax=1,
        text=stress_corr.round(2),
        texttemplate='%{text}'
    ))
    heatmap_fig.update_layout(template="plotly_dark")

    # Zeitverlauf
    if "Date" in df_p.columns:
        time_fig = px.line(
            df_p,
            x="Date",
            y=[stress_col, variable],
            template="plotly_dark"
        )
        time_fig.update_layout(height=300)
    else:
        time_fig = go.Figure()

    # Histogramm
    histogram_fig = px.histogram(
        df_p,
        x=variable,
        nbins=20,
        template="plotly_dark"
    )

    # Info
    info_text = (
        f"Gesamtwerte: {total_before} | Verwendet: {total_after} | Entfernt: {total_removed} "
        f"(Stress=0: {removed_stress}, Variable=0: {removed_var})"
    )

    return scatter_fig, box_fig, heatmap_fig, time_fig, histogram_fig, info_text


if __name__ == "__main__":
    app.run(debug=True)