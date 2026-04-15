import pandas as pd
import plotly.graph_objects as go
from dash import Dash, html, dcc, Input, Output
import os

# Lade die Excel-Datei
table = 'table'
excel_files = [f for f in os.listdir(table) if f.endswith(('.xlsx', '.xls'))]

if not excel_files:
    raise FileNotFoundError("Keine Excel-Datei gefunden")

df = pd.read_excel(os.path.join(table, excel_files[0]))

# Spalten identifizieren
stress_col = [c for c in df.columns if 'stress' in c.lower()][0]
participant_col = 'Participant_ID'

numeric_cols = df.select_dtypes(include='number').columns.tolist()
numeric_cols.remove(stress_col)

participants = sorted(df[participant_col].dropna().unique().tolist())
participants.append("Durchschnitt")

# Dash
app = Dash(__name__)

# Layout
app.layout = html.Div([

    html.H1("Stress Analyse Dashboard"),

    # Sidebar
    html.Div([

        html.Label("Proband auswählen"),
        dcc.Dropdown(
            id="participant",
            options=[{"label": p, "value": p} for p in participants],
            value=participants[0]
        ),

        html.Br(),

        html.Label("Variable auswählen"),
        dcc.Dropdown(
            id="variable",
            options=[{"label": v, "value": v} for v in numeric_cols],
            value=numeric_cols[0]
        ),

    ], style={
        "width": "20%",
        "display": "inline-block",
        "verticalAlign": "top",
        "padding": "20px",
        "backgroundColor": "#f4f4f4",
        "borderRight": "1px solid #ccc"
    }),

    # Main
    html.Div([

        dcc.Graph(id="scatter"),
        dcc.Graph(id="boxplot"),
        dcc.Graph(id="heatmap"),

        html.Div(
            id="info-box",
            style={
                "marginTop": "15px",
                "padding": "10px",
                "border": "1px solid #ccc",
                "backgroundColor": "#fafafa",
                "width": "80%"
            }
        )

    ], style={
        "width": "75%",
        "display": "inline-block",
        "padding": "20px"
    })

])

# Callback
@app.callback(
    Output("scatter", "figure"),
    Output("boxplot", "figure"),
    Output("heatmap", "figure"),
    Output("info-box", "children"),
    Input("participant", "value"),
    Input("variable", "value")
)
def update_dashboard(participant, variable):

    # Probanden filtern 
    if participant == "Durchschnitt":
        df_p = df.copy()
    else:
        df_p = df[df[participant_col] == participant]

    df_temp = df_p[[stress_col, variable]].dropna()

    total_before = len(df_temp)

    # Nullwerte entfernen und zählen
    removed_stress = (df_temp[stress_col] == 0).sum()
    removed_var = (df_temp[variable] == 0).sum()

    df_filtered = df_temp[
        (df_temp[stress_col] != 0) &
        (df_temp[variable] != 0)
    ]

    total_after = len(df_filtered)
    total_removed = total_before - total_after

    # Scatterplot Stress vs Variable
    scatter_fig = go.Figure()

    scatter_fig.add_trace(go.Scatter(
        x=df_filtered[stress_col],
        y=df_filtered[variable],
        mode='markers',
        marker=dict(size=8, opacity=0.7)
    ))

    scatter_fig.update_layout(
        title=f"{variable} vs Stress ({participant})",
        xaxis_title="Stress",
        yaxis_title=variable,
        template="plotly_white"
    )

    # Boxplot nach Stress
    box_fig = go.Figure()

    if len(df_filtered) > 0:
        df_filtered['stress_bin'] = pd.cut(df_filtered[stress_col], bins=5).astype(str)

        box_fig.add_trace(go.Box(
            x=df_filtered['stress_bin'],
            y=df_filtered[variable]
        ))

    box_fig.update_layout(
        title=f"Verteilung von {variable} nach Stress",
        xaxis_title="Stress Gruppen",
        yaxis_title=variable
    )

    # Heatmap der Korrelationen
    corr = df.select_dtypes(include='number').corr()
    stress_corr = corr[[stress_col]].drop(stress_col)

    heatmap_fig = go.Figure(go.Heatmap(
        z=stress_corr.values,
        x=[stress_col],
        y=stress_corr.index,
        colorscale='RdBu_r',
        zmin=-1,
        zmax=1,
        text=stress_corr.round(2),
        texttemplate='%{text}',
        colorbar=dict(title="r", len=0.4)
    ))

    heatmap_fig.update_layout(title="Stress-Korrelationen")

    # Info-Text
    info_text = (
        f"Gesamtwerte: {total_before} | "
        f"Verwendet: {total_after} | "
        f"Entfernt: {total_removed} "
        f"(Stress=0: {removed_stress}, Variable=0: {removed_var})"
    )

    return scatter_fig, box_fig, heatmap_fig, info_text


if __name__ == "__main__":
    app.run(debug=True)