# =============================================================================
# Stress Analysis Dashboard (FINAL CLEAN VERSION)
# =============================================================================

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# -----------------------------------------------------------------------------
# LOAD DATA
# -----------------------------------------------------------------------------
table_folder = 'table'
excel_files = [f for f in os.listdir(table_folder) if f.endswith(('.xlsx', '.xls'))]
df = pd.read_excel(os.path.join(table_folder, excel_files[0]))

# -----------------------------------------------------------------------------
# DETECT STRESS
# -----------------------------------------------------------------------------
stress_col = [c for c in df.columns if 'stress' in c.lower()][0]
participant_col = 'Participant_ID'

numeric_cols = df.select_dtypes(include='number').columns.tolist()
numeric_cols.remove(stress_col)

participants = sorted(df[participant_col].dropna().unique().tolist())
participants.append("Durchschnitt")

# -----------------------------------------------------------------------------
# FIGURE
# -----------------------------------------------------------------------------
fig = make_subplots(
    rows=2, cols=2,
    subplot_titles=(
        "Zusammenhang (Scatterplot)",
        "Verteilung nach Stress (Boxplot)",
        "Stress-Korrelationen",
        ""
    ),
    horizontal_spacing=0.08,
    vertical_spacing=0.15
)

traces = []
trace_map = {}
valid_pairs = []

# -----------------------------------------------------------------------------
# CREATE TRACES (ALLE Kombinationen zulassen!)
# -----------------------------------------------------------------------------
for participant in participants:

    df_p = df if participant == "Durchschnitt" else df[df[participant_col] == participant]

    for var in numeric_cols:

        df_temp = df_p[[stress_col, var]].dropna()

        if len(df_temp) == 0:
            continue

        key = (participant, var)
        valid_pairs.append(key)

        # robuste Korrelation
        if len(df_temp) < 2 or df_temp[stress_col].std() == 0 or df_temp[var].std() == 0:
            corr = None
        else:
            corr = df_temp[stress_col].corr(df_temp[var])

        label = f"{participant}"
        if corr is not None:
            label += f" (r={corr:.2f})"
        else:
            label += " (wenig Daten)"

        # SCATTER
        fig.add_trace(go.Scatter(
            x=df_temp[stress_col],
            y=df_temp[var],
            mode='markers',
            visible=False,
            name=label,
            marker=dict(size=8, opacity=0.6)
        ), row=1, col=1)

        traces.append(fig.data[-1])

        # BOXPLOT
        df_temp['stress_bin'] = pd.cut(df_temp[stress_col], bins=5).astype(str)

        fig.add_trace(go.Box(
            x=df_temp['stress_bin'],
            y=df_temp[var],
            visible=False,
            showlegend=False
        ), row=1, col=2)

        traces.append(fig.data[-1])

        trace_map[key] = len(traces) - 2

# -----------------------------------------------------------------------------
# HEATMAP (Farbskala FIXED)
# -----------------------------------------------------------------------------
corr_matrix = df.select_dtypes(include='number').corr()
stress_corr = corr_matrix[[stress_col]].drop(stress_col)

fig.add_trace(go.Heatmap(
    z=stress_corr.values,
    x=[stress_col],
    y=stress_corr.index,
    colorscale='RdBu_r',
    zmin=-1,
    zmax=1,
    text=stress_corr.round(2),
    texttemplate='%{text}',
    colorbar=dict(
        title="r",
        len=0.4,         # klein
        thickness=12,    # schmal
        x=0.92,          # direkt neben Heatmap
        y=0.2            # vertikal passend
    )
), row=2, col=1)

# -----------------------------------------------------------------------------
# INITIAL STATE
# -----------------------------------------------------------------------------
current_participant, current_var = valid_pairs[0]

def update_visibility(p, v):
    vis = [False] * len(traces)
    for (pp, vv) in valid_pairs:
        if pp == p and vv == v:
            idx = trace_map[(pp, vv)]
            vis[idx] = True
            vis[idx + 1] = True
    return vis

initial_vis = update_visibility(current_participant, current_var)

for i, v in enumerate(initial_vis):
    traces[i].visible = v

# -----------------------------------------------------------------------------
# DROPDOWN: PARTICIPANTS
# -----------------------------------------------------------------------------
participant_buttons = []

for p in participants:
    participant_buttons.append(dict(
        label=p,
        method="update",
        args=[{"visible": update_visibility(p, current_var)}]
    ))

# -----------------------------------------------------------------------------
# DROPDOWN: VARIABLES
# -----------------------------------------------------------------------------
variable_buttons = []

for var in numeric_cols:
    variable_buttons.append(dict(
        label=var,
        method="update",
        args=[
            {"visible": update_visibility(current_participant, var)},
            {
                "yaxis.title.text": var,
                "yaxis2.title.text": var
            }
        ]
    ))

# -----------------------------------------------------------------------------
# LAYOUT
# -----------------------------------------------------------------------------
fig.update_layout(
    title="Stress Analyse Dashboard",
    template='plotly_white',
    height=950,
    width=1400,

    updatemenus=[
        dict(buttons=participant_buttons, x=0.2, y=1.15),
        dict(buttons=variable_buttons, x=0.6, y=1.15)
    ]
)

# -----------------------------------------------------------------------------
# AXES
# -----------------------------------------------------------------------------
fig.update_xaxes(title_text="Stress Level", row=1, col=1)
fig.update_xaxes(title_text="Stress", row=1, col=2)

fig.update_yaxes(title_text=current_var, row=1, col=1)
fig.update_yaxes(title_text=current_var, row=1, col=2)

# -----------------------------------------------------------------------------
# LEGENDE (unten rechts)
# -----------------------------------------------------------------------------
fig.add_annotation(
    text=(
        "<b>Legende & Interpretation</b><br><br>"

        "<b>Scatterplot:</b><br>"
        "Jeder Punkt entspricht einer Messung.<br>"
        "X-Achse = Stresslevel<br>"
        "Y-Achse = gewählte Variable<br>"
        "→ zeigt direkte Zusammenhänge<br><br>"

        "<b>Boxplot:</b><br>"
        "Gruppierung nach Stresslevel.<br>"
        "Linie = Median<br>"
        "Box = mittlere 50% der Werte<br>"
        "→ zeigt Verteilung & Streuung<br><br>"

        "<b>Korrelation (r):</b><br>"
        "+1 = starker positiver Zusammenhang<br>"
        "-1 = starker negativer Zusammenhang<br>"
        "0 = kein Zusammenhang<br><br>"

        "<b>Interpretation:</b><br>"
        "Steigende Punkte → Variable steigt mit Stress<br>"
        "Fallende Punkte → Variable sinkt mit Stress<br>"
        "Starke Streuung → schwacher Zusammenhang"
    ),
    x=1.02,
    y=0,
    xref="paper",
    yref="paper",
    showarrow=False,
    align="left",
    bordercolor="black",
    borderwidth=1,
    bgcolor="white",
    font=dict(size=11)
)
# -----------------------------------------------------------------------------
# KALENDER FUNKTION
# -----------------------------------------------------------------------------
df['Date'] = pd.to_datetime(df['Date'])
df['Month'] = df['Date'].dt.to_period('M')

def create_calendar(df_input, month):

    df_m = df_input[df_input['Month'] == month]

    if df_m.empty:
        return go.Figure()

    df_m['Day'] = df_m['Date'].dt.day
    df_m['Weekday'] = df_m['Date'].dt.weekday

    import numpy as np

    first_day = df_m['Date'].min().replace(day=1)
    first_weekday = first_day.weekday()
    days_in_month = month.days_in_month

    weeks = int(np.ceil((first_weekday + days_in_month) / 7))
    z = np.full((weeks, 7), np.nan)

    for _, row in df_m.iterrows():
        d = row['Day']
        w = (first_weekday + d - 1) // 7
        wd = row['Weekday']

        z[w, wd] = row[stress_col]

    fig_cal = go.Figure(data=go.Heatmap(
        z=z,
        x=['Mo','Di','Mi','Do','Fr','Sa','So'],
        y=[f"W{i+1}" for i in range(weeks)],
        colorscale='RdYlGn_r',
        zmin=df[stress_col].min(),
        zmax=df[stress_col].max(),
        colorbar=dict(title="Stress")
    ))

    fig_cal.update_layout(title=f"Kalender {month}")

    return fig_cal

# -----------------------------------------------------------------------------
# SHOW
# -----------------------------------------------------------------------------
fig.show()