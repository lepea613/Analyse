import pandas as pd
import numpy as np
import plotly.graph_objects as go

# Daten laden
df = pd.read_excel(r"C:\Studium\EMI\Analyse\table\SleepStress_P1P2P6_daily.xlsx")
# Datum vorbereiten
df["Date"] = pd.to_datetime(df["Date"])

# Spaltennamen kürzen
df = df.rename(columns={
    "Stress Score_STRESS_SCORE":                                                  "stress_score",
    "daily_heart_rate_variability_average heart rate variability milliseconds":   "hrv",
    "daily_resting_heart_rate_beats per minute":                                  "resting_hr",
    "daily_respiratory_rate_breaths per minute":                                  "resp_rate",
    "daily_sleep_temperature_derivations_nightly temperature celsius":            "skin_temp",
    "daily_oxygen_saturation_average percentage":                                 "spo2",
    "sleep_score_overall_score":                                                  "sleep_score",
    "sleep_score_restlessness":                                                   "restlessness",
    "UserSleeps_minutes_asleep":                                                  "minutes_asleep",
    "TagebuchEntries_Schlafqualitat":                                             "sleep_quality_diary",
})

# Schlafdauer in Stunden
df["sleep_hours"] = df["minutes_asleep"] / 60

# Kalender-Keys
iso = df["Date"].dt.isocalendar()
df["year_week"] = iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
df["weekday"] = df["Date"].dt.weekday

# Metriken mit Farbskalen und Wertebereichen
metrics = {
    "Stress Score":        ("stress_score",        "RdYlGn_r",  0,   100),
    "HRV (ms)":            ("hrv",                 "RdYlGn",   10,    50),
    "Ruheherzrate":        ("resting_hr",          "RdYlGn_r", 45,    85),
    "Atemfrequenz":        ("resp_rate",           "RdYlGn_r", 10,    18),
    "Hauttemperatur (°C)": ("skin_temp",           "RdBu_r",   -2,     2),
    "SpO₂ (%)":            ("spo2",                "RdYlGn",   92,   100),
    "Sleep Score":         ("sleep_score",         "RdYlGn",   40,   100),
    "Unruhe":              ("restlessness",        "RdYlGn_r",  0,   0.5),
    "Schlafdauer (h)":     ("sleep_hours",         "Blues",     4,    10),
}

# Probanden + Gesamt
participants = ["Apple", "Orange", "Pear"]
views = participants + ["Alle (Ø)"]

def build_matrix(data, metric_col):
    """Baut Kalendermatrix (7 Wochentage × n Wochen) für einen Datensatz."""
    year_weeks = sorted(data["year_week"].unique())
    yw_index = {yw: i for i, yw in enumerate(year_weeks)}

    matrix = np.full((7, len(year_weeks)), np.nan)
    hover = [["" for _ in range(len(year_weeks))] for _ in range(7)]

    for _, row in data.iterrows():
        if row["year_week"] not in yw_index:
            continue
        w = yw_index[row["year_week"]]
        d = int(row["weekday"])
        matrix[d, w] = row[metric_col] if pd.notna(row[metric_col]) else np.nan

        def fmt(val, unit=""):
            return f"{val:.1f}{unit}" if pd.notna(val) else "–"

        hover[d][w] = (
            f"<b>{row['Date'].strftime('%d.%m.%Y')}</b><br>"
            f"Stress Score: {fmt(row['stress_score'])}<br>"
            f"HRV: {fmt(row['hrv'],' ms')}<br>"
            f"Ruheherzrate: {fmt(row['resting_hr'],' bpm')}<br>"
            f"Atemfrequenz: {fmt(row['resp_rate'],' /min')}<br>"
            f"SpO₂: {fmt(row['spo2'],' %')}<br>"
            f"Sleep Score: {fmt(row['sleep_score'])}<br>"
            f"Schlafdauer: {fmt(row['sleep_hours'],' h')}<br>"
            f"Unruhe: {fmt(row['restlessness'])}"
        )

    return matrix, hover, year_weeks, yw_index

def get_month_ticks(data, yw_index):
    """Monatsnamen auf X-Achse."""
    tmp = data.copy()
    tmp["month_label"] = tmp["Date"].dt.strftime("%b %Y")
    ticks = (
        tmp.sort_values("Date")
        .groupby("month_label", sort=False)
        .first()
        .reset_index()[["month_label", "year_week"]]
    )
    vals = [yw_index[yw] for yw in ticks["year_week"] if yw in yw_index]
    texts = ticks["month_label"].tolist()[:len(vals)]
    return vals, texts

# Für "Alle (Ø)": Tagesdurchschnitt über alle Probanden
df_mean = df.groupby("Date")[
    [m[0] for m in metrics.values()]
].mean().reset_index()
iso_mean = df_mean["Date"].dt.isocalendar()
df_mean["year_week"] = iso_mean["year"].astype(str) + "-W" + iso_mean["week"].astype(str).str.zfill(2)
df_mean["weekday"] = df_mean["Date"].dt.weekday

datasets = {p: df[df["Participant_ID"] == p].copy() for p in participants}
datasets["Alle (Ø)"] = df_mean

# Traces bauen: eine Trace pro (View × Metrik)-Kombination
traces = []
trace_map = {}  # (view, metric) → trace-index

for view in views:
    data = datasets[view]
    matrix_data, hover_data, year_weeks, yw_index = build_matrix(data, list(metrics.values())[0][0])
    tick_vals, tick_text = get_month_ticks(data, yw_index)

    for m_label, (col, colorscale, zmin, zmax) in metrics.items():
        matrix, hover, yw, yw_idx = build_matrix(data, col)
        tv, tt = get_month_ticks(data, yw_idx)

        idx = len(traces)
        trace_map[(view, m_label)] = idx

        traces.append(go.Heatmap(
            z=matrix,
            text=hover,
            hoverinfo="text",
            colorscale=colorscale,
            zmin=zmin,
            zmax=zmax,
            x=list(range(len(yw))),
            y=["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"],
            colorbar=dict(title=m_label),
            visible=False,
            meta={"tickvals": tv, "ticktext": tt},
        ))

# Erste Ansicht sichtbar machen
first_view = views[0]
first_metric = list(metrics.keys())[0]
traces[trace_map[(first_view, first_metric)]].visible = True

# Dropdown: Probanden
view_buttons = []
for view in views:
    first_m = list(metrics.keys())[0]
    visibility = [False] * len(traces)
    active_idx = trace_map[(view, first_m)]
    visibility[active_idx] = True

    tv = traces[active_idx].meta["tickvals"]
    tt = traces[active_idx].meta["ticktext"]

    view_buttons.append(dict(
        label=view,
        method="update",
        args=[
            {"visible": visibility},
            {
                "title": f"Stress-Kalender – {view} | {first_m}",
                "xaxis.tickvals": tv,
                "xaxis.ticktext": tt,
            }
        ]
    ))

# Dropdown: Metriken (Standard: erster Proband)
metric_buttons = []
for m_label in metrics.keys():
    visibility = [False] * len(traces)
    active_idx = trace_map[(first_view, m_label)]
    visibility[active_idx] = True

    tv = traces[active_idx].meta["tickvals"]
    tt = traces[active_idx].meta["ticktext"]

    metric_buttons.append(dict(
        label=m_label,
        method="update",
        args=[
            {"visible": visibility},
            {
                "title": f"Stress-Kalender – {first_view} | {m_label}",
                "xaxis.tickvals": tv,
                "xaxis.ticktext": tt,
            }
        ]
    ))

# Layout
fig = go.Figure(data=traces)
fig.update_layout(
    title=f"Stress-Kalender – {first_view} | {first_metric}",
    paper_bgcolor="#1e1e2e",
    plot_bgcolor="#1e1e2e",
    font=dict(color="#cdd6f4"),
    xaxis=dict(
        tickmode="array",
        tickvals=traces[trace_map[(first_view, first_metric)]].meta["tickvals"],
        ticktext=traces[trace_map[(first_view, first_metric)]].meta["ticktext"],
        title="",
        gridcolor="#313244",
    ),
    yaxis=dict(title="", gridcolor="#313244"),
    updatemenus=[
        dict(
            buttons=view_buttons,
            direction="down",
            x=0.0,
            xanchor="left",
            y=1.12,
            yanchor="top",
            bgcolor="#313244",
            bordercolor="#cdd6f4",
            font=dict(color="#cdd6f4"),
        ),
        dict(
            buttons=metric_buttons,
            direction="down",
            x=1.0,
            xanchor="right",
            y=1.12,
            yanchor="top",
            bgcolor="#313244",
            bordercolor="#cdd6f4",
            font=dict(color="#cdd6f4"),
        ),
    ],
    margin=dict(t=100, b=40, l=60, r=40),
)

fig.show()