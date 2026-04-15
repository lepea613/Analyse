# =============================================================================
# Sleep Analysis Dashboard
# Visualizes relationships between sleep quality, heart metrics, and sleep scores
# =============================================================================

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# -----------------------------------------------------------------------------
# Data Loading
# -----------------------------------------------------------------------------
# Find the first Excel file in the table folder
table_folder = 'table'
excel_files = [f for f in os.listdir(table_folder) if f.endswith(('.xlsx', '.xls'))]
if not excel_files:
    raise FileNotFoundError(f"No Excel files found in '{table_folder}' folder")
df = pd.read_excel(os.path.join(table_folder, excel_files[0]))

# -----------------------------------------------------------------------------
# CHART 1: Subjective Sleep Quality vs. Objective Sleep Score
# Compares diary-based sleep quality ratings (1-10) with Fitbit sleep scores
# -----------------------------------------------------------------------------
df_valid1 = df[['TagebuchEntries_Schlafqualitat', 'sleep_score_overall_score']].dropna()

# Group by sleep quality rating and calculate median sleep score for each level
medians = df_valid1.groupby('TagebuchEntries_Schlafqualitat')['sleep_score_overall_score'].median().reset_index()
medians.columns = ['Schlafqualitat', 'Median_Sleep_Score']

# Calculate deviation from expected score (assumption: quality 1-10 maps to score 10-100)
medians['Expected_Score'] = medians['Schlafqualitat'] * 10
medians['Deviation'] = medians['Median_Sleep_Score'] - medians['Expected_Score']

# Build detailed hover tooltips showing actual vs expected values
hover_text1 = [
    f"Schlafqualität: {row['Schlafqualitat']:.0f}<br>"
    f"Median Sleep Score: {row['Median_Sleep_Score']:.1f}<br>"
    f"Erwarteter Score: {row['Expected_Score']:.0f}<br>"
    f"Abweichung: {row['Deviation']:+.1f}"
    for _, row in medians.iterrows()
]

# -----------------------------------------------------------------------------
# CHART 2: Heart Metrics vs. Sleep Metrics (Interactive Comparison)
# Allows switching between different heart and sleep metric combinations
# -----------------------------------------------------------------------------

# X-axis options: Heart/cardiovascular metrics
x_axis_options = {
    'Ruheherzfrequenz (bpm)': 'daily_resting_heart_rate_beats per minute',
    'HRV (ms)': 'daily_heart_rate_variability_average heart rate variability milliseconds'
}

# Y-axis options: Sleep quality/duration metrics  
y_axis_options = {
    'Sleep Score (Overall)': 'sleep_score_overall_score',
    'Schlafdauer (min)': 'UserSleeps_minutes_asleep',
    'Tiefschlaf (min)': 'sleep_score_deep_sleep_in_minutes'
}

# Prepare data with all required columns (drop rows with any missing values)
chart2_cols = list(x_axis_options.values()) + list(y_axis_options.values())
df_valid2 = df[chart2_cols].dropna()

# -----------------------------------------------------------------------------
# CHART 3: Sleep Metrics Correlation Heatmap
# Shows pairwise correlations between different sleep-related metrics
# -----------------------------------------------------------------------------
sleep_metrics_cols = ['sleep_score_restlessness', 'sleep_score_deep_sleep_in_minutes', 
                      'sleep_score_overall_score', 'UserSleeps_minutes_asleep']
df_valid3 = df[sleep_metrics_cols].dropna()

# Use readable column names for display in the heatmap
df_valid3_display = df_valid3.copy()
df_valid3_display.columns = ['Restlessness', 'Deep Sleep (min)', 'Overall Score', 'Total Sleep (min)']

# Compute Pearson correlation matrix
corr_matrix = df_valid3_display.corr()

# -----------------------------------------------------------------------------
# Calculate Correlation Coefficients for Chart Titles
# -----------------------------------------------------------------------------
correlation1 = df_valid1['TagebuchEntries_Schlafqualitat'].corr(df_valid1['sleep_score_overall_score'])

# -----------------------------------------------------------------------------
# Create Figure with Subplot Layout
# Row 1: Two scatter plots (Chart 1 and Chart 2)
# Row 2: One wide heatmap spanning both columns (Chart 3)
# -----------------------------------------------------------------------------
fig = make_subplots(
    rows=2, cols=2,
    subplot_titles=(
        f'Schlafqualität vs. Sleep Score (r={correlation1:.3f})',
        '',  # Chart 2 title is empty - metrics shown via dropdown and axis labels
        'Schlafmetriken: Korrelationsmatrix'
    ),
    horizontal_spacing=0.1,
    vertical_spacing=0.15,
    specs=[[{}, {}], [{"colspan": 2}, None]]
)

# -----------------------------------------------------------------------------
# CHART 1 TRACES: Scatter plot with median trend line
# -----------------------------------------------------------------------------
# Individual data points (semi-transparent to show density)
fig.add_trace(go.Scatter(
    x=df_valid1['TagebuchEntries_Schlafqualitat'],
    y=df_valid1['sleep_score_overall_score'],
    mode='markers',
    marker=dict(size=10, color='steelblue', opacity=0.3),
    name='Einzelne Werte',
    hovertemplate='Schlafqualität: %{x:.0f}<br>Sleep Score: %{y:.0f}<extra></extra>',
    legendgroup='chart1'
), row=1, col=1)

# Median trend line connecting median sleep scores per quality level
fig.add_trace(go.Scatter(
    x=medians['Schlafqualitat'],
    y=medians['Median_Sleep_Score'],
    mode='markers+lines',
    marker=dict(size=14, color='darkblue', line=dict(width=2, color='white')),
    line=dict(width=2, color='darkblue'),
    text=hover_text1,
    hoverinfo='text',
    name='Median Sleep Score',
    legendgroup='chart1'
), row=1, col=1)

# Reference line showing perfect linear relationship (quality * 10 = score)
fig.add_trace(go.Scatter(
    x=[1, 10],
    y=[10, 100],
    mode='lines',
    line=dict(dash='dash', color='gray', width=1),
    name='Erwartete Linie',
    hoverinfo='skip',
    legendgroup='chart1'
), row=1, col=1)

# -----------------------------------------------------------------------------
# CHART 2 TRACES: Generate all metric combinations (controlled by dropdown)
# Each combination creates 2 traces: scatter points + median trend line
# Only the first combination is visible initially; dropdown toggles visibility
# -----------------------------------------------------------------------------
trace_index = 0
chart2_traces = []

for x_label, x_col in x_axis_options.items():
    for y_label, y_col in y_axis_options.items():
        corr = df_valid2[x_col].corr(df_valid2[y_col])
        visible = (trace_index == 0)
        
        # Scatter plot showing all individual data points
        fig.add_trace(go.Scatter(
            x=df_valid2[x_col],
            y=df_valid2[y_col],
            mode='markers',
            marker=dict(size=10, color='coral', opacity=0.3),
            name=f'{x_label} vs {y_label}',
            hovertemplate=f'{x_label}: %{{x:.1f}}<br>{y_label}: %{{y:.1f}}<extra></extra>',
            legendgroup='chart2',
            visible=visible,
            showlegend=False
        ), row=1, col=2)
        
        # Bin x-values into 10 groups and compute median x and y for each bin
        # This creates a smoothed trend line through the scattered data
        df_temp = df_valid2[[x_col, y_col]].copy()
        df_temp['x_binned'] = pd.cut(df_temp[x_col], bins=10, labels=False)
        medians_chart2 = df_temp.groupby('x_binned').agg({
            x_col: 'median',
            y_col: 'median'
        }).dropna().reset_index(drop=True)
        
        # Hover tooltips for median points
        hover_text_median = [
            f"{x_label}: {row[x_col]:.1f}<br>{y_label}: {row[y_col]:.1f}<br>(Median)"
            for _, row in medians_chart2.iterrows()
        ]
        
        # Median trend line with markers
        fig.add_trace(go.Scatter(
            x=medians_chart2[x_col],
            y=medians_chart2[y_col],
            mode='markers+lines',
            marker=dict(size=12, color='darkred', line=dict(width=2, color='white')),
            line=dict(width=2, color='darkred'),
            text=hover_text_median,
            hoverinfo='text',
            name=f'Median: {x_label} vs {y_label}',
            legendgroup='chart2',
            visible=visible
        ), row=1, col=2)
        
        chart2_traces.append({
            'index': trace_index,
            'x_label': x_label,
            'y_label': y_label,
            'correlation': corr
        })
        trace_index += 1

# -----------------------------------------------------------------------------
# CHART 3 TRACE: Correlation Heatmap
# -----------------------------------------------------------------------------
# Pre-format correlation values as text annotations for each cell
annotations_text = [[f'{corr_matrix.iloc[i, j]:.2f}' for j in range(len(corr_matrix.columns))] 
                    for i in range(len(corr_matrix.index))]

fig.add_trace(go.Heatmap(
    z=corr_matrix.values,
    x=corr_matrix.columns.tolist(),
    y=corr_matrix.index.tolist(),
    colorscale='RdBu_r',
    zmin=-1,
    zmax=1,
    text=annotations_text,
    texttemplate='%{text}',
    textfont=dict(size=12, color='white'),
    hovertemplate='%{y} vs %{x}<br>Korrelation: %{z:.3f}<extra></extra>',
    colorbar=dict(title='Korrelation', x=1.02, len=0.4, y=0.2),
    showscale=True,
    name='Korrelation'
), row=2, col=1)

# -----------------------------------------------------------------------------
# Global Layout Settings
# -----------------------------------------------------------------------------
fig.update_layout(
    hovermode='closest',
    template='plotly_white',
    width=1400,
    height=900,
    showlegend=True
)

# -----------------------------------------------------------------------------
# CHART 2 DROPDOWN MENU
# Creates buttons for each X/Y metric combination with correlation displayed
# Trace visibility: Chart1 (3) + Chart2 (6 combinations × 2 traces each) + Heatmap (1)
# -----------------------------------------------------------------------------
num_y_options = len(y_axis_options)
num_chart2_traces = len(x_axis_options) * len(y_axis_options)
combined_buttons = []
for i, (x_label, x_col) in enumerate(x_axis_options.items()):
    for j, (y_label, y_col) in enumerate(y_axis_options.items()):
        trace_idx = i * num_y_options + j
        corr = chart2_traces[trace_idx]['correlation']
        
        # Build visibility array: [Chart1 traces, Chart2 traces, Heatmap]
        visible = [True, True, True]  # Chart 1: always visible
        for k in range(num_chart2_traces):
            visible.append(k == trace_idx)  # Chart 2 scatter: show only selected
            visible.append(k == trace_idx)  # Chart 2 median: show only selected
        visible.append(True)  # Heatmap: always visible
        
        combined_buttons.append(dict(
            label=f'{x_label} vs {y_label} (r={corr:.2f})',
            method='update',
            args=[
                {'visible': visible},
                {'xaxis2.title.text': x_label, 'yaxis2.title.text': y_label}
            ]
        ))

# Position dropdown menu above Chart 2
fig.update_layout(
    updatemenus=[
        dict(
            buttons=combined_buttons,
            direction='down',
            showactive=True,
            x=0.55,
            xanchor='left',
            y=1.0,
            yanchor='bottom',
            bgcolor='white',
            bordercolor='darkgray',
            borderwidth=1,
            font=dict(size=10)
        )
    ]
)

# -----------------------------------------------------------------------------
# Axis Configuration
# -----------------------------------------------------------------------------
# Chart 1: Fixed axis labels and ranges
fig.update_xaxes(title_text='Schlafqualität (Tagebuch)', range=[1, 10], dtick=1, row=1, col=1)
fig.update_yaxes(title_text='Sleep Score (Overall)', range=[30, 100], row=1, col=1)

# Chart 2: Dynamic axis labels (updated via dropdown, initialized to first option)
first_x_label = list(x_axis_options.keys())[0]
first_y_label = list(y_axis_options.keys())[0]
fig.update_xaxes(title_text=first_x_label, row=1, col=2)
fig.update_yaxes(title_text=first_y_label, row=1, col=2)

# -----------------------------------------------------------------------------
# Display the Dashboard
# -----------------------------------------------------------------------------
fig.show()
