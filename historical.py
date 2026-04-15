# =============================================================================
# Historical Sleep Dashboard
# Monthly calendar view showing sleep metrics per day per participant
# =============================================================================

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import numpy as np

# -----------------------------------------------------------------------------
# Data Loading
# -----------------------------------------------------------------------------
table_folder = 'table'
excel_files = [f for f in os.listdir(table_folder) if f.endswith(('.xlsx', '.xls'))]
if not excel_files:
    raise FileNotFoundError(f"No Excel files found in '{table_folder}' folder")
df = pd.read_excel(os.path.join(table_folder, excel_files[0]))

df['Date'] = pd.to_datetime(df['Date'])

# -----------------------------------------------------------------------------
# Prepare Data for Monthly Calendar View
# -----------------------------------------------------------------------------
participants = sorted(df['Participant_ID'].unique().tolist())
df['YearMonth'] = df['Date'].dt.to_period('M')
df['DayOfWeek'] = df['Date'].dt.dayofweek  # Monday=0, Sunday=6
df['Day'] = df['Date'].dt.day
df['DateStr'] = df['Date'].dt.strftime('%Y-%m-%d')

months = sorted(df['YearMonth'].unique())
num_months = len(months)
num_participants = len(participants)

day_labels = ['Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa', 'Su']

# Metrics to display
metrics = [
    {'col': 'sleep_score_overall_score', 'label': 'Sleep Score', 'zmin': 40, 'zmax': 100, 'unit': ''},
    {'col': 'UserSleeps_minutes_asleep', 'label': 'Minutes Asleep', 'zmin': 300, 'zmax': 540, 'unit': ' min'}
]

# Color scale (red to green)
colorscale = [
    [0, 'rgb(165, 0, 38)'],
    [0.25, 'rgb(215, 48, 39)'],
    [0.5, 'rgb(253, 174, 97)'],
    [0.75, 'rgb(166, 217, 106)'],
    [1, 'rgb(0, 104, 55)']
]

# -----------------------------------------------------------------------------
# Create Figure: Rows = Participants, Columns = Months x Metrics + 2 Scatter plots
# -----------------------------------------------------------------------------
num_calendar_cols = num_months * len(metrics)
num_cols = num_calendar_cols + 2  # +2 for scatter plots

subplot_titles = []
for p in participants:
    for metric in metrics:
        for m in months:
            subplot_titles.append(f'{p} - {metric["label"]} - {m.strftime("%b")}')
    # Add titles for scatter plots
    subplot_titles.append(f'{p} - Sleep Score Trend')
    subplot_titles.append(f'{p} - Minutes Asleep Trend')

fig = make_subplots(
    rows=num_participants,
    cols=num_cols,
    subplot_titles=subplot_titles,
    horizontal_spacing=0.04,
    vertical_spacing=0.08,
    column_widths=[1, 1, 1, 1, 1.5, 1.5]  # Scatter plots slightly wider
)

for p_idx, participant in enumerate(participants):
    # Calendar heatmaps
    for met_idx, metric in enumerate(metrics):
        for m_idx, month in enumerate(months):
            col_idx = met_idx * num_months + m_idx + 1
            
            df_filtered = df[(df['Participant_ID'] == participant) & (df['YearMonth'] == month)]
            
            first_day = month.to_timestamp()
            first_weekday = first_day.dayofweek
            days_in_month = month.days_in_month
            num_weeks = (first_weekday + days_in_month + 6) // 7
            
            z_matrix = np.full((num_weeks, 7), np.nan)
            text_matrix = [[''] * 7 for _ in range(num_weeks)]
            
            for day in range(1, days_in_month + 1):
                date = first_day + pd.Timedelta(days=day - 1)
                weekday = date.dayofweek
                week_num = (first_weekday + day - 1) // 7
                
                day_data = df_filtered[df_filtered['Day'] == day]
                if not day_data.empty:
                    value = day_data[metric['col']].values[0]
                    date_str = day_data['DateStr'].values[0]
                    if pd.notna(value):
                        z_matrix[week_num, weekday] = value
                        text_matrix[week_num][weekday] = f"{date_str}<br>{value:.0f}{metric['unit']}"
                    else:
                        text_matrix[week_num][weekday] = f"{date_str}<br>No data"
                else:
                    text_matrix[week_num][weekday] = f"{date.strftime('%Y-%m-%d')}<br>No data"
            
            # Show colorbar for last participant, last month of each metric (at bottom)
            show_colorbar = (p_idx == num_participants - 1 and m_idx == num_months - 1)
            colorbar_x = 0.13 if met_idx == 0 else 0.42
            
            fig.add_trace(go.Heatmap(
                z=z_matrix,
                x=day_labels,
                y=[f'W{i+1}' for i in range(num_weeks)],
                colorscale=colorscale,
                zmin=metric['zmin'],
                zmax=metric['zmax'],
                text=text_matrix,
                hovertemplate='%{text}<extra></extra>',
                showscale=show_colorbar,
                colorbar=dict(
                    title=dict(text=metric['label'], font=dict(size=9), side='right'),
                    len=0.3,
                    thickness=15,
                    y=-0.15,
                    yanchor='top',
                    x=colorbar_x,
                    xanchor='center',
                    orientation='h',
                    tickfont=dict(size=8)
                ) if show_colorbar else None,
                name=f'{participant} {metric["label"]} {month}'
            ), row=p_idx + 1, col=col_idx)
    
    # Scatter plots for this participant (by day of week)
    df_participant = df[df['Participant_ID'] == participant].copy()
    
    scatter_colors = ['steelblue', 'coral']
    day_order = [0, 1, 2, 3, 4, 5, 6]  # Monday to Sunday
    
    for s_idx, metric in enumerate(metrics):
        scatter_col = num_calendar_cols + s_idx + 1
        df_valid = df_participant[['DayOfWeek', metric['col']]].dropna()
        
        # Add jitter to x-axis for better visibility
        jitter = np.random.uniform(-0.2, 0.2, len(df_valid))
        
        # Individual data points
        fig.add_trace(go.Scatter(
            x=df_valid['DayOfWeek'] + jitter,
            y=df_valid[metric['col']],
            mode='markers',
            marker=dict(size=6, color=scatter_colors[s_idx], opacity=0.5),
            name=f'{participant} {metric["label"]}',
            hovertemplate=f'%{{x:.0f}}<br>{metric["label"]}: %{{y:.0f}}{metric["unit"]}<extra></extra>',
            showlegend=False
        ), row=p_idx + 1, col=scatter_col)
        
        # Median per day of week (color matches data points)
        if len(df_valid) > 0:
            medians_by_day = df_valid.groupby('DayOfWeek')[metric['col']].median().reindex(day_order)
            fig.add_trace(go.Scatter(
                x=list(range(7)),
                y=medians_by_day.values,
                mode='markers+lines',
                marker=dict(size=8, color=scatter_colors[s_idx], line=dict(width=1, color='white')),
                line=dict(color=scatter_colors[s_idx], width=2),
                name=f'Median',
                hovertemplate=f'%{{x}}<br>Median: %{{y:.0f}}{metric["unit"]}<extra></extra>',
                showlegend=False
            ), row=p_idx + 1, col=scatter_col)

# -----------------------------------------------------------------------------
# Layout Configuration
# -----------------------------------------------------------------------------
fig.update_layout(
    title=dict(text='Sleep Metrics - Monthly Calendar', font=dict(size=16)),
    template='plotly_white',
    height=180 * num_participants + 120,
    width=280 * num_cols + 80,
    showlegend=False,
    margin=dict(l=40, r=60, t=80, b=80)
)

# Make subplot titles smaller
for annotation in fig['layout']['annotations']:
    annotation['font'] = dict(size=9)

# Shift the Minutes Asleep charts (columns 3 and 4) to the right by compressing Sleep Score charts
shift_amount = 0.03
for p_idx in range(num_participants):
    for met_idx, metric in enumerate(metrics):
        for m_idx in range(num_months):
            col_idx = met_idx * num_months + m_idx + 1
            axis_num = p_idx * num_cols + col_idx
            axis_name = f'xaxis{axis_num}' if axis_num > 1 else 'xaxis'
            if axis_name in fig['layout']:
                domain = list(fig['layout'][axis_name]['domain'])
                if met_idx == 0:  # Sleep Score columns - compress slightly left
                    domain[1] = domain[1] - shift_amount / 2
                else:  # Minutes Asleep columns - shift right
                    domain[0] = domain[0] + shift_amount / 2
                fig['layout'][axis_name]['domain'] = domain
            # Also adjust annotation (subplot title)
            annot_idx = p_idx * num_cols + col_idx - 1
            if annot_idx < len(fig['layout']['annotations']):
                if met_idx == 1:
                    fig['layout']['annotations'][annot_idx]['x'] += shift_amount / 2

for p_idx in range(num_participants):
    for col_idx in range(num_calendar_cols):
        fig.update_xaxes(tickfont=dict(size=8), row=p_idx + 1, col=col_idx + 1)
        fig.update_yaxes(tickfont=dict(size=8), autorange='reversed', row=p_idx + 1, col=col_idx + 1)
    # Scatter plot axes (day of week)
    for s_idx in range(2):
        scatter_col = num_calendar_cols + s_idx + 1
        fig.update_xaxes(
            tickfont=dict(size=8), 
            tickvals=list(range(7)), 
            ticktext=day_labels,
            range=[-0.5, 6.5],
            row=p_idx + 1, col=scatter_col
        )
        fig.update_yaxes(tickfont=dict(size=8), row=p_idx + 1, col=scatter_col)

# -----------------------------------------------------------------------------
# Display the Dashboard
# -----------------------------------------------------------------------------
fig.show()
