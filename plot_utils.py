import os
import plotly.graph_objects as plotly_go
from plotly.subplots import make_subplots

# PPTXにあわせたカラー設定
COLOR_BAR = "rgb(179, 226, 131)"
COLOR_LINE = "rgb(105, 175, 230)"
FONT_FAMILY = "Arial, Helvetica, sans-serif"  # Standard fonts usually supported without issue by kaleido

def ensure_dir(filepath):
    dirname = os.path.dirname(filepath)
    if dirname:
        os.makedirs(dirname, exist_ok=True)

def _apply_common_layout(fig, width, height):
    fig.update_layout(
        template="simple_white",
        width=width,
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        font=dict(family=FONT_FAMILY, size=10),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5
        )
    )
    # yaxis gridlines and discrete xaxis
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
    fig.update_xaxes(type='category')
    return fig

def save_bar_chart(categories, bar_data, bar_name, filepath, width=400, height=280):
    ensure_dir(filepath)
    fig = plotly_go.Figure()
    fig.add_trace(plotly_go.Bar(
        x=categories,
        y=bar_data,
        name=bar_name,
        marker_color=COLOR_BAR
    ))
    _apply_common_layout(fig, width, height)
    fig.update_layout(showlegend=False)
    fig.write_image(filepath, scale=2)

def save_combo_chart(categories, bar_data, bar_name, line_data, line_name, filepath, width=400, height=280):
    ensure_dir(filepath)
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        plotly_go.Bar(x=categories, y=bar_data, name=bar_name, marker_color=COLOR_BAR),
        secondary_y=False,
    )
    fig.add_trace(
        plotly_go.Scatter(x=categories, y=line_data, name=line_name, mode='lines+markers', line=dict(color=COLOR_LINE, width=2)),
        secondary_y=True,
    )
    
    _apply_common_layout(fig, width, height)
    fig.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray', secondary_y=False)
    fig.update_yaxes(showgrid=False, secondary_y=True)
    fig.write_image(filepath, scale=2)

def save_multi_line_chart(categories, series_dict, filepath, width=700, height=250):
    ensure_dir(filepath)
    fig = plotly_go.Figure()
    
    for name, data in series_dict.items():
        fig.add_trace(plotly_go.Scatter(
            x=categories,
            y=data,
            name=name,
            mode='lines+markers'
        ))
        
    _apply_common_layout(fig, width, height)
    fig.update_layout(
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.3,
            xanchor="center",
            x=0.5
        )
    )
    fig.write_image(filepath, scale=2)
