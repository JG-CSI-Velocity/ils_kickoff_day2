"""Charts for Analyses 5-10: NSF/OD Stratification."""

import plotly.graph_objects as go

from ils_kickoff.settings import ChartConfig


def chart_nsf_volume(df, config: ChartConfig) -> go.Figure:
    """Analyses 5-6: Horizontal bar of NSF volume stratification."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=data["NSF Bin"],
        x=data["# of Accounts"],
        name="# of Accounts",
        orientation="h",
        marker_color=colors[0],
    ))

    fig.add_trace(go.Bar(
        y=data["NSF Bin"],
        x=data["Total OD/NSF Items"],
        name="Total OD/NSF Items",
        orientation="h",
        marker_color=colors[1],
    ))

    fig.update_layout(
        template=config.theme,
        barmode="group",
        xaxis_title="Count",
        yaxis_title="NSF/OD Items Bin",
        yaxis=dict(autorange="reversed"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=40, l=80),
    )

    return fig


def chart_nsf_pay_ratio(df, config: ChartConfig) -> go.Figure:
    """Analyses 7-8: Bar + line of NSF pay ratio by bracket."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=data["NSF Bin"],
        y=data["# of Accounts"],
        name="# of Accounts",
        marker_color=colors[0],
        yaxis="y",
    ))

    if "% Pay Rate" in data.columns:
        fig.add_trace(go.Scatter(
            x=data["NSF Bin"],
            y=data["% Pay Rate"],
            name="% Pay Rate",
            mode="lines+markers",
            marker=dict(color=colors[2], size=8),
            line=dict(color=colors[2], width=2),
            yaxis="y2",
        ))

    fig.update_layout(
        template=config.theme,
        xaxis_title="NSF/OD Items Bin",
        yaxis=dict(title="# of Accounts", side="left"),
        yaxis2=dict(title="% Pay Rate", side="right", overlaying="y", range=[0, 105]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=40),
    )

    return fig


def chart_nsf_full_metrics(df, config: ChartConfig) -> go.Figure:
    """Analyses 9-10: Multi-metric bar of NSF + deposits + swipes."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=data["NSF Bin"],
        y=data["# of Accounts"],
        name="# of Accounts",
        marker_color=colors[0],
    ))

    if "Avg $$ Dep/Month" in data.columns:
        fig.add_trace(go.Bar(
            x=data["NSF Bin"],
            y=data["Avg $$ Dep/Month"],
            name="Avg $$ Dep/Month",
            marker_color=colors[2],
        ))

    if "Average of Swipes" in data.columns:
        fig.add_trace(go.Bar(
            x=data["NSF Bin"],
            y=data["Average of Swipes"],
            name="Avg Swipes",
            marker_color=colors[3],
        ))

    fig.update_layout(
        template=config.theme,
        barmode="group",
        xaxis_title="NSF/OD Items Bin",
        yaxis_title="Value",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=40),
    )

    return fig
