"""Charts for Analyses 1-2: Account Status and Account Type."""

import plotly.graph_objects as go

from ils_kickoff.settings import ChartConfig


def chart_account_status(df, config: ChartConfig) -> go.Figure:
    """Analysis 1: Bar chart of accounts and pay ratio by Account Status."""
    # Exclude Grand Total row from chart
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=data["Account Status"],
        y=data["# of Accounts"],
        name="# of Accounts",
        marker_color=colors[0],
        yaxis="y",
    ))

    if "Pay Ratio" in data.columns:
        fig.add_trace(go.Scatter(
            x=data["Account Status"],
            y=data["Pay Ratio"],
            name="Pay Ratio",
            mode="lines+markers",
            marker=dict(color=colors[1], size=8),
            line=dict(color=colors[1], width=2),
            yaxis="y2",
        ))

    fig.update_layout(
        template=config.theme,
        xaxis_title="Account Status",
        yaxis=dict(title="# of Accounts", side="left"),
        yaxis2=dict(title="Pay Ratio", side="right", overlaying="y", range=[0, 1.1]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=40),
    )

    return fig


def chart_account_type(df, config: ChartConfig) -> go.Figure:
    """Analysis 2: Grouped bar chart of Personal vs Business accounts."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=data.iloc[:, 0],
        y=data["# of Accounts"],
        name="# of Accounts",
        marker_color=colors[0],
    ))

    fig.add_trace(go.Bar(
        x=data.iloc[:, 0],
        y=data["# of Items"],
        name="# of Items",
        marker_color=colors[1],
    ))

    fig.update_layout(
        template=config.theme,
        barmode="group",
        xaxis_title="Account Type",
        yaxis_title="Count",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=40),
    )

    return fig
