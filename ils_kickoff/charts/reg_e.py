"""Charts for Analyses 13, 15: Reg E Summary and Historical."""

import plotly.graph_objects as go

from ils_kickoff.settings import ChartConfig


def chart_reg_e_summary(df, config: ChartConfig) -> go.Figure:
    """Analysis 13: Simple bar chart of Reg E opt-in distribution."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=data["Reg E Flag"],
        y=data["# of Accounts"],
        marker_color=colors[0],
    ))

    fig.update_layout(
        template=config.theme,
        xaxis_title="Reg E Flag",
        yaxis_title="# of Accounts",
        showlegend=False,
        margin=dict(t=60, b=40),
    )

    return fig


def chart_reg_e_historical(df, config: ChartConfig) -> go.Figure:
    """Analysis 15: Grouped bar of Reg E by Year Opened with opt-in % line."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=data["Year Opened"].astype(str),
        y=data["# of Accounts"],
        name="# of Accounts",
        marker_color=colors[0],
        yaxis="y",
    ))

    if "Opt In %" in data.columns:
        fig.add_trace(go.Scatter(
            x=data["Year Opened"].astype(str),
            y=data["Opt In %"],
            name="Opt In %",
            mode="lines+markers",
            marker=dict(color=colors[2], size=6),
            line=dict(color=colors[2], width=2),
            yaxis="y2",
        ))

    fig.update_layout(
        template=config.theme,
        xaxis_title="Year Opened",
        xaxis=dict(tickangle=-45),
        yaxis=dict(title="# of Accounts", side="left"),
        yaxis2=dict(title="Opt In %", side="right", overlaying="y", range=[0, 105]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=60),
    )

    return fig
