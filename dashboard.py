from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# --------------------------------
# PAGE SETUP
# --------------------------------
st.set_page_config(
    page_title="CRM Campaign Dashboard",
    layout="wide",
)


DATABASE_FILE = "database.xlsx"
DATA_SHEET = "audience_daily_snapshot"

BASE_COLUMNS = [
    "date",
    "customer_type",
    "campaign_name",
    "active_audience",
    "new_imports",
    "sent",
    "delivered",
    "opens",
    "open_rate",
    "clicks",
    "click_rate",
    "unsubscribes",
    "unsubscribe_rate",
    "bounces",
    "bounce_rate",
]

EMAIL_BENCHMARKS = {
    "open_rate": ("Open Rate Benchmark", 25),
    "click_rate": ("Click Rate Benchmark", 5),
    "unsubscribe_rate": ("Unsubscribe Warning", 1),
    "bounce_rate": ("Bounce Warning", 2),
}

PERFORMANCE_COLUMNS = [
    "sent",
    "delivered",
    "opens",
    "open_rate",
    "clicks",
    "click_rate",
    "unsubscribes",
    "unsubscribe_rate",
    "bounces",
    "bounce_rate",
]

EMAIL_ENGAGEMENT_SEGMENTS = ["Build", "EOI", "Reservation"]


# --------------------------------
# DATA CLEANING HELPERS
# --------------------------------
def pct(numerator, denominator):
    if denominator == 0:
        return 0
    return numerator / denominator * 100


def normalize_rate(value):
    if pd.isna(value):
        return pd.NA
    value = float(value)
    if value <= 1:
        return value * 100
    return value


def clean_rate(row, rate_col, numerator_col, denominator_col):
    if pd.notna(row[rate_col]):
        return normalize_rate(row[rate_col])
    if row[numerator_col] and row[denominator_col]:
        return pct(row[numerator_col], row[denominator_col])
    return pd.NA


def weighted_rate(data, rate_col, weight_col):
    rate_data = data[[rate_col, weight_col]].dropna()
    rate_data = rate_data[rate_data[weight_col] > 0]
    if rate_data.empty:
        return 0
    return (rate_data[rate_col] * rate_data[weight_col]).sum() / rate_data[weight_col].sum()


def normalize_columns(df):
    df = df.copy()
    df.columns = [str(col).strip().lower().replace(" ", "_") for col in df.columns]

    rename_map = {
        "opene_rate": "open_rate",
        "opened": "opens",
        "audience": "active_audience",
        "recipients": "sent",
    }
    return df.rename(columns=rename_map)


def detect_channel(row):
    campaign_name = str(row.get("campaign_name", "")).lower()
    customer_type = str(row.get("customer_type", "")).lower()
    if "sms" in campaign_name or "sms" in customer_type:
        return "SMS"
    return "Email"


def clean_segment(value):
    value = str(value).strip()
    value = value.replace("_SMS", "").replace("_sms", "")
    value = value.replace("EOI&build", "EOI & Build")
    value = value.replace("EOI&Build", "EOI & Build")
    return value


def ensure_columns(df):
    for col in BASE_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df


def clean_database(df):
    df = normalize_columns(df)
    df = ensure_columns(df)

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["customer_type"] = df["customer_type"].astype(str).str.strip()
    df["campaign_name"] = df["campaign_name"].fillna("").astype(str).str.strip()
    df["has_performance_data"] = df[PERFORMANCE_COLUMNS].notna().any(axis=1)

    count_columns = [
        "active_audience",
        "new_imports",
        "sent",
        "delivered",
        "opens",
        "clicks",
        "unsubscribes",
        "bounces",
    ]
    rate_columns = ["open_rate", "click_rate", "unsubscribe_rate", "bounce_rate"]
    for col in count_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in rate_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["date"])
    df = df[df["customer_type"].str.lower().ne("nan")]
    df = df[df["customer_type"].str.strip().ne("")]

    df["channel"] = df.apply(detect_channel, axis=1)
    df["audience_segment"] = df["customer_type"].apply(clean_segment)

    sms_rows = df["channel"].eq("SMS")
    df.loc[sms_rows, "new_imports"] = df.loc[sms_rows, "active_audience"]

    df["campaign_name"] = df.apply(
        lambda row: row["campaign_name"]
        if row["campaign_name"]
        else f"{row['channel']} campaign - {row['audience_segment']} - {row['date'].strftime('%Y-%m-%d')}",
        axis=1,
    )

    df["sent"] = df.apply(
        lambda row: row["sent"]
        if row["sent"] or not row["has_performance_data"]
        else row["active_audience"],
        axis=1,
    )
    df["delivered"] = df.apply(
        lambda row: row["delivered"]
        if row["delivered"] or not row["has_performance_data"]
        else max(row["sent"] - row["bounces"], 0),
        axis=1,
    )

    df["open_rate"] = df.apply(lambda row: clean_rate(row, "open_rate", "opens", "delivered"), axis=1)
    df["click_rate"] = df.apply(lambda row: clean_rate(row, "click_rate", "clicks", "delivered"), axis=1)
    df["unsubscribe_rate"] = df.apply(
        lambda row: clean_rate(row, "unsubscribe_rate", "unsubscribes", "delivered"),
        axis=1,
    )
    df["bounce_rate"] = df.apply(lambda row: clean_rate(row, "bounce_rate", "bounces", "sent"), axis=1)

    df["date_tag"] = df["date"].dt.strftime("%Y-%m-%d")
    return df.sort_values(["date", "channel", "audience_segment"]).reset_index(drop=True)


# --------------------------------
# DATA LOADING
# --------------------------------
@st.cache_data
def make_sample_database():
    today = date.today()
    first_wednesday = today - timedelta(days=(today.weekday() - 2) % 7 + 35)
    dates = pd.date_range(first_wednesday, periods=6, freq="7D")

    rows = []
    for week_index, campaign_date in enumerate(dates):
        for customer_type, base in [("Reservation", 68), ("Build", 120), ("EOI", 820)]:
            sent = base + week_index * 14
            delivered = sent - (week_index % 2)
            opens = int(delivered * (0.32 + week_index % 3 * 0.03))
            clicks = int(delivered * (0.09 + week_index % 2 * 0.02))
            unsubscribes = 0 if customer_type != "EOI" else week_index % 4
            rows.append(
                {
                    "date": campaign_date,
                    "customer_type": customer_type,
                    "campaign_name": f"EDM #{week_index + 1}: What to Expect ({customer_type})",
                    "active_audience": sent,
                    "new_imports": week_index % 3,
                    "sent": sent,
                    "delivered": delivered,
                    "opens": opens,
                    "clicks": clicks,
                    "unsubscribes": unsubscribes,
                    "bounces": sent - delivered,
                }
            )

        for customer_type, base in [("Reservation_SMS", 68), ("EOI&build", 752)]:
            sent = base + week_index * 10
            delivered = int(sent * 0.975)
            clicks = int(delivered * (0.38 + week_index % 3 * 0.02))
            unsubscribes = 0 if "Reservation" in customer_type else week_index + 1
            rows.append(
                {
                    "date": campaign_date,
                    "customer_type": customer_type,
                    "campaign_name": f"SMS #{week_index + 1} What to Expect ({clean_segment(customer_type)})",
                    "active_audience": sent,
                    "new_imports": week_index % 2,
                    "sent": sent,
                    "delivered": delivered,
                    "opens": 0,
                    "clicks": clicks,
                    "unsubscribes": unsubscribes,
                    "bounces": sent - delivered,
                }
            )

    return pd.DataFrame(rows)


@st.cache_data
def load_database(file_modified_at=None):
    path = Path(DATABASE_FILE)
    if path.exists():
        return pd.read_excel(path, sheet_name=DATA_SHEET)
    return make_sample_database()


# --------------------------------
# SUMMARY HELPERS
# --------------------------------
def latest_audience(data):
    if data.empty:
        return 0

    email_data = data[data["channel"] == "Email"]
    sms_data = data[data["channel"] == "SMS"]

    email_audience = 0
    if not email_data.empty:
        email_audience = (
            email_data.sort_values("date")
            .groupby("audience_segment", as_index=False)
            .tail(1)["active_audience"]
            .sum()
        )

    sms_audience = sms_data["active_audience"].sum()
    return email_audience + sms_audience


def total_metrics(data):
    sent = data["sent"].sum()
    delivered = data["delivered"].sum()
    opens = data["opens"].sum()
    clicks = data["clicks"].sum()
    unsubscribes = data["unsubscribes"].sum()
    bounces = data["bounces"].sum()

    return {
        "active_audience": latest_audience(data),
        "new_imports": data["new_imports"].sum(),
        "sent": sent,
        "delivered": delivered,
        "opens": opens,
        "clicks": clicks,
        "unsubscribes": unsubscribes,
        "bounces": bounces,
        "open_rate": weighted_rate(data, "open_rate", "delivered"),
        "click_rate": weighted_rate(data, "click_rate", "delivered"),
        "unsubscribe_rate": weighted_rate(data, "unsubscribe_rate", "delivered"),
        "bounce_rate": weighted_rate(data, "bounce_rate", "sent"),
    }


def weekly_summary(data):
    summary = (
        data.groupby(["date", "date_tag", "audience_segment"], as_index=False)
        .agg(
            active_audience=("active_audience", "max"),
            new_imports=("new_imports", "sum"),
            sent=("sent", "sum"),
            delivered=("delivered", "sum"),
            opens=("opens", "sum"),
            clicks=("clicks", "sum"),
            unsubscribes=("unsubscribes", "sum"),
            bounces=("bounces", "sum"),
            open_rate=("open_rate", lambda values: weighted_rate(data.loc[values.index], "open_rate", "delivered")),
            click_rate=("click_rate", lambda values: weighted_rate(data.loc[values.index], "click_rate", "delivered")),
            unsubscribe_rate=(
                "unsubscribe_rate",
                lambda values: weighted_rate(data.loc[values.index], "unsubscribe_rate", "delivered"),
            ),
            bounce_rate=("bounce_rate", lambda values: weighted_rate(data.loc[values.index], "bounce_rate", "sent")),
        )
        .sort_values(["date", "audience_segment"])
    )
    return summary


def weekly_total_summary(data):
    if data.empty:
        return pd.DataFrame()

    return (
        data.groupby(["date", "date_tag"], as_index=False)
        .agg(
            sent=("sent", "sum"),
            delivered=("delivered", "sum"),
            opens=("opens", "sum"),
            clicks=("clicks", "sum"),
            unsubscribes=("unsubscribes", "sum"),
            bounces=("bounces", "sum"),
            open_rate=("open_rate", lambda values: weighted_rate(data.loc[values.index], "open_rate", "delivered")),
            click_rate=("click_rate", lambda values: weighted_rate(data.loc[values.index], "click_rate", "delivered")),
            unsubscribe_rate=(
                "unsubscribe_rate",
                lambda values: weighted_rate(data.loc[values.index], "unsubscribe_rate", "delivered"),
            ),
            bounce_rate=("bounce_rate", lambda values: weighted_rate(data.loc[values.index], "bounce_rate", "sent")),
        )
        .sort_values("date")
    )


def date_order(data):
    return data.sort_values("date")["date_tag"].drop_duplicates().tolist()


# --------------------------------
# CHART HELPERS
# --------------------------------
def add_email_benchmark(fig, metric):
    if metric not in EMAIL_BENCHMARKS:
        return fig
    label, value = EMAIL_BENCHMARKS[metric]
    fig.add_hline(
        y=value,
        line_dash="dash",
        line_color="#8A94A6",
        annotation_text=f"{label}: {value}%",
        annotation_position="top left",
    )
    return fig


def render_rate_trend(data, metric, title, benchmark=False, chart_key=None):
    data = data.dropna(subset=[metric]).copy()
    if data.empty:
        st.info(f"No data for {title}.")
        return

    fig = px.line(
        data,
        x="date_tag",
        y=metric,
        color="audience_segment",
        markers=True,
        title=title,
        category_orders={"date_tag": date_order(data)},
    )
    fig.update_xaxes(title="Week")
    fig.update_yaxes(title="Rate", ticksuffix="%")
    if benchmark:
        fig = add_email_benchmark(fig, metric)
    st.plotly_chart(fig, width="stretch", key=chart_key or f"rate_{metric}_{title}")


def render_funnel(data, channel, chart_key):
    metrics = total_metrics(data)
    if channel == "Email":
        labels = ["Sent", "Delivered", "Opened", "Clicked", "Unsubscribed"]
        values = [
            metrics["sent"],
            metrics["delivered"],
            metrics["opens"],
            metrics["clicks"],
            metrics["unsubscribes"],
        ]
    else:
        labels = ["Sent", "Delivered", "Clicked", "Unsubscribed"]
        values = [
            metrics["sent"],
            metrics["delivered"],
            metrics["clicks"],
            metrics["unsubscribes"],
        ]

    fig = go.Figure(go.Funnel(y=labels, x=values, textinfo="value+percent initial"))
    fig.update_layout(title=f"{channel} Funnel Snapshot", height=420)
    st.plotly_chart(fig, width="stretch", key=chart_key)


def render_volume_trend(data, channel, chart_key):
    summary = weekly_summary(data)
    if summary.empty:
        st.info(f"No {channel} weekly volume data for the selected filters.")
        return

    volume_data = summary.melt(
        id_vars=["date", "date_tag", "audience_segment"],
        value_vars=["sent", "delivered", "clicks", "unsubscribes"],
        var_name="metric",
        value_name="count",
    )
    metric_labels = {
        "sent": "Sent",
        "delivered": "Delivered",
        "clicks": "Clicks",
        "unsubscribes": "Unsubscribes",
    }
    volume_data["metric"] = volume_data["metric"].map(metric_labels)

    fig = px.bar(
        volume_data,
        x="date_tag",
        y="count",
        color="metric",
        facet_col="audience_segment",
        title=f"{channel} Weekly Volume Detail",
        category_orders={"date_tag": date_order(summary)},
        barmode="group",
    )
    fig.update_xaxes(title="Week")
    fig.update_yaxes(title="Count")
    st.plotly_chart(fig, width="stretch", key=chart_key)


def render_weekly_detail_table(data, channel):
    summary = weekly_summary(data).sort_values(["date", "audience_segment"], ascending=[False, True])
    if summary.empty:
        st.info(f"No {channel} weekly detail rows for the selected filters.")
        return

    columns = [
        "date",
        "audience_segment",
        "sent",
        "delivered",
        "clicks",
        "click_rate",
        "unsubscribes",
        "unsubscribe_rate",
        "bounce_rate",
    ]
    if channel == "Email":
        columns.insert(5, "opens")
        columns.insert(6, "open_rate")

    st.subheader("Weekly Detail")
    st.dataframe(summary[columns], width="stretch", hide_index=True)


def previous_period_range(start_date, end_date):
    period_days = (end_date - start_date).days + 1
    previous_end = start_date - timedelta(days=1)
    previous_start = previous_end - timedelta(days=period_days - 1)
    return previous_start, previous_end


def metric_delta_labels(data, comparison_data=None):
    if data.empty or comparison_data is None or comparison_data.empty:
        return {}

    current = total_metrics(data)
    previous = total_metrics(comparison_data)
    labels = {}
    rate_metrics = {"open_rate", "click_rate", "unsubscribe_rate", "bounce_rate"}

    for metric, current_value in current.items():
        delta = current_value - previous[metric]
        if metric in rate_metrics:
            labels[metric] = f"{delta:+.2f} pts vs previous period"
        else:
            labels[metric] = f"{delta:+,.0f} vs previous period"
    return labels


def render_primary_kpis(data, channel=None, show_delta=False, comparison_data=None):
    metrics = total_metrics(data)
    deltas = metric_delta_labels(data, comparison_data) if show_delta else {}

    if channel == "SMS":
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}", deltas.get("active_audience"))
        kpi2.metric("Click Rate", f"{metrics['click_rate']:.1f}%", deltas.get("click_rate"))
        kpi3.metric(
            "Unsubscribe Rate",
            f"{metrics['unsubscribe_rate']:.2f}%",
            deltas.get("unsubscribe_rate"),
            delta_color="inverse",
        )
        kpi4.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%", deltas.get("bounce_rate"), delta_color="inverse")
    elif channel is None:
        email_data = data[data["channel"] == "Email"]
        email_metrics = total_metrics(email_data)
        if email_data.empty:
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}", deltas.get("active_audience"))
            kpi2.metric("Click Rate", f"{metrics['click_rate']:.1f}%", deltas.get("click_rate"))
            kpi3.metric(
                "Unsubscribe Rate",
                f"{metrics['unsubscribe_rate']:.2f}%",
                deltas.get("unsubscribe_rate"),
                delta_color="inverse",
            )
            kpi4.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%", deltas.get("bounce_rate"), delta_color="inverse")
        else:
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            comparison_email_data = None
            if comparison_data is not None:
                comparison_email_data = comparison_data[comparison_data["channel"] == "Email"]
            email_deltas = metric_delta_labels(email_data, comparison_email_data) if show_delta else {}
            kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}", deltas.get("active_audience"))
            kpi2.metric("Email Open Rate", f"{email_metrics['open_rate']:.1f}%", email_deltas.get("open_rate"))
            kpi3.metric("Click Rate", f"{metrics['click_rate']:.1f}%", deltas.get("click_rate"))
            kpi4.metric(
                "Unsubscribe Rate",
                f"{metrics['unsubscribe_rate']:.2f}%",
                deltas.get("unsubscribe_rate"),
                delta_color="inverse",
            )
            kpi5.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%", deltas.get("bounce_rate"), delta_color="inverse")
    else:
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}", deltas.get("active_audience"))
        kpi2.metric("Open Rate", f"{metrics['open_rate']:.1f}%", deltas.get("open_rate"))
        kpi3.metric("Click Rate", f"{metrics['click_rate']:.1f}%", deltas.get("click_rate"))
        kpi4.metric(
            "Unsubscribe Rate",
            f"{metrics['unsubscribe_rate']:.2f}%",
            deltas.get("unsubscribe_rate"),
            delta_color="inverse",
        )
        kpi5.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%", deltas.get("bounce_rate"), delta_color="inverse")

    raw1, raw2, raw3, raw4 = st.columns(4)
    raw1.metric("Total Sent", f"{metrics['sent']:,.0f}", deltas.get("sent"))
    raw2.metric("Delivered", f"{metrics['delivered']:,.0f}", deltas.get("delivered"))
    raw3.metric("Clicks", f"{metrics['clicks']:,.0f}", deltas.get("clicks"))
    raw4.metric("New Imports", f"{metrics['new_imports']:,.0f}", deltas.get("new_imports"))


def trend_phrase(summary, metric, label, precision=1):
    metric_data = summary.dropna(subset=[metric]).sort_values("date")
    if metric_data.empty:
        return f"{label}: no recent rate data"

    latest = metric_data.iloc[-1][metric]
    if len(metric_data) == 1:
        return f"{label}: {latest:.{precision}f}%"

    previous = metric_data.iloc[-2][metric]
    delta = latest - previous
    direction = "up" if delta >= 0 else "down"
    return f"{label}: {latest:.{precision}f}% ({direction} {abs(delta):.{precision}f} pts vs previous week)"


def format_rate(value, precision=1):
    if pd.isna(value):
        return "N/A"
    return f"{value:.{precision}f}%"


def format_point_delta(value, precision=1):
    if pd.isna(value):
        return None
    return f"{value:+.{precision}f} pts vs previous week"


def latest_email_engagement_comparison(data):
    email_data = data[
        (data["channel"] == "Email")
        & (data["audience_segment"].isin(EMAIL_ENGAGEMENT_SEGMENTS))
    ].copy()
    email_data = email_data[email_data[["open_rate", "click_rate"]].notna().any(axis=1)].copy()
    if email_data.empty:
        return pd.DataFrame(), None, None

    result_weeks = (
        email_data.groupby(["date", "date_tag"], as_index=False)
        .agg(result_sections=("audience_segment", "nunique"))
        .sort_values("date")
    )
    result_weeks = result_weeks[result_weeks["result_sections"] == len(EMAIL_ENGAGEMENT_SEGMENTS)]
    if len(result_weeks) < 2:
        return pd.DataFrame(), None, None

    previous_week = result_weeks.iloc[-2]
    latest_week = result_weeks.iloc[-1]

    summary = weekly_summary(
        email_data[email_data["date_tag"].isin([previous_week["date_tag"], latest_week["date_tag"]])]
    )
    current = summary[summary["date_tag"] == latest_week["date_tag"]].copy()
    previous = summary[summary["date_tag"] == previous_week["date_tag"]].copy()

    comparison = current.merge(
        previous,
        on="audience_segment",
        suffixes=("_current", "_previous"),
        how="outer",
    )
    for metric in ["open_rate", "click_rate", "unsubscribe_rate"]:
        comparison[f"{metric}_change"] = comparison[f"{metric}_current"] - comparison[f"{metric}_previous"]

    return comparison, latest_week, previous_week


def render_email_engagement_change(data):
    comparison, latest_week, previous_week = latest_email_engagement_comparison(data)

    st.markdown("**Email Engagement Change by Audience Section**")
    if comparison.empty:
        st.info("Need at least two Email result weeks for Build, EOI, or Reservation to compare engagement changes.")
        return

    st.caption(
        f"Comparing latest Email result week {latest_week['date_tag']} "
        f"with previous result week {previous_week['date_tag']}."
    )

    section_columns = st.columns(3)
    for column, segment in zip(section_columns, EMAIL_ENGAGEMENT_SEGMENTS):
        segment_rows = comparison[comparison["audience_segment"] == segment]
        with column:
            st.markdown(f"**{segment}**")
            if segment_rows.empty:
                st.info("No result in one of the comparison weeks.")
                continue

            row = segment_rows.iloc[0]
            st.metric(
                "Open Rate",
                format_rate(row["open_rate_current"]),
                format_point_delta(row["open_rate_change"]),
            )
            st.metric(
                "Click Rate",
                format_rate(row["click_rate_current"]),
                format_point_delta(row["click_rate_change"]),
            )
            st.metric(
                "Unsubscribe Rate",
                format_rate(row["unsubscribe_rate_current"], precision=2),
                format_point_delta(row["unsubscribe_rate_change"], precision=2),
                delta_color="inverse",
            )

    detail = comparison.copy()
    table_columns = [
        "audience_segment",
        "delivered_current",
        "open_rate_previous",
        "open_rate_current",
        "open_rate_change",
        "click_rate_previous",
        "click_rate_current",
        "click_rate_change",
        "unsubscribe_rate_previous",
        "unsubscribe_rate_current",
        "unsubscribe_rate_change",
    ]
    detail = detail[[col for col in table_columns if col in detail.columns]]
    detail = detail.rename(
        columns={
            "audience_segment": "Audience Section",
            "delivered_current": "Delivered",
            "open_rate_previous": "Previous Open Rate",
            "open_rate_current": "Latest Open Rate",
            "open_rate_change": "Open Rate Change",
            "click_rate_previous": "Previous Click Rate",
            "click_rate_current": "Latest Click Rate",
            "click_rate_change": "Click Rate Change",
            "unsubscribe_rate_previous": "Previous Unsubscribe Rate",
            "unsubscribe_rate_current": "Latest Unsubscribe Rate",
            "unsubscribe_rate_change": "Unsubscribe Rate Change",
        }
    )
    st.dataframe(detail, width="stretch", hide_index=True)


def best_segment_phrase(summary, channel):
    if summary.empty:
        return "No segment performance data available yet."

    latest_date = summary["date"].max()
    latest_segments = summary[summary["date"] == latest_date].copy()
    if latest_segments.empty:
        return "No segment performance data available yet."

    if channel == "Email":
        best_row = latest_segments.sort_values(["click_rate", "open_rate"], ascending=[False, False]).iloc[0]
        return (
            f"Best current segment: {best_row['audience_segment']} "
            f"({format_rate(best_row['click_rate'])} click rate, {format_rate(best_row['open_rate'])} open rate)."
        )

    best_row = latest_segments.sort_values("click_rate", ascending=False).iloc[0]
    return f"Best current segment: {best_row['audience_segment']} ({format_rate(best_row['click_rate'])} click rate)."


def executive_action_items(metrics, channel):
    if channel == "Email":
        if metrics["bounce_rate"] > EMAIL_BENCHMARKS["bounce_rate"][1]:
            return (
                "List quality needs attention.",
                "Clean or suppress bounced contacts before the next send, then review source quality for recent imports.",
                "warning",
            )
        if metrics["unsubscribe_rate"] > EMAIL_BENCHMARKS["unsubscribe_rate"][1]:
            return (
                "Audience fatigue risk is rising.",
                "Reduce frequency for weaker segments and make the next message more clearly tied to customer intent.",
                "warning",
            )
        if metrics["open_rate"] < EMAIL_BENCHMARKS["open_rate"][1]:
            return (
                "Top-of-funnel engagement is the constraint.",
                "Test subject line and preheader first; keep the same offer so the next read is a clean comparison.",
                "warning",
            )
        if metrics["click_rate"] < EMAIL_BENCHMARKS["click_rate"][1]:
            return (
                "Email is being opened, but action is weak.",
                "Tighten the primary CTA and landing-page match before adding more campaign volume.",
                "warning",
            )
        return (
            "Email performance is healthy.",
            "Scale the best-performing segment first, then use the detailed tab to identify which message pattern to repeat.",
            "success",
        )

    if metrics["unsubscribe_rate"] > 1:
        return (
            "SMS opt-out risk needs attention.",
            "Slow the cadence for broad segments and reserve SMS for time-sensitive or high-intent messages.",
            "warning",
        )
    if metrics["click_rate"] < 20:
        return (
            "SMS response is below the expected action level.",
            "Retest send timing and CTA wording before expanding the next SMS batch.",
            "warning",
        )
    return (
        "SMS is working as an action channel.",
        "Keep SMS focused on clear next steps, and use the detailed tab to find segments that can take more volume.",
        "success",
    )


def render_executive_summary(performance_data, channel):
    if performance_data.empty:
        return

    metrics = total_metrics(performance_data)
    weekly_totals = weekly_total_summary(performance_data)
    segment_summary = weekly_summary(performance_data)
    title, action, status = executive_action_items(metrics, channel)

    st.markdown("**Executive Summary**")
    summary_col, action_col, watch_col = st.columns(3)

    with summary_col:
        if status == "success":
            st.success(title)
        else:
            st.warning(title)
        st.caption(best_segment_phrase(segment_summary, channel))

    with action_col:
        st.info(action)

    with watch_col:
        if channel == "Email":
            st.caption(trend_phrase(weekly_totals, "open_rate", "Open rate"))
            st.caption(trend_phrase(weekly_totals, "click_rate", "Click rate"))
            st.caption(trend_phrase(weekly_totals, "unsubscribe_rate", "Unsubscribe rate", precision=2))
        else:
            st.caption(trend_phrase(weekly_totals, "click_rate", "Click rate"))
            st.caption(trend_phrase(weekly_totals, "unsubscribe_rate", "Unsubscribe rate", precision=2))
            st.caption(trend_phrase(weekly_totals, "bounce_rate", "Bounce rate", precision=2))


def render_campaign_ranking(data, channel):
    st.subheader("Top Performing Campaigns")
    data = data[data["has_performance_data"]].copy()
    if data.empty:
        st.info(f"No {channel} campaign data for the selected filters.")
        return

    ranking = data.copy()
    if channel == "Email":
        ranking = ranking.sort_values(["click_rate", "open_rate"], ascending=[False, False])
        columns = [
            "date",
            "audience_segment",
            "campaign_name",
            "sent",
            "delivered",
            "open_rate",
            "click_rate",
            "unsubscribe_rate",
            "bounce_rate",
        ]
    else:
        ranking = ranking.sort_values("click_rate", ascending=False)
        columns = [
            "date",
            "audience_segment",
            "campaign_name",
            "sent",
            "delivered",
            "clicks",
            "click_rate",
            "unsubscribe_rate",
        ]

    st.dataframe(ranking[columns], width="stretch", hide_index=True)


# --------------------------------
# PAGE SECTIONS
# --------------------------------
def apply_segment_filter(data, label, key):
    segment_options = sorted(data["audience_segment"].dropna().unique())
    if not segment_options:
        return data

    selected_segments = st.multiselect(
        label,
        segment_options,
        default=segment_options,
        key=key,
    )
    return data[data["audience_segment"].isin(selected_segments)].copy()


def apply_channel_segment_filter(data, key_prefix):
    channel_options = sorted(data["channel"].dropna().unique())
    if not channel_options:
        return data

    selected_channel = st.radio(
        "Channel",
        channel_options,
        horizontal=True,
        key=f"{key_prefix}_channel",
    )
    channel_data = data[data["channel"] == selected_channel].copy()
    return apply_segment_filter(
        channel_data,
        f"{selected_channel} audience segment",
        f"{key_prefix}_{selected_channel.lower()}_segments",
    )


def apply_channel_segment_filter_pair(data, comparison_data, key_prefix):
    channel_options = sorted(data["channel"].dropna().unique())
    if not channel_options:
        return data, comparison_data

    selected_channel = st.radio(
        "Channel",
        channel_options,
        horizontal=True,
        key=f"{key_prefix}_channel",
    )
    channel_data = data[data["channel"] == selected_channel].copy()
    comparison_channel_data = comparison_data[comparison_data["channel"] == selected_channel].copy()

    segment_options = sorted(channel_data["audience_segment"].dropna().unique())
    if not segment_options:
        return channel_data, comparison_channel_data

    selected_segments = st.multiselect(
        f"{selected_channel} audience segment",
        segment_options,
        default=segment_options,
        key=f"{key_prefix}_{selected_channel.lower()}_segments",
    )
    filtered_data = channel_data[channel_data["audience_segment"].isin(selected_segments)].copy()
    filtered_comparison = comparison_channel_data[
        comparison_channel_data["audience_segment"].isin(selected_segments)
    ].copy()
    return filtered_data, filtered_comparison


def render_executive_overview(data, comparison_data=None, selected_period=None, comparison_period=None, full_data=None):
    st.subheader("Executive Overview")
    if data.empty:
        st.info("No data for the selected week range.")
        return

    render_email_engagement_change(full_data if full_data is not None else data)

    if comparison_data is None:
        comparison_data = data.iloc[0:0].copy()

    data, comparison_data = apply_channel_segment_filter_pair(data, comparison_data, "overview")
    if data.empty:
        st.info("No data for the selected overview filters.")
        return

    selected_channel = data["channel"].iloc[0]

    st.markdown("**Executive KPIs**")
    render_primary_kpis(
        data,
        channel=selected_channel,
        show_delta=True,
        comparison_data=comparison_data,
    )
    if not comparison_data.empty:
        if selected_period is None:
            selected_period = (data["date"].min().date(), data["date"].max().date())
        if comparison_period is None:
            comparison_period = (comparison_data["date"].min().date(), comparison_data["date"].max().date())
        current_period_label = f"{selected_period[0].strftime('%Y-%m-%d')} to {selected_period[1].strftime('%Y-%m-%d')}"
        previous_period_label = (
            f"{comparison_period[0].strftime('%Y-%m-%d')} to "
            f"{comparison_period[1].strftime('%Y-%m-%d')}"
        )
        st.info(
            f"KPI deltas show the change from the previous comparable period "
            f"({previous_period_label}) to the selected period ({current_period_label}). "
            "Percentage metrics are shown as percentage-point changes; volume metrics are shown as count changes."
        )
    else:
        st.info(
            "KPI deltas compare the selected period with the previous period of the same length. "
            "No delta is shown because the database does not yet have enough earlier data for this selection."
        )

    performance_data = data[data["has_performance_data"]].copy()
    if performance_data.empty:
        st.info(f"No {selected_channel} performance metrics yet for the selected week range.")
    else:
        render_executive_summary(performance_data, selected_channel)
        render_funnel(performance_data, selected_channel, f"overview_{selected_channel.lower()}_funnel")

    st.markdown("**Weekly Trends**")
    summary = weekly_summary(performance_data)

    row1_left, row1_right = st.columns(2)
    with row1_left:
        if selected_channel == "Email":
            render_rate_trend(
                summary,
                "open_rate",
                "Open Rate Trend",
                benchmark=True,
                chart_key="overview_open_rate",
            )
        else:
            render_rate_trend(
                summary,
                "click_rate",
                "SMS Click Rate Trend",
                chart_key="overview_sms_click_rate",
            )
    with row1_right:
        if selected_channel == "Email":
            render_rate_trend(
                summary,
                "click_rate",
                "Click Rate Trend",
                benchmark=True,
                chart_key="overview_click_rate",
            )
        else:
            render_rate_trend(
                summary,
                "unsubscribe_rate",
                "SMS Unsubscribe Rate Trend",
                chart_key="overview_sms_unsubscribe_rate",
            )

    if selected_channel == "Email":
        row2_left, row2_right = st.columns(2)
        with row2_left:
            render_rate_trend(
                summary,
                "unsubscribe_rate",
                "Unsubscribe Rate Trend",
                benchmark=True,
                chart_key="overview_unsubscribe_rate",
            )
        with row2_right:
            render_rate_trend(
                summary,
                "bounce_rate",
                "Bounce Rate Trend",
                benchmark=True,
                chart_key="overview_bounce_rate",
            )


def render_email_performance(data):
    st.subheader("Email Performance")
    if data.empty:
        st.info("No Email data for the selected week range.")
        return

    data = apply_segment_filter(data, "Email audience segment", "email_performance_segments")
    if data.empty:
        st.info("No Email data for the selected audience segment.")
        return

    render_primary_kpis(data, channel="Email")
    performance_data = data[data["has_performance_data"]].copy()
    if performance_data.empty:
        st.info("No Email performance metrics yet for the selected week range. Audience totals still use the latest snapshot.")
        return

    st.markdown("**Email Segment Trends**")

    summary = weekly_summary(performance_data)
    render_rate_trend(summary, "open_rate", "Open Rate Trend", benchmark=True, chart_key="email_open_rate")
    render_rate_trend(summary, "click_rate", "Click Rate Trend", benchmark=True, chart_key="email_click_rate")
    render_rate_trend(
        summary,
        "unsubscribe_rate",
        "Unsubscribe Rate Trend",
        benchmark=True,
        chart_key="email_unsubscribe_rate",
    )
    render_rate_trend(summary, "bounce_rate", "Bounce Rate Trend", benchmark=True, chart_key="email_bounce_rate")

    render_volume_trend(performance_data, "Email", "email_volume_detail")
    render_weekly_detail_table(performance_data, "Email")
    render_campaign_ranking(performance_data, "Email")


def render_sms_performance(data):
    st.subheader("SMS Performance")
    if data.empty:
        st.info("No SMS data for the selected week range.")
        return

    data = apply_segment_filter(data, "SMS audience segment", "sms_performance_segments")
    if data.empty:
        st.info("No SMS data for the selected audience segment.")
        return

    render_primary_kpis(data, channel="SMS")
    performance_data = data[data["has_performance_data"]].copy()
    if performance_data.empty:
        st.info("No SMS performance metrics yet for the selected week range. Audience totals still use SMS active audience as the new-user batch size.")
        return

    st.markdown("**SMS Segment Trends**")

    summary = weekly_summary(performance_data)
    render_rate_trend(summary, "click_rate", "SMS Click Rate Trend", chart_key="sms_click_rate")
    render_rate_trend(
        summary,
        "unsubscribe_rate",
        "SMS Unsubscribe Rate Trend",
        chart_key="sms_unsubscribe_rate",
    )

    render_volume_trend(performance_data, "SMS", "sms_volume_detail")
    render_weekly_detail_table(performance_data, "SMS")
    render_campaign_ranking(performance_data, "SMS")


def render_audience_movement(data):
    st.subheader("Audience Movement")
    if data.empty:
        st.info("No audience data for the selected week range.")
        return

    st.caption("Total audience uses the Email audience base. SMS active audience is treated as the new-user batch size.")

    data = data[data["audience_segment"].isin(["Build", "EOI", "Reservation"])].copy()
    if data.empty:
        st.info("No Build, EOI, or Reservation audience data for the selected week range.")
        return

    data = apply_segment_filter(data, "Audience segment", "audience_segments")
    if data.empty:
        st.info("No audience data for the selected section filters.")
        return

    email_audience = (
        data[data["channel"] == "Email"]
        .groupby(["date", "date_tag", "audience_segment"], as_index=False)
        .agg(
            active_audience=("active_audience", "max"),
        )
        .sort_values(["date", "audience_segment"])
    )

    if email_audience.empty:
        st.info("No Email audience base data for the selected filters.")
        return

    order = date_order(email_audience)
    total_audience = (
        email_audience.groupby(["date", "date_tag"], as_index=False)
        .agg(active_audience=("active_audience", "sum"))
        .sort_values("date")
    )
    total_audience["previous_audience"] = total_audience["active_audience"].shift(1)
    total_audience["audience_change"] = (
        total_audience["active_audience"] - total_audience["previous_audience"]
    )

    latest_audience_row = total_audience.iloc[-1]

    unsubscribes = (
        data.groupby(["date", "date_tag", "channel"], as_index=False)
        .agg(
            unsubscribes=("unsubscribes", "sum"),
            delivered=("delivered", "sum"),
        )
        .sort_values(["date", "channel"])
    )
    unsubscribes["unsubscribe_rate"] = unsubscribes.apply(
        lambda row: pct(row["unsubscribes"], row["delivered"]),
        axis=1,
    )
    unsubscribes["unsubscribe_type"] = unsubscribes["channel"] + " Unsubscribes"

    latest_unsubscribes = (
        unsubscribes.sort_values("date")
        .groupby("channel", as_index=False)
        .tail(1)
        .set_index("channel")
    )

    kpi1, kpi2, kpi3 = st.columns(3)
    audience_delta_value = latest_audience_row["audience_change"]
    kpi1.metric(
        "Total Audience",
        f"{latest_audience_row['active_audience']:,.0f}",
        None if pd.isna(audience_delta_value) else f"{audience_delta_value:+,.0f} vs previous week",
    )
    for column, channel in [(kpi2, "Email"), (kpi3, "SMS")]:
        if channel in latest_unsubscribes.index:
            latest_row = latest_unsubscribes.loc[channel]
            column.metric(
                f"{channel} Unsubscribes",
                f"{latest_row['unsubscribes']:,.0f}",
                f"{latest_row['unsubscribe_rate']:.2f}% rate",
                delta_color="off",
            )
        else:
            column.metric(f"{channel} Unsubscribes", "0", "No data", delta_color="off")

    audience_delta = total_audience.dropna(subset=["previous_audience"]).copy()
    audience_delta["change_type"] = audience_delta["audience_change"].apply(
        lambda value: "Increase" if value >= 0 else "Decrease"
    )

    if audience_delta.empty:
        st.info("Select at least two weeks to see audience change.")
    else:
        fig = px.bar(
            audience_delta,
            x="date_tag",
            y="audience_change",
            color="change_type",
            title="Total Audience Change",
            category_orders={"date_tag": order},
            color_discrete_map={
                "Increase": "#2E7D32",
                "Decrease": "#C62828",
            },
            hover_data={
                "previous_audience": ":,.0f",
                "active_audience": ":,.0f",
                "audience_change": "+,.0f",
                "change_type": False,
            },
        )
        fig.add_hline(y=0, line_color="#8A94A6", line_width=1)
        fig.update_xaxes(title="Week")
        fig.update_yaxes(title="Audience Change")
        st.plotly_chart(fig, width="stretch", key="audience_weekly_change")

    fig = px.line(
        total_audience,
        x="date_tag",
        y="active_audience",
        markers=True,
        title="Total Audience Trend",
        category_orders={"date_tag": order},
    )
    fig.update_xaxes(title="Week")
    fig.update_yaxes(title="Active Audience")
    st.plotly_chart(fig, width="stretch", key="audience_active_trend")

    fig = px.bar(
        unsubscribes,
        x="date_tag",
        y="unsubscribes",
        color="unsubscribe_type",
        barmode="group",
        title="Email vs SMS Unsubscribes",
        category_orders={"date_tag": order},
        color_discrete_map={
            "Email Unsubscribes": "#1565C0",
            "SMS Unsubscribes": "#6A1B9A",
        },
    )
    fig.update_xaxes(title="Week")
    fig.update_yaxes(title="Unsubscribes")
    st.plotly_chart(fig, width="stretch", key="audience_unsubscribes_by_channel")

    rate_data = unsubscribes[unsubscribes["delivered"] > 0].copy()
    if rate_data.empty:
        st.info("No delivered counts available to calculate unsubscribe rate.")
    else:
        fig = px.line(
            rate_data,
            x="date_tag",
            y="unsubscribe_rate",
            color="unsubscribe_type",
            markers=True,
            title="Email vs SMS Unsubscribe Rate",
            category_orders={"date_tag": order},
            color_discrete_map={
                "Email Unsubscribes": "#1565C0",
                "SMS Unsubscribes": "#6A1B9A",
            },
        )
        fig.update_xaxes(title="Week")
        fig.update_yaxes(title="Unsubscribe Rate", ticksuffix="%")
        st.plotly_chart(fig, width="stretch", key="audience_unsubscribe_rate_by_channel")

    st.dataframe(
        email_audience.sort_values(["date", "audience_segment"], ascending=[False, True]),
        width="stretch",
        hide_index=True,
    )


# --------------------------------
# APP
# --------------------------------
st.title("CRM Campaign Dashboard")

with st.expander("Database structure", expanded=False):
    st.markdown(
        f"""
This dashboard uses one Excel database: `{DATABASE_FILE}`.

Each row is one weekly campaign result for one audience segment and one channel.
The weekly `date` is the cohort tag. Email and SMS are separated automatically.

Expected columns:

`date`, `customer_type`, `campaign_name`, `active_audience`, `new_imports`, `sent`, `delivered`,
`opens`, `open_rate`, `clicks`, `click_rate`, `unsubscribes`, `unsubscribe_rate`, `bounces`, `bounce_rate`

Cleaning rules:

- `opene_rate` is normalized to `open_rate`
- blank rates stay blank unless count data is available to calculate them
- rows with no send/performance fields are treated as audience snapshots, not 0% campaign results
- for Email, `active_audience` is the current total audience, so KPI totals use the latest selected snapshot per segment
- for SMS, `active_audience` is the new-user batch size; `new_imports` is automatically aligned to the same number
- blank `sent` uses `active_audience` only when the row has campaign performance data
- blank `delivered` uses `sent - bounces` only when the row has campaign performance data
- SMS rows are detected when `campaign_name` or `customer_type` contains `SMS`

Update cadence:

- Email audience snapshots update on Wednesday
- SMS audience snapshots update on Monday, Wednesday, and Friday
        """
    )

# --------------------------------
# SIDEBAR FILTERS
# --------------------------------
st.sidebar.header("Filters")

if st.sidebar.button("Force reload data"):
    st.cache_data.clear()
    st.rerun()

database_path = Path(DATABASE_FILE)
database_modified_at = database_path.stat().st_mtime if database_path.exists() else None
raw_df = load_database(database_modified_at)
df = clean_database(raw_df)

if not database_path.exists():
    st.warning(f"`{DATABASE_FILE}` was not found. Showing sample data.")

if df.empty:
    st.error("No valid rows found. Please check that the database has valid dates and customer types.")
    st.stop()

week_lookup = (
    df[["date", "date_tag"]]
    .drop_duplicates()
    .sort_values("date")
    .reset_index(drop=True)
)
week_options = week_lookup["date_tag"].tolist()

if len(week_options) == 1:
    selected_week_tags = week_options
    st.sidebar.selectbox("Week", week_options, index=0)
else:
    default_period_length = min(4, max(1, len(week_options) // 2))
    default_start = max(0, len(week_options) - default_period_length)
    selected_week_range = st.sidebar.select_slider(
        "Week range",
        options=week_options,
        value=(week_options[default_start], week_options[-1]),
    )
    start_index = week_options.index(selected_week_range[0])
    end_index = week_options.index(selected_week_range[1])
    if start_index > end_index:
        start_index, end_index = end_index, start_index
    selected_week_tags = week_options[start_index : end_index + 1]

st.sidebar.caption(f"{selected_week_tags[0]} -> {selected_week_tags[-1]}")

week_filtered = df[df["date_tag"].isin(selected_week_tags)].copy()
selected_period_start = week_filtered["date"].min().date()
selected_period_end = week_filtered["date"].max().date()
comparison_period_start, comparison_period_end = previous_period_range(
    selected_period_start,
    selected_period_end,
)
previous_week_filtered = df[
    (df["date"].dt.date >= comparison_period_start)
    & (df["date"].dt.date <= comparison_period_end)
].copy()

latest_tag = week_filtered["date"].max()
if pd.notna(latest_tag):
    st.caption(f"Latest weekly date tag in current filters: {latest_tag.strftime('%Y-%m-%d')}")


# --------------------------------
# TABS
# --------------------------------
overview_tab, email_tab, sms_tab, audience_tab, raw_tab = st.tabs(
    [
        "Executive Overview",
        "Email Performance",
        "SMS Performance",
        "Audience Movement",
        "Raw Database",
    ]
)

with overview_tab:
    render_executive_overview(
        week_filtered,
        previous_week_filtered,
        selected_period=(selected_period_start, selected_period_end),
        comparison_period=(comparison_period_start, comparison_period_end),
        full_data=df,
    )

with email_tab:
    render_email_performance(week_filtered[week_filtered["channel"] == "Email"].copy())

with sms_tab:
    render_sms_performance(week_filtered[week_filtered["channel"] == "SMS"].copy())

with audience_tab:
    render_audience_movement(week_filtered)

with raw_tab:
    st.subheader("Raw Database")
    raw_filtered = apply_channel_segment_filter(week_filtered, "raw")
    st.dataframe(
        raw_filtered.sort_values(["date", "channel", "audience_segment"], ascending=[False, True, True]),
        width="stretch",
        hide_index=True,
    )
    st.download_button(
        "Download filtered CSV",
        raw_filtered.to_csv(index=False),
        "crm_campaign_filtered.csv",
        "text/csv",
    )
