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


# --------------------------------
# DATA CLEANING HELPERS
# --------------------------------
def pct(numerator, denominator):
    if denominator == 0:
        return 0
    return numerator / denominator * 100


def normalize_rate(value):
    if pd.isna(value):
        return 0
    value = float(value)
    if value <= 1:
        return value * 100
    return value


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

    numeric_columns = [
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
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

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

    df["open_rate"] = df.apply(
        lambda row: pct(row["opens"], row["delivered"])
        if row["opens"] and row["delivered"]
        else normalize_rate(row["open_rate"]),
        axis=1,
    )
    df["click_rate"] = df.apply(
        lambda row: pct(row["clicks"], row["delivered"])
        if row["clicks"] and row["delivered"]
        else normalize_rate(row["click_rate"]),
        axis=1,
    )
    df["unsubscribe_rate"] = df.apply(
        lambda row: pct(row["unsubscribes"], row["delivered"])
        if row["unsubscribes"] and row["delivered"]
        else normalize_rate(row["unsubscribe_rate"]),
        axis=1,
    )
    df["bounce_rate"] = df.apply(
        lambda row: pct(row["bounces"], row["sent"])
        if row["bounces"] and row["sent"]
        else normalize_rate(row["bounce_rate"]),
        axis=1,
    )

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
def load_database():
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
        "open_rate": pct(opens, delivered),
        "click_rate": pct(clicks, delivered),
        "unsubscribe_rate": pct(unsubscribes, delivered),
        "bounce_rate": pct(bounces, sent),
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
        )
        .sort_values(["date", "audience_segment"])
    )
    summary["open_rate"] = summary.apply(lambda row: pct(row["opens"], row["delivered"]), axis=1)
    summary["click_rate"] = summary.apply(lambda row: pct(row["clicks"], row["delivered"]), axis=1)
    summary["unsubscribe_rate"] = summary.apply(lambda row: pct(row["unsubscribes"], row["delivered"]), axis=1)
    summary["bounce_rate"] = summary.apply(lambda row: pct(row["bounces"], row["sent"]), axis=1)
    return summary


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


def render_primary_kpis(data, channel=None):
    metrics = total_metrics(data)

    if channel == "SMS":
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}")
        kpi2.metric("Click Rate", f"{metrics['click_rate']:.1f}%")
        kpi3.metric("Unsubscribe Rate", f"{metrics['unsubscribe_rate']:.2f}%")
        kpi4.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%")
    elif channel is None:
        email_data = data[data["channel"] == "Email"]
        email_metrics = total_metrics(email_data)
        if email_data.empty:
            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}")
            kpi2.metric("Click Rate", f"{metrics['click_rate']:.1f}%")
            kpi3.metric("Unsubscribe Rate", f"{metrics['unsubscribe_rate']:.2f}%")
            kpi4.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%")
        else:
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}")
            kpi2.metric("Email Open Rate", f"{email_metrics['open_rate']:.1f}%")
            kpi3.metric("Click Rate", f"{metrics['click_rate']:.1f}%")
            kpi4.metric("Unsubscribe Rate", f"{metrics['unsubscribe_rate']:.2f}%")
            kpi5.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%")
    else:
        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("Total Audience", f"{metrics['active_audience']:,.0f}")
        kpi2.metric("Open Rate", f"{metrics['open_rate']:.1f}%")
        kpi3.metric("Click Rate", f"{metrics['click_rate']:.1f}%")
        kpi4.metric("Unsubscribe Rate", f"{metrics['unsubscribe_rate']:.2f}%")
        kpi5.metric("Bounce Rate", f"{metrics['bounce_rate']:.2f}%")

    raw1, raw2, raw3, raw4 = st.columns(4)
    raw1.metric("Total Sent", f"{metrics['sent']:,.0f}")
    raw2.metric("Delivered", f"{metrics['delivered']:,.0f}")
    raw3.metric("Clicks", f"{metrics['clicks']:,.0f}")
    raw4.metric("New Imports", f"{metrics['new_imports']:,.0f}")


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


def render_executive_overview(data):
    st.subheader("Executive Overview")
    if data.empty:
        st.info("No data for the selected week range.")
        return

    data = apply_channel_segment_filter(data, "overview")
    if data.empty:
        st.info("No data for the selected overview filters.")
        return

    selected_channel = data["channel"].iloc[0]

    st.markdown("**Executive KPIs**")
    render_primary_kpis(data, channel=selected_channel)

    performance_data = data[data["has_performance_data"]].copy()
    if performance_data.empty:
        st.info(f"No {selected_channel} performance metrics yet for the selected week range.")
    else:
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

    render_funnel(performance_data, "Email", "email_performance_funnel")

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

    render_funnel(performance_data, "SMS", "sms_performance_funnel")

    summary = weekly_summary(performance_data)
    render_rate_trend(summary, "click_rate", "SMS Click Rate Trend", chart_key="sms_click_rate")
    render_rate_trend(
        summary,
        "unsubscribe_rate",
        "SMS Unsubscribe Rate Trend",
        chart_key="sms_unsubscribe_rate",
    )

    render_campaign_ranking(performance_data, "SMS")


def render_audience_movement(data):
    st.subheader("Audience Movement")
    if data.empty:
        st.info("No audience data for the selected week range.")
        return

    data = apply_channel_segment_filter(data, "audience")
    if data.empty:
        st.info("No audience data for the selected section filters.")
        return

    audience = (
        data.groupby(["date", "date_tag", "channel", "audience_segment"], as_index=False)
        .agg(
            active_audience=("active_audience", "max"),
            new_imports=("new_imports", "sum"),
            unsubscribes=("unsubscribes", "sum"),
        )
        .sort_values(["date", "channel", "audience_segment"])
    )
    order = date_order(audience)

    audience_delta = audience.sort_values(["channel", "audience_segment", "date"]).copy()
    audience_delta["previous_audience"] = audience_delta.groupby(
        ["channel", "audience_segment"]
    )["active_audience"].shift(1)
    audience_delta["audience_change"] = (
        audience_delta["active_audience"] - audience_delta["previous_audience"]
    )
    audience_delta = audience_delta.dropna(subset=["previous_audience"]).copy()
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
            color="audience_segment",
            barmode="group",
            pattern_shape="change_type",
            title="Weekly Audience Change",
            category_orders={"date_tag": order},
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
        audience,
        x="date_tag",
        y="active_audience",
        color="audience_segment",
        line_dash="channel",
        markers=True,
        title="Active Audience Trend",
        category_orders={"date_tag": order},
    )
    fig.update_xaxes(title="Week")
    st.plotly_chart(fig, width="stretch", key="audience_active_trend")

    movement = audience.melt(
        id_vars=["date_tag", "channel", "audience_segment"],
        value_vars=["new_imports", "unsubscribes"],
        var_name="metric",
        value_name="contacts",
    )
    fig = px.bar(
        movement,
        x="date_tag",
        y="contacts",
        color="metric",
        barmode="group",
        pattern_shape="channel",
        title="New Imports vs Unsubscribes",
        category_orders={"date_tag": order},
    )
    fig.update_xaxes(title="Week")
    st.plotly_chart(fig, width="stretch", key="audience_imports_unsubscribes")

    st.dataframe(
        audience.sort_values(["date", "channel", "audience_segment"], ascending=[False, True, True]),
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
- blank rates are calculated from counts
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

raw_df = load_database()
df = clean_database(raw_df)

if not Path(DATABASE_FILE).exists():
    st.warning(f"`{DATABASE_FILE}` was not found. Showing sample data.")


# --------------------------------
# SIDEBAR FILTERS
# --------------------------------
st.sidebar.header("Filters")

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
    default_start = max(0, len(week_options) - 4)
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
    render_executive_overview(week_filtered)

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
