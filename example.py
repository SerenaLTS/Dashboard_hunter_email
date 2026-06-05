import streamlit as st
import pandas as pd
import plotly.express as px
import os

# --------------------------------
# PAGE SETUP
# --------------------------------
st.set_page_config(
    page_title="Dealer Weekly Enquiry Dashboard",
    layout="wide"
)

# --------------------------------
# LOAD DATA
# --------------------------------
@st.cache_data
def load_data():
    file_path = "JAC_Daily_Enquiry_With_Website_REBUILT.xlsx"

    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path, parse_dates=["week_start"])
    else:
        df = pd.read_csv(file_path, parse_dates=["week_start"])

    # 统一列名（内部用更简短的名字，原始列会被重命名）
    rename_map = {
        "Enquiries for the week [Walk In]": "walk_in",
        "Enquiries for the week [Phone]": "phone",
        "Enquiries for the week [Website]": "website",
        "Enquiries for the week [Carsales]": "carsales",
        "Test Drives ": "test_drives",
        "Sold [Oasis]": "sold_oasis",
        "Enquiries for the week [Autotrader / Gumtree]": "autotrader_gumtree",
        "Enquiries for the week [Social Media]": "social_media",
        "Sold [Haven]": "sold_haven",
        "Sold [Osprey X]": "sold_osprey_x",
        "Dealer_Name": "dealer_name",
        "Dealer_State": "dealer_state",
        "Dealer_Region": "dealer_region",
    }
    df = df.rename(columns=rename_map)

    # ---- ✅ 基础清洗：保证 key & 字段存在 ----
    if "Dealer_Code" not in df.columns:
        df["Dealer_Code"] = ""

    df["Dealer_Code"] = df["Dealer_Code"].astype(str).str.strip()
    df["dealer_name"] = df.get("dealer_name", "").astype(str).str.strip()
    df["dealer_state"] = df.get("dealer_state", "").astype(str).str.strip()
    # --------------------------------
    # ❌ REMOVE ROWS WITH NO DEALER INFO
    # --------------------------------
    # 判断 dealer 信息是否全部为空
    no_dealer_mask = (
        (df["Dealer_Code"].isna() | (df["Dealer_Code"].str.strip() == "") | (df["Dealer_Code"].str.lower() == "nan"))
        &
        (df["dealer_name"].isna() | (df["dealer_name"].str.strip() == "") | (df["dealer_name"].str.lower() == "nan"))
        &
        (df["dealer_state"].isna() | (df["dealer_state"].str.strip() == "") | (df["dealer_state"].str.lower() == "nan"))
    )

    df = df[~no_dealer_mask].copy()
    df["week_start"] = pd.to_datetime(df["week_start"], errors="coerce").dt.normalize()

    # ---- ✅ carsales actual（来自 merge 后的列）----
    if "carsales_actual" not in df.columns:
        df["carsales_actual"] = 0

    df["carsales_actual"] = pd.to_numeric(df["carsales_actual"], errors="coerce").fillna(0)
    if "carsales" not in df.columns:
        df["carsales"] = 0
    df["carsales"] = pd.to_numeric(df["carsales"], errors="coerce").fillna(0)

    # ✅ NEW RULE：
    # 只要该 week_start 存在 carsales_actual>0，则当周 carsales 只用虚拟 dealer（000000 / carsales / national）
    carsales_code = "000000"
    is_carsales_row = (
        (df["Dealer_Code"].str.zfill(6) == carsales_code)
        | (df["dealer_name"].str.lower().str.strip() == "carsales")
    )

    # 某周是否有 actual
    week_has_actual = df.groupby("week_start")["carsales_actual"].transform("max") > 0

    # 先默认用原 carsales
    carsales_effective = df["carsales"].copy()

    # 如果某周有 actual：全部 dealer 的 carsales 清零
    carsales_effective = carsales_effective.mask(week_has_actual, 0)

    # 如果某周有 actual：仅 carsales 虚拟 dealer 取 carsales_actual
    carsales_effective = carsales_effective.mask(week_has_actual & is_carsales_row, df["carsales_actual"])

    # 覆盖 carsales 列（后续所有图表/总和都会自动用新的 carsales）
    df["carsales"] = pd.to_numeric(carsales_effective, errors="coerce").fillna(0)

    # ------------ Orders / Sold 逻辑 ------------
    order_cols = ["sold_oasis", "sold_haven", "sold_osprey_x"]
    for c in order_cols:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["total_order"] = df[order_cols].sum(axis=1)

    # ------------ Website warm leads 逻辑 ------------
    warm_cols = [
        "book_test_drive",
        "enquire_special_tiktok",
        "enquire_special_non_tiktok",
        "showroom_enquiry",
        "trade_in",
    ]
    for c in warm_cols:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["website_warm_total"] = df[warm_cols].sum(axis=1)
    df["website"] = df["website_warm_total"]

    # ------------ total enquiry ------------
    enquiry_cols = [
        "walk_in",
        "phone",
        "website",
        "carsales",  # ✅ 已按规则被覆盖
        "autotrader_gumtree",
        "social_media",
    ]
    for c in enquiry_cols:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["total_enquiry"] = df[enquiry_cols].sum(axis=1)

    df["week_label"] = df["week_start"].dt.strftime("Week of %d/%m/%Y")
    return df

df = load_data()
# ✅ Compute weeks that used Carsales Actual (global, before expander)
carsales_code = "000000"
is_carsales_row_df = (
    (df["Dealer_Code"].astype(str).str.zfill(6) == carsales_code)
    | (df["dealer_name"].astype(str).str.lower().str.strip() == "carsales")
)
actual_weeks = (
    df.loc[is_carsales_row_df]
      .groupby("week_start")["carsales_actual"]
      .max()
)
actual_weeks = actual_weeks[actual_weeks > 0].index.sort_values()

if len(actual_weeks) > 0:
    actual_weeks_str = ", ".join([pd.to_datetime(w).strftime("%d/%m/%Y") for w in actual_weeks])
else:
    actual_weeks_str = "None"

st.title("📊 Dealer Weekly Enquiry Dashboard")

with st.expander("Data sources & definitions"):
    st.markdown(f"""
    **Data sources**
    - Dealer weekly enquiry report (Walk-in, Phone, Website, Carsales, Social)
    - Digital Dealer (website, warm leads)
    - Carsales Actual (virtual dealer: Dealer_Code=000000, dealer_name=carsales, dealer_state=national)

    **Carsales Actual weeks**
    - {actual_weeks_str}

    **Notes**
    - If a week has Carsales Actual, that week uses ONLY the Carsales (virtual dealer) value for the carsales channel.
    - Filters apply to all charts below
    """)

# --------------------------------
# SIDEBAR FILTERS
# --------------------------------
st.sidebar.header("Filters")

if st.sidebar.button("🔄 Force reload data"):
    st.cache_data.clear()

min_week = df["week_start"].min().date()
max_week = df["week_start"].max().date()

date_mode = st.sidebar.radio(
    "Time range (by week_start)",
    ["Custom range", "Last 7 days", "Last 30 days", "Last 90 days"],
    index=0
)

if date_mode == "Custom range":
    date_range = st.sidebar.date_input(
        "Select week_start range",
        value=(min_week, max_week)
    )
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = date_range
        end_date = date_range
else:
    days_lookup = {"Last 7 days": 7, "Last 30 days": 30, "Last 90 days": 90}
    days = days_lookup[date_mode]
    start_date = max_week - pd.Timedelta(days=days - 1)
    if start_date < min_week:
        start_date = min_week
    end_date = max_week
    st.sidebar.info(f"{start_date} → {end_date}")

states = sorted(df["dealer_state"].dropna().unique())
select_all_states = st.sidebar.checkbox("Select All States", value=True)
if select_all_states:
    selected_states = states
else:
    selected_states = st.sidebar.multiselect("Select State(s)", states, default=states)

dealers = sorted(df["dealer_name"].dropna().unique())
select_all_dealers = st.sidebar.checkbox("Select All Dealers", value=True)
if select_all_dealers:
    selected_dealers = dealers
else:
    selected_dealers = st.sidebar.multiselect("Select Dealer(s)", dealers, default=dealers)

website_view = st.sidebar.radio(
    "Website channel view",
    ["Combined website (warm total)", "Split website warm types"],
    index=0
)

# --------------------------------
# APPLY FILTERS
# --------------------------------
filtered = df.copy()
start_ts = pd.to_datetime(start_date)
end_ts = pd.to_datetime(end_date)

filtered = filtered[
    (filtered["week_start"] >= start_ts)
    & (filtered["week_start"] <= end_ts)
]

if selected_states:
    filtered = filtered[filtered["dealer_state"].isin(selected_states)]

if selected_dealers:
    filtered = filtered[filtered["dealer_name"].isin(selected_dealers)]

# --------------------------------
# KPIs
# --------------------------------
k1, k2, k3, k4 = st.columns(4)

total_enquiries = int(filtered["total_enquiry"].sum()) if not filtered.empty else 0
total_orders = int(filtered["total_order"].sum()) if (not filtered.empty and "total_order" in filtered.columns) else 0
active_dealers = filtered["dealer_name"].nunique()
num_states = filtered["dealer_state"].nunique()

k1.metric("Total Enquiries", total_enquiries)
k2.metric("Total Orders", total_orders)
k3.metric("Active Dealers", active_dealers)
k4.metric("States", num_states)

st.markdown(f"**Current time range:** {start_date} → {end_date}")
st.markdown("---")

# --------------------------------
# CHARTS
# --------------------------------
if not filtered.empty:
    base_non_web = [
        c for c in ["walk_in", "phone", "carsales", "autotrader_gumtree", "social_media"]
        if c in filtered.columns
    ]
    warm_cols = [
        c for c in [
            "book_test_drive",
            "enquire_special_tiktok",
            "enquire_special_non_tiktok",
            "showroom_enquiry",
            "trade_in",
        ]
        if c in filtered.columns
    ]

    if website_view == "Combined website (warm total)":
        enquiry_cols_base = base_non_web + (["website"] if "website" in filtered.columns else [])
    else:
        enquiry_cols_base = base_non_web + warm_cols

    # ✅ weekly chart 专用（先复制一份）
    enquiry_cols_weekly = enquiry_cols_base.copy()
    
    # ✅ week-level flag: does this week use carsales actual?
    carsales_code = "000000"
    is_carsales_row = (
        (filtered["Dealer_Code"].astype(str).str.zfill(6) == carsales_code)
        | (filtered["dealer_name"].astype(str).str.lower().str.strip() == "carsales")
    )
    week_has_actual = (
        filtered.loc[is_carsales_row]
        .groupby("week_start")["carsales_actual"]
        .max()
        .gt(0)
    )

    weekly_channel = (
        filtered
        .groupby("week_start")[enquiry_cols_weekly]
        .sum()
        .reset_index()
        .sort_values("week_start")
    )

    # ✅ attach flag to weekly_channel
    weekly_channel["has_carsales_actual"] = (
        weekly_channel["week_start"].map(week_has_actual).fillna(False).astype(bool)
    )

    # ✅ split carsales into two series for legend + color control
    # ✅ split carsales into two series for legend + color control
    if "carsales" in weekly_channel.columns:
        weekly_channel["carsales_actual_used"] = weekly_channel["carsales"].where(weekly_channel["has_carsales_actual"], 0)
        weekly_channel["carsales_reported"] = weekly_channel["carsales"].where(~weekly_channel["has_carsales_actual"], 0)

        if "carsales" in enquiry_cols_weekly:
            idx = enquiry_cols_weekly.index("carsales")
            enquiry_cols_weekly.pop(idx)
            enquiry_cols_weekly[idx:idx] = ["carsales_reported", "carsales_actual_used"]

    weekly_channel["week_start_str"] = weekly_channel["week_start"].dt.strftime("%d/%m/%Y")
    weekly_channel["total_for_chart"] = weekly_channel[enquiry_cols_weekly].sum(axis=1)
    weekly_channel = weekly_channel[weekly_channel["total_for_chart"] > 0]

    if not weekly_channel.empty:
        weekly_melted = weekly_channel.melt(
            id_vars="week_start_str",
            value_vars=enquiry_cols_weekly,
            var_name="channel",
            value_name="enquiries"
        )

        channel_label_map = {
            "walk_in": "Walk In",
            "phone": "Phone",
            "website": "Website (warm)",
            "carsales_reported": "Carsales (Reported)",
            "carsales_actual_used": "Carsales (Actual)",
            "autotrader_gumtree": "Autotrader / Gumtree",
            "social_media": "Social Media",
            "book_test_drive": "Book Test Drive",
            "enquire_special_tiktok": "Enquire Special (TikTok)",
            "enquire_special_non_tiktok": "Enquire Special (Non-TikTok)",
            "showroom_enquiry": "Showroom Enquiry",
            "trade_in": "Trade-in Enquiry",
        }

        weekly_melted["channel_label"] = weekly_melted["channel"].map(channel_label_map).fillna(weekly_melted["channel"])

        color_map = {
            "Carsales (Actual)": "#800080"   # 紫色
        }

        fig_week = px.bar(
            weekly_melted,
            x="week_start_str",
            y="enquiries",
            color="channel_label",
            title="Total Enquiries by Week",
            labels={
                "week_start_str": "Week Start",
                "enquiries": "Enquiries",
                "channel_label": "Channel",
            },
            color_discrete_map=color_map
        )
        st.plotly_chart(fig_week, use_container_width=True)

        # ---------- Total Orders by Week ----------
        order_cols = [c for c in ["sold_oasis", "sold_haven", "sold_osprey_x"] if c in filtered.columns]
        order_label_map = {
            "sold_oasis": "Sold (Oasis)",
            "sold_haven": "Sold (Haven)",
            "sold_osprey_x": "Sold (Osprey X)",
        }

        weekly_orders = (
            filtered
            .groupby("week_start")[order_cols]
            .sum()
            .reset_index()
            .sort_values("week_start")
        )

        if not weekly_orders.empty:
            weekly_orders["week_start_str"] = weekly_orders["week_start"].dt.strftime("%d/%m/%Y")
            weekly_orders["total_orders_for_chart"] = weekly_orders[order_cols].sum(axis=1)
            weekly_orders = weekly_orders[weekly_orders["total_orders_for_chart"] > 0]

        if not weekly_orders.empty:
            orders_melted = weekly_orders.melt(
                id_vars="week_start_str",
                value_vars=order_cols,
                var_name="order_type",
                value_name="orders"
            )
            orders_melted["order_label"] = orders_melted["order_type"].map(order_label_map).fillna(orders_melted["order_type"])

            fig_orders_week = px.bar(
                orders_melted,
                x="week_start_str",
                y="orders",
                color="order_label",
                title="Total Orders by Week",
                labels={
                    "week_start_str": "Week Start",
                    "orders": "Orders",
                    "order_label": "Sold Type",
                },
            )
            st.plotly_chart(fig_orders_week, use_container_width=True)
        else:
            st.info("No weeks with non-zero orders in the selected filters.")

        st.subheader("Weekly Summary (Enquiries + Orders)")

        weekly_table = weekly_channel.copy()
        weekly_table["Week_Start"] = weekly_table["week_start_str"]
        weekly_table["Total_Enquiries"] = weekly_table[enquiry_cols_weekly].sum(axis=1)

        weekly_orders_for_table = (
            filtered
            .groupby("week_start")[order_cols]
            .sum()
            .reset_index()
        )
        weekly_table = weekly_table.merge(weekly_orders_for_table, on="week_start", how="left")
        weekly_table[order_cols] = weekly_table[order_cols].fillna(0)
        weekly_table["Total_Orders"] = weekly_table[order_cols].sum(axis=1)

        weekly_table = weekly_table.drop(
            columns=["week_start", "week_start_str", "total_for_chart"],
            errors="ignore"
        )

        cols_order = ["Week_Start", "Total_Enquiries", "Total_Orders"] + order_cols + enquiry_cols_weekly
        weekly_table = weekly_table[cols_order]

        st.dataframe(weekly_table, use_container_width=True)

    else:
        st.info("No weeks with non-zero enquiries in the selected filters.")

    # ---------- Total Enquiries by Dealer ----------
    dealer_totals = (
        filtered
        .groupby("dealer_name")["total_enquiry"]
        .sum()
        .reset_index()
        .sort_values("total_enquiry", ascending=False)
    )

    fig_dealer = px.bar(
        dealer_totals,
        x="dealer_name",
        y="total_enquiry",
        title="Total Enquiries by Dealer",
        labels={"dealer_name": "Dealer", "total_enquiry": "Enquiries"},
    )
    fig_dealer.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_dealer, use_container_width=True)

    # ---------- Total Orders by Dealer ----------
    if "total_order" in filtered.columns:
        dealer_order_totals = (
            filtered
            .groupby("dealer_name")["total_order"]
            .sum()
            .reset_index()
            .sort_values("total_order", ascending=False)
        )

        fig_dealer_orders = px.bar(
            dealer_order_totals,
            x="dealer_name",
            y="total_order",
            title="Total Orders by Dealer",
            labels={"dealer_name": "Dealer", "total_order": "Orders"},
        )
        fig_dealer_orders.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_dealer_orders, use_container_width=True)

    # --------------------------------
    # Dealer view over time
    # --------------------------------
    st.markdown("### 🔍 Dealer view over time")

    dealer_for_time = st.selectbox(
        "Select a dealer (weekly trend)",
        options=sorted(filtered["dealer_name"].unique()),
        key="dealer_time_view"
    )

    dealer_time_df = (
        filtered[filtered["dealer_name"] == dealer_for_time]
        .groupby("week_start")[enquiry_cols_base]
        .sum()
        .reset_index()
        .sort_values("week_start")
    )

    if not dealer_time_df.empty:
        dealer_time_df["week_start_str"] = dealer_time_df["week_start"].dt.strftime("%d/%m/%Y")
        dealer_time_df["total"] = dealer_time_df[enquiry_cols_base].sum(axis=1)

        dealer_melted = dealer_time_df.melt(
            id_vars="week_start_str",
            value_vars=enquiry_cols_base,
            var_name="channel",
            value_name="enquiries"
        )

        dealer_melted["channel_label"] = dealer_melted["channel"].map(channel_label_map).fillna(dealer_melted["channel"])

        fig_dealer_time = px.bar(
            dealer_melted,
            x="week_start_str",
            y="enquiries",
            color="channel_label",
            title=f"Weekly Enquiries — {dealer_for_time}",
            labels={
                "week_start_str": "Week Start",
                "enquiries": "Enquiries",
                "channel_label": "Channel",
            },
        )
        st.plotly_chart(fig_dealer_time, use_container_width=True)

        dealer_time_table = dealer_time_df.drop(columns=["week_start"])
        dealer_time_table.insert(0, "Week", dealer_time_df["week_start_str"])
        st.dataframe(dealer_time_table, use_container_width=True)

        # orders overtime
        order_cols = [c for c in ["sold_oasis", "sold_haven", "sold_osprey_x"] if c in filtered.columns]
        dealer_orders_df = (
            filtered[filtered["dealer_name"] == dealer_for_time]
            .groupby("week_start")[order_cols]
            .sum()
            .reset_index()
            .sort_values("week_start")
        )

        if not dealer_orders_df.empty:
            dealer_orders_df["week_start_str"] = dealer_orders_df["week_start"].dt.strftime("%d/%m/%Y")
            dealer_orders_df["total_order"] = dealer_orders_df[order_cols].sum(axis=1)

            dealer_orders_melted = dealer_orders_df.melt(
                id_vars="week_start_str",
                value_vars=order_cols,
                var_name="order_type",
                value_name="orders"
            )
            dealer_orders_melted["order_label"] = dealer_orders_melted["order_type"].map(order_label_map).fillna(dealer_orders_melted["order_type"])

            fig_dealer_orders_time = px.bar(
                dealer_orders_melted,
                x="week_start_str",
                y="orders",
                color="order_label",
                title=f"Weekly Orders — {dealer_for_time}",
                labels={
                    "week_start_str": "Week Start",
                    "orders": "Orders",
                    "order_label": "Sold Type",
                },
            )
            st.plotly_chart(fig_dealer_orders_time, use_container_width=True)

            dealer_orders_table = dealer_orders_df.drop(columns=["week_start"])
            dealer_orders_table.insert(0, "Week", dealer_orders_df["week_start_str"])
            st.dataframe(dealer_orders_table, use_container_width=True)
        else:
            st.info("No orders for this dealer in the selected time range.")
    else:
        st.info("No data for this dealer in the selected time range.")

    # ---------- Total Enquiries by State ----------
    state_totals = (
        filtered
        .groupby("dealer_state")["total_enquiry"]
        .sum()
        .reset_index()
        .sort_values("total_enquiry", ascending=False)
    )

    fig_state = px.bar(
        state_totals,
        x="dealer_state",
        y="total_enquiry",
        title="Total Enquiries by State",
        labels={"dealer_state": "State", "total_enquiry": "Enquiries"},
    )
    st.plotly_chart(fig_state, use_container_width=True)

    # --------------------------------
    # State view over time
    # --------------------------------
    st.markdown("### 🌏 State view over time")

    state_for_time = st.selectbox(
        "Select a state (weekly trend)",
        options=sorted(filtered["dealer_state"].dropna().unique()),
        key="state_time_view"
    )

    state_time_df = (
        filtered[filtered["dealer_state"] == state_for_time]
        .groupby("week_start")[enquiry_cols_base]
        .sum()
        .reset_index()
        .sort_values("week_start")
    )

    if not state_time_df.empty:
        state_time_df["week_start_str"] = state_time_df["week_start"].dt.strftime("%d/%m/%Y")
        state_time_df["total"] = state_time_df[enquiry_cols_base].sum(axis=1)

        state_melted = state_time_df.melt(
            id_vars="week_start_str",
            value_vars=enquiry_cols_base,
            var_name="channel",
            value_name="enquiries"
        )

        state_melted["channel_label"] = state_melted["channel"].map(channel_label_map).fillna(state_melted["channel"])

        fig_state_time = px.bar(
            state_melted,
            x="week_start_str",
            y="enquiries",
            color="channel_label",
            title=f"Weekly Enquiries — State: {state_for_time}",
            labels={
                "week_start_str": "Week Start",
                "enquiries": "Enquiries",
                "channel_label": "Channel",
            },
        )
        st.plotly_chart(fig_state_time, use_container_width=True)

        state_time_table = state_time_df.drop(columns=["week_start"])
        state_time_table.insert(0, "Week", state_time_df["week_start_str"])
        st.dataframe(state_time_table, use_container_width=True)

        # orders overtime (state)
        order_cols = [c for c in ["sold_oasis", "sold_haven", "sold_osprey_x"] if c in filtered.columns]
        state_orders_df = (
            filtered[filtered["dealer_state"] == state_for_time]
            .groupby("week_start")[order_cols]
            .sum()
            .reset_index()
            .sort_values("week_start")
        )

        if not state_orders_df.empty:
            state_orders_df["week_start_str"] = state_orders_df["week_start"].dt.strftime("%d/%m/%Y")
            state_orders_df["total_order"] = state_orders_df[order_cols].sum(axis=1)

            state_orders_melted = state_orders_df.melt(
                id_vars="week_start_str",
                value_vars=order_cols,
                var_name="order_type",
                value_name="orders"
            )
            state_orders_melted["order_label"] = state_orders_melted["order_type"].map(order_label_map).fillna(state_orders_melted["order_type"])

            fig_state_orders_time = px.bar(
                state_orders_melted,
                x="week_start_str",
                y="orders",
                color="order_label",
                title=f"Weekly Orders — State: {state_for_time}",
                labels={
                    "week_start_str": "Week Start",
                    "orders": "Orders",
                    "order_label": "Sold Type",
                },
            )
            st.plotly_chart(fig_state_orders_time, use_container_width=True)

            state_orders_table = state_orders_df.drop(columns=["week_start"])
            state_orders_table.insert(0, "Week", state_orders_df["week_start_str"])
            st.dataframe(state_orders_table, use_container_width=True)
        else:
            st.info("No orders for this state in the selected time range.")
    else:
        st.info("No data for this state in the selected time range.")
else:
    st.warning("No data for the selected filters.")

# --------------------------------
# DETAIL TABLE
# --------------------------------
st.subheader("Filtered Records (RAW)")

if not filtered.empty:
    if "dealer_name" in filtered.columns:
        df_detail = filtered.sort_values(["week_start", "dealer_name"])
    else:
        df_detail = filtered.sort_values(["week_start"])
    st.dataframe(df_detail, use_container_width=True)
else:
    st.write("No rows match the current filters.")