"""Prepare Digital Dealer enquiries for weekly audience movement reporting."""

from pathlib import Path
import re

import pandas as pd


SOURCE_PATTERN = "digital_dealer_enquiries_*.csv"
OUTPUT_DIR = Path("processed")

FORM_GROUP_MAP = {
    "hunterupdates": ("JAC Hunter Build", "Build", 2),
    "register your interest jac hunter": ("JAC Hunter", "EOI", 1),
    "reserve your build": ("JAC Hunter Reserve", "Reservation", 3),
}

EVENT_COLUMNS = [
    "lead_id",
    "lead_datetime",
    "week_start",
    "week_end",
    "week_complete",
    "email",
    "customer_group",
    "segment_code",
    "group_stage",
    "form",
    "name",
    "surname",
    "phone",
    "location",
    "postcode",
    "state",
    "status",
    "deposit",
    "source_url",
    "submission_url",
]


def normalize_email(value):
    value = str(value).strip().lower()
    if not value or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", value):
        return pd.NA
    return value


def parse_lead_datetime(values):
    extracted = values.astype(str).str.extract(
        r"(?P<time>\d{1,2}:\d{2}\s*[ap]m)\s*-\s*"
        r"(?P<date>\d{1,2}/\d{1,2}/\d{2,4})",
        flags=re.IGNORECASE,
    )
    return pd.to_datetime(
        extracted["date"] + " " + extracted["time"],
        format="%d/%m/%y %I:%M%p",
        errors="coerce",
    )


def clean_events(raw):
    data = raw.copy()
    data["form_key"] = data["Form"].astype(str).str.strip().str.lower()
    unknown_forms = sorted(set(data["form_key"]) - set(FORM_GROUP_MAP))
    if unknown_forms:
        raise ValueError(f"Unmapped Form value(s): {unknown_forms}")

    mapping = data["form_key"].map(FORM_GROUP_MAP)
    data["customer_group"] = mapping.str[0]
    data["segment_code"] = mapping.str[1]
    data["group_stage"] = mapping.str[2]
    data["email"] = data["Email"].map(normalize_email)
    data["lead_datetime"] = parse_lead_datetime(data["Lead Time"])

    data["week_start"] = data["lead_datetime"].dt.normalize() - pd.to_timedelta(
        data["lead_datetime"].dt.weekday, unit="D"
    )
    data["week_end"] = data["week_start"] + pd.Timedelta(days=6)
    today = pd.Timestamp.now(tz="Australia/Sydney").tz_localize(None).normalize()
    data["week_complete"] = data["week_end"] < today

    rename_map = {
        "Lead ID": "lead_id",
        "Form": "form",
        "Name": "name",
        "Surname": "surname",
        "Phone": "phone",
        "Location": "location",
        "Postcode": "postcode",
        "State": "state",
        "Status": "status",
        "Deposit": "deposit",
        "Source URL": "source_url",
        "Submission URL": "submission_url",
    }
    data = data.rename(columns=rename_map)
    data["lead_id"] = pd.to_numeric(data["lead_id"], errors="coerce").astype("Int64")

    invalid = data["email"].isna() | data["lead_datetime"].isna()
    rejected = data.loc[invalid].copy()
    events = data.loc[~invalid, EVENT_COLUMNS].copy()
    events = events.sort_values(["email", "lead_datetime", "lead_id"]).reset_index(drop=True)

    # Repeated submissions to the same group are events, but not group changes.
    events["previous_group"] = events.groupby("email")["customer_group"].shift()
    events["is_first_seen"] = events["previous_group"].isna()
    events["is_group_change"] = (
        events["previous_group"].notna()
        & events["previous_group"].ne(events["customer_group"])
    )
    return events, rejected


def make_transitions(events):
    transitions = events.loc[events["is_group_change"]].copy()
    transitions = transitions.rename(
        columns={
            "previous_group": "from_group",
            "customer_group": "to_group",
            "lead_datetime": "transition_datetime",
        }
    )
    columns = [
        "email",
        "transition_datetime",
        "week_start",
        "week_end",
        "week_complete",
        "from_group",
        "to_group",
        "lead_id",
        "form",
    ]
    transitions = transitions[columns].sort_values(["transition_datetime", "email"])
    summary = (
        transitions.groupby(
            ["week_start", "week_end", "week_complete", "from_group", "to_group"],
            as_index=False,
        )
        .agg(people_moved=("email", "nunique"), transition_events=("email", "size"))
        .sort_values(["week_start", "from_group", "to_group"])
    )
    return transitions, summary


def make_weekly_snapshots(events):
    first_week = events["week_start"].min()
    last_week = events["week_start"].max()
    weeks = pd.date_range(first_week, last_week, freq="7D")
    groups = [item[0] for item in FORM_GROUP_MAP.values()]
    groups = list(dict.fromkeys(groups))
    today = pd.Timestamp.now(tz="Australia/Sydney").tz_localize(None).normalize()

    rows = []
    for week_start in weeks:
        week_end = week_start + pd.Timedelta(days=6)
        cutoff = week_end + pd.Timedelta(days=1)
        observed = events.loc[events["lead_datetime"] < cutoff]
        current = observed.groupby("email", as_index=False).tail(1)
        counts = current["customer_group"].value_counts()
        first_seen = events.loc[
            events["is_first_seen"] & events["week_start"].eq(week_start),
            ["email", "customer_group"],
        ]
        new_counts = first_seen["customer_group"].value_counts()

        for group in groups:
            rows.append(
                {
                    "week_start": week_start,
                    "week_end": week_end,
                    "week_complete": week_end < today,
                    "customer_group": group,
                    "active_people": int(counts.get(group, 0)),
                    "new_people": int(new_counts.get(group, 0)),
                }
            )
    return pd.DataFrame(rows)


def find_source():
    candidates = sorted(Path(".").glob(SOURCE_PATTERN), key=lambda p: p.stat().st_mtime)
    if not candidates:
        raise FileNotFoundError(f"No file matching {SOURCE_PATTERN}")
    return candidates[-1]


def main():
    source = find_source()
    raw = pd.read_csv(source, dtype=str, keep_default_na=False)
    events, rejected = clean_events(raw)
    transitions, transition_summary = make_transitions(events)
    snapshots = make_weekly_snapshots(events)

    OUTPUT_DIR.mkdir(exist_ok=True)
    events.to_csv(OUTPUT_DIR / "digital_dealer_events_clean.csv", index=False)
    snapshots.to_csv(OUTPUT_DIR / "digital_dealer_weekly_snapshot.csv", index=False)
    transitions.to_csv(OUTPUT_DIR / "digital_dealer_transitions.csv", index=False)
    transition_summary.to_csv(
        OUTPUT_DIR / "digital_dealer_weekly_transitions.csv", index=False
    )
    if not rejected.empty:
        rejected.to_csv(OUTPUT_DIR / "digital_dealer_rejected.csv", index=False)

    print(f"Source: {source}")
    print(f"Clean events: {len(events):,}")
    print(f"Unique emails: {events['email'].nunique():,}")
    print(f"Group changes: {len(transitions):,}")
    print(f"Rejected rows: {len(rejected):,}")
    print(f"Outputs: {OUTPUT_DIR.resolve()}")


if __name__ == "__main__":
    main()
