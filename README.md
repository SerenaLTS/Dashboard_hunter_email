# Hunter CRM Campaign Dashboard

Streamlit dashboard for weekly Hunter CRM Email and SMS campaign reporting.

## Run locally

```bash
python3 -m streamlit run dashboard.py
```

The app reads `database.xlsx` from the project root.

## Digital Dealer audience processing

Run:

```bash
python3 process_digital_dealer.py
```

The script reads the newest `digital_dealer_enquiries_*.csv` file and writes:

- `processed/digital_dealer_events_clean.csv`: normalized enquiry events, keyed by
  lowercase email.
- `processed/digital_dealer_weekly_snapshot.csv`: each group's audience at the end
  of every Monday–Sunday week.
- `processed/digital_dealer_transitions.csv`: person-level group changes.
- `processed/digital_dealer_weekly_transitions.csv`: weekly group-to-group counts.

Form mapping:

| Digital Dealer Form | Customer group |
| --- | --- |
| `hunterUpdates` | `JAC Hunter Build` |
| `Register Your Interest JAC Hunter` | `JAC Hunter` |
| `Reserve Your Build` | `JAC Hunter Reserve` |

The current partial week is retained and marked with `week_complete = False`.
