# CRM Journey Dashboard Data Template

This dashboard needs three CSV exports. Start with weekly exports if daily is hard, but daily snapshots will make audience movement much clearer.

## 1. Campaign Performance CSV

One row per sent campaign, customer type, journey stage, and channel.

Required columns:

| Column | Example | Notes |
| --- | --- | --- |
| `send_date` | `2026-05-28` | Date the campaign/SMS was sent. |
| `customer_type` | `Reservation` | One of the three fixed customer types. |
| `journey_stage` | `What to Expect #2` | The automation step/stage this send belongs to. |
| `campaign_name` | `EDM #2: What to Expect (Reservation)` | Mailchimp campaign name. |
| `channel` | `Email` | Use `Email` or `SMS`. |
| `recipients` | `867` | Number targeted/sent to. |
| `delivered` | `842` | Delivered count. For email, sent minus bounces is OK. |
| `opens_unique` | `274` | Unique opens. For SMS, use `0` unless you track opens. |
| `clicks_unique` | `75` | Unique clicks. |
| `unsubscribes` | `4` | Unsubscribes caused by this send. |
| `bounces` | `25` | Bounced/failed delivery count. |
| `revenue` | `0` | Optional revenue, use `0` if not tracked. |

Useful source: Mailchimp campaign report export, plus manual mapping to customer type and journey stage if the campaign name does not already encode it.

## 2. Audience Daily Snapshot CSV

One row per date and customer type.

Required columns:

| Column | Example | Notes |
| --- | --- | --- |
| `date` | `2026-05-28` | Snapshot date. |
| `customer_type` | `Build` | One of the three customer types. |
| `active_audience` | `752` | Current subscribed/marketable contacts in this type. |
| `new_imports` | `18` | Contacts newly added/imported on this date. |
| `entered_journey` | `18` | New contacts who entered the active journey stage. |
| `unsubscribes` | `3` | Contacts who unsubscribed on this date. |

This is the dataset that answers: "How many people of this audience type did we have this week, and how is that changing?"

## 3. Journey Stage Daily CSV

One row per date, customer type, and journey stage.

Required columns:

| Column | Example | Notes |
| --- | --- | --- |
| `date` | `2026-05-28` | Snapshot date. |
| `customer_type` | `EOI` | One of the three customer types. |
| `journey_stage` | `Pre-Production Drive` | Fixed automation stage. |
| `contacts` | `164` | Contacts currently assigned to that stage. |

This is the dataset that shows whether each type of customer is moving through the fixed automation flow as expected.

## Automation Tracking Rule

Newly imported customers should enter from the current active journey stage onward only. Do not backfill them into emails already sent earlier in the flow.

Recommended setup:

1. Keep the welcome/introduction email as a separate automation triggered immediately when a new contact is added.
2. Store each contact's `customer_type`, `journey_entry_date`, and `journey_entry_stage`.
3. Store a weekly or daily `active_stage` for each customer type, for example `Reservation -> What to Expect #2`.
4. When importing new contacts, set their `journey_entry_stage` to the current `active_stage`.
5. Use exclusion rules or segment logic so a contact only receives campaigns where the stage order is greater than or equal to their `journey_entry_stage`.

## Minimum Manual Collection Workflow

1. Export campaign reports from Mailchimp after each send.
2. Add or confirm `customer_type` and `journey_stage` for each campaign.
3. Export audience/contact counts by customer type at least once per week.
4. Log imports and unsubscribes by date and customer type.
5. Keep a simple journey-stage mapping sheet so the dashboard knows which contacts are in each fixed flow stage.
