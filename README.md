# E-Commerce Product Analytics & Conversion Optimization

A complete end-to-end analytics project for e-commerce funnel analysis, KPI segmentation, and conversion optimization.

## Project Structure

```
ecommerce-analytics/
├── generate_data.py          # Generates synthetic event dataset (~28k events)
├── analysis.py               # Runs SQL analysis → builds Excel dashboard
├── sql/
│   └── queries.sql           # 10 documented reference SQL queries
├── output/
│   ├── events.csv            # Generated dataset
│   └── ecommerce_dashboard.xlsx   # Final Excel dashboard
└── README.md
```

## Quick Start

```bash
# 1. Install dependency
pip install openpyxl

# 2. Generate synthetic dataset
python generate_data.py

# 3. Build Excel dashboard
python analysis.py
```

## Dataset Schema

| Column | Type | Description |
|---|---|---|
| event_id | string | Unique event ID |
| user_id | string | Anonymized user ID |
| session_id | string | Session identifier |
| event_type | string | `product_view`, `add_to_cart`, `purchase` |
| product_id | string | Product identifier |
| category | string | Electronics, Apparel, Home, Beauty, Sports |
| brand | string | 15 synthetic brands |
| price | float | Product price ($) |
| price_band | string | `<$25`, `$25-$75`, `$75-$200`, `$200+` |
| user_type | string | `new` / `returning` |
| timestamp | datetime | Event timestamp |
| revenue | float | Revenue (purchase events only) |

## Excel Dashboard Sheets

| Sheet | Contents |
|---|---|
| 📊 Overview KPIs | 10 KPI metric cards + funnel summary table |
| 🔽 Funnel by Category | Full funnel metrics per category + bar chart |
| 🗂️ Segmentation | Brand, price band, user type breakdown |
| 📉 Drop-off Analysis | View→Cart and Cart→Purchase drop-off ranked tables |
| 🧪 A/B Test Framework | Sample size calculator + 10 prioritized hypotheses |
| 💡 Recommendations | 10 prioritized action items with impact estimates |
| 📋 Raw Data | First 5,000 events with auto-filter |
| 📈 Daily Trend | Day-by-day funnel metrics + chart |

## Key Findings (Sample Dataset)

- **Beauty** has the highest overall CVR (~18%)
- **Electronics** has the highest cart abandonment (high-ticket hesitation)
- **Returning users** convert ~15% better than new users
- **$75–$200 price band** drives the most revenue volume
- Top A/B test priority: BNPL options for Electronics checkout
