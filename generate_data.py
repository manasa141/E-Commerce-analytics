import random
import csv
import sqlite3
from datetime import datetime, timedelta

# ── Seed for reproducibility ──────────────────────────────────────────────────
random.seed(42)

# ── Master data ───────────────────────────────────────────────────────────────
CATEGORIES = {
    "Electronics": ["Samsung", "Apple", "Sony", "LG", "Bose"],
    "Apparel":     ["Nike", "Adidas", "Zara", "H&M", "Levi's"],
    "Home":        ["IKEA", "Dyson", "Philips", "Whirlpool", "Cuisinart"],
    "Beauty":      ["L'Oreal", "Maybelline", "Neutrogena", "MAC", "NYX"],
    "Sports":      ["Nike", "Adidas", "Under Armour", "Puma", "Reebok"],
}

PRICE_RANGES = {
    "Electronics": (30, 1500),
    "Apparel":     (15, 300),
    "Home":        (20, 800),
    "Beauty":      (8, 120),
    "Sports":      (20, 400),
}

# Conversion rates tuned per category (view→cart, cart→purchase)
CONV = {
    "Electronics": (0.28, 0.38),
    "Apparel":     (0.32, 0.45),
    "Home":        (0.22, 0.35),
    "Beauty":      (0.35, 0.55),
    "Sports":      (0.30, 0.42),
}

N_USERS    = 3000
N_SESSIONS = 12000
START_DATE = datetime(2024, 7, 18)
END_DATE   = datetime(2024, 10, 15)

def random_date():
    delta = END_DATE - START_DATE
    return START_DATE + timedelta(seconds=random.randint(0, int(delta.total_seconds())))

def price_band(p):
    if p < 25:   return "<$25"
    if p < 75:   return "$25-$75"
    if p < 200:  return "$75-$200"
    return "$200+"

# ── Generate users ────────────────────────────────────────────────────────────
user_ids   = [f"U{str(i).zfill(5)}" for i in range(1, N_USERS + 1)]
user_types = {uid: random.choice(["new", "returning"]) for uid in user_ids}
# returning users convert slightly better
returning_boost = {uid: (1.15 if user_types[uid] == "returning" else 1.0) for uid in user_ids}

# ── Generate products ─────────────────────────────────────────────────────────
products = []
pid = 1
for cat, brands in CATEGORIES.items():
    lo, hi = PRICE_RANGES[cat]
    for brand in brands:
        for _ in range(12):   # 12 products per brand
            price = round(random.uniform(lo, hi), 2)
            products.append({
                "product_id": f"P{str(pid).zfill(4)}",
                "category":   cat,
                "brand":      brand,
                "price":      price,
                "price_band": price_band(price),
            })
            pid += 1

product_by_cat = {}
for p in products:
    product_by_cat.setdefault(p["category"], []).append(p)

# ── Generate events ───────────────────────────────────────────────────────────
events = []
eid = 1

for s in range(N_SESSIONS):
    uid        = random.choice(user_ids)
    sid        = f"S{str(s+1).zfill(6)}"
    utype      = user_types[uid]
    boost      = returning_boost[uid]
    session_ts = random_date()

    # Pick a category (Electronics weighted higher for volume)
    cat = random.choices(
        list(CATEGORIES.keys()),
        weights=[30, 25, 20, 15, 10],
        k=1
    )[0]

    # Products viewed this session (1–4)
    n_views = random.randint(1, 4)
    viewed  = random.sample(product_by_cat[cat], min(n_views, len(product_by_cat[cat])))

    view_to_cart_rate, cart_to_pur_rate = CONV[cat]
    # Apply slight random noise per session
    v2c = min(view_to_cart_rate * boost * random.uniform(0.7, 1.3), 0.95)
    c2p = min(cart_to_pur_rate  * boost * random.uniform(0.7, 1.3), 0.95)

    for i, prod in enumerate(viewed):
        ts = session_ts + timedelta(minutes=i * random.randint(1, 5))

        # Product view event (always)
        events.append({
            "event_id":   f"E{str(eid).zfill(7)}",
            "user_id":    uid,
            "session_id": sid,
            "event_type": "product_view",
            "product_id": prod["product_id"],
            "category":   prod["category"],
            "brand":      prod["brand"],
            "price":      prod["price"],
            "price_band": prod["price_band"],
            "user_type":  utype,
            "timestamp":  ts.strftime("%Y-%m-%d %H:%M:%S"),
            "revenue":    0.0,
        })
        eid += 1

        # Add-to-cart event
        if random.random() < v2c:
            ts2 = ts + timedelta(minutes=random.randint(1, 3))
            events.append({
                "event_id":   f"E{str(eid).zfill(7)}",
                "user_id":    uid,
                "session_id": sid,
                "event_type": "add_to_cart",
                "product_id": prod["product_id"],
                "category":   prod["category"],
                "brand":      prod["brand"],
                "price":      prod["price"],
                "price_band": prod["price_band"],
                "user_type":  utype,
                "timestamp":  ts2.strftime("%Y-%m-%d %H:%M:%S"),
                "revenue":    0.0,
            })
            eid += 1

            # Purchase event
            if random.random() < c2p:
                ts3 = ts2 + timedelta(minutes=random.randint(2, 10))
                events.append({
                    "event_id":   f"E{str(eid).zfill(7)}",
                    "user_id":    uid,
                    "session_id": sid,
                    "event_type": "purchase",
                    "product_id": prod["product_id"],
                    "category":   prod["category"],
                    "brand":      prod["brand"],
                    "price":      prod["price"],
                    "price_band": prod["price_band"],
                    "user_type":  utype,
                    "timestamp":  ts3.strftime("%Y-%m-%d %H:%M:%S"),
                    "revenue":    prod["price"],
                })
                eid += 1

# ── Write CSV ──────────────────────────────────────────────────────────────────
FIELDS = ["event_id","user_id","session_id","event_type","product_id",
          "category","brand","price","price_band","user_type","timestamp","revenue"]

output_path = "output/events.csv"
import os; os.makedirs("output", exist_ok=True)

with open(output_path, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=FIELDS)
    writer.writeheader()
    writer.writerows(events)

print(f"[OK] Generated {len(events):,} events -> {output_path}")
print(f"     Users: {N_USERS:,}  |  Sessions: {N_SESSIONS:,}  |  Products: {len(products)}")

# ── Quick sanity via SQLite ────────────────────────────────────────────────────
conn = sqlite3.connect(":memory:")
conn.execute("""
    CREATE TABLE events (
        event_id TEXT, user_id TEXT, session_id TEXT, event_type TEXT,
        product_id TEXT, category TEXT, brand TEXT, price REAL,
        price_band TEXT, user_type TEXT, timestamp TEXT, revenue REAL
    )
""")
conn.executemany(
    "INSERT INTO events VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
    [(e["event_id"],e["user_id"],e["session_id"],e["event_type"],
      e["product_id"],e["category"],e["brand"],e["price"],
      e["price_band"],e["user_type"],e["timestamp"],e["revenue"]) for e in events]
)

views = conn.execute("SELECT COUNT(*) FROM events WHERE event_type='product_view'").fetchone()[0]
carts = conn.execute("SELECT COUNT(*) FROM events WHERE event_type='add_to_cart'").fetchone()[0]
purch = conn.execute("SELECT COUNT(*) FROM events WHERE event_type='purchase'").fetchone()[0]
rev   = conn.execute("SELECT SUM(revenue) FROM events WHERE event_type='purchase'").fetchone()[0]

print(f"\nFunnel Sanity Check:")
print(f"   Product Views : {views:,}")
print(f"   Add-to-Cart   : {carts:,}  ({carts/views*100:.1f}% of views)")
print(f"   Purchases     : {purch:,}  ({purch/views*100:.1f}% of views)")
print(f"   Total Revenue : ${rev:,.2f}")
conn.close()
