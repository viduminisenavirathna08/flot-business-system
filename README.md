# 24K Shine FLOT - Business Management System

A web app to manage orders, call-center leads, stock, parcel tracking and the daily / monthly budget of **24K Shine FLOT** (jewelry e-commerce, Kegalle, Sri Lanka).

Built with **Flask**, **MySQL / MariaDB** (PyMySQL), **Excel** files (pandas / openpyxl) and a **Tailwind CSS** front end.

---

## Features

### Orders and call center
- Add orders manually or upload an Excel file (with or without a header row). Duplicates (same phone + same items) are skipped and shown in a list.
- One order can have **many items** (item code, quantity, price). Item name and price are filled from `itemcost.xlsx`; the price can be changed.
- Edit any order (customer, items, delivery method).
- Call status: New Lead, Confirmed, Cancelled, Callback, No Answer, Postponed (with a postpone date).
- Delivery method dropdown: **Post Office / Speed Post / Courier**.
- Row colors (light): Confirmed = green, Cancelled = red, Postponed = blue, others = yellow.
- Date range filter with counts, print, and Excel download.
- Every change is also written to `orders_log.xlsx`.

### Items and stock
- Items and stock levels in `items.xlsx` (Items and Availability tab).
- Stock is reduced when an order is Confirmed.

### Parcel tracking and returns (Track Orders tab)
- Links to SL Post and Domex tracking pages.
- **Barcode pool**: add Post Office / Courier barcodes (single codes or ranges such as `FL000001-FL000100`). A Confirmed parcel automatically gets the next available barcode of its delivery type. Barcodes can be edited.
- Parcel status: In Transit, Delivered, Returned. Use a **barcode reader** (scan + Enter) or paste many barcodes to mark parcels Delivered / Returned.
- "Scheduled on Hold" counter for postponed orders.

### Budget
- Budget per parcel: item cost, packing cost, delivery cost, extra cost, ad share, profit (all editable per parcel).
- Item cost and packing cost per **item code and quantity** from `itemcost.xlsx` (edit in Cost Settings, or upload Excel).
- Delivery cost per method (Post Office, Speed Post, Courier) and default costs.
- Daily ad cost, divided between that day's parcels.
- Monthly expenses and savings (rent etc.), spread over the days of the month.
- **Returned parcels**: no income; packing, delivery, ad and extra cost count as a loss.
- Daily profit, monthly summary, and reminders in Notifications.

### Notifications
- Low stock (below minimum stock), budget not updated for several days, postponed orders coming up (3 days before).

---

## Project structure

```
Project/
|-- project.py            Flask app (API + pages)
|-- templates/            order.html, login.html, customer.html, product.html
|-- static/               logo.jpg
|-- schema.sql            Database + orders table
|-- budget_tables.sql     Budget tables
|-- requirements.txt
|-- .env                  Your secrets (NOT in git)
|-- items.xlsx            Items and stock
|-- itemcost.xlsx         Item cost / packing cost / selling price per code + quantity
`-- orders_log.xlsx       Auto-generated copy of all orders
```

`itemcost.xlsx` columns: `Code, Name, Category, Quantity, Selling price, Item cost, Packing cost` (one row per code and quantity).

---

## Setup (Windows, XAMPP)

1. **Install** Python 3 and XAMPP. In XAMPP Control Panel start **MySQL**.
2. **Install packages**
   ```
   pip install -r requirements.txt
   ```
3. **Create the database**: open `http://localhost/phpmyadmin`, go to **Import**, import `schema.sql`, then import `budget_tables.sql`.
   (`order_items` and `barcode_pool` tables and the parcel status columns are created automatically when the server starts.)
4. **Create `.env`** in the project folder:
   ```
   SECRET_KEY=change_me_to_a_long_random_text
   ADMIN_USERNAME=admin
   ADMIN_PASSWORD=change_me
   DB_HOST=localhost
   DB_USER=root
   DB_PASSWORD=
   DB_NAME=business_management_system_db
   ```
   Put your real MySQL root password in `DB_PASSWORD` (XAMPP's default is empty). `.env` must never be pushed to GitHub.
5. **Run**
   ```
   cd Project
   python project.py
   ```
6. Open `http://127.0.0.1:5000/login` and sign in with the admin user from `.env`.

Every time you start the app, start **MySQL** in XAMPP first.

### Common errors
| Error | Meaning / fix |
|-------|---------------|
| `1045 Access denied` | Wrong `DB_PASSWORD` in `.env`. Restart the server after changing it. |
| `2003 / 10061 Can't connect` | MySQL is not running. Start it in XAMPP. |
| `1049 Unknown database` / tables missing | Import `schema.sql` and `budget_tables.sql`. |
| `orders_log.xlsx` not updating | The file is open in Excel. Close it. |
| `itemcost.xlsx is open in Excel` | Close the file and save again. |

---

## How the budget is calculated

- Parcel = a **Confirmed** order. Day = order date. Sale = order amount (sum of item prices).
- Item cost = from `itemcost.xlsx` for the code and quantity (if a quantity row is missing, the quantity 1 cost x quantity is used). Packing = highest packing cost of the items in the parcel, or the default.
- Delivery cost = by delivery method, or the default.
- Ad share = the day's ad cost / number of parcels that day.
- Profit = sale - item - packing - delivery - extra - ad share.
- Monthly expenses and savings are divided by the days of the month for the daily net profit.
- Returned parcel: sale = 0 and item cost = 0; other costs stay as a loss.
- Costs saved on a parcel (edited in the Daily Profit table) override the defaults.

---

## Notes

- Excel files that hold customer data (order exports, courier / post office sheets, `orders_log.xlsx`) are listed in `.gitignore` and are not uploaded to GitHub.
- Use a **private** repository for this project.

## Update GitHub

```
git add .
git commit -m "describe the change"
git push
```
