# 🏪 NeelKanth IMS
### Inventory Management System

A lightweight, full-featured **Inventory Management System** built with Flask and SQLite. Track stock levels, get low-stock alerts, analyse inventory value by category, and export reports — all from a clean web UI.

🔗 **Live Demo:** [neelkanth-ims.onrender.com](https://neelkanth-ims.onrender.com)

---

## ✨ Features

| Feature | Description |
|---|---|
| 📦 **Full CRUD** | Add, edit, and delete inventory items |
| 📊 **Dashboard Stats** | Total items, low-stock count, out-of-stock count, total inventory value |
| 🔍 **Search & Filter** | Filter by name, supplier, notes, or category; sort by any column |
| 📉 **Low Stock Alerts** | Browser notifications when items fall below threshold |
| 📈 **Analytics** | Category breakdown, top-value items, stock health, 14-day activity trend |
| 📤 **CSV Export** | Download full inventory as a formatted CSV |
| 📊 **Excel Export** | Multi-sheet XLSX report — Inventory, By Category, Low Stock |
| 🪵 **Activity Log** | Tracks every add / update / delete with timestamps |
| 🚀 **Boot Stream** | SSE-powered animated startup sequence |

---

## 🗂️ Project Structure

```
NeelKanth-IMS/
├── app.py               # Flask application (routes, DB logic, exports)
├── templates/
│   └── index.html       # Single-page frontend (HTML + JS)
├── requirements.txt     # Python dependencies
├── render.yaml          # Render deployment config
├── Procfile             # Gunicorn start command
├── runtime.txt          # Python version pin
└── .gitignore
```

---

## 🛠️ Tech Stack

- **Backend:** Python 3.11, Flask 3.x
- **Database:** SQLite (via Python's built-in `sqlite3`)
- **Frontend:** Vanilla HTML / CSS / JavaScript (no frameworks)
- **Export:** `openpyxl` for Excel, `csv` stdlib for CSV
- **Server:** Gunicorn (production), Flask dev server (local)
- **Streaming:** Server-Sent Events (SSE) for boot sequence

---

## 🚀 Local Setup

### 1. Clone the repo

```bash
git clone https://github.com/TRINEXOR/NeelKanth-IMS.git
cd NeelKanth-IMS
```

### 2. Create a virtual environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the app

```bash
python app.py
```

Open your browser at **http://127.0.0.1:5000**

> The database (`inventory.db`) is created automatically on first run and seeded with 10 sample items.

---

## 🌐 Deployment

### Render (current)

The repo includes a `render.yaml` — just connect it to Render and it deploys automatically.

```yaml
startCommand: gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
```

> ⚠️ **Note:** Render's free tier uses an ephemeral filesystem. The SQLite database resets on each redeploy or restart. For persistent data, consider adding a Render Persistent Disk or migrating to PostgreSQL.

### PythonAnywhere (recommended for persistence)

1. Upload files via the Files tab or `git clone` in the Bash console
2. Create a virtual environment and `pip install -r requirements.txt`
3. Configure a new Web App pointing to `app.py`
4. SQLite data **persists** between restarts on PythonAnywhere's free tier ✅

### Railway

1. Push to GitHub → connect repo on [railway.app](https://railway.app)
2. Set start command: `gunicorn app:app --bind 0.0.0.0:$PORT`
3. Deploy — Railway provides $5 free credits monthly

---

## 📡 API Reference

| Method | Endpoint | Description |
|---|---|---|
| `GET` | `/api/items` | List all items (supports `search`, `category`, `low_only`, `sort`, `order`) |
| `POST` | `/api/items` | Create a new item |
| `PUT` | `/api/items/<id>` | Update an item |
| `DELETE` | `/api/items/<id>` | Delete an item |
| `GET` | `/api/stats` | Dashboard summary stats |
| `GET` | `/api/categories` | List all categories |
| `GET` | `/api/analytics` | Full analytics payload |
| `GET` | `/api/log` | Last 50 activity log entries |
| `GET` | `/api/export/csv` | Download CSV report |
| `GET` | `/api/export/excel` | Download multi-sheet XLSX report |
| `GET` | `/api/alerts/low-stock` | Low-stock items for notifications |
| `POST` | `/api/alerts/config` | Enable / disable browser notifications |
| `POST` | `/api/alerts/ack` | Acknowledge notification |
| `GET` | `/api/boot-stream` | SSE boot sequence stream |

---

## 🗃️ Database Schema

### `items`
| Column | Type | Description |
|---|---|---|
| `id` | INTEGER PK | Auto-increment |
| `name` | TEXT | Item name |
| `category` | TEXT | Category (default: `General`) |
| `quantity` | INTEGER | Current stock quantity |
| `unit` | TEXT | Unit of measure (kg, pcs, L, …) |
| `low_stock` | INTEGER | Low-stock threshold |
| `price` | REAL | Unit price (₹) |
| `supplier` | TEXT | Supplier name (optional) |
| `notes` | TEXT | Free-text notes (optional) |
| `updated_at` | TEXT | ISO 8601 timestamp of last update |

### `activity_log`
Tracks every item add, update, and delete with `item_id`, `item_name`, `action`, `details`, and `timestamp`.

### `alert_config`
Single-row config for browser notification state (`enabled`, `last_notified`).

---

## 🌱 Seed Data

On first run, if the database is empty, 10 sample items are automatically inserted across categories: Grains, Oils, Canned Goods, Legumes, Condiments, Baby, and Hygiene.

---

## 📄 License

This project is open source. Feel free to fork and adapt for your own use.

---

<p align="center">Built with ❤️ by <a href="https://github.com/TRINEXOR">TRINEXOR</a></p>
