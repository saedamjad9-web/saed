from pathlib import Path

from flask import Flask, abort, render_template

from invoice_parser import parse_all, parse_invoice

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

app = Flask(__name__)


@app.route("/")
def dashboard():
    invoices = parse_all(DATA_DIR)

    totals = {
        "usd_value": sum(inv["total"]["usd"]["total_value"] for inv in invoices),
        "usd_cost": sum(inv["total"]["usd"]["total_cost"] for inv in invoices),
        "usd_margin": sum(inv["total"]["usd"]["margin"] for inv in invoices),
        "iqd_value": sum(inv["total"]["iqd"]["total_value"] for inv in invoices),
        "iqd_cost": sum(inv["total"]["iqd"]["total_cost"] for inv in invoices),
        "iqd_margin": sum(inv["total"]["iqd"]["margin"] for inv in invoices),
        "count": len(invoices),
    }
    totals["usd_margin_pct"] = (
        totals["usd_margin"] / totals["usd_value"] if totals["usd_value"] else 0
    )

    daily: dict = {}
    for inv in invoices:
        key = str(inv["order_date"])
        d = daily.setdefault(
            key,
            {
                "date": key,
                "invoices": 0,
                "usd_value": 0.0,
                "usd_cost": 0.0,
                "usd_margin": 0.0,
                "iqd_value": 0.0,
                "iqd_cost": 0.0,
                "iqd_margin": 0.0,
            },
        )
        d["invoices"] += 1
        d["usd_value"] += inv["total"]["usd"]["total_value"]
        d["usd_cost"] += inv["total"]["usd"]["total_cost"]
        d["usd_margin"] += inv["total"]["usd"]["margin"]
        d["iqd_value"] += inv["total"]["iqd"]["total_value"]
        d["iqd_cost"] += inv["total"]["iqd"]["total_cost"]
        d["iqd_margin"] += inv["total"]["iqd"]["margin"]

    daily_rows = sorted(daily.values(), key=lambda r: r["date"], reverse=True)
    chart_rows = sorted(daily.values(), key=lambda r: r["date"])
    chart = {
        "labels": [r["date"] for r in chart_rows],
        "usd_value": [r["usd_value"] for r in chart_rows],
        "usd_cost": [r["usd_cost"] for r in chart_rows],
        "usd_margin": [r["usd_margin"] for r in chart_rows],
        "iqd_value": [r["iqd_value"] for r in chart_rows],
        "iqd_cost": [r["iqd_cost"] for r in chart_rows],
        "iqd_margin": [r["iqd_margin"] for r in chart_rows],
    }

    return render_template(
        "dashboard.html",
        invoices=invoices,
        totals=totals,
        daily_rows=daily_rows,
        chart=chart,
    )


@app.route("/invoice/<file_name>")
def invoice_detail(file_name: str):
    path = (DATA_DIR / file_name).resolve()
    if not path.is_file() or path.parent != DATA_DIR.resolve():
        abort(404)
    invoice = parse_invoice(path)
    return render_template("invoice.html", inv=invoice)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
