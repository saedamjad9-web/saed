from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "invoices.csv"

app = Flask(__name__)


def load_invoices() -> pd.DataFrame:
    if DATA_FILE.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(DATA_FILE)
    else:
        df = pd.read_csv(DATA_FILE)

    df["date"] = pd.to_datetime(df["date"]).dt.date
    for col in ["sale_usd", "sale_iqd", "cost_usd", "cost_iqd"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["profit_usd"] = df["sale_usd"] - df["cost_usd"]
    df["profit_iqd"] = df["sale_iqd"] - df["cost_iqd"]
    return df


def filter_by_range(df: pd.DataFrame, start, end) -> pd.DataFrame:
    if start:
        df = df[df["date"] >= pd.to_datetime(start).date()]
    if end:
        df = df[df["date"] <= pd.to_datetime(end).date()]
    return df


@app.route("/")
def dashboard():
    df = load_invoices()

    start = request.args.get("start") or ""
    end = request.args.get("end") or ""
    df = filter_by_range(df, start, end)

    daily = (
        df.groupby("date")
        .agg(
            sale_usd=("sale_usd", "sum"),
            sale_iqd=("sale_iqd", "sum"),
            cost_usd=("cost_usd", "sum"),
            cost_iqd=("cost_iqd", "sum"),
            profit_usd=("profit_usd", "sum"),
            profit_iqd=("profit_iqd", "sum"),
            invoices=("invoice_id", "count"),
        )
        .reset_index()
        .sort_values("date")
    )

    totals = {
        "sale_usd": float(df["sale_usd"].sum()),
        "sale_iqd": float(df["sale_iqd"].sum()),
        "cost_usd": float(df["cost_usd"].sum()),
        "cost_iqd": float(df["cost_iqd"].sum()),
        "profit_usd": float(df["profit_usd"].sum()),
        "profit_iqd": float(df["profit_iqd"].sum()),
        "invoices": int(len(df)),
    }

    chart = {
        "labels": [d.isoformat() for d in daily["date"]],
        "sale_usd": [float(v) for v in daily["sale_usd"]],
        "cost_usd": [float(v) for v in daily["cost_usd"]],
        "profit_usd": [float(v) for v in daily["profit_usd"]],
        "sale_iqd": [float(v) for v in daily["sale_iqd"]],
        "cost_iqd": [float(v) for v in daily["cost_iqd"]],
        "profit_iqd": [float(v) for v in daily["profit_iqd"]],
    }

    invoices = df.sort_values(["date", "invoice_id"], ascending=[False, False]).to_dict(orient="records")
    daily_rows = daily.sort_values("date", ascending=False).to_dict(orient="records")

    return render_template(
        "dashboard.html",
        totals=totals,
        invoices=invoices,
        daily_rows=daily_rows,
        chart=chart,
        start=start,
        end=end,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
