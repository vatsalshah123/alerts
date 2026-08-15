"""
Portfolio P&L Alert — Telegram + Email (both free)
-----------------------------------------------------
Reads holdings from portfolio.xlsx, fetches live NSE/BSE prices,
computes P&L, and sends the summary to BOTH Telegram (nicely aligned
table via monospace) and Email (real HTML table via Brevo, works great
in Hotmail/Outlook/Gmail).

SETUP:
1. pip install yfinance requests openpyxl
2. Telegram: create a bot via @BotFather, get BOT_TOKEN + CHAT_ID (see README)
3. Email: create a free Brevo account (brevo.com), verify a sender email,
   get an API key. No App Passwords / 2FA hassle needed. (see README)
4. Fill in credentials below (or set as env vars / GitHub Secrets)
5. Test: python portfolio_pnl_alert.py
6. Schedule hourly (see README)

portfolio.xlsx format (first sheet, header row required):
| symbol   | exchange | qty | buy_price |
|----------|----------|-----|-----------|
| RELIANCE | NSE      | 10  | 2450.50   |
| TCS      | BSE      | 5   | 3600.00   |
"""

import os
import sys
from datetime import datetime

import requests
import yfinance as yf
from openpyxl import load_workbook
import json

# ---------------- CONFIG ----------------
# All read from environment variables (set as GitHub Secrets in Actions).
# Falls back to placeholder strings for local editing/testing.
# Local secrets helper: if a secrets.local.json file exists in the repo root,
# load its keys into environment variables (unless already set). This lets
# developers keep a local ignored file with credentials for testing.
LOCAL_SECRETS_FILE = "secrets.local.json"
if os.path.exists(LOCAL_SECRETS_FILE):
    try:
        with open(LOCAL_SECRETS_FILE, "r", encoding="utf-8") as _f:
            _local = json.load(_f)
        for _k, _v in _local.items():
            if _v is None:
                continue
            # do not overwrite already-set environment variables
            if not os.environ.get(_k):
                os.environ[_k] = str(_v)
    except Exception as _e:
        print(f"Warning: failed to load {LOCAL_SECRETS_FILE}: {_e}", file=sys.stderr)

PORTFOLIO_FILE = "portfolio.xlsx"
# ------------------------------------------

EXCHANGE_SUFFIX = {"NSE": ".NS", "BSE": ".BO"}


def load_portfolio(path):
    """Reads holdings from the first sheet of an .xlsx file.
    Expects header row: symbol | exchange | qty | buy_price"""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip().lower() if c else "" for c in rows[0]]
    col = {name: idx for idx, name in enumerate(header)}

    required = {"symbol", "exchange", "qty", "buy_price"}
    missing = required - set(col)
    if missing:
        raise ValueError(f"portfolio.xlsx is missing column(s): {missing}")

    holdings = []
    for r in rows[1:]:
        if not r or r[col["symbol"]] in (None, ""):
            continue  # skip blank rows
        holdings.append({
            "symbol": str(r[col["symbol"]]).strip(),
            "exchange": str(r[col["exchange"]]).strip(),
            "qty": float(r[col["qty"]]),
            "buy_price": float(r[col["buy_price"]]),
        })
    return merge_duplicate_holdings(holdings)


def merge_duplicate_holdings(holdings):
    """Combines multiple rows for the same stock (symbol + exchange) into
    one, using a quantity-weighted average buy price."""
    merged = {}
    order = []  # preserves first-seen order for consistent report ordering

    for h in holdings:
        key = (h["symbol"].upper(), h["exchange"].upper())
        if key not in merged:
            merged[key] = {"symbol": h["symbol"], "exchange": h["exchange"], "qty": 0.0, "cost": 0.0}
            order.append(key)
        merged[key]["qty"] += h["qty"]
        merged[key]["cost"] += h["qty"] * h["buy_price"]

    result = []
    for key in order:
        m = merged[key]
        avg_buy_price = m["cost"] / m["qty"] if m["qty"] else 0.0
        result.append({
            "symbol": m["symbol"],
            "exchange": m["exchange"],
            "qty": m["qty"],
            "buy_price": round(avg_buy_price, 4),
        })
    return result


def fetch_price(symbol, exchange):
    ticker = symbol.strip().upper() + EXCHANGE_SUFFIX[exchange.upper()]
    t = yf.Ticker(ticker)

    # fast_info reflects the latest traded price during market hours,
    # more current than the daily history bar.
    try:
        price = t.fast_info.get("last_price")
        if price:
            return round(float(price), 2)
    except Exception:
        pass

    # Fallback: most recent 1-minute bar (still current during market hours)
    data = t.history(period="1d", interval="1m")
    if data.empty:
        # Market closed (weekend/holiday) - fall back to last daily close
        data = t.history(period="5d")
        if data.empty:
            raise ValueError(f"No price data for {ticker}")
    return round(float(data["Close"].iloc[-1]), 2)


def compute_rows(holdings):
    """Fetch prices and compute P&L for each holding. Returns rows + totals."""
    rows = []
    total_invested = 0.0
    total_current = 0.0

    for h in holdings:
        try:
            ltp = fetch_price(h["symbol"], h["exchange"])
        except Exception as e:
            rows.append({"symbol": h["symbol"], "error": str(e)})
            continue

        invested = h["qty"] * h["buy_price"]
        current = h["qty"] * ltp
        pnl = current - invested
        pnl_pct = (pnl / invested) * 100 if invested else 0

        total_invested += invested
        total_current += current

        rows.append({
            "symbol": h["symbol"],
            "exchange": h["exchange"],
            "qty": h["qty"],
            "buy_price": h["buy_price"],
            "ltp": ltp,
            "invested": invested,
            "current": current,
            "pnl": pnl,
            "pnl_pct": pnl_pct,
        })

    total_pnl = total_current - total_invested
    total_pct = (total_pnl / total_invested) * 100 if total_invested else 0
    totals = {
        "invested": total_invested,
        "current": total_current,
        "pnl": total_pnl,
        "pnl_pct": total_pct,
    }
    return rows, totals


def build_telegram_message(rows, totals):
    timestamp = datetime.now().strftime("%d %b, %I:%M %p")

    lines = [f"📊 *Portfolio P&L*", f"_{timestamp}_", ""]

    for r in rows:
        if "error" in r:
            lines.append(f"⚠️ *{r['symbol']}* — price fetch failed")
            lines.append("")
            continue
        arrow = "🟢" if r["pnl"] >= 0 else "🔴"
        lines.append(f"{arrow} *{r['symbol']}* ({r['exchange']})")
        lines.append(f"    Qty {r['qty']:g}  •  Buy ₹{r['buy_price']:,.2f}  •  LTP ₹{r['ltp']:,.2f}")
        lines.append(f"    Total Investment ₹{r['invested']:,.2f}  •  Market Value ₹{r['current']:,.2f}")
        lines.append(f"    P&L ₹{r['pnl']:+,.2f} ({r['pnl_pct']:+.2f}%)")
        lines.append("")

    total_arrow = "🟢" if totals["pnl"] >= 0 else "🔴"
    lines.append("──────────────")
    lines.append(f"{total_arrow} *Total P&L: ₹{totals['pnl']:+,.2f} ({totals['pnl_pct']:+.2f}%)*")
    lines.append(f"Total Investment: ₹{totals['invested']:,.2f}")
    lines.append(f"Market Value: ₹{totals['current']:,.2f}")

    return "\n".join(lines)


def build_email_html(rows, totals):
    timestamp = datetime.now().strftime("%d %b %Y, %I:%M %p")

    row_html = ""
    for r in rows:
        if "error" in r:
            row_html += (
                f"<tr><td style='padding:8px;border-bottom:1px solid #eee;'>{r['symbol']}</td>"
                f"<td colspan='8' style='padding:8px;border-bottom:1px solid #eee;color:#999;'>"
                f"price fetch failed</td></tr>"
            )
            continue
        color = "#1a7f37" if r["pnl"] >= 0 else "#cf222e"
        bg = "#f0fdf4" if r["pnl"] >= 0 else "#fef2f2"
        row_html += f"""
        <tr style="background:{bg};">
          <td style="padding:8px;border-bottom:1px solid #eee;">{r['symbol']} <span style="color:#888;font-size:12px;">({r['exchange']})</span></td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:right;">{r['qty']:g}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:right;">₹{r['buy_price']:,.2f}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:right;">₹{r['ltp']:,.2f}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:right;">₹{r['invested']:,.2f}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:right;">₹{r['current']:,.2f}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:right;color:{color};font-weight:600;">₹{r['pnl']:+,.2f}</td>
          <td style="padding:8px;border-bottom:1px solid #eee;text-align:right;color:{color};font-weight:600;">{r['pnl_pct']:+.2f}%</td>
        </tr>"""

    total_color = "#1a7f37" if totals["pnl"] >= 0 else "#cf222e"

    html = f"""
    <html><body style="font-family:Segoe UI,Arial,sans-serif;background:#f6f6f6;padding:20px;">
      <div style="max-width:820px;margin:auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.1);">
        <div style="background:#1a1a2e;color:#fff;padding:16px 20px;">
          <h2 style="margin:0;font-size:18px;">📊 Portfolio P&amp;L</h2>
          <div style="font-size:12px;color:#aaa;margin-top:4px;">{timestamp}</div>
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <thead>
            <tr style="background:#f0f0f5;text-align:right;">
              <th style="padding:8px;text-align:left;">Stock</th>
              <th style="padding:8px;text-align:right;">Qty</th>
              <th style="padding:8px;text-align:right;">Buy Price</th>
              <th style="padding:8px;text-align:right;">LTP</th>
              <th style="padding:8px;text-align:right;">Total Investment</th>
              <th style="padding:8px;text-align:right;">Market Value</th>
              <th style="padding:8px;text-align:right;">P&amp;L</th>
              <th style="padding:8px;text-align:right;">%</th>
            </tr>
          </thead>
          <tbody>
            {row_html}
            <tr style="background:#f0f0f5;font-weight:700;">
              <td style="padding:10px 8px;">TOTAL</td>
              <td></td>
              <td></td>
              <td></td>
              <td style="padding:10px 8px;text-align:right;">₹{totals['invested']:,.2f}</td>
              <td style="padding:10px 8px;text-align:right;">₹{totals['current']:,.2f}</td>
              <td style="padding:10px 8px;text-align:right;color:{total_color};">₹{totals['pnl']:+,.2f}</td>
              <td style="padding:10px 8px;text-align:right;color:{total_color};">{totals['pnl_pct']:+.2f}%</td>
            </tr>
          </tbody>
        </table>
      </div>
    </body></html>
    """
    return html


def send_telegram(text):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    resp = requests.post(url, data={"chat_id": CHAT_ID, "text": text, "parse_mode": "Markdown"})
    resp.raise_for_status()


def send_email(html_body):
    subject = f"Portfolio P&L — {datetime.now().strftime('%d %b, %I:%M %p')}"
    url = "https://api.brevo.com/v3/smtp/email"
    headers = {
        "api-key": BREVO_API_KEY,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    payload = {
        "sender": {"email": EMAIL_FROM},
        "to": [{"email": EMAIL_TO}],
        "subject": subject,
        "htmlContent": html_body,
    }
    resp = requests.post(url, json=payload, headers=headers)
    resp.raise_for_status()


def main():
    holdings = load_portfolio(PORTFOLIO_FILE)
    rows, totals = compute_rows(holdings)

    telegram_text = build_telegram_message(rows, totals)
    email_html = build_email_html(rows, totals)

    print(telegram_text)  # for local testing visibility

    try:
        send_telegram(telegram_text)
        print("✅ Telegram sent.")
    except Exception as e:
        print(f"⚠️ Telegram failed: {e}", file=sys.stderr)

    try:
        send_email(email_html)
        print("✅ Email sent.")
    except Exception as e:
        print(f"⚠️ Email failed: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()