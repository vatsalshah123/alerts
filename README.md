# Portfolio P&L Alert — Telegram + Email, Free, Cloud-Scheduled

Runs on GitHub's servers on a schedule — your laptop can be off. Sends
both a Telegram message (aligned table) and an HTML email (styled table,
looks great in Hotmail/Outlook/Gmail) every run.

## 1. Set up Telegram (2 minutes)
1. Telegram → search **@BotFather** → `/newbot` → follow prompts → copy the **token**
2. Search **@userinfobot** → message it → it replies with your **Chat ID**
3. Message your new bot once (e.g. "hi") to activate it

## 2. Set up Email (Brevo — free API, sends to any address incl. Hotmail/Outlook)
1. Sign up free at https://www.brevo.com (no credit card needed)
2. Go to Senders, Domains & Dedicated IPs → Senders → add and verify a sender
   email (an email you own — Brevo sends a verification link to it)
3. Go to Settings → SMTP & API → API Keys → **Generate a new API key** →
   copy it (shown once)
4. Note the API key, your verified sender email, and the Hotmail/Outlook
   address you want the report sent to
   (Free tier: 300 emails/day — more than enough for hourly alerts)

## 3. Put this in a GitHub repo
1. Create a free GitHub account: https://github.com/join
2. Create a new repo (private is fine, free)
3. Upload these files, keeping the folder structure:
   - `portfolio_pnl_alert.py`
   - `portfolio.xlsx`
   - `.github/workflows/pnl-alert.yml`

## 4. Add your holdings
Open `portfolio.xlsx` and fill in your actual stocks, quantities, buy prices —
one row per holding. Keep the header row (`symbol | exchange | qty | buy_price`)
exactly as is. `exchange` must be `NSE` or `BSE`.

## 5. Add secrets (Settings → Secrets and variables → Actions → New repository secret)
| Secret name | Value |
|---|---|
| `BOT_TOKEN` | your Telegram bot token |
| `CHAT_ID` | your Telegram chat ID |
| `BREVO_API_KEY` | your Brevo API key |
| `EMAIL_FROM` | your verified Brevo sender email |
| `EMAIL_TO` | your Hotmail/Outlook address |

## 6. Test it
Actions tab → "Portfolio P&L Alert" → **Run workflow**. You should get a
Telegram message and an email within a minute.

## 7. Runs automatically
Hourly, 9:15 AM–3:15 PM IST, Monday–Friday (edit the `cron:` line in the
workflow file to change — crontab.guru helps, remember GitHub uses UTC).

## Notes
- Both channels are completely free — no message limits, no trial watermarks.
- Yahoo Finance data has a slight delay (~15 min), not tick-by-tick real-time.
- Telegram message uses a monospace block so the columns stay aligned on any phone.
- Email uses a real HTML table with green/red shading per row — renders properly
  in Outlook.com, Hotmail, and Gmail.