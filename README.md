# Consulting Client Onboarding Workflow

WeChat message -> contract filled in -> emailed with signing instructions ->
signed copy comes back -> calendar invite sent.

## One-time setup (about 15 minutes)

1. Install Python 3.9+ and the libraries:
   `pip install -r requirements.txt`
2. Copy `config.example.json` to `config.json` and fill in your email address,
   default hourly rate, and your Teams meeting link.
3. Create an app password for your email (Gmail: Google Account -> Security ->
   2-Step Verification -> App passwords) and set it in your terminal:
   - Mac: `export ONBOARD_EMAIL_PASSWORD="xxxx xxxx xxxx xxxx"`
   - Windows: `setx ONBOARD_EMAIL_PASSWORD "xxxx xxxx xxxx xxxx"`
4. Optional, for the AI step: set `ANTHROPIC_API_KEY` the same way.
   Without it, the script still finds the email and asks you to type the name.

## Every new client

```
python onboard.py new --wechat "<paste WeChat message>" --session "2026-10-21 20:00"
```
Add `--rate 150`, `--duration 45`, or `--first/--last/--email` as needed.
Add `--dry-run` first time to check the PDF in `contracts/` without sending.

## When the signed contract comes back

```
python onboard.py watch      # checks inbox, saves signed PDF, sends invite
python onboard.py invite isla@example.com   # or send the invite by hand
python onboard.py list       # see everyone's status
```
To run `watch` automatically every 15 minutes on a Mac: `crontab -e` and add
`*/15 * * * * cd /path/to/consulting-workflow && /usr/bin/python3 onboard.py watch`
