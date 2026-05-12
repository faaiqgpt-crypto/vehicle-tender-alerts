import requests
import json
import pandas as pd
import time
import os
import io
import smtplib
from datetime import datetime, timezone, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ============================================================
#  SETTINGS
# ============================================================
FIRST_RUN = False

HISTORY_FILE      = "tender_history_master.xlsx"
MASTER_90DAY_FILE = "vehicle_tenders_90day_master.xlsx"
os.makedirs("tender_documents", exist_ok=True)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL   = "claude-sonnet-4-20250514"

EMAIL_FROM     = "faaiqgpt@gmail.com"
EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
EMAIL_TO       = ["faaiq@halfwaygo.co.za", "faaiqdavids@gmail.com"]

WHATSAPP_ENABLED = True
WHATSAPP_API_KEY = "2220740"
WHATSAPP_NUMBER  = "27836888820"

KEYWORDS = [
    "vehicle","fleet","truck","car","bakkie","bus","transport",
    "logistics","leasing","rental","maintenance","repair",
    "servicing","telematics","fuel"
]
EXCLUDE_KEYWORDS = ["carport"]

OCP_BASE = "https://data.open-contracting.org/en/publication/143/download?name="

# ============================================================
#  DATES — always UTC-aware so they compare with the API data
# ============================================================
today     = datetime.now(timezone.utc)
date_from = today - timedelta(days=90)
date_to_str   = today.strftime("%Y-%m-%d")
date_from_str = date_from.strftime("%Y-%m-%d")

# ============================================================
#  LOAD HISTORY
# ============================================================
hist_cols = ['ocid','tender_title','buyer_name','province','category',
             'status','status_label','date_published','expiry_date',
             'vehicle_score','description','ai_summary','date_found']
if not FIRST_RUN and os.path.exists(HISTORY_FILE):
    try:
        master_df = pd.read_excel(HISTORY_FILE, engine='openpyxl')
        for c in hist_cols:
            if c not in master_df.columns: master_df[c] = "N/A"
        master_df['date_found'] = pd.to_datetime(master_df['date_found'], errors='coerce')
        seen_ids = set(master_df['ocid'].astype(str))
        print(f"History: {len(seen_ids)} previously seen IDs")
    except Exception as e:
        print(f"History load failed ({e}) — fresh start")
        master_df, seen_ids = pd.DataFrame(columns=hist_cols), set()
else:
    print("FIRST_RUN — treating all 90-day tenders as new")
    master_df, seen_ids = pd.DataFrame(columns=hist_cols), set()

# ============================================================
#  DOWNLOAD
# ============================================================
years_needed = sorted(set([date_from.year, today.year]))
print(f"\nFetching tenders from {date_from_str} to {date_to_str}")
print(f"Years to download: {years_needed}\n")

def download_year(year):
    url = f"{OCP_BASE}{year}.xlsx"
    print(f"  Downloading {year}.xlsx ...", end=" ", flush=True)
    try:
        r = requests.get(url, timeout=120, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code == 200 and len(r.content) > 500:
            xl = pd.ExcelFile(io.BytesIO(r.content), engine='openpyxl')
            df = xl.parse(xl.sheet_names[0])
            print(f"OK ({len(r.content)//1024} KB, {len(df)} rows)")
            return df
        print(f"FAILED (HTTP {r.status_code})")
        return None
    except Exception as e:
        print(f"ERROR: {e}")
        return None

all_frames = []
for yr in years_needed:
    df_yr = download_year(yr)
    if df_yr is not None:
        df_yr['_source_year'] = yr
        all_frames.append(df_yr)
    time.sleep(1)

if not all_frames:
    print("\nCRITICAL: Could not download data.")
    exit()

df = pd.concat(all_frames, ignore_index=True)
print(f"\nTotal rows: {len(df)}")

# ============================================================
#  COLUMN MAPPING
#  Exact names confirmed from the API output:
#    tender_title, buyer_name, tender_status,
#    tender_description, tender_tenderPeriod_endDate, date, ocid
# ============================================================
df['tender_title']   = df.get('tender_title',   pd.Series("No Title", index=df.index)).fillna("No Title").astype(str).str.strip()
df['description']    = df.get('tender_description', pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
df['buyer_name']     = df.get('buyer_name',     pd.Series("Unknown Dept", index=df.index)).fillna("Unknown Dept").astype(str).str.strip()
df['expiry_date']    = df.get('tender_tenderPeriod_endDate', pd.Series("Not Specified", index=df.index)).fillna("Not Specified").astype(str).str.strip()
df['date_published'] = df.get('date',           pd.Series("Not Specified", index=df.index)).fillna("Not Specified").astype(str).str.strip()
df['status']         = df.get('tender_status',  pd.Series("unknown", index=df.index)).fillna("unknown").astype(str).str.lower().str.strip()
df['ocid']           = df.get('ocid',           pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
df['province']       = df.get('tender_province', pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
df['category']       = df.get('tender_category', pd.Series("", index=df.index)).fillna("").astype(str).str.strip()

# Fill blank ocids
mask = df['ocid'] == ''
df.loc[mask, 'ocid'] = "row_" + df[mask].index.astype(str)

# ============================================================
#  DATE FILTER — strip timezone from parsed dates so comparison works
# ============================================================
for col in ['date_published', 'expiry_date']:
    parsed = pd.to_datetime(df[col], errors='coerce', utc=True)   # parse as UTC
    df[f'{col}_dt'] = parsed.dt.tz_localize(None)                  # strip tz → naive

# Compare against naive datetimes
date_from_naive = date_from.replace(tzinfo=None)
date_to_naive   = today.replace(tzinfo=None)

in_window = (
    (df['date_published_dt'].between(date_from_naive, date_to_naive)) |
    (df['expiry_date_dt'].between(date_from_naive, date_to_naive))    |
    ((df['expiry_date_dt'] >= date_from_naive) & (df['status'] == 'active'))
)
df_window = df[in_window].copy()
print(f"Rows in 90-day window: {len(df_window)}")

if df_window.empty:
    print("WARNING: date filter removed everything — using full dataset")
    df_window = df.copy()

STATUS_MAP = {
    'active':'🟢 Open', 'open':'🟢 Open',
    'complete':'🔴 Closed', 'closed':'🔴 Closed', 'awarded':'🔴 Closed',
    'cancelled':'⚫ Cancelled', 'unsuccessful':'🔴 Unsuccessful',
    'planning':'🔵 Planned', 'planned':'🔵 Planned',
    'unknown':'⚪ Unknown',
}
df_window['status_label'] = df_window['status'].map(STATUS_MAP).fillna('⚪ Unknown')
print(f"Status breakdown: {df_window['status'].value_counts().to_dict()}")

# ============================================================
#  KEYWORD FILTER
# ============================================================
df_window['combined_text'] = (df_window['tender_title'] + " " + df_window['description']).str.lower()

excl = df_window['combined_text'].apply(lambda x: any(e in x for e in EXCLUDE_KEYWORDS))
print(f"Excluded ({EXCLUDE_KEYWORDS}): {excl.sum()}")
df_window = df_window[~excl].copy()

df_window['vehicle_score'] = df_window['combined_text'].apply(
    lambda x: sum(k in x for k in KEYWORDS)
)
print(f"Score distribution: {df_window['vehicle_score'].value_counts().sort_index().to_dict()}")

relevant_df = df_window[df_window['vehicle_score'] >= 1].drop_duplicates(subset=['ocid']).copy()
print(f"\nRelevant vehicle tenders: {len(relevant_df)}")

if relevant_df.empty:
    print("WARNING: 0 results after filtering.")
    exit()

new_tenders = relevant_df[~relevant_df['ocid'].isin(seen_ids)].copy()
print(f"New (not in history): {len(new_tenders)}")

# ============================================================
#  AI SUMMARIES
# ============================================================
def get_ai_summaries(tenders_df):
    if tenders_df.empty or ANTHROPIC_API_KEY == "YOUR_ANTHROPIC_API_KEY_HERE":
        if ANTHROPIC_API_KEY == "YOUR_ANTHROPIC_API_KEY_HERE":
            print("WARNING: No Anthropic API key — skipping AI summaries")
        return {}
    all_summaries = {}
    rows       = list(tenders_df.iterrows())
    BATCH_SIZE = 15   # small batches to stay well within token limits
    for b in range(0, len(rows), BATCH_SIZE):
        batch  = rows[b:b+BATCH_SIZE]
        items, ocids = [], []
        for i, (_, row) in enumerate(batch, 1):
            # Use only the title if description is very long — keeps prompt short
            title = str(row['tender_title'])[:120]
            desc  = str(row['description'])[:200].strip()
            if desc:
                items.append(f"{i}. {title} | {desc}")
            else:
                items.append(f"{i}. {title}")
            ocids.append(str(row['ocid']))
        prompt = (
            "For each numbered tender below write ONE summary sentence (max 20 words): "
            "what is being procured and by which department. "
            "Return ONLY a JSON array of strings in the same order. No markdown, no extra text.\n\n"
            + "\n".join(items)
            """
            For each numbered tender below:

            1. Summarize the procurement in under 20 words
            2. State whether it is:
               - HIGHLY VEHICLE RELATED
               - POSSIBLY VEHICLE RELATED
               - LOW VEHICLE RELEVANCE
            
            Return ONLY a JSON array like:
            
            [
              {
                "summary": "...",
                "relevance": "HIGHLY VEHICLE RELATED"
              }
            ]
            """
        )
        try:
            r = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={"x-api-key": ANTHROPIC_API_KEY,
                         "anthropic-version": "2023-06-01",
                         "content-type": "application/json"},
                json={"model": ANTHROPIC_MODEL, "max_tokens": 1024,
                      "messages": [{"role": "user", "content": prompt}]},
                timeout=60,
            )
            if r.status_code != 200:
                print(f"  AI batch {b//BATCH_SIZE+1} error {r.status_code}: {r.text[:300]}")
                time.sleep(2)
                continue
            raw = r.json()["content"][0]["text"].strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"): raw = raw[4:]
            summaries = json.loads(raw.strip())
            all_summaries.update(dict(zip(ocids, summaries)))
            print(f"  AI batch {b//BATCH_SIZE+1}: {len(summaries)} summaries OK")
            time.sleep(0.5)
        except Exception as e:
            print(f"  AI batch error: {e}")
    return all_summaries

print("\nGenerating AI summaries...")
ai_summaries = get_ai_summaries(new_tenders)

# AI fallback: use full description if no AI summary available
def resolve_summary(row):
    ai = ai_summaries.get(str(row['ocid']), "").strip()
    if ai:
        return ai
    desc = str(row.get('description', '')).strip()
    return desc if desc else ""

new_tenders['ai_summary'] = new_tenders.apply(resolve_summary, axis=1)
relevant_df['ai_summary'] = relevant_df.apply(resolve_summary, axis=1)

# ============================================================
#  FORMAT EXPIRY DATE as YYYY/MM/DD and apply EXPIRED status
# ============================================================
def fmt_date(val):
    try:
        return pd.to_datetime(val, utc=True).strftime("%Y/%m/%d")
    except Exception:
        return str(val)

relevant_df['expiry_display'] = relevant_df['expiry_date'].apply(fmt_date)
new_tenders['expiry_display'] = new_tenders['expiry_date'].apply(fmt_date)

# Mark expired: expiry date is before today (i.e. yesterday or earlier)
yesterday = today.replace(tzinfo=None) - timedelta(days=1)

def apply_expired(row):
    try:
        exp_dt = pd.to_datetime(row['expiry_display'], format="%Y/%m/%d")
        if exp_dt <= yesterday:
            return 'expired', '🔴 Expired'
    except Exception:
        pass
    return row['status'], row['status_label']

exp_results = relevant_df.apply(apply_expired, axis=1, result_type='expand')
relevant_df[['status', 'status_label']] = exp_results

# ============================================================
#  SAVE MASTER CSV + EXCEL
# ============================================================
master_cols = ['ocid','tender_title','buyer_name','province','category',
               'status','status_label','date_published',
               'expiry_display','vehicle_score','description','ai_summary']
new_master  = relevant_df[[c for c in master_cols if c in relevant_df.columns]].copy()
new_master  = new_master.rename(columns={'expiry_display': 'expiry_date'})

if not FIRST_RUN and os.path.exists(MASTER_90DAY_FILE):
    try:
        old = pd.read_excel(MASTER_90DAY_FILE, engine='openpyxl')
        if 'ai_summary' in old.columns:
            old_sums = old.set_index('ocid')['ai_summary'].dropna().to_dict()
            new_master['ai_summary'] = new_master.apply(
                lambda r: r['ai_summary'] or old_sums.get(r['ocid'], ""), axis=1)
    except Exception: pass

with pd.ExcelWriter(MASTER_90DAY_FILE, engine='openpyxl') as writer:
    new_master.to_excel(writer, sheet_name='90-Day Master', index=False)
print(f"\nOK  Master file: {len(new_master)} rows -> {MASTER_90DAY_FILE}")

open_count    = new_master['status'].isin(['active','open']).sum()
expired_count = (new_master['status'] == 'expired').sum()
closed_count  = new_master['status'].isin(['complete','closed','awarded','unsuccessful','cancelled']).sum()
other_count   = len(new_master) - open_count - expired_count - closed_count

excel_name = f"VehicleTenders_90day_{date_to_str}.xlsx"
try:
    sorted_m = new_master.sort_values(
        by='status', key=lambda s: s.map({'active':0,'open':0,'planning':1,'planned':1,'expired':3}).fillna(2))
    with pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
        sorted_m.to_excel(writer, sheet_name='All Tenders', index=False)
        new_master[new_master['status'].isin(['active','open'])].to_excel(
            writer, sheet_name='Open Only', index=False)
        new_master[new_master['status'].isin(
            ['complete','closed','awarded','unsuccessful','cancelled','expired'])].to_excel(
            writer, sheet_name='Closed & Expired', index=False)
    print(f"OK  Excel: {excel_name}  ({len(new_master)} rows)")
except Exception as e:
    print(f"FAILED Excel: {e}")
    excel_name = None

# ============================================================
#  HTML EMAIL
# ============================================================
html = f"""
<html><body style="font-family:Arial,sans-serif;color:#222;max-width:900px;margin:auto">
<h2 style="color:#1a56db">Vehicle Tender Report &mdash; {today.strftime('%d %B %Y')}</h2>
<p style="color:#888;font-size:12px">
  Period: {date_from_str} to {date_to_str} &nbsp;|&nbsp;
  Source: data.open-contracting.org (SA National Treasury)
  {"&nbsp;|&nbsp;<b style='color:#e65100'>FIRST RUN</b>" if FIRST_RUN else ""}
</p>
<table style="border-collapse:separate;border-spacing:6px;width:100%;margin-bottom:24px">
  <tr>
    <td style="background:#e8f5e9;padding:14px;border-radius:8px;text-align:center">
      <div style="font-size:28px;font-weight:bold;color:#2e7d32">{len(new_master)}</div>
      <div style="font-size:12px;color:#555">Total (90d)</div></td>
    <td style="background:#e3f2fd;padding:14px;border-radius:8px;text-align:center">
      <div style="font-size:28px;font-weight:bold;color:#1565c0">{open_count}</div>
      <div style="font-size:12px;color:#555">🟢 Open</div></td>
    <td style="background:#fce4ec;padding:14px;border-radius:8px;text-align:center">
      <div style="font-size:28px;font-weight:bold;color:#c62828">{closed_count}</div>
      <div style="font-size:12px;color:#555">🔴 Closed</div></td>
    <td style="background:#fff8e1;padding:14px;border-radius:8px;text-align:center">
      <div style="font-size:28px;font-weight:bold;color:#b71c1c">{expired_count}</div>
      <div style="font-size:12px;color:#555">🔴 Expired</div></td>
    <td style="background:#f3f3f3;padding:14px;border-radius:8px;text-align:center">
      <div style="font-size:28px;font-weight:bold;color:#757575">{other_count}</div>
      <div style="font-size:12px;color:#555">🔵 Other</div></td>
  </tr>
</table>
<h3 style="color:#1a56db;border-bottom:2px solid #1a56db;padding-bottom:4px">
  All {len(new_master)} Vehicle Tenders — Last 90 Days</h3>
<table style="width:100%;border-collapse:collapse;font-size:13px">
  <thead><tr style="background:#1a56db;color:white">
    <th style="padding:9px 10px;text-align:left;width:28%">Tender Title</th>
    <th style="padding:9px 10px;text-align:left;width:20%">Department</th>
    <th style="padding:9px 10px;text-align:left;width:10%">Province</th>
    <th style="padding:9px 10px;text-align:left;width:9%">Status</th>
    <th style="padding:9px 10px;text-align:left;width:10%">Closes</th>
    <th style="padding:9px 10px;text-align:left;width:23%">AI Summary</th>
  </tr></thead><tbody>
"""

table_data = new_master.sort_values(
    by='status', key=lambda s: s.map({'active':0,'open':0,'planning':1,'planned':1}).fillna(2))

for i, (_, row) in enumerate(table_data.iterrows()):
    expiry  = str(row.get('expiry_date',''))
    summary = str(row.get('ai_summary',''))
    bg      = "#ffffff" if i % 2 == 0 else "#f2f7ff"
    try:
        days = (pd.to_datetime(expiry) - today.replace(tzinfo=None)).days
        if 0 <= days <= 7: bg = "#fff3e0"
    except Exception: pass
    html += (f"<tr style='background:{bg}'>"
             f"<td style='padding:7px 10px'>{row['tender_title']}</td>"
             f"<td style='padding:7px 10px'>{row['buyer_name']}</td>"
             f"<td style='padding:7px 10px'>{row.get('province','')}</td>"
             f"<td style='padding:7px 10px;white-space:nowrap'>{row.get('status_label','⚪')}</td>"
             f"<td style='padding:7px 10px;white-space:nowrap'>{expiry}</td>"
             f"<td style='padding:7px 10px;color:#555;font-style:italic'>{summary}</td>"
             f"</tr>")

html += ("</tbody></table>"
         "<p style='font-size:11px;color:#aaa;margin-top:6px'>"
         "Amber rows = closing within 7 days. Full data in attached Excel.</p>"
         "<p style='font-size:11px;color:#ccc;margin-top:16px'>"
         "data.open-contracting.org + Claude AI | Exclusions: carport</p></body></html>")

# ============================================================
#  SEND EMAIL
# ============================================================
msg            = MIMEMultipart()
msg['From']    = EMAIL_FROM
msg['To']      = ", ".join(EMAIL_TO)
msg['Subject'] = (f"{'[FIRST RUN] ' if FIRST_RUN else ''}"
                  f"Vehicle Tenders: {len(new_master)} Total | "
                  f"{open_count} Open / {closed_count} Closed (90d)")
msg.attach(MIMEText(html, 'html'))
if excel_name and os.path.exists(excel_name):
    with open(excel_name, "rb") as f:
        part = MIMEApplication(f.read(), Name=excel_name)
        part['Content-Disposition'] = f'attachment; filename="{excel_name}"'
        msg.attach(part)
try:
    srv = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    srv.login(EMAIL_FROM, EMAIL_PASSWORD)
    srv.send_message(msg)
    srv.quit()
    print("OK  Email sent")
except Exception as e:
    print(f"FAILED  Email: {e}")

# ============================================================
#  WHATSAPP
# ============================================================
def send_wa(text):
    if not WHATSAPP_ENABLED: return
    try:
        r = requests.get("https://api.callmebot.com/whatsapp.php",
                         params={"phone":WHATSAPP_NUMBER,"text":text,"apikey":WHATSAPP_API_KEY},
                         timeout=15)
        print("OK  WhatsApp" if r.status_code==200 else f"FAILED  WhatsApp {r.status_code}")
    except Exception as e: print(f"FAILED  WhatsApp: {e}")

top = (new_master[new_master['status'].isin(['active','open'])].iloc[0]
       if not new_master[new_master['status'].isin(['active','open'])].empty
       else (new_master.iloc[0] if not new_master.empty else None))
send_wa(
    f"{'[FIRST RUN] ' if FIRST_RUN else ''}Vehicle Tenders 90d\n"
    f"Total: {len(new_master)} | Open: {open_count} | Closed: {closed_count}\n"
    + (f"\nTop: {top['tender_title'][:50]}\n{top['buyer_name']}\nCloses: {top.get('expiry_date','')}"
       if top is not None else "No tenders found")
)

# ============================================================
#  UPDATE HISTORY FILE
# ============================================================
if not relevant_df.empty:
    new_hist = relevant_df[[c for c in master_cols if c in relevant_df.columns]].copy()
    new_hist  = new_hist.rename(columns={'expiry_display': 'expiry_date'})
    new_hist['date_found'] = today.strftime("%Y/%m/%d")

    if not master_df.empty:
        combined = pd.concat([master_df, new_hist]).drop_duplicates(subset=['ocid'])
    else:
        combined = new_hist

    with pd.ExcelWriter(HISTORY_FILE, engine='openpyxl') as writer:
        combined.sort_values('date_found', ascending=False).to_excel(
            writer, sheet_name='History', index=False)
    print(f"OK  History updated: {len(combined)} total records -> {HISTORY_FILE}")

print(f"""
--- FINAL SUMMARY ---
Total raw rows          : {len(df)}
After 90-day filter     : {len(df_window)}
After keyword filter    : {len(relevant_df)}
  🟢 Open               : {open_count}
  🔴 Closed             : {closed_count}
  🔴 Expired            : {expired_count}
  🔵 Other/Unknown      : {other_count}
Daily Excel             : {excel_name}
90-Day Master           : {MASTER_90DAY_FILE}
History File            : {HISTORY_FILE}
DONE
""")
