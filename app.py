# --- All imports and setup ---
import streamlit as st
import pandas as pd
import smtplib, os, json
from email.message import EmailMessage
from datetime import datetime
import time
import imaplib
from concurrent.futures import ThreadPoolExecutor
from time import perf_counter
import urllib.parse
import uuid
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config("📧 Email Campaign App", layout="wide")

# --- Google Sheet Setup ---
SHEET_NAME = "CampaignHistory"
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource
def get_google_sheet():
    from google.oauth2.service_account import Credentials
    import gspread

    # Load credentials from Streamlit secrets
    service_account_info = st.secrets["google_service_account"]
    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPE
    )

    gc = gspread.authorize(credentials)
    sheet = gc.open(SHEET_NAME).sheet1

    # Check if headers exist, if not, create them
    headers = sheet.row_values(1)
    if not headers:
        sheet.insert_row(["timestamp", "campaign_name", "subject", "total", "delivered", "failed"], 1)
    
    return sheet


def append_to_sheet(data_dict):
    sheet = get_google_sheet()
    if sheet:
        sheet.append_row(list(data_dict.values()), value_input_option="USER_ENTERED")

def load_campaigns_from_sheet():
    sheet = get_google_sheet()
    if not sheet:
        return []
    records = sheet.get_all_records()
    return records

# --- Uptime Check ---
params = st.query_params
if "ping" in params:
    st.write("✅ App is alive!")
    st.stop()

# Folders
os.makedirs("campaign_results", exist_ok=True)
os.makedirs("campaign_resume", exist_ok=True)

# Functions
def load_campaigns():
    return load_campaigns_from_sheet()

def log_campaign(metadata):
    with open("campaigns.json", "w") as f:
        json.dump(load_campaigns_from_sheet() + [metadata], f, indent=2)
    append_to_sheet(metadata)

def save_resume_point(timestamp, data, last_sent_index):
    with open(f"campaign_resume/{timestamp}.json", "w") as f:
        json.dump({
            "data": data,
            "last_sent_index": last_sent_index
        }, f)

def load_resume_point(timestamp):
    try:
        with open(f"campaign_resume/{timestamp}.json") as f:
            return json.load(f)
    except Exception:
        return None

def generate_email_html(full_name, recipient_email=None, subject=None):
    if not full_name or str(full_name).lower() == "nan":
        name_part = "Visitor"
    else:
        name_part = f"{full_name}"

    event_url = "https://www.eventbrite.com/e/bournemouth-b2b-growth-expo-2025-book-your-stand-at-the-premier-b2b-show-tickets-815224106247?aff=oddtdtcreator"
    encoded_event_url = urllib.parse.quote(event_url, safe='')
    
    email_for_tracking = recipient_email if recipient_email else "unknown@example.com"
    encoded_subject = urllib.parse.quote(subject or "No Subject", safe='')
    tracking_link = f"https://tracking-oyzi.onrender.com/track/click?email={email_for_tracking}&url={encoded_event_url}&subject={encoded_subject}"
    tracking_pixel = f'<img src="https://tracking-oyzi.onrender.com/track/open?email={email_for_tracking}&subject={encoded_subject}" width="1" height="1" style="display:none;">'
    unsubscribe_link = f"https://unsubscribe-thsj.onrender.com/unsubscribe?email={email_for_tracking}"

    return f"""
    <html>
      <body style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
       {tracking_pixel} <!-- move tracking pixel HIGH UP for better load rate -->

        <h2 style="color: #2E86C1;">🎉 Thank You for Making Milton Keynes B2B Growth Expo 2025 a Huge Success!</h2>

        <p>Dear <strong>{name_part}</strong>,</p>

        <p>On behalf of the entire <strong>B2B Growth Expo</strong> team, we would like to extend our heartfelt thanks to you for attending the <strong>Milton Keynes B2B Growth Expo 2025</strong> on 23rd April.</p>

        <p>Your energy, enthusiasm, and engagement helped make this event a truly remarkable success!</p>

        <p>Over 100+ exhibitors, inspiring speakers, and unlimited networking opportunities were made even more special by your presence.</p>

        <p>We hope you left with new ideas, valuable contacts, and exciting opportunities to fuel your business journey ahead.</p>

        <h3 style="color: #2E86C1; margin-top: 30px;">What’s Next?</h3>

        <p>We are thrilled to announce that our next big event is already on the horizon:</p>

        <ul style="line-height: 1.6;">
          <li><strong>Event:</strong> Bournemouth B2B Growth Expo 2025</li>
          <li><strong>Date:</strong> Thursday, 3rd July 2025</li>
          <li><strong>Venue:</strong> Citygate Centre, 138a Holdenhurst Road, Bournemouth BH8 8AS</li>
          <li><strong>Time:</strong> 10:00 AM – 4:30 PM</li>
        </ul>

        <p>Get ready for another powerful day of learning, networking, and growing your business with an even bigger line-up of exhibitors, workshops, and keynote speakers!</p>

        <p style="margin-top: 20px;"><strong>Early Bird Offer – 50% OFF!</strong><br/>
          Don't miss out on this exclusive deal!<br/>
          <strong>Use Code at Checkout:</strong> <span style="background-color: #F7DC6F; padding: 3px 6px; border-radius: 3px; font-weight: bold;">BCTC50</span>
        </p>

        <p style="margin-top: 10px;">
          <a href="{tracking_link}" target="_blank" style="background-color: #2E86C1; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; display: inline-block; font-weight: bold;">
            Book My Stand Now
          </a>
        </p>

        <p style="margin-top: 25px;"><em>Offer valid for a limited time—act fast and secure your spot at the South’s biggest B2B growth event!</em></p>

        <br/>
        <p>Once again, thank you for being an important part of the Milton Keynes B2B Growth Expo community.</p>

        <p>We can’t wait to see you again soon — building, growing, and succeeding together!</p>

        <br/>
        <p>Best regards,</p>
        <p>
          Mike Randell<br/>
          Marketing Executive | B2B Growth Expo<br/>
          <a href="mailto:mike@miltonkeynesexpo.com">mike@miltonkeynesexpo.com</a><br/>
          (+44) 03303 209 609
        </p>

        <p style="font-size: 12px; color: #888; text-align: center; margin-top: 20px;">
          If you no longer wish to receive emails from us, you can <a href="{unsubscribe_link}" style="color: #2E86C1; text-decoration: none;">unsubscribe here</a>.
        </p>

      </body>
    </html>
    """

def send_email(sender_email, sender_password, row, subject):
    try:
        server = smtplib.SMTP("mail.miltonkeynesexpo.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = row['email']
        msg.set_content(generate_email_html(row['full_name'], row['email'], subject), subtype='html')

        server.send_message(msg)

        try:
            imap = imaplib.IMAP4_SSL("mail.miltonkeynesexpo.com")
            imap.login(sender_email, sender_password)
            imap.append('INBOX.Sent', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
            imap.logout()
        except Exception as e:
            return (row['email'], f"✅ Delivered (⚠️ Failed to save to Sent: {e})")

        server.quit()
        return (row['email'], "✅ Delivered")
    except Exception as e:
        return (row['email'], f"❌ Failed: {e}")

def send_delivery_report(sender_email, sender_password, report_file):
    try:
        msg = EmailMessage()
        msg['Subject'] = "Delivery Report for Email Campaign"
        msg['From'] = sender_email
        msg['To'] = "b2bgrowthexpo@gmail.com"
        msg.set_content("Please find the attached delivery report for the recent email campaign.")

        with open(report_file, 'rb') as file:
            msg.add_attachment(file.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(report_file))

        server = smtplib.SMTP("mail.miltonkeynesexpo.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()

        st.success("📤 Delivery report emailed to **b2bgrowthexpo@gmail.com**")
    except Exception as e:
        st.error(f"❌ Could not send delivery report: {e}")

# --- UI Starts Here ---
st.title("📨 Automated Email Campaign Manager")

with st.expander("📜 View Past Campaigns"):
    for c in reversed(load_campaigns()):
        name = c.get("campaign_name", "")
        timestamp = c.get("timestamp", "")
        label = f"📧 {name} {timestamp}" if name else f"🕒 {timestamp}"
        st.markdown(f"**{label}** | 👥 {c['total']} | ✅ {c['delivered']} | ❌ {c['failed']}")

st.header("📤 Send Email Campaign")
sender_email = st.text_input("Sender Email", value="mike@miltonkeynesexpo.com")
sender_password = st.text_input("Password", type="password")
subject = st.text_input("Email Subject")
campaign_name = st.text_input("Campaign Name", placeholder="e.g. MK Expo – VIP Invite List")
file = st.file_uploader("Upload CSV with `email`, `full name` columns")

st.subheader("📧 Preview of Email:")
st.components.v1.html(generate_email_html("Sarah Johnson"), height=500, scrolling=True)

resume_data = None
resume_choice = False
latest_resume = None

if os.path.exists("campaign_resume"):
    files = sorted(os.listdir("campaign_resume"), reverse=True)
    if files:
        latest_resume = files[0]
        resume_data = load_resume_point(latest_resume.replace(".json", ""))
        if resume_data:
            resume_choice = st.checkbox(f"🔁 Resume Last Campaign ({latest_resume})")

if st.button("🚀 Start Campaign"):
    if not subject or not sender_email or not sender_password:
        st.warning("Please fill in all fields.")
        st.stop()

    if resume_choice and resume_data:
        df = pd.DataFrame(resume_data["data"])
        timestamp = latest_resume.replace(".json", "")
        last_sent_index = resume_data["last_sent_index"]
        df = df.iloc[last_sent_index:]
        st.success("🔄 Resuming previous campaign from saved point...")
    else:
        if not file:
            st.warning("Please upload a CSV file.")
            st.stop()

        df = pd.read_csv(file)
        df.columns = df.columns.str.strip().str.lower()
        col_map = {col: "email" if "email" in col else "full_name" for col in df.columns if "email" in col or "name" in col}
        df.rename(columns=col_map, inplace=True)

        if not {"email", "full_name"}.issubset(df.columns):
            st.error("CSV must contain `email` and `full name` columns.")
            st.stop()

        df = df[["email", "full_name"]].dropna(subset=["email"]).drop_duplicates(subset="email")
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        last_sent_index = 0

    total = len(df)
    delivered, failed = 0, 0
    delivery_report = []

    progress = st.progress(0)
    status_text = st.empty()
    time_text = st.empty()

    start_time = perf_counter()

    with ThreadPoolExecutor(max_workers=20) as executor:
        futures = []

        for i, row in df.iterrows():
            status_text.text(f"📤 Sending email to {row['email']}")
            futures.append(executor.submit(send_email, sender_email, sender_password, row, subject))

        for i, future in enumerate(futures):
            email, result = future.result()
            if "✅" in result:
                delivered += 1
            else:
                failed += 1

            delivery_report.append({"email": email, "status": result})
            progress.progress((i + 1) / total)
            save_resume_point(timestamp, df.to_dict(orient="records"), i + 1)

            # Estimate time
            elapsed = perf_counter() - start_time
            avg_per_email = elapsed / (i + 1)
            remaining = avg_per_email * (total - (i + 1))
            mins, secs = divmod(remaining, 60)
            time_text.text(f"⏳ Estimated time left: {int(mins)}m {int(secs)}s")

    # Final time and summary
    duration = perf_counter() - start_time
    final_mins, final_secs = divmod(duration, 60)
    avg_per_email = duration / total
    estimated_total = avg_per_email * total
    est_mins, est_secs = divmod(estimated_total, 60)

    time_text.markdown(
        f"""
        ### ✅ Campaign Finished!
        - ⏱️ **Actual Time Taken:** {int(final_mins)}m {int(final_secs)}s  
        - ⏳ **Originally Estimated Time:** {int(est_mins)}m {int(est_secs)}s
        """
    )

    with st.container():
        st.markdown("### 📊 Campaign Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Emails", total)
        col2.metric("✅ Delivered", delivered)
        col3.metric("❌ Failed", failed)

        metadata = {
            "timestamp": timestamp,
            "campaign_name": campaign_name,
            "subject": subject,
            "total": total,
            "delivered": delivered,
            "failed": failed
        }

        log_campaign(metadata)

        report_df = pd.DataFrame(delivery_report)
        report_filename = f"campaign_results/report_{campaign_name.replace(' ', '_')}_{timestamp}.csv"
        report_df.to_csv(report_filename, index=False)

        st.download_button("📥 Download Delivery Report", data=report_df.to_csv(index=False), file_name=os.path.basename(report_filename), mime="text/csv")

        send_delivery_report(sender_email, sender_password, report_filename)
