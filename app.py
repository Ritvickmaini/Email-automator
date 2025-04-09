import streamlit as st
import pandas as pd
import smtplib, os, json
from email.message import EmailMessage
from datetime import datetime
import time
import imaplib

st.set_page_config("📧 Email Campaign App", layout="wide")

# ✅ Uptime Robot check
params = st.query_params
if "ping" in params:
    st.write("✅ App is alive!")
    st.stop()

# Create folders if missing
os.makedirs("campaign_results", exist_ok=True)
os.makedirs("campaign_resume", exist_ok=True)

# Load previous campaign logs
def load_campaigns():
    if os.path.exists("campaigns.json"):
        try:
            with open("campaigns.json") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return []
    return []

# Save campaign metadata
def log_campaign(metadata):
    campaigns = load_campaigns()
    campaigns.append(metadata)
    with open("campaigns.json", "w") as f:
        json.dump(campaigns, f, indent=2)

# Resume checkpoint storage
def save_resume_point(timestamp, batch_index, data):
    with open(f"campaign_resume/{timestamp}.json", "w") as f:
        json.dump({
            "batch_index": batch_index,
            "data": data
        }, f)

def load_resume_point(timestamp):
    try:
        with open(f"campaign_resume/{timestamp}.json") as f:
            return json.load(f)
    except:
        return None

# Email template with personalization (no company name)
def generate_email_html(full_name):
    return f"""
    <html>
      <body style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
        <h2 style="color: #2E86C1;">You're Invited to the Milton Keynes Business Expo 2025!</h2>
        <p>Dear <strong>{full_name}</strong>,</p>
        <p>We are excited to invite you to the Milton Keynes Business Expo — one of the region’s largest networking events bringing together hundreds of businesses under one roof.</p>

        <ul style="line-height: 1.6;">
          <li><strong>Date:</strong> 23rd April, 2025</li>
          <li><strong>Time:</strong> 10:00 AM – 4:30 PM</li>
          <li><strong>Venue:</strong> The Ridgeway Centre, MK12 5TH, United Kingdom</li>
        </ul>

        <p style="margin-top: 20px;">🎟️ <strong>Grab your free visitor ticket:</strong><br/>
        <a href="https://www.eventbrite.com/e/998974207747" target="_blank" style="display: inline-block; background-color: #2E86C1; color: white; padding: 12px 20px; text-decoration: none; border-radius: 5px; margin-top: 10px;">
          Register on Eventbrite
        </a></p>

        <p style="margin-top: 30px;">If you’re interested in exhibiting or getting involved as a keynote speaker, seminar host, or panelist — simply reply to this email or contact us at the number provided below.</p>

        <br/>
        <p>Best regards,</p>
        <p>
          Mike Randell<br/>
          Marketing Executive | B2B Growth Expo<br/>
          <a href="mailto:mike@miltonkeynesexpo.com">mike@miltonkeynesexpo.com</a><br/>
          (+44) 03303 209 609
        </p>
      </body>
    </html>
    """

# Send a batch of emails
def send_emails_batch(sender_email, sender_password, batch_df, subject):
    import random
    results = []

    try:
        server = smtplib.SMTP("mail.miltonkeynesexpo.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
    except Exception as e:
        st.error(f"❌ SMTP Error: {e}")
        return [(row['email'], f"❌ SMTP Error: {e}") for _, row in batch_df.iterrows()]

    # Try IMAP for saving to Sent
    try:
        imap = imaplib.IMAP4_SSL("mail.miltonkeynesexpo.com")
        imap.login(sender_email, sender_password)
    except:
        imap = None

    for _, row in batch_df.iterrows():
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = row['email']
        html_body = generate_email_html(row['full_name'])
        msg.set_content(html_body, subtype='html')

        try:
            server.send_message(msg)
            results.append((row['email'], "✅ Delivered"))

            if imap:
                imap.append('INBOX.Sent', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())

            delay = round(random.uniform(5, 12), 2)
            st.write(f"⏳ Waiting {delay} seconds...")
            time.sleep(delay)

        except Exception as e:
            results.append((row['email'], f"❌ Failed: {e}"))

    server.quit()
    if imap:
        imap.logout()
    return results

# --- Streamlit UI ---
st.title("📨 Automated Email Campaign Manager")

# View logs
with st.expander("📜 View Past Campaigns"):
    for c in reversed(load_campaigns()):
        st.markdown(f"**🕒 {c['timestamp']}** | **📧 {c['subject']}** | 👥 {c['total']} | ✅ {c['delivered']} | ❌ {c['failed']}")

# Inputs
st.header("📤 Send Email Campaign")
sender_email = st.text_input("Sender Email", value="mike@miltonkeynesexpo.com")
sender_password = st.text_input("Password", type="password")
subject = st.text_input("Subject")
file = st.file_uploader("Upload CSV (email, full name)")

# Preview
st.subheader("📧 Email Preview:")
st.components.v1.html(generate_email_html("Sarah Johnson"), height=500, scrolling=True)

# Resume prompt
resume_available = False
resume_data = None
resume_choice = False

if os.path.exists("campaign_resume"):
    files = sorted(os.listdir("campaign_resume"), reverse=True)
    if files:
        latest = files[0]
        resume_data = load_resume_point(latest.replace(".json", ""))
        if resume_data:
            resume_available = True
            resume_choice = st.checkbox(f"🔁 Resume Last Campaign ({latest}) from Batch {resume_data['batch_index'] + 1}")

# Send emails
if st.button("🚀 Start Campaign"):
    if not subject or not sender_email or not sender_password:
        st.warning("Please fill in all fields.")
        st.stop()

    if resume_choice and resume_data:
        df = pd.DataFrame(resume_data["data"])
        timestamp = latest.replace(".json", "")
        start_batch = resume_data["batch_index"] + 1
        st.success(f"Resuming campaign from batch {start_batch + 1}...")
    else:
        if not file:
            st.warning("Please upload a CSV.")
            st.stop()

        df = pd.read_csv(file)
        df.columns = df.columns.str.strip().str.lower()
        col_map = {col: "email" if "email" in col else "full_name" for col in df.columns if "email" in col or "name" in col}
        df.rename(columns=col_map, inplace=True)

        if not {'email', 'full_name'}.issubset(df.columns):
            st.error("CSV must have columns for email and full name.")
            st.stop()

        df = df[["email", "full_name"]].dropna().drop_duplicates(subset="email")
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        start_batch = 0

    total = len(df)
    delivered, failed = 0, 0

    for i in range(start_batch, (total // 100) + 1):
        batch = df.iloc[i*100 : (i+1)*100]
        if batch.empty:
            continue

        st.info(f"Sending batch {i + 1} ({len(batch)} emails)...")
        results = send_emails_batch(sender_email, sender_password, batch, subject)

        # Report
        df_result = pd.DataFrame(results, columns=["Email", "Status"])
        report_file = f"campaign_results/campaign_{timestamp}_batch_{i+1}.json"
        with open(report_file, "w") as f:
            json.dump(results, f, indent=2)

        st.dataframe(df_result)
        st.download_button(
            f"📥 Download Batch {i+1} Report",
            df_result.to_csv(index=False).encode(),
            file_name=f"delivery_report_batch_{i+1}.csv",
            mime="text/csv"
        )

        delivered += sum(1 for _, s in results if "Delivered" in s)
        failed += sum(1 for _, s in results if "Failed" in s)

        # Save resume point
        save_resume_point(timestamp, i, df.to_dict(orient="records"))

    log_campaign({
        "timestamp": timestamp,
        "subject": subject,
        "total": total,
        "delivered": delivered,
        "failed": failed
    })

    st.success(f"✅ Campaign Finished — Delivered: {delivered}, Failed: {failed}")
