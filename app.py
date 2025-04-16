# --- All imports and setup unchanged ---
import streamlit as st
import pandas as pd
import smtplib, os, json
from email.message import EmailMessage
from datetime import datetime
import time
import imaplib
from concurrent.futures import ThreadPoolExecutor
from time import perf_counter

st.set_page_config("📧 Email Campaign App", layout="wide")

# ✅ Uptime Robot check
params = st.query_params
if "ping" in params:
    st.write("✅ App is alive!")
    st.stop()

# Create folders if missing
os.makedirs("campaign_results", exist_ok=True)
os.makedirs("campaign_resume", exist_ok=True)

def load_campaigns():
    if os.path.exists("campaigns.json"):
        try:
            with open("campaigns.json") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return []
    return []

def log_campaign(metadata):
    campaigns = load_campaigns()
    campaigns.append(metadata)
    with open("campaigns.json", "w") as f:
        json.dump(campaigns, f, indent=2)

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

def generate_email_html(full_name):
    return f"""
    <html>
      <body style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
        <p>Dear <strong>{full_name}</strong>,</p>
        <p>We’re thrilled to have you registered for the Milton Keynes B2B Growth Expo happening on <strong>23rd April</strong> at <strong>The Ridgeway Centre</strong> – and we’ve lined up a powerhouse of assured rewards and giveaways just for showing up and engaging!</p>
        <p><strong>Here’s what’s waiting for you:</strong></p>
        <ol>
          <li style="margin-bottom: 10px;"><strong>25 Assured Business Leads</strong><br>Boost your sales pipeline instantly with 25 qualified leads from our Sales Lead Machine.</li>
          <li style="margin-bottom: 10px;"><strong>Free Brand New Digital Speaker</strong><br>Walk away with a high-quality digital speaker, gifted to you by B2B Growth Hub.</li>
          <li style="margin-bottom: 10px;"><strong>Free Expo Stand Next Time – Just Refer & Win</strong><br>Refer 10 businesses to register and attend. If they enter your name as reference, you get a complimentary stand next time!<br>
          <a href="https://www.eventbrite.com/e/milton-keynes-b2b-growth-expo-23rd-april-2025-free-visitor-ticket-tickets-998974207747?aff=REFERAFRIEND" target="_blank">Register Your Ticket</a></li>
          <li style="margin-bottom: 10px;"><strong>Visit & Win – Brand New Sofa Set!</strong><br>Visit 50 stands and enter our lucky draw to win a stylish new sofa set.</li>
          <li style="margin-bottom: 10px;"><strong>£50 Cash with Tide Bank</strong><br>Open a business account at the expo and get £50 cash from B2B Growth Hub.</li>
          <li style="margin-bottom: 10px;"><strong>£50 Cash with Worldpay</strong><br>Set up a terminal during the event and get another £50.</li>
          <li style="margin-bottom: 10px;"><strong>Free Book – Vision to Victory</strong><br>Claim your free copy of <em>Vision to Victory</em> by Santosh Kumar.</li>
          <li style="margin-bottom: 10px;"><strong>Free Annual Website Hosting</strong><br>Enjoy 1 year of free hosting from Visualytes.</li>
          <li style="margin-bottom: 10px;"><strong>Free Business Listing worth £450</strong><br>Get a Diamond package listing on our directory for free.</li>
        </ol>
        <p>Your name is already on the list – now all you have to do is show up and claim what’s yours!</p>
        <p><strong>Event Details:</strong><br>
        Date: Tuesday, 23rd April<br>
        Location: The Ridgeway Centre, Featherstone Rd, Wolverton Mill, Milton Keynes MK12 5TH</p>
        <p>If you’d like to schedule a quick meeting with me, use this link:<br>
        <a href="https://tidycal.com/nagendra/b2b-discovery-call" target="_blank">Book a Discovery Call</a></p>
        <p>Thanks & Regards,</p>
        <p>
          Nagendra Mishra<br/>
          Director | B2B Growth Hub<br/>
          <a href="mailto:nagendra@b2bgrowthexpo.com">nagendra@b2bgrowthexpo.com</a><br/>
          +44 7913 027482<br/>
          www.b2bgrowthhub.com
        </p>
        <p style="margin-top: 30px; font-size: 0.9em; color: #888;">
          If you don't want to hear from me again, please let me know.
        </p>
      </body>
    </html>
    """

def send_email(sender_email, sender_password, row, subject, smtp_server):
    try:
        server = smtplib.SMTP(smtp_server, 587)
        server.starttls()
        server.login(sender_email, sender_password)

        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = row['email']
        msg.set_content(generate_email_html(row['full_name']), subtype='html')

        server.send_message(msg)

        try:
            imap = imaplib.IMAP4_SSL(smtp_server)
            imap.login(sender_email, sender_password)
            imap.append('INBOX.Sent', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
            imap.logout()
        except Exception as e:
            return (row['email'], f"✅ Delivered (⚠️ Failed to save to Sent: {e})")

        server.quit()
        return (row['email'], "✅ Delivered")
    except Exception as e:
        return (row['email'], f"❌ Failed: {e}")

def send_delivery_report(sender_email, sender_password, report_file, smtp_server):
    try:
        msg = EmailMessage()
        msg['Subject'] = "Delivery Report for Email Campaign"
        msg['From'] = sender_email
        msg['To'] = "b2bgrowthexpo@gmail.com"
        msg.set_content("Please find the attached delivery report for the recent email campaign.")

        with open(report_file, 'rb') as file:
            msg.add_attachment(file.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(report_file))

        server = smtplib.SMTP(smtp_server, 587)
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
        timestamp = c["timestamp"]
        label = f"📧 {name} {timestamp}" if name else f"🕒 {timestamp}"
        st.markdown(f"**{label}** | 👥 {c['total']} | ✅ {c['delivered']} | ❌ {c['failed']}")

st.header("📤 Send Email Campaign")

server_options = {
    "Mike (miltonkeynesexpo.com)": ("mail.miltonkeynesexpo.com", "mike@miltonkeynesexpo.com"),
    "Nagendra (b2bgrowthexpo.com)": ("mail.b2bgrowthexpo.com", "nagendra@b2bgrowthexpo.com"),
}

selected_identity = st.selectbox("Select Sender Identity", list(server_options.keys()))
smtp_server, sender_email = server_options[selected_identity]

st.text_input("Sender Email", value=sender_email, disabled=True)
sender_password = st.text_input("Password", type="password")
subject = st.text_input("Email Subject")
campaign_name = st.text_input("Campaign Name", placeholder="e.g. MK Expo – VIP Invite List")
file = st.file_uploader("Upload CSV with `email`, `full name` columns")

st.subheader("📧 Preview of Email:")
st.components.v1.html(generate_email_html("Sarah Johnson"), height=800, scrolling=True)

# Resume section unchanged
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

        df = df[["email", "full_name"]].dropna().drop_duplicates(subset="email")
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        last_sent_index = 0

    total = len(df)
    delivered, failed = 0, 0
    delivery_report = []

    progress = st.progress(0)
    status_text = st.empty()
    time_text = st.empty()

    start_time = perf_counter()

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = []

        for i, row in df.iterrows():
            status_text.text(f"📤 Sending email to {row['email']}")
            futures.append(executor.submit(send_email, sender_email, sender_password, row, subject, smtp_server))

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

    duration = perf_counter() - start_time
    final_mins, final_secs = divmod(duration, 60)
    avg_per_email = duration / total
    est_mins, est_secs = divmod(avg_per_email * total, 60)

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

        log_campaign({
            "timestamp": timestamp,
            "campaign_name": campaign_name,
            "subject": subject,
            "total": total,
            "delivered": delivered,
            "failed": failed
        })

        report_df = pd.DataFrame(delivery_report)
        report_filename = f"campaign_results/report_{campaign_name.replace(' ', '_')}_{timestamp}.csv"
        report_df.to_csv(report_filename, index=False)

        st.download_button("📥 Download Delivery Report", data=report_df.to_csv(index=False), file_name=os.path.basename(report_filename), mime="text/csv")

        send_delivery_report(sender_email, sender_password, report_filename, smtp_server)
