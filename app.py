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
        <p>We’re thrilled to have you registered for the <strong>Milton Keynes B2B Growth Expo</strong> happening on <strong>23rd April 2025</strong> at <strong>Ridgeway Centre</strong> – and we’ve lined up a powerhouse of assured rewards and giveaways just for showing up and engaging!</p>

        <h3>Here’s what’s waiting for you:</h3>
        <ol style="padding-left: 20px;">
          <li style="margin-bottom: 15px;"><strong>25 Assured Business Leads</strong> – Boost your sales pipeline instantly with 25 qualified leads from our Sales Lead Machine.</li>
          <li style="margin-bottom: 15px;"><strong>Free Brand New Digital Speaker</strong> – Walk away with a high-quality digital speaker, gifted to you by B2B Growth Hub.</li>
          <li style="margin-bottom: 15px;"><strong>Free Expo Stand Next Time – Just Refer & Win</strong> – Refer 10 friends/businesses to register and attend the expo. Ask them to enter your full name as the reference, and you get a complimentary stand at our next big show!<br/>
            <a href="https://www.eventbrite.com/e/milton-keynes-b2b-growth-expo-23rd-april-2025-free-visitor-ticket-tickets-998974207747?aff=REFERAFRIEND">
               Refer Friends
            </a>
          </li>
          <li style="margin-bottom: 15px;"><strong>Visit & Win – Brand New Sofa Set!</strong> – Visit 50 stands on the day and enter our lucky draw to win a stylish new sofa set for your home or office.</li>
          <li style="margin-bottom: 15px;"><strong>£50 Cash with Tide Bank</strong> – Open a free business bank account with Tide at the expo and get £50 cash from B2B Growth Hub – no strings attached!</li>
          <li style="margin-bottom: 15px;"><strong>£50 Cash with Worldpay</strong> – Set up a payment terminal with Worldpay during the event and take home another £50 cash reward from B2B Growth Hub.</li>
          <li style="margin-bottom: 15px;"><strong>Free Book – Business Bible for Entrepreneurs: Vision to Victory</strong> – Grab your complimentary copy of Vision to Victory by Santosh Kumar – a must-read for entrepreneurs and growth-driven minds.</li>
          <li style="margin-bottom: 15px;"><strong>Free Annual Website Hosting</strong> – Get 1-year FREE hosting for your website courtesy of our generous sponsor, Visualytes.</li>
          <li style="margin-bottom: 15px;"><strong>Free Business Listing on our Directory worth £450</strong> – Get a free business listing on our directory on the Diamond package on the spot.</li>
        </ol>

        <p>Your name is already on the list – now all you have to do is turn up and claim what’s yours!</p>

        <h4>📅 Mark your calendar:</h4>
        <ul>
          <li><strong>Date:</strong> Tuesday, 23rd April</li>
          <li><strong>Location:</strong> The Ridgeway Centre, Featherstone Rd, Wolverton Mill, Milton Keynes MK12 5TH</li>
        </ul>

        <p>Make connections, grow your business, and walk away with more than just inspiration!</p>

        <p>If you would like to schedule a meeting with me at your convenient time, please use the diary link below:<br/>
        <a href="https://tidycal.com/nagendra/b2b-discovery-call">Book a Call with Me</a></p>

        <br/>
        <p>Thanks & Regards,</p>
        <div style="margin-top: 10px; font-size: 14px; line-height: 1.6;">
          <strong>Mike Randell</strong><br/>
          Marketing Executive | B2B Growth Expo<br/>
          Mo: +44 7913 027482<br/>
          Email: <a href="mailto:mike@miltonkeynesexpo.com">mike@miltonkeynesexpo.com</a><br/>
          <a href="https://www.b2bgrowthexpo.com">www.b2bgrowthexpo.com</a>
        </div>
        <br/>
        <p style="font-size: 12px; color: gray;">If you don’t want to hear from me again, please let me know.</p>
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
        msg.set_content(generate_email_html(row['full_name']), subtype='html')

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
        timestamp = c["timestamp"]
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

    # Styled summary card
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

        send_delivery_report(sender_email, sender_password, report_filename)
