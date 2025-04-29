import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from io import BytesIO
import time

st.set_page_config(page_title="Custom Mass Email Sender", page_icon="📧")
st.title("📧 Custom Mass Email Sender")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    sheet = st.selectbox("Select sheet", sheet_names)

    df = pd.read_excel(uploaded_file, sheet_name=sheet)
    st.success(f"Loaded sheet with {len(df)} rows and {len(df.columns)} columns.")

    st.subheader("Compose Email")

    columns = df.columns.tolist()
    placeholder_area = st.empty()

    subject = st.text_input("✉️ Email Subject", "")
    body = st.text_area("📝 Email Body", height=300, placeholder="Type your email here...")

    st.markdown("**Click to insert fields into your email:**")
    col_buttons = st.columns(4)
    for idx, field in enumerate(columns):
        with col_buttons[idx % 4]:
            if st.button(field):
                body += f" {{{field}}} "
                placeholder_area.text_area("📝 Email Body", value=body, height=300, key="body_updated")

    st.subheader("📎 Attachments")
    attachments = st.file_uploader("Upload one or more attachments", accept_multiple_files=True)

    st.subheader("🔐 SMTP Server Settings")
    sender_email = st.text_input("Sender Email")
    sender_password = st.text_input("Sender Password", type="password")
    smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP Port", value=587)

    email_column = st.selectbox("Select the column with recipient Email addresses", columns)

    if st.button("Preview Sample Emails"):
        st.subheader("🔎 Preview (First 3 Rows)")
        preview_rows = df.head(3)
        for idx, row in preview_rows.iterrows():
            try:
                sample_subject = subject.format(**row.to_dict())
                sample_body = body.format(**row.to_dict())
                st.write(f"**To:** {row[email_column]}")
                st.write(f"**Subject:** {sample_subject}")
                st.write(sample_body)
                st.markdown("---")
            except Exception as e:
                st.error(f"Error formatting preview for row {idx}: {e}")

    if st.button("🚀 Send Emails"):
        if not sender_email or not sender_password:
            st.error("Sender credentials are required.")
        else:
            error_rows = []
            success_count = 0
            progress_bar = st.progress(0)

            try:
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)

                total = len(df)
                for idx, row in df.iterrows():
                    recipient = row.get(email_column)
                    if pd.isna(recipient) or not str(recipient).strip():
                        error_rows.append((idx, "Missing Email"))
                        continue

                    try:
                        msg = MIMEMultipart()
                        msg['From'] = sender_email
                        msg['To'] = recipient
                        msg['Subject'] = subject.format(**row.to_dict())

                        formatted_body = body.format(**row.to_dict())
                        msg.attach(MIMEText(formatted_body, 'plain'))

                        for file in attachments:
                            part = MIMEApplication(file.read(), Name=file.name)
                            part['Content-Disposition'] = f'attachment; filename="{file.name}"'
                            msg.attach(part)
                            file.seek(0)

                        server.sendmail(sender_email, recipient, msg.as_string())
                        success_count += 1

                    except Exception as e:
                        error_rows.append((idx, str(e)))

                    progress_bar.progress((idx + 1) / total)

                server.quit()

                st.success(f"✅ Emails successfully sent: {success_count}")
                st.warning(f"⚠️ Errors: {len(error_rows)}")

                if error_rows:
                    error_df = pd.DataFrame(error_rows, columns=["RowIndex", "Error"])
                    csv = error_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "📄 Download Error Report",
                        data=csv,
                        file_name="error_report.csv",
                        mime="text/csv"
                    )

            except Exception as e:
                st.error(f"❌ Failed to connect to SMTP server: {e}")