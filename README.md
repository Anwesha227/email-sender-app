# 📧 Custom Mass Email Sender (Streamlit App)

Send personalized mass emails directly from a spreadsheet easily.

## Features
- Upload Excel (.xlsx) and select a sheet
- Compose emails with dynamic `{ColumnName}` placeholders
- Insert placeholders with click-buttons
- Upload multiple attachments
- Preview sample emails before sending
- Progress bar while sending
- Error report for skipped/failed emails (downloadable)

## How to Deploy

### Local Setup
```bash
pip install streamlit pandas openpyxl
streamlit run app.py
```

### Streamlit Community Cloud (Free Hosting)
1. Create a GitHub repo and upload this project (`app.py`, `requirements.txt`, `README.md`)
2. Go to [https://streamlit.io/cloud](https://streamlit.io/cloud)
3. New App → Connect your GitHub repo → Select `app.py`
4. Click Deploy! 🚀

You will get a public website instantly!

---

⚠️ **Gmail Users:** If using Gmail with 2FA, create an [App Password](https://support.google.com/accounts/answer/185833) and use it instead of your normal password.