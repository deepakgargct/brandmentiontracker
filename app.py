import streamlit as st
import pandas as pd
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import serpapi

# --- CORE FUNCTIONS ---

def search_brand_mentions(brand_name, api_key, num_results=10):
    """Searches the web for recent brand mentions using the user's SerpApi key."""
    client = serpapi.Client(api_key=api_key)
    results_list = []
    
    try:
        search_results = client.search({
            "engine": "google",
            "q": brand_name,
            "num": num_results,
            "hl": "en",
            "gl": "us"
        })
        
        if "organic_results" in search_results:
            for r in search_results["organic_results"]:
                results_list.append({
                    "Title": r.get("title", ""),
                    "URL": r.get("link", ""),
                    "Snippet": r.get("snippet", "")
                })
    except Exception as e:
        st.error(f"SerpApi Search failed. Please check if your API key is valid. Error: {e}")
        
    return results_list

def generate_excel(dataframe):
    """Converts a pandas DataFrame to an Excel file in memory."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Brand Mentions')
    return output.getvalue()

def send_email_with_attachment(recipient_email, brand_name, excel_data):
    """Sends the generated Excel file via SMTP."""
    # The app still needs its own email credentials to send the message out
    sender_email = st.secrets.get("EMAIL_SENDER", "your_email@gmail.com")
    sender_password = st.secrets.get("EMAIL_PASSWORD", "your_app_password")
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"Your Brand Mention Report: {brand_name}"
    
    body = f"Hello,\n\nPlease find the recent web mentions for {brand_name} attached.\n\nBest,\nYour App"
    msg.attach(MIMEText(body, 'plain'))
    
    attachment = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    attachment.set_payload(excel_data)
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', f'attachment; filename="{brand_name}_Report.xlsx"')
    msg.attach(attachment)
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Failed to send email: {e}")
        return False

# --- STREAMLIT UI ---

st.set_page_config(page_title="Brand Monitor", page_icon="🔍")
st.title("🔍 Brand Monitor")
st.write("Enter a brand name and your SerpApi key to scan the web for mentions and generate an Excel report.")

with st.form("brand_form"):
    brand = st.text_input("Brand / Company Name:", placeholder="e.g., Apple, Tesla")
    email = st.text_input("Delivery Email Address:", placeholder="name@example.com")
    
    # New input for the user's API key
    user_api_key = st.text_input("Your SerpApi Key:", type="password", help="Get your free key at serpapi.com")
    
    submit_button = st.form_submit_button("Generate Report")

if submit_button:
    # Updated validation to require all three fields
    if not brand or not email or not user_api_key:
        st.warning("Please fill out the brand name, email address, and your SerpApi key.")
    else:
        with st.status(f"Tracking mentions for **{brand}**...", expanded=True) as status:
            
            st.write("🔎 Searching the web...")
            # Pass the user's API key into the search function
            raw_mentions = search_brand_mentions(brand, user_api_key, num_results=15)
            
            if not raw_mentions:
                status.update(label="No mentions found or invalid API key.", state="error")
                st.stop()
                
            df = pd.DataFrame(raw_mentions)
            
            st.write("📊 Generating Excel report...")
            excel_file = generate_excel(df)
            
            st.write("📧 Sending email...")
            email_sent = send_email_with_attachment(email, brand, excel_file)
            
            if email_sent:
                status.update(label="Report successfully generated and emailed!", state="complete", expanded=False)
            else:
                status.update(label="Generated report, but email failed to send.", state="error")
        
        st.subheader("Data Preview")
        st.dataframe(df, use_container_width=True)
        
        st.download_button(
            label="📥 Download Excel File Directly",
            data=excel_file,
            file_name=f"{brand}_Mentions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
