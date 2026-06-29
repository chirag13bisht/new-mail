import pandas as pd
import requests
import msal
import time
import random
import os
import re
from datetime import datetime, timedelta

# ==========================================
# 1. CONFIGURATION
# ==========================================
SENDER_EMAIL = "admin@mobifirst.co"

# Azure App Credentials
CLIENT_ID     = "dabde45a-de62-44f1-a15d-1572327a9302"
CLIENT_SECRET = "Hh38Q~ffccw6d6WnCNRlIrnwljKZHn7ucZi13aqO"
TENANT_ID     = "f14f07b0-a186-41e6-a3b4-19cfd15af98c"

# File Paths (Using Railway Volume)
DATABASE_FILE  = "/data/MobiFirst_Master_List.xlsx"
TEMPLATES_FILE = "/data/templates.xlsx"

BATCH_SIZE          = 499       # 499 BCC + 1 TO = 500 max limit
BATCH_DELAY_SECONDS = 600       # 10 minutes between batches
DAILY_LIMIT         = 10000     # 20 batches per day maximum

BOUNCE_KEYWORDS = [
    "undeliverable", "delivery has failed", "couldn't be delivered", 
    "suspects your message is spam", "rejected", "permanent error",
    "wasn't able to deliver"
]

# ==========================================
# 2. GRAPH API & HELPERS
# ==========================================
_token_cache = {"token": None, "expires_at": 0}

def get_access_token() -> str:
    if _token_cache["token"] and time.time() < _token_cache["expires_at"] - 60:
        return _token_cache["token"]

    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" not in result:
        raise Exception(f"❌ Token error: {result.get('error_description')}")

    _token_cache["token"] = result["access_token"]
    _token_cache["expires_at"] = time.time() + result.get("expires_in", 3600)
    return _token_cache["token"]

def send_email_batch_bcc(recipients: list, subject: str, body: str) -> bool:
    try:
        token = get_access_token()
        bcc_list = [{"emailAddress": {"address": email}} for email in recipients]
        
        response = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={
                "message": {
                    "subject": subject,
                    "body": {"contentType": "HTML", "content": body},
                    "toRecipients": [{"emailAddress": {"address": SENDER_EMAIL}}],
                    "bccRecipients": bcc_list
                },
                "saveToSentItems": "true"
            },
            timeout=60
        )
        return response.status_code == 202
    except Exception as e:
        print(f"    API Error: {e}")
        return False

def check_and_delete_bounces(df):
    """Scans inbox for recent automated bounce receipts and deletes them."""
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Sweeping inbox for instant bounces...")
    try:
        token = get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        
        # Check for emails from the last 2 hours (catches instant and slightly delayed bounces)
        two_hours_ago = (datetime.utcnow() - timedelta(hours=2)).strftime('%Y-%m-%dT%H:%M:%SZ')
        query_filter = f"receivedDateTime ge {two_hours_ago}"
        url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/messages?$filter={query_filter}&$select=id,subject,from,bodyPreview&$top=50"
        
        deleted_count = 0
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            messages = response.json().get('value', [])
            email_regex = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            
            for msg in messages:
                msg_id = msg['id']
                subject = msg.get('subject', '').lower()
                body_text = msg.get('bodyPreview', '').lower()
                
                sender_info = msg.get('from', {}).get('emailAddress', {})
                sender_email_address = sender_info.get('address', '').lower()
                sender_name = sender_info.get('name', '').lower()
                
                is_automated = ('postmaster' in sender_email_address or 'mailer-daemon' in sender_email_address or 'microsoft outlook' in sender_name)
                is_bounce_text = any(kw in body_text or kw in subject for kw in BOUNCE_KEYWORDS)
                
                if is_automated and is_bounce_text:
                    # Extract the failed email from the bounce message body
                    found_emails = set(re.findall(email_regex, body_text))
                    for failed_email in found_emails:
                        failed_email = failed_email.lower().strip()
                        if failed_email != SENDER_EMAIL.lower() and failed_email in df['Email Address'].str.lower().values:
                            df.loc[df['Email Address'].str.lower() == failed_email, 'Status'] = 'Bounced'
                            print(f"      -> Marked Bounced: {failed_email}")
                    
                    # Delete the receipt
                    requests.delete(f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/messages/{msg_id}", headers=headers)
                    deleted_count += 1
            
            if deleted_count > 0:
                print(f"    🗑️ Cleaned {deleted_count} bounce receipts from inbox.")
            else:
                print("    ✅ Inbox clear. No bounces found.")
                
    except Exception as e:
        print(f"    Bounce sweep error: {e}")
        
    return df

# ==========================================
# 3. MAIN ENGINE
# ==========================================
def main():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting Batch Engine...")
    total_sent_today = 0

    while True:
        # 1. Load the Database
        try:
            df = pd.read_excel(DATABASE_FILE)
        except Exception as e:
            print(f"❌ Cannot read Master List: {e}")
            break

        # 2. Dynamically Load Templates (allows hot-swapping templates while running)
        try:
            templates_df = pd.read_excel(TEMPLATES_FILE)
            templates = templates_df.to_dict('records')
            if not templates:
                raise ValueError("Template file is empty.")
        except Exception as e:
            print(f"❌ Cannot read templates.xlsx: {e}")
            time.sleep(60)
            continue

        pending_df = df[df['Status'].str.lower() == 'pending']
        if pending_df.empty:
            print("🎉 Campaign Complete! No pending emails left.")
            break

        if total_sent_today >= DAILY_LIMIT:
            print(f"⚠️ Daily limit reached. Pausing for 24 hours...")
            time.sleep(86400)
            total_sent_today = 0
            continue

        # 3. Process Batch
        batch = pending_df.head(BATCH_SIZE)
        recipients = batch['Email Address'].dropna().tolist()
        template = random.choice(templates)

        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Sending batch of {len(recipients)} via BCC...")
        print(f"    Using Template: {template.get('Subject', 'No Subject')}")
        
        raw_body = str(template.get('Body', ''))
        
        # Automatically convert Excel newlines and spaces to HTML
        html_safe_body = raw_body.replace('\n', '<br>').replace('  ', '&nbsp;&nbsp;')
        
        success = send_email_batch_bcc(recipients, template.get('Subject', ''), html_safe_body)

        if success:
            df.loc[batch.index, 'Status'] = 'Sent'
            total_sent_today += len(recipients)
            print("  ✅ Batch Sent Successfully.")
        else:
            df.loc[batch.index, 'Status'] = 'Failed'
            print("  ❌ Batch Failed.")

        # Save initial progress
        df.to_excel(DATABASE_FILE, index=False)

        # 4. Wait briefly for immediate bounces to hit the inbox, then sweep
        print("  ⏳ Waiting 30 seconds for potential immediate bounces...")
        time.sleep(30)
        df = check_and_delete_bounces(df)
        
        # Save again if any bounces were recorded
        df.to_excel(DATABASE_FILE, index=False)
        
        # 5. Sleep for the remainder of the interval
        sleep_time = max(0, BATCH_DELAY_SECONDS - 30)
        print(f"⏳ Waiting {sleep_time // 60} minutes and {sleep_time % 60} seconds for next batch...\n")
        time.sleep(sleep_time)

if __name__ == "__main__":
    main()