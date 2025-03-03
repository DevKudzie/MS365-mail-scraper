import msal
import requests
import os
import pandas as pd
from PIL import Image
import pytesseract
import fitz
import base64
import re
import pytz
from datetime import datetime
import argparse
import time
import traceback
from transformers import AutoModelForVision2Seq, AutoProcessor
import torch
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Function to authenticate and fetch emails
def access_emails_with_delegated_permissions(start_date, end_date, category_filter=None, limit=1000):
    # Get Microsoft Graph credentials from environment variables
    client_id = os.environ.get("MS_CLIENT_ID")
    tenant_id = os.environ.get("MS_TENANT_ID")
    
    if not client_id or not tenant_id:
        raise ValueError("Microsoft Graph credentials not found in .env file. Please check your .env file contains MS_CLIENT_ID and MS_TENANT_ID variables.")
    
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    
    app = msal.PublicClientApplication(client_id, authority=authority)
    device_flow = app.initiate_device_flow(scopes=["Mail.Read"])

    print("Please visit the following URL and enter the code:", device_flow["verification_uri"])
    print("Code:", device_flow["user_code"])

    result = app.acquire_token_by_device_flow(device_flow)

    if "access_token" not in result:
        print("Failed to acquire token:", result)
        return [], None

    # Ensure that the token is a valid JWT token
    access_token = result['access_token']
    print("Token retrieved successfully. Proceeding with API requests.")

    headers = {"Authorization": f"Bearer {access_token}"}

    # Build the filter query
    filter_query = f"receivedDateTime ge {start_date}T00:00:00Z and receivedDateTime le {end_date}T23:59:59Z"
    if category_filter:
        filter_query += f" and categories/any(c:c eq '{category_filter}')"

    params = {"$top": limit, "$filter": filter_query}

    response = requests.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers, params=params)

    if response.status_code == 200:
        emails = response.json().get("value", [])
        return emails, headers  # Return emails and headers for attachments request
    else:
        print(f"Error fetching emails: {response.status_code}, {response.text}")
        return [], None

# Function to extract text from PDF/Image using primary OCR method
def extract_text_with_primary_ocr(file_path):
    text = ""
    try:
        # Try to extract text from PDF
        doc = fitz.open(file_path)
        for page in doc:
            text += page.get_text("text")  # Extract text from PDF pages
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")
        text = None

    # If no text found, try OCR on image
    if not text or text.strip() == "":
        print("No text found in PDF, trying primary OCR on image...")
        try:
            img = Image.open(file_path)  # This could be an image, try OCR
            text = pytesseract.image_to_string(img)
        except Exception as e:
            print(f"Error during primary OCR: {e}")
            text = None
            
    return text

# Function to extract text using the advanced GOT-OCR model
def extract_text_with_got_ocr(file_path):
    print("Attempting advanced OCR with GOT-OCR-2.0 model...")
    try:
        # Initialize the model and processor
        device = "cuda" if torch.cuda.is_available() else "cpu"
        model = AutoModelForVision2Seq.from_pretrained("stepfun-ai/GOT-OCR-2.0-hf").to(device)
        processor = AutoProcessor.from_pretrained("stepfun-ai/GOT-OCR-2.0-hf")
        
        # Load and process the image
        image = Image.open(file_path)
        pixel_values = processor(images=image, return_tensors="pt").pixel_values.to(device)
        
        # Generate OCR results
        generated_ids = model.generate(
            pixel_values=pixel_values,
            max_length=512,
        )
        text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]
        return text
    except Exception as e:
        print(f"Error during advanced OCR: {e}")
        traceback.print_exc()
        return None

# Function to process email attachments (PDF/Image)
def extract_text_from_attachment(attachment):
    print(f"Extracting text from attachment: {attachment['name']}")

    if 'contentBytes' in attachment:
        content = base64.b64decode(attachment['contentBytes'])
        temp_file_path = "temp_attachment.pdf"
        with open(temp_file_path, "wb") as f:
            f.write(content)
    else:
        print("Error: 'contentBytes' key not found in attachment. Skipping this attachment.")
        return "Failed to extract attachment content."

    # Process the PDF or image for OCR
    text = extract_text_with_primary_ocr(temp_file_path)

    # If primary OCR fails, try advanced OCR
    if not text or text.strip() == "" or "Failed to extract" in text:
        print("Primary OCR failed, trying advanced OCR...")
        text = extract_text_with_got_ocr(temp_file_path)

    if text and text.strip() != "":
        print("Text extracted from attachment.")
        return text
    else:
        print("Failed to extract text from attachment with both OCR methods.")
        return "Failed to extract text from attachment."

# Function to fetch attachments for each email
def fetch_attachments(email_id, headers):
    attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/attachments"
    attachments_response = requests.get(attachments_url, headers=headers)

    if attachments_response.status_code == 200:
        return attachments_response.json().get("value", [])
    else:
        print(f"Error fetching attachments for {email_id}: {attachments_response.status_code}, {attachments_response.text}")
        return []

def process_emails(start_date, end_date, output_file="email_data.xlsx", batch_size=200, category_filter=None, limit=1000):
    """
    Process emails within a specified date range and save to Excel.
    
    Parameters:
    - start_date (str): Start date in ISO format (e.g., "2023-01-01").
    - end_date (str): End date in ISO format (e.g., "2023-01-31").
    - output_file (str): Name of the Excel file to save results.
    - batch_size (int): Number of records to process before saving progress.
    - category_filter (str, optional): Category to filter emails by.
    - limit (int): Maximum number of emails to retrieve.
    """
    # Set the timezone to UTC
    utc_tz = pytz.timezone('UTC')
    
    # Fetch emails within the specific date range
    emails, headers = access_emails_with_delegated_permissions(
        start_date=start_date,
        end_date=end_date,
        category_filter=category_filter,
        limit=limit
    )

    if not emails:
        print("No emails found.")
        return

    print(f"Retrieved {len(emails)} emails. Processing...")
    
    # Prepare data structure for storing email information
    email_data = []
    
    # Track progress
    processed_count = 0
    batch_count = 0
    
    # Process each email
    for email in emails:
        try:
            processed_count += 1
            
            sender = email.get('from', {}).get('emailAddress', {}).get('address', 'Not found')
            sender_name = email.get('from', {}).get('emailAddress', {}).get('name', 'Not found')
            subject = email.get('subject', 'Not found')
            date_str = email.get('receivedDateTime', 'Not found')
            
            # Convert date string to datetime object
            try:
                if date_str.endswith("Z"):
                    date_str = date_str.replace("Z", "+00:00")
                email_datetime = datetime.fromisoformat(date_str)
                formatted_date = email_datetime.strftime('%Y-%m-%d %H:%M:%S')
            except ValueError:
                print(f"Invalid datetime format for email: {date_str}")
                formatted_date = date_str
            
            print(f"Processing email {processed_count}/{len(emails)}")
            print(f"From: {sender} ({sender_name})")
            print(f"Subject: {subject}")
            print(f"Date: {formatted_date}")
            
            # Get email body
            email_body = email.get('body', {}).get('content', 'No body content')
            email_body_clean = re.sub(r'<[^>]+>', '', email_body)  # Remove HTML tags
            
            # Prepare email data dictionary
            email_info = {
                'Sender Email': sender,
                'Sender Name': sender_name,
                'Subject': subject,
                'Date': formatted_date,
                'Body': email_body_clean,
                'Attachment Content': '',
                'Has Attachments': 'No',
                'Attachment Names': '',
                'Attachment Count': 0
            }
            
            # Fetch attachments for this email
            email_id = email.get('id')
            attachments = fetch_attachments(email_id, headers)
            
            if attachments:
                email_info['Has Attachments'] = 'Yes'
                email_info['Attachment Count'] = len(attachments)
                attachment_names = []
                attachment_texts = []
                
                for attachment in attachments:
                    attachment_name = attachment.get('name', 'Unnamed')
                    attachment_names.append(attachment_name)
                    
                    # Extract text from the attachment
                    attachment_text = extract_text_from_attachment(attachment)
                    if attachment_text:
                        attachment_texts.append(f"[{attachment_name}]: {attachment_text}")
                
                email_info['Attachment Names'] = ', '.join(attachment_names)
                email_info['Attachment Content'] = '\n\n'.join(attachment_texts)
            
            # Add email info to the dataset
            email_data.append(email_info)
            
            # Save progress at regular intervals
            if processed_count % batch_size == 0:
                batch_count += 1
                temp_filename = f"{os.path.splitext(output_file)[0]}_batch{batch_count}{os.path.splitext(output_file)[1]}"
                pd.DataFrame(email_data).to_excel(temp_filename, index=False)
                print(f"Saved progress: {processed_count}/{len(emails)} emails processed. Batch saved to {temp_filename}")
        
        except Exception as e:
            print(f"Error processing email: {e}")
            traceback.print_exc()
            continue
    
    # Save final results
    final_df = pd.DataFrame(email_data)
    final_df.to_excel(output_file, index=False)
    print(f"Processing complete. {processed_count} emails processed. Data saved to {output_file}")
    
    # Return summary info
    return {
        "total_emails": len(emails),
        "processed_emails": processed_count,
        "with_attachments": sum(1 for email in email_data if email['Has Attachments'] == 'Yes'),
        "output_file": output_file
    }

def main():
    parser = argparse.ArgumentParser(description='Microsoft 365 Email Scraper')
    parser.add_argument('--start-date', required=True, help='Start date in ISO format (YYYY-MM-DD)')
    parser.add_argument('--end-date', required=True, help='End date in ISO format (YYYY-MM-DD)')
    parser.add_argument('--category', help='Optional category filter')
    parser.add_argument('--output', default='email_data.xlsx', help='Output Excel file name')
    parser.add_argument('--batch-size', type=int, default=200, help='Number of emails to process before saving progress')
    parser.add_argument('--limit', type=int, default=1000, help='Maximum number of emails to retrieve')
    
    args = parser.parse_args()
    
    summary = process_emails(
        start_date=args.start_date,
        end_date=args.end_date,
        output_file=args.output,
        batch_size=args.batch_size,
        category_filter=args.category,
        limit=args.limit
    )
    
    print("\nSummary:")
    print(f"Total emails retrieved: {summary['total_emails']}")
    print(f"Emails processed: {summary['processed_emails']}")
    print(f"Emails with attachments: {summary['with_attachments']}")
    print(f"Output saved to: {summary['output_file']}")

if __name__ == "__main__":
    main()