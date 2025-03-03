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
from transformers import AutoModel, AutoProcessor, AutoTokenizer
import torch
from dotenv import load_dotenv
import logging
import sys
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Configure logging
def setup_logging(log_level=logging.INFO):
    """Set up logging configuration"""
    logger = logging.getLogger('ms365_scraper')
    logger.setLevel(log_level)
    
    # Create console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    
    # Create file handler
    file_handler = logging.FileHandler('email_scraper.log')
    file_handler.setLevel(log_level)
    
    # Create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    
    # Add handlers to logger
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    
    return logger

# Initialize logger
logger = setup_logging()

# Load environment variables from .env file
load_dotenv()
logger.info("Environment variables loaded from .env file")

# Configure requests with retry capability
def create_requests_session(retries=3, backoff_factor=0.5, status_forcelist=(500, 502, 503, 504)):
    """Create a requests session with retry capabilities"""
    session = requests.Session()
    retry = Retry(
        total=retries,
        read=retries,
        connect=retries,
        backoff_factor=backoff_factor,
        status_forcelist=status_forcelist,
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session

# Function to authenticate and fetch emails
def access_emails_with_delegated_permissions(start_date, end_date, category_filter=None, limit=1000):
    """
    Authenticate with Microsoft Graph API and fetch emails within a date range.
    
    Args:
        start_date (str): Start date in ISO format (YYYY-MM-DD)
        end_date (str): End date in ISO format (YYYY-MM-DD)
        category_filter (str, optional): Category to filter by
        limit (int): Maximum number of emails to retrieve
        
    Returns:
        tuple: (list of emails, headers dict) or ([], None) on failure
    """
    logger.info(f"Starting email fetch process for date range: {start_date} to {end_date}")
    
    # Get Microsoft Graph credentials from environment variables
    client_id = os.environ.get("MS_CLIENT_ID")
    tenant_id = os.environ.get("MS_TENANT_ID")
    
    if not client_id or not tenant_id:
        logger.error("Microsoft Graph credentials not found in .env file")
        raise ValueError("Microsoft Graph credentials not found in .env file. Please check your .env file contains MS_CLIENT_ID and MS_TENANT_ID variables.")
    
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    logger.info(f"Authenticating with Microsoft Graph using authority: {authority}")
    
    app = msal.PublicClientApplication(client_id, authority=authority)
    logger.info("Initiating device flow for authentication")
    device_flow = app.initiate_device_flow(scopes=["Mail.Read"])

    print("Please visit the following URL and enter the code:", device_flow["verification_uri"])
    print("Code:", device_flow["user_code"])
    logger.info(f"Authentication URL: {device_flow['verification_uri']}, Code: {device_flow['user_code']}")

    result = app.acquire_token_by_device_flow(device_flow)

    if "access_token" not in result:
        logger.error(f"Failed to acquire token: {result}")
        print("Failed to acquire token:", result)
        return [], None

    # Ensure that the token is a valid JWT token
    access_token = result['access_token']
    logger.info("Token retrieved successfully")
    print("Token retrieved successfully. Proceeding with API requests.")

    headers = {"Authorization": f"Bearer {access_token}"}

    # Ensure date format is correct (YYYY-MM-DD)
    try:
        # Parse and reformat dates to ensure they have proper formatting
        start_date_obj = datetime.strptime(start_date, "%Y-%m-%d")
        end_date_obj = datetime.strptime(end_date, "%Y-%m-%d")
        
        formatted_start_date = start_date_obj.strftime("%Y-%m-%d")
        formatted_end_date = end_date_obj.strftime("%Y-%m-%d")
        
        # Build the filter query with properly formatted dates
        filter_query = f"receivedDateTime ge {formatted_start_date}T00:00:00Z and receivedDateTime le {formatted_end_date}T23:59:59Z"
        logger.info(f"Date filter query: {filter_query}")
    except ValueError as e:
        logger.error(f"Error parsing dates: {e}")
        print(f"Error parsing dates: {e}")
        print("Please ensure dates are in YYYY-MM-DD format.")
        return [], None

    if category_filter:
        filter_query += f" and categories/any(c:c eq '{category_filter}')"
        logger.info(f"Added category filter: {category_filter}")

    params = {"$top": limit, "$filter": filter_query}
    logger.info(f"Request parameters: {params}")

    # Use session with retry capability
    session = create_requests_session()
    
    try:
        logger.info("Sending request to Microsoft Graph API")
        response = session.get("https://graph.microsoft.com/v1.0/me/messages", headers=headers, params=params)
        
        if response.status_code == 200:
            emails = response.json().get("value", [])
            logger.info(f"Retrieved {len(emails)} emails")
            return emails, headers  # Return emails and headers for attachments request
        elif response.status_code == 504:
            logger.error(f"Gateway timeout error (504). This could be due to a large volume of emails or temporary server issues.")
            print("Gateway timeout error (504). Trying with a shorter date range may help.")
            return [], None
        else:
            logger.error(f"Error fetching emails: {response.status_code}, {response.text}")
            print(f"Error fetching emails: {response.status_code}, {response.text}")
            return [], None
    except requests.exceptions.RequestException as e:
        logger.error(f"Request exception during email fetch: {e}")
        print(f"Network error: {e}")
        return [], None

# Function to extract text from PDF/Image using primary OCR method
def extract_text_with_primary_ocr(file_path):
    """Extract text from a PDF or image using PyMuPDF and pytesseract"""
    logger.info(f"Attempting primary OCR extraction on {file_path}")
    text = ""
    try:
        # Try to extract text from PDF
        doc = fitz.open(file_path)
        for page in doc:
            text += page.get_text("text")  # Extract text from PDF pages
        logger.debug(f"PDF text extraction yielded {len(text)} characters")
    except Exception as e:
        logger.error(f"Error extracting text from PDF: {e}")
        text = None

    # If no text found, try OCR on image
    if not text or text.strip() == "":
        logger.info("No text found in PDF, trying primary OCR on image")
        try:
            img = Image.open(file_path)  # This could be an image, try OCR
            text = pytesseract.image_to_string(img)
            logger.debug(f"Image OCR extraction yielded {len(text)} characters")
        except Exception as e:
            logger.error(f"Error during primary OCR: {e}")
            text = None
            
    return text

# Function to extract text using the advanced GOT-OCR model
def extract_text_with_got_ocr(file_path):
    """Extract text using the GOT-OCR2_0 model from HuggingFace"""
    logger.info(f"Attempting advanced OCR with GOT-OCR2_0 model on {file_path}")
    try:
        # Initialize the model and processor
        device = "cuda" if torch.cuda.is_available() else "cpu"
        logger.info(f"Using device: {device} for GOT-OCR")
        
        logger.info("Loading GOT-OCR model (this may take some time on first run)")
        # Use the correct model ID and parameters
        tokenizer = AutoTokenizer.from_pretrained(
            "stepfun-ai/GOT-OCR2_0", 
            trust_remote_code=True,
        )
        
        model = AutoModel.from_pretrained(
            "stepfun-ai/GOT-OCR2_0", 
            trust_remote_code=True,
            low_cpu_mem_usage=True,
            device_map=device,
            pad_token_id=tokenizer.eos_token_id
        )
        
        # Check if the file is a PDF or an image
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.pdf':
            logger.info(f"Processing PDF file with GOT-OCR: {file_path}")
            # Use PyMuPDF to convert PDF to images
            full_text = []
            
            # Open the PDF
            pdf_document = fitz.open(file_path)
            total_pages = len(pdf_document)  # Store page count before closing
            
            # Process each page
            for page_num in range(total_pages):
                logger.info(f"Processing page {page_num+1}/{total_pages} of PDF")
                page = pdf_document.load_page(page_num)
                
                # Convert page to an image
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                
                # Save the image to a temporary file
                temp_img_path = f"temp_page_{page_num}.png"
                pix.save(temp_img_path)
                
                # Process the image with GOT-OCR
                try:
                    # Process the image with GOT-OCR using the chat method as per documentation
                    page_text = model.chat(tokenizer, temp_img_path, ocr_type='ocr')
                    if page_text:
                        full_text.append(page_text)
                    logger.info(f"Extracted {len(page_text) if page_text else 0} characters from page {page_num+1}")
                except Exception as e:
                    logger.warning(f"Error processing PDF page {page_num+1}: {e}")
                
                # Remove temporary image file
                try:
                    os.remove(temp_img_path)
                except:
                    pass
            
            # Close the PDF
            pdf_document.close()
            
            # Combine text from all pages
            text = "\n".join(full_text)
            logger.info(f"GOT-OCR extraction complete for PDF, yielded {len(text)} characters from {total_pages} pages")
        else:
            # Process a regular image file
            logger.info(f"Processing image file with GOT-OCR: {file_path}")
            # Use the chat method as documented in the model page
            text = model.chat(tokenizer, file_path, ocr_type='ocr')
            logger.info(f"GOT-OCR extraction complete, yielded {len(text) if text else 0} characters")
        
        return text
    except Exception as e:
        logger.error(f"Error during advanced OCR: {e}")
        traceback.print_exc()
        return None

# Function to process email attachments (PDF/Image)
def extract_text_from_attachment(attachment):
    """Extract text from an email attachment using multiple OCR methods if needed"""
    attachment_name = attachment.get('name', 'Unnamed')
    logger.info(f"Extracting text from attachment: {attachment_name}")

    if 'contentBytes' in attachment:
        content = base64.b64decode(attachment['contentBytes'])
        temp_file_path = "temp_attachment.pdf"
        with open(temp_file_path, "wb") as f:
            f.write(content)
        logger.debug(f"Attachment content saved to {temp_file_path}")
    else:
        logger.error("'contentBytes' key not found in attachment")
        print("Error: 'contentBytes' key not found in attachment. Skipping this attachment.")
        return "Failed to extract attachment content."

    # Process the PDF or image for OCR
    text = extract_text_with_primary_ocr(temp_file_path)

    # If primary OCR fails, try advanced OCR
    if not text or text.strip() == "" or "Failed to extract" in text:
        logger.info("Primary OCR failed, trying advanced OCR...")
        text = extract_text_with_got_ocr(temp_file_path)

    if text and text.strip() != "":
        logger.info("Text extracted from attachment successfully")
        return text
    else:
        logger.warning("Failed to extract text from attachment with both OCR methods")
        return "Failed to extract text from attachment."

# Function to fetch attachments for each email
def fetch_attachments(email_id, headers):
    """Fetch attachments for a specific email"""
    logger.info(f"Fetching attachments for email: {email_id}")
    
    attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/attachments"
    
    # Use session with retry capability
    session = create_requests_session()
    
    try:
        attachments_response = session.get(attachments_url, headers=headers)
        
        if attachments_response.status_code == 200:
            attachments = attachments_response.json().get("value", [])
            logger.info(f"Retrieved {len(attachments)} attachments")
            return attachments
        else:
            logger.error(f"Error fetching attachments: {attachments_response.status_code}, {attachments_response.text}")
            print(f"Error fetching attachments for {email_id}: {attachments_response.status_code}, {attachments_response.text}")
            return []
    except requests.exceptions.RequestException as e:
        logger.error(f"Request exception during attachment fetch: {e}")
        print(f"Network error while fetching attachments: {e}")
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
    logger.info(f"Starting email processing for date range: {start_date} to {end_date}")
    logger.info(f"Output file: {output_file}, Batch size: {batch_size}, Category filter: {category_filter}, Limit: {limit}")
    
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
        logger.warning("No emails found or error occurred during fetching")
        print("No emails found.")
        # Return an empty summary dictionary instead of None
        return {
            "total_emails": 0,
            "processed_emails": 0,
            "with_attachments": 0,
            "output_file": output_file
        }

    logger.info(f"Retrieved {len(emails)} emails. Processing...")
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
                logger.error(f"Invalid datetime format for email: {date_str}")
                formatted_date = date_str
            
            logger.info(f"Processing email {processed_count}/{len(emails)}: {subject} from {sender}")
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
                logger.info(f"Processing {len(attachments)} attachments for email")
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
            else:
                logger.info("No attachments found for this email")
            
            # Add email info to the dataset
            email_data.append(email_info)
            
            # Save progress at regular intervals
            if processed_count % batch_size == 0:
                batch_count += 1
                temp_filename = f"{os.path.splitext(output_file)[0]}_batch{batch_count}{os.path.splitext(output_file)[1]}"
                logger.info(f"Saving progress batch {batch_count} to {temp_filename}")
                pd.DataFrame(email_data).to_excel(temp_filename, index=False)
                print(f"Saved progress: {processed_count}/{len(emails)} emails processed. Batch saved to {temp_filename}")
        
        except Exception as e:
            logger.error(f"Error processing email: {e}")
            traceback.print_exc()
            continue
    
    # Save final results
    logger.info(f"Processing complete. Saving final results to {output_file}")
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
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    
    args = parser.parse_args()
    
    # Set up logging level based on debug flag
    if args.debug:
        logger.setLevel(logging.DEBUG)
        logger.info("Debug logging enabled")
    
    try:
        # Validate date formats
        datetime.strptime(args.start_date, "%Y-%m-%d")
        datetime.strptime(args.end_date, "%Y-%m-%d")
        
        logger.info("Starting email processing with command line arguments")
        summary = process_emails(
            start_date=args.start_date,
            end_date=args.end_date,
            output_file=args.output,
            batch_size=args.batch_size,
            category_filter=args.category,
            limit=args.limit
        )
        
        logger.info("Email processing completed successfully")
        print("\nSummary:")
        print(f"Total emails retrieved: {summary['total_emails']}")
        print(f"Emails processed: {summary['processed_emails']}")
        print(f"Emails with attachments: {summary['with_attachments']}")
        print(f"Output saved to: {summary['output_file']}")
    
    except ValueError as e:
        logger.error(f"Error: {e}")
        print(f"Error: {e}")
        print("Please ensure dates are in YYYY-MM-DD format.")
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        print(f"Unexpected error: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()