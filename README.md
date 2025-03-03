# Microsoft 365 Email Scraper

A Python script for scraping emails from Microsoft 365 accounts and saving the content (including attachments) to Excel.

## Features

- Authenticate with Microsoft 365 using the Microsoft Authentication Library (MSAL)
- Fetch emails within a specified date range
- Optional category filtering
- Extract email metadata (sender, subject, date)
- Extract email body content
- Extract text from attachments (PDF, images) using OCR
- Save results to Excel
- Periodic saving of progress to prevent data loss
- Advanced OCR capabilities using GOT-OCR-2.0 model as a fallback

## Requirements

- Python 3.6+
- Required Python packages (install via `pip install -r requirements.txt`):
  - msal
  - requests
  - pandas
  - Pillow
  - pytesseract
  - PyMuPDF (fitz)
  - transformers
  - torch
  - openpyxl
  - pytz
  - python-dotenv

## Setup

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/ms365-mail-scraper.git
   cd ms365-mail-scraper
   ```

2. Install required packages:
   ```
   pip install -r requirements.txt
   ```

3. Install Tesseract OCR:
   - For Windows: Download and install from [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)
   - For macOS: `brew install tesseract`
   - For Ubuntu: `sudo apt install tesseract-ocr`

4. Set up Microsoft Graph credentials:
   - Register an application in the Azure Portal
   - Set the required permissions (Mail.Read)
   - Create a `.env` file in the project root with your credentials:
     ```
     MS_CLIENT_ID=your_client_id
     MS_TENANT_ID=your_tenant_id
     ```

## Usage

Run the script with the required date range parameters:

```
python app.py --start-date 2023-01-01 --end-date 2023-01-31
```

### Command-Line Arguments

- `--start-date`: Start date in ISO format (YYYY-MM-DD) (required)
- `--end-date`: End date in ISO format (YYYY-MM-DD) (required)
- `--category`: Optional category filter (e.g., "Important")
- `--output`: Output Excel file name (default: email_data.xlsx)
- `--batch-size`: Number of emails to process before saving progress (default: 200)
- `--limit`: Maximum number of emails to retrieve (default: 1000)

### Examples

Fetch emails from January 2023 with a category filter:
```
python app.py --start-date 2023-01-01 --end-date 2023-01-31 --category "Important"
```

Fetch emails from Q1 2023 and customize output file:
```
python app.py --start-date 2023-01-01 --end-date 2023-03-31 --output q1_emails.xlsx
```

Fetch more emails with different batch save frequency:
```
python app.py --start-date 2023-01-01 --end-date 2023-12-31 --limit 5000 --batch-size 500
```

## Output Format

The script generates an Excel file with the following columns:

- `Sender Email`: Email address of the sender
- `Sender Name`: Name of the sender
- `Subject`: Email subject
- `Date`: Date and time when the email was received
- `Body`: Email body content (HTML tags removed)
- `Has Attachments`: Yes/No indicating if the email has attachments
- `Attachment Names`: Comma-separated list of attachment file names
- `Attachment Count`: Number of attachments
- `Attachment Content`: Text content extracted from attachments

## OCR Capabilities

The script uses two OCR methods to extract text from attachments:

1. **Primary OCR**: Uses PyMuPDF for PDF text extraction and pytesseract for image OCR
2. **Advanced OCR**: Uses the GOT-OCR-2.0 model from HuggingFace as a fallback when primary OCR fails

## License

MIT License

## Troubleshooting

- **Authentication Issues**: Make sure your client ID and tenant ID are correctly set in the `.env` file.
- **OCR Problems**: Ensure Tesseract is properly installed and accessible in your PATH.
- **Memory Issues**: If processing large volumes of emails, consider reducing the `--limit` parameter or increasing your system's available memory.
- **Hugging Face Model Download**: The first time the script runs, it will download the GOT-OCR model, which might take some time depending on your internet connection.