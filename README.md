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
- Advanced OCR capabilities using GOT-OCR2_0 model as a fallback
- Comprehensive logging system for tracking progress and troubleshooting
- Auto-retry mechanism for handling network issues and API timeouts

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
  - verovio==4.3.1
  - tiktoken==0.6.0
  - accelerate

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

## Running in Google Colab

You can also run this script in Google Colab, which provides free GPU acceleration that can be useful for the advanced OCR functionality.

### Using the Included Colab Notebook

The easiest way to get started with Colab is to use the included notebook:

1. Download the `ms365_email_scraper.ipynb` file from this repository
2. Upload it to [Google Colab](https://colab.research.google.com)
3. Follow the step-by-step instructions in the notebook

The notebook contains all necessary setup steps and examples for running the script in Colab.

### Manual Setup in Colab

If you prefer to set up manually, follow these steps:

1. Create a new Colab notebook at [colab.research.google.com](https://colab.research.google.com)

2. Install the required dependencies:
   ```python
   !apt-get update
   !apt-get install -y tesseract-ocr
   !pip install msal requests pandas Pillow pytesseract PyMuPDF transformers torch openpyxl pytz python-dotenv verovio==4.3.1 tiktoken==0.6.0 accelerate
   ```

3. Upload the script or clone the repository:
   ```python
   !git clone https://github.com/yourusername/ms365-mail-scraper.git
   %cd ms365-mail-scraper
   ```

4. Create a .env file directly in Colab:
   ```python
   %%writefile .env
   MS_CLIENT_ID="your_client_id"
   MS_TENANT_ID="your_tenant_id"
   ```

5. Optional: Connect to Google Drive to save output files permanently:
   ```python
   from google.colab import drive
   drive.mount('/content/drive')
   ```

6. Run the script with arguments:
   ```python
   !python app.py --start-date 2023-01-01 --end-date 2023-01-31 --output "/content/drive/My Drive/email_data.xlsx"
   ```

### Important Notes for Colab

- **Date Formatting**: Always use the YYYY-MM-DD format for dates (e.g., 2023-01-01 instead of 2023-1-1). The script now validates date formats to prevent API errors.
  
- **First Run**: The first time you run the script, it will download the GOT-OCR model, which might take several minutes.

- **Session Timeouts**: Google Colab sessions have time limits. For processing large volumes of emails, use the `--batch-size` parameter to save progress regularly.

### GPU Acceleration

To use GPU acceleration for the OCR model in Colab:

1. Go to Runtime > Change runtime type
2. Select "GPU" under Hardware accelerator
3. Click "Save"

This will make the advanced OCR functionality run much faster.

### Handling Authentication

When the script runs in Colab, it will display a URL and code for Microsoft authentication:

1. Copy the URL and open it in a new browser tab
2. Enter the code displayed in the Colab output
3. Follow the authentication steps for your Microsoft account
4. Return to Colab to continue the script execution

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
- `--debug`: Enable debug level logging (more verbose output)

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

Enable debug logging for troubleshooting:
```
python app.py --start-date 2023-01-01 --end-date 2023-01-31 --debug
```

## Logging

The script includes a comprehensive logging system that records information at different levels:

- **INFO**: Normal operation information (default level)
- **WARNING**: Potential issues that don't stop execution
- **ERROR**: Errors that affect specific operations
- **DEBUG**: Detailed information for troubleshooting (enabled with `--debug` flag)

Logs are written to both the console and a file named `email_scraper.log` in the project directory.

## Handling Timeouts and Errors

The script includes automatic retry mechanisms for handling common API issues:

- **504 Gateway Timeout**: If you encounter this error, the script will provide a specific message. Try reducing your date range or limit to process fewer emails at once.
- **Network Issues**: The script will automatically retry failed requests up to 3 times with exponential backoff.
- **Large Volume of Emails**: For mailboxes with thousands of emails, use smaller date ranges and the `--batch-size` parameter to save progress regularly.

## OCR Capabilities

The script uses two OCR methods to extract text from attachments:

1. **Primary OCR**: Uses PyMuPDF for PDF text extraction and pytesseract for image OCR
2. **Advanced OCR**: Uses the GOT-OCR2_0 model from HuggingFace as a fallback when primary OCR fails

### Advanced OCR (GOT-OCR2_0) Requirements

The GOT-OCR2_0 model requires additional dependencies:
- transformers (version 4.37.2 or newer)
- verovio (version 4.3.1)
- tiktoken (version 0.6.0)
- accelerate (version 0.28.0)

These will be installed automatically by the setup script or Colab notebook.

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

## License

MIT License

## Troubleshooting

- **Authentication Issues**: Make sure your client ID and tenant ID are correctly set in the `.env` file.
- **OCR Problems**: Ensure Tesseract is properly installed and accessible in your PATH.
- **Memory Issues**: If processing large volumes of emails, consider reducing the `--limit` parameter or increasing your system's available memory.
- **Hugging Face Model Download**: The first time the script runs, it will download the GOT-OCR model, which might take some time depending on your internet connection.
- **Advanced OCR Issues**: If you encounter errors with the GOT-OCR2_0 model, try updating your transformers package to the latest version: `pip install --upgrade transformers`