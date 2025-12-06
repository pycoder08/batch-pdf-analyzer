# gemini-batch-pdfs
## Description
This script automates the analysis of student assignments by leveraging Google Sheets, Google Drive, and Google's Gemini API. It reads assignment links directly from a spreadsheet, fetches the student's PDF from Google Drive, and sends it to Gemini for analysis or processing based on your custom instructions.

### Features
- **Google Sheets Integration:** Reads student names and file links directly from your spreadsheet
- **Intelligent Name Extraction:** Features a parsing algorithm to identify student names. It automatically isolates names by detecting and stripping common file prefixes (like dates or assignment titles in parentheses) and truncating trailing metadata like Student IDs or course codes.
- **Smart State Awareness:** Automatically skips files that have already been uploaded or analyzed. You can stop and restart the script at any time without losing progress or duplicating work.
- **Cost-Effective:** Uses a synchronous approach designed to stay within the free tier limits of the Gemini API.
- **Structured Output:** Saves all AI responses to a local JSON file (`responses.json`) for easy data processing.
- **PDF Report Generation:** Optionally converts the AI's text analysis into clean, formatted PDF reports for each student.

## Requirements
- Python 3.8+
- A Google Cloud Project with the following APIs enabled:
    - Google Drive API
    - Google Sheets API

## Setup

1. **Repository Setup**
    - Clone this repository and install required python libraries.
    ```
    pip install -r requirements.txt 
    ```

2. **Get a Gemini API Key**
    - Next, obtain a Google Gemini API key for free [here](https://aistudio.google.com/app/apikey) (note that your usage will be limited on the free tier as per Google's policy, but it should be more than enough for personal use).
    - You will need this for your configuration in step 4

3. **Setup Google Cloud**
    - Select Your Google Cloud Project
        - Go to the [Google Cloud Console](console.cloud.google.com) and select the project you used for your API key (or make a new one).
    - Enable the Google Drive API: In the top search bar, search for and enable the "Google Drive API".
        -  Configure the OAuth Consent Screen:
        -  Navigate to APIs & Services > OAuth consent screen.
        -  Set User Type to External and click Create.
        -  Fill in the required fields (App name, User support email, Developer contact email).
        -  Click Save and Continue through the Scopes and Optional Info pages. You don't need to add anything.
    -  Create OAuth Credentials:
        - Navigate to APIs & Services > Credentials.
        - Click + CREATE CREDENTIALS and select OAuth client ID.
        - Set the Application type to Desktop app.
        - Click Create.
    - Download and Rename:
        - In the pop-up, click DOWNLOAD JSON.
        - Rename the downloaded file to exactly credentials.json.
        - Place it in the root folder of this project.
    - Add Your Test User:
        - Go back to the OAuth consent screen page.
        - Under "Test users," click + ADD USERS.
        - Enter the Google email address you will use to run the script and click Save.

4. **Configure Enviornment Variables**
    - Create a '.env' file in the root directory as specified in `.env.example`:
    ```
    # --- Google API Configuration ---
    # The ID of the Google Drive folder to read files from
    FOLDER_ID=
    # The ID of the Google Sheet to read links from
    SPREADSHEET_ID=
    # The specific sheet and range to read (e.g., 'Sheet1!A1:A')
    SHEET_RANGE=
    
    # --- Gemini API Configuration ---
    GEMINI_MODEL=gemini-1.5-flash
    # The main prompt for analyzing student feedback
    ANALYSIS_PROMPT_FILE=prompts/analysis_prompt.txt
    # The prompt for transcribing (OCR) handwritten PDFs
    OCR_PROMPT_FILE=prompts/ocr_prompt.txt
    
    # --- Output Configuration ---
    # Set to "True" to convert markdown analyses to PDFs
    CONVERT_TO_PDF=True
    # The local folder to save PDF outputs to
    OUTPUT_FOLDER=output
    ```
    - *Tip: The SPREADSHEET_ID is the long string of letters and numbers in your Google Sheet URL.*

5. **Configure Prompts**
    - Create a file in the root directory called `prompts`
    - Inside that folder, create a text file named `analysis_prompt.txt`
    - Put the instructions you want the AI to follow in that file



## Usage

1. **Populate the Spreadsheet**
   Ensure the column specified in your `.env` file (e.g., `C2:C`) contains the **Google Drive links** to the student PDF files. The script will automatically extract the student's name from the file name itself, so you do not need to enter names manually in the sheet.

2. **Refine Your Prompt**
   To change how the AI analyzes the files (e.g., changing from grading to transcription), simply edit the text inside `prompts/analysis_prompt.txt`.

3. **Run the Script**
   Execute the script from your terminal:
   ```
   python gemini_batch.py
   ```
