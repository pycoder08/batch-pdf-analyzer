# gemini-batch-pdfs
## Description
This script automates analysis of PDF documents by leveraging Google's Generative Language and Google Drive APIs. It accesses all PDFs within a specified Google Drive folder, processes each one with a custom prompt using the Gemini API, and saves the structured analysis locally.

### Features

- Uses synchronous approach to avoid API rate limits, meaning it is usable without any charge
- Saves all of Gemini's responses to a single JSON file for easy access
- Optionally converts all responses to individual PDFs saved locally with markdown support

## Requirements
- Python 3.8+
- A Google Cloud account

## Setup

To begin, clone the repository and install required python libraries. 

Next, you will need a Google Gemini API key, which you can obtain for free [here](https://aistudio.google.com/app/apikey) (note that your usage will be limited on the free tier as per Google's policy, but it should be more than enough for personal use).

Once you have your key, save it as an envoirnment variable named 'GEMINI_API_KEY' exactly by creating a .env file and adding the following command:

`GEMINI_API_KEY:"your key here"`


Select Your Google Cloud Project: Go to the [Google Cloud Console](console.cloud.google.com) and select the project you used for your API key.

Enable the Google Drive API: In the top search bar, search for and enable the "Google Drive API".

1. Configure the OAuth Consent Screen:

    Navigate to APIs & Services > OAuth consent screen.

    Set User Type to External and click Create.

    Fill in the required fields (App name, User support email, Developer contact email).

    Click Save and Continue through the Scopes and Optional Info pages. You don't need to add anything.

2. Create OAuth Credentials:

    Navigate to APIs & Services > Credentials.

    Click + CREATE CREDENTIALS and select OAuth client ID.

    Set the Application type to Desktop app.

    Click Create.

3. Download and Rename:

    In the pop-up, click DOWNLOAD JSON.

    Rename the downloaded file to exactly credentials.json.

    Place it in the root folder of this project.

4. Add Your Test User:

    Go back to the OAuth consent screen page.

    Under "Test users," click + ADD USERS.

    Enter the Google email address you will use to run the script and click Save.

## Usage

In the script, configure the constants to adjust functionality: 
- `FOLDER_ID` constant: Set it to the ID of the folder you want to analyze (the string of numbers and letters at the end of a google drive link).
- `PROMPT`: Edit the prompt that will be passed to gemini with every file
- `GEMINI_MODEL`: Controls what version of Gemini to use
- `CONVERT_TO_PDF`: Controls whether to convert Gemini's responses to markdown-enabled PDFs or not
- `OUTPUT_FOLDER`: Controls where to save the output PDFs if you decide to make them

Finally, run the script.
Gemini will proceed to process the files one by one, and once finished, save each response into a file called `responses.json`. The first time you run the file you will have to authenticate your account. As it runs it will update you on its progress.
