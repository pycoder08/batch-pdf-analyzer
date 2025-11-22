import time
import json
import os
import tempfile
import re
import sys
from typing import Optional, List, Dict, Any

from dotenv import load_dotenv
import markdown
from fpdf import FPDF
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from google.genai.types import UploadFileConfig
from google import genai
from google.api_core import exceptions


# --- Configuration Loading ---

def load_config():
    """Loads all configuration from the .env file."""
    load_dotenv()

    # Check for required Google API credentials
    if not os.path.exists('credentials.json'):
        print(
            "ERROR: 'credentials.json' not found. Please download it from Google Cloud and place it in the root folder.")
        sys.exit(1)

    # Load file-based prompts
    try:
        with open(os.getenv("ANALYSIS_PROMPT_FILE", "prompts/analysis_prompt.txt"), 'r') as f:
            analysis_prompt = f.read()
        with open(os.getenv("OCR_PROMPT_FILE", "prompts/ocr_prompt.txt"), 'r') as f:
            ocr_prompt = f.read()
    except FileNotFoundError as e:
        print(f"ERROR: Prompt file not found. Make sure {e.filename} exists.")
        sys.exit(1)

    config = {
        "SCOPES": [
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ],
        "FOLDER_ID": os.getenv("FOLDER_ID"),
        "SPREADSHEET_ID": os.getenv("SPREADSHEET_ID"),
        "SHEET_RANGE": os.getenv("SHEET_RANGE"),
        "GEMINI_MODEL": os.getenv("GEMINI_MODEL", "gemini-1.5-flash"),
        "ANALYSIS_PROMPT": analysis_prompt,
        "OCR_PROMPT": ocr_prompt,
        "CONVERT_TO_PDF": os.getenv("CONVERT_TO_PDF", "True").lower() == "true",
        "OUTPUT_FOLDER": os.getenv("OUTPUT_FOLDER", "output")
    }

    # Validate critical config
    if not config["SPREADSHEET_ID"] or not config["SHEET_RANGE"]:
        print("Warning: SPREADSHEET_ID or SHEET_RANGE not set in .env. Sheet-related functions will fail.")

    return config


# --- Main Application Logic ---

def main():
    """Main execution flow for the PDF analysis tool."""

    config = load_config()

    # --- Initialize API Services ---
    print("Initializing Google services...")
    gemini_client = genai.Client()
    drive_service = get_service("drive", "v3", config["SCOPES"])
    sheets_service = get_service("sheets", "v4", config["SCOPES"])

    if not drive_service or not sheets_service:
        print("Failed to initialize Google services. Exiting.")
        return

    # --- CHOOSE YOUR WORKFLOW ---
    # Uncomment the workflow you want to run.

    # Workflow 1: Read links from Sheet, process, and analyze
    print("Starting Workflow 1: Process from Google Sheet")
    links = get_sheet_data(sheets_service, config["SPREADSHEET_ID"], config["SHEET_RANGE"])
    if links:
        uploaded_pdfs = process_files_from_list(gemini_client, drive_service, links)
        if uploaded_pdfs:
            analyze_pdfs(gemini_client, config["ANALYSIS_PROMPT"], uploaded_pdfs, config["GEMINI_MODEL"])

    # Workflow 2: Update names in the Google Sheet
    # print("Starting Workflow 2: Update Sheet Names")
    # update_sheet(drive_service, sheets_service, config["SPREADSHEET_ID"], config["SHEET_RANGE"])

    # Workflow 3: Convert existing 'responses.json' to PDFs
    # print("Starting Workflow 3: Convert JSON to PDF")
    # try:
    #     with open('responses.json', 'r') as f:
    #         responses = json.load(f)
    #     analyses_to_pdf(responses, config["OUTPUT_FOLDER"])
    # except FileNotFoundError:
    #     print("ERROR: 'responses.json' not found. Run analysis first.")

    # Workflow 4: Run OCR on a single test file
    # print("Starting Workflow 4: OCR Single File")
    # test_links = ["https://drive.google.com/file/d/1NZ1T9atC_eVb9I92IqDwX7Wl9pXOcP2q/view"] # Example link
    # ocr_pdfs = process_files_from_list(gemini_client, drive_service, test_links)
    # if ocr_pdfs:
    #     analyze_pdfs(gemini_client, config["OCR_PROMPT"], ocr_pdfs, config["GEMINI_MODEL"])

    print("Main script finished.")


def get_service(api_name: str, version: str, scopes: List[str]) -> Optional[Resource]:
    """
    Initializes and returns an authenticated Google API service object.
    Handles token creation, storage, and refresh.
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            print("No valid credentials found. Starting auth flow...")
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', scopes)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            print("Credentials saved to 'token.json'.")

    try:
        service = build(api_name, version, credentials=creds)
        print(f"Successfully connected to {api_name} v{version}.")
        return service
    except HttpError as error:
        print(f"An error occurred building the service: {error}")
        return None


def process_files_from_list(gemini_client: genai.Client, drive_service: Resource, file_list: List[str]) -> List[
    Dict[str, Any]]:
    """Downloads files from a Drive link list, uploads to Gemini, and returns a list of processed file info."""
    if not drive_service:
        print("Google Drive service not available.")
        return []
    if not file_list:
        print("No files to process.")
        return []

    print("Fetching list of existing files from Gemini...")
    gemini_files_map = {}
    try:
        for f in gemini_client.files.list():
            gemini_files_map[f.display_name] = f
        print(f"Found {len(gemini_files_map)} files already on Gemini.")
    except Exception as e:
        print(f"Could not list Gemini files: {e}. Will attempt to upload all.")

    uploaded_pdfs = []
    for link in file_list:
        file_id = extract_file_id(link)
        if not file_id:
            print(f"Invalid link: {link}, skipped.")
            continue

        temp_file_path = None  # Ensure temp_file_path is defined
        try:
            file_metadata = drive_service.files().get(fileId=file_id, fields="name").execute()
            file_name = file_metadata.get("name")
            if not file_name:
                print(f"  - File ID {file_id} has no name. Skipping.")
                continue

            student_name = extract_student_name(file_name)
            if not student_name:
                print(f"  - Could not parse name from: {file_name}, using placeholder.")
                student_name = "Unknown Student"

            gemini_file = None
            if file_name in gemini_files_map:
                print(f"  - File {file_name} already uploaded to Gemini, using existing.")
                gemini_file = gemini_files_map[file_name]
            else:
                print(f"  - Processing {file_name}...")
                request = drive_service.files().get_media(fileId=file_id)

                # Use a temp file for download/upload
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_file:
                    temp_file_path = temp_file.name
                    downloader = MediaIoBaseDownload(temp_file, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                        print(f"    Downloading {file_name}... {int(status.progress() * 100)}%.", end='\r')

                print(f"\n    Uploading {file_name} to Gemini...")
                gemini_file = gemini_client.files.upload(
                    file=temp_file_path,
                    config=UploadFileConfig(display_name=file_name, mime_type='application/pdf')
                )
                print(f"    Upload complete.")
                time.sleep(1)  # Avoid rate limits

            uploaded_pdfs.append({
                "file_id": file_id,
                "file_name": file_name,
                "student_name": student_name,
                "gemini_file": gemini_file
            })

        except HttpError as error:
            print(f"! HTTP Error processing {file_name or link}: {error}, skipped.")
        except Exception as e:
            print(f"! Unexpected error with {file_name or link}: {e}, skipped.")
        finally:
            # Clean up the temp file
            if temp_file_path and os.path.exists(temp_file_path):
                os.remove(temp_file_path)

    return uploaded_pdfs


def analyze_pdfs(gemini_client: genai.Client, prompt: str, uploaded_pdfs: List[Dict], gemini_model: str):
    """Analyzes a list of uploaded PDFs using Gemini and saves responses."""
    all_responses = []
    for up in uploaded_pdfs:
        try:
            print(f"Analyzing {up['file_name']}...")
            result = gemini_client.models.generate_content(
                model=gemini_model,
                contents=[{"text": prompt}, {"file_data": {"file_uri": up["gemini_file"].uri}}],
            )
            print("  - Analysis Complete")

            analysis_text = result.text.strip() if result.text else "No analysis available."

            all_responses.append({
                "file_name": up['file_name'],
                "file_id": up['file_id'],
                "student_name": up['student_name'],
                "analysis": analysis_text
            })

            # Optional: Clean up file from Gemini after analysis
            # print(f"  - Cleaning up file {up['gemini_file'].name} from Gemini.")
            # gemini_client.files.delete(name=up['gemini_file'].name)

            time.sleep(1)  # Avoid rate limits

        except (exceptions.GoogleAPICallError, exceptions.ResourceExhausted) as error:
            print(f"! API Error analyzing {up['file_name']}: {error}, skipping.")
            time.sleep(5)  # Back off on rate limits/errors
        except Exception as error:
            print(f"! Unknown Error analyzing {up['file_name']}: {error}, skipping.")

    # Save responses to json
    print(f"\nSaving {len(all_responses)} responses to responses.json...")
    try:
        with open('responses.json', 'w') as f:
            json.dump(all_responses, f, indent=4)
        print("Successfully saved responses.")
    except IOError as e:
        print(f"ERROR: Could not write to responses.json: {e}")

    return all_responses


def get_sheet_data(sheets_service: Resource, spreadsheet_id: str, sheet_range: str) -> List[str]:
    """Reads hyperlink data from a specific range in a Google Sheet."""
    if not sheets_service:
        print("Google Sheets service not available.")
        return []

    print(f"Reading links from Sheet ID: {spreadsheet_id}, Range: {sheet_range}")
    try:
        result = sheets_service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            ranges=sheet_range,
            includeGridData=True,
            fields="sheets/data/rowData/values/hyperlink"
        ).execute()

        links_list = []
        sheets = result.get("sheets", [])
        if not sheets:
            print("No sheets found in the spreadsheet.")
            return []

        row_data = sheets[0].get("data", [{}])[0].get("rowData", [])
        if not row_data:
            print("No row data found in sheet range.")
            return []

        for row in row_data:
            cells = row.get("values", [])
            if cells and cells[0].get("hyperlink"):
                links_list.append(cells[0]["hyperlink"])

        print(f"Found {len(links_list)} links in the spreadsheet.")
        return links_list

    except HttpError as error:
        print(f"An error occurred reading from the sheet: {error}")
        return []


def update_sheet(drive_service: Resource, sheets_service: Resource, spreadsheet_id: str, sheet_range: str):
    """
    Reads links from the sheet, extracts first/last names, and overwrites the sheet
    with the data in the format [firstname, lastname, link].
    """
    print("Reading data from spreadsheet for name update...")
    links_list = get_sheet_data(sheets_service, spreadsheet_id, sheet_range)
    if not links_list:
        print("No data found to update.")
        return

    updated_values = [["First Name", "Last Name", "Link"]]  # Header row
    print("Processing file names...")
    for link in links_list:
        if not link:
            continue

        file_id = extract_file_id(link)
        if not file_id:
            updated_values.append(["ERROR", "Invalid link", link])
            continue

        try:
            file_metadata = drive_service.files().get(fileId=file_id, fields="name").execute()
            file_name = file_metadata.get("name")

            student_name = extract_student_name(file_name)
            if not student_name:
                updated_values.append(["ERROR", "Could not parse name", link])
                continue

            name_parts = student_name.split(" ", 1)
            first_name = name_parts[0]
            last_name = name_parts[1] if len(name_parts) > 1 else ""

            updated_values.append([first_name, last_name, link])
            print(f"  + Processed: {first_name} {last_name}")

        except HttpError as error:
            print(f"An error occurred processing link {link}: {error}")
            updated_values.append(["ERROR", "API error", link])

    print("Updating spreadsheet with processed data...")
    try:
        # Clear existing data
        sheets_service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id, range=sheet_range
        ).execute()

        # Write new data
        result = sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_range,
            valueInputOption="RAW",
            body={"values": updated_values}
        ).execute()
        print(f"Spreadsheet updated with {result.get('updatedRows')} rows.")
    except HttpError as error:
        print(f"An error occurred while updating the spreadsheet: {error}")


# --- Helper Functions ---

def analyses_to_pdf(responses: List[Dict], output_folder: str):
    """Converts a list of analysis responses to individual PDF files."""
    os.makedirs(output_folder, exist_ok=True)
    print(f"Converting {len(responses)} responses to PDF in '{output_folder}'...")

    for i, data in enumerate(responses):
        original_filename = data.get("file_name", f"response_{i}")
        analysis_text = data.get('analysis', 'No analysis available.')
        student_name = data.get('student_name', 'Unknown Student')

        pdf_filename = clean_filename(original_filename)
        pdf_filepath = os.path.join(output_folder, pdf_filename)

        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font('helvetica', '', 12)

            # Encode text to handle potential special characters
            cleaned_text = analysis_text.encode('latin-1', 'replace').decode('latin-1')

            # Create markdown, then convert to HTML for fpdf
            markdown_content = f"# {student_name}\n\n*{original_filename}*\n\n---\n\n{cleaned_text}"
            html_content = markdown.markdown(markdown_content)

            pdf.write_html(html_content)
            pdf.output(pdf_filepath)
            print(f"  - Created PDF for '{original_filename}'")

        except Exception as e:
            print(f"  - FAILED to create PDF for '{original_filename}'. Reason: {e}")


def clean_filename(filename: str) -> str:
    """Removes invalid file characters and ensures a .pdf extension."""
    name_without_ext = os.path.splitext(filename)[0]
    sanitized_name = re.sub(r'[\\/*?:"<>|]', "", name_without_ext)
    return f"{sanitized_name}.pdf"


def extract_file_id(link: str) -> Optional[str]:
    """Extracts the Google Drive file ID from various link formats."""
    match = re.search(r"(?:file\/d\/|open\?id=|uc\?id=)([a-zA-Z0-9-_]+)", link)
    if match:
        return match.group(1)
    else:
        print(f"  - Could not find ID in link: {link}")
        return None


def extract_student_name(text: str) -> Optional[str]:
    """Attempts to extract a student's name from a complex file name string."""
    try:
        # This is based on your logic: "Delete everything up to and including the ending parenthesis"
        try:
            processed_text = text[text.rindex(')') + 1:]
        except ValueError:
            print(f"  - No ')' in filename '{text}', using full name.")
            processed_text = text

        # "When we find the first ID code containing numbers, we delete it and everything after it"
        # Split by common delimiters
        parts = re.split(r'[_\s-]+', processed_text)
        name_parts = []
        for part in parts:
            if part.isalnum() and not part.isalpha() and any(char.isdigit() for char in part):
                # This part is alphanumeric and not purely alpha, assume it's an ID
                break
            if part:  # Append non-empty parts
                name_parts.append(part)

        # Final cleanup
        full_name = " ".join(name_parts).strip()

        # Remove any leading non-alpha characters
        if full_name and not full_name[0].isalpha():
            full_name = full_name[1:].strip()

        return full_name if full_name else None

    except Exception as e:
        print(f"Unexpected error extracting name from '{text}': {e}")
        return None


if __name__ == "__main__":
    main()