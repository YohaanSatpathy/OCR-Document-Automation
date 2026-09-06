import os
import time
import re
from pathlib import Path
import tempfile
import ocrmypdf
from pypdf import PdfReader


# Folder setup
BASE_DIR = Path(__file__).resolve().parent
SAMPLES_DIR = BASE_DIR / "samples"
ARCHIVE_DIR = BASE_DIR / "archive"

SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)


def download_pdf_attachments_and_process(folder_path):
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items.Restrict(
        "[Subject] = 'Scanned from a Xerox Multifunction Printer' "
        "AND [UnRead] = True"
    )

    # Make a list before changing email status
    messages = list(messages)

    for message in messages:
        attachments = message.Attachments
        successful = True

        for attachment in attachments:

            if attachment.FileName.lower().endswith(".pdf"):

                file_path = os.path.join(
                    folder_path,
                    attachment.FileName
                )

                try:
                    attachment.SaveAsFile(file_path)
                    print(f"Saved: {attachment.FileName}")

                    if not process_pdf(file_path):
                        successful = False

                except Exception as e:
                    print(f"Failed to process file: {e}")
                    successful = False

                finally:
                    if os.path.exists(file_path):
                        os.remove(file_path)

        if successful:
            message.UnRead = False
            print("Email processed successfully.")


def process_pdf(file_path):

    file_path = Path(file_path)

    try:
        # Keep OCR output in a temporary folder
        with tempfile.TemporaryDirectory() as temp_dir:

            ocr_output = Path(temp_dir) / "OCR.pdf"

            extracted_text = perform_ocr(
                file_path,
                ocr_output
            )

        print(f"Extracted text: {extracted_text}")

        new_filename = extract_information_from_text(
            extracted_text
        )

        if not new_filename:
            print("Failed to extract invoice information.")
            return False

        new_filepath = ARCHIVE_DIR / (
            new_filename + ".pdf"
        )

        # Prevent overwriting an existing file
        if new_filepath.exists():
            print(f"File already exists: {new_filepath.name}")
            return False

        os.replace(file_path, new_filepath)

        print(f"Renamed file to: {new_filepath.name}")
        return True

    except Exception as e:
        print(f"Failed to process {file_path.name}: {e}")
        return False


def perform_ocr(file_path, output_path):

    ocrmypdf.ocr(
        str(file_path),
        str(output_path),
        skip_text=True,
        deskew=True,
        progress_bar=False
    )

    print(f"OCR completed for {file_path.name}")

    extracted_text = ""

    with open(output_path, "rb") as f:
        reader = PdfReader(f)

        for page in reader.pages:
            page_text = page.extract_text()

            if page_text:
                extracted_text += page_text + "\n"

    return extracted_text


def extract_information_from_text(text):

    # OCR can create inconsistent spaces and line breaks
    text = re.sub(r"\s+", " ", text)

    order_match = re.search(
        r"Order\s*Number\s*[:.]?\s*(\d{5,10})",
        text,
        re.IGNORECASE
    )

    invoice_match = re.search(
        r"Invoice\s*Number\s*[:.]?\s*(\d{5,10})",
        text,
        re.IGNORECASE
    )

    if order_match:
        order_number = order_match.group(1)
        print(f"Found Order Number: {order_number}")
    else:
        order_number = None

    if invoice_match:
        invoice_number = invoice_match.group(1)
        print(f"Found Invoice Number: {invoice_number}")
    else:
        invoice_number = None

    if order_number and invoice_number:
        return f"WO{order_number}- INV{invoice_number}"

    return None


def process_folder(folder):

    folder = Path(folder)
    pdf_files = list(folder.glob("*.pdf"))

    if not pdf_files:
        print("No PDF files found.")
        return

    for pdf_file in pdf_files:
        print(f"\nProcessing: {pdf_file.name}")
        process_pdf(pdf_file)


def main():

    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--folder",
        action="store_true",
        help="Process PDFs from the samples folder"
    )

    args = parser.parse_args()

    if args.folder:
        print(f"Processing: {SAMPLES_DIR}")
        process_folder(SAMPLES_DIR)
        return

    folder_path = tempfile.gettempdir()

    print("Starting Outlook monitoring...")

    while True:

        try:
            download_pdf_attachments_and_process(
                folder_path
            )

            print("Waiting 30 seconds...")
            time.sleep(30)

        except KeyboardInterrupt:
            print("Program stopped.")
            break

        except Exception as e:
            print(f"Error: {e}")
            time.sleep(30)


if __name__ == "__main__":
    main()
