# OCR-Document-Automation



**WHY**



A multifunction printer or dedicated scanner emails scanned invoices as attachments and gives them a default name. If manually done, a person would have to find the email, open the document, read the invoice number and order number, then rename the document with the format WOXXXXXXX- INVXXXXXXX and save it in a specific folder. Doing this for one document would take 20 seconds. Doing it for 100 documents would take 33.33 minutes. That's time spent on doing a job that can be very easily automated and for which resources can be allocated towards better directions. 

This script does that job automatically, and the copy it saves is text-searchable rather than a flat image. Originally written to automate document filing at a manufacturing company, this version builds on that and refines it.



**HOW IT WORKS**


Outlook inbox
     │  match subject, unread only
     ▼
save PDF attachment ──► temporary folder
     │
     ▼
OCR  (ocrmypdf + Tesseract)  ──►  searchable PDF + text
     │
     ▼
find "Invoice Number" and "Order Number" in the text
     │
     ▼
save as  WO<order>- INV<invoice>.pdf
     │
     ▼
mark the email as read

The email is marked read only after the file is safely saved. If, for any reason, OCR or parsing fails, the message is left unread and retried on the next pass rather than simply discarded.


**TRY IT WITHOUT OUTLOOK**

The pipeline is the same whether the PDFs are delivered over email or placed directly in a folder, so the same script can be used to process either:

bash
python converter_modified.py --folder

That will read PDFs from Desktop/samples and write renamed copies to Desktop/archive, deleting the originals if the process succeeds, or leaving them in samples/ if it fails.

With the included samples you should end up with exactly: 

DOC_1.pdf  ->  WO0197006- INV0346424.pdf
DOC_2.pdf  ->  WO0190112- INV0331011.pdf
DOC_3.pdf  ->  WO0195580- INV0319271.pdf

To point it somewhere else:

bash
python converter_modified.py --folder "C:/Users/You/Documents/scans"



**SAMPLE DOCUMENTS**


samples/ contains synthetic invoices generated with the help of AI. Company names, addresses, part numbers, prices, and total amounts are all fictional. This repository does not contain any real documents.

The sample PDFs are designed to be image-only with no text layer to simulate a real-world document scan as you'd get from a physical scanner. This way the pipeline can be tested on real PDFs that also require OCR to be read.



**INSTALLATION**


bash
pip install -r requirements.txt

ocrmypdf also needs two programs installed separately — it is not pure Python, and pip install alone is not enough:

  Program	                Windows      	          macOS	                        Linux
Tesseract OCR	   UB Mannheim installer	  brew install tesseract	     apt install tesseract-ocr
Ghostscript	          ghostscript.com	   brew install ghostscript	     apt install ghostscript

pywin32 is only needed for Outlook mode and is specific to Windows; the --folder mode does not need it, but currently the script is written to always import it. IF THIS SCRIPT IS RUN ON MAC, DELETE THE LINE PYWIN32.



**USAGE**


_Command	What it does_

python converter_modified.py --folder	              Process Desktop/samples, no Outlook needed
python converter_modified.py --folder "<path>"	      Process any folder of PDFs
python converter_modified.py --once	                One pass over the Outlook inbox, then stop
python converter_modified.py	                        Keep watching the inbox, checking every 30 seconds
python converter_modified.py --once --include-read	  Also process messages already marked as read



**CONFIGURATION**


The settings block at the top of converter_modified.py:

python
SAMPLES_DIR = DESKTOP / "samples"            # PDFs are read from here
ARCHIVE_DIR = DESKTOP / "samples modified"   # renamed copies are saved here
SUBJECT = "Scanned from a Xerox Multifunction Printer"
POLL_SECONDS = 30

SUBJECT must match the printer's email subject exactly. The Desktop folder is auto-detected, including the OneDrive-redirected version that Windows uses when the account is on OneDrive. If your scans deliver with a different subject simply alter the SUBJECT string.



**FIELD EXTRACTION**


The invoice and order numbers are extracted with two regular expressions:

INVOICE_RE = re.compile(r"Invoice\s*Number\s*[:.]?\s*(\d{5,10})", re.I)
ORDER_RE   = re.compile(r"Order\s*Number\s*[:.]?\s*(\d{5,10})",   re.I)

Matching is done across the entire document with the whitespace collapsed, not line by line, since OCR is not reliably line-ordered. \s eats up any whitespace, [:.]? handles cases where OCR incorrectly reads a colon as a period or misses it entirely, and {5,10} prevents dates and page numbers from matching by accident. The match stops at the first non-digit, so a printed 0383810-IN becomes 0383810 without any special handling of the suffix.



**LIMITATIONS**


1. Read messages are skipped. If you open an email in Outlook before the script sees it, it will be ignored. --include-read is the opposite of this behavior. Tracking message IDs in a state file would avoid this entirely at the cost of disk usage.

2. The subject must match exactly. A changed SUBJECT will fail.

3. Only one invoice per PDF. A PDF with multiple invoices gets filed under the first set of numbers.

4. Polling, not push. The inbox is checked on a schedule rather than notified of new messages.



**DEVELOPMENT NOTES**


The original version of this script was written in 2024 to solve a document filing problem at work. This public version was rewritten in 2026 with the help of AI tools. The OCR logic in particular was made more robust to handle a wider variety of scanner outputs, several bugs were fixed, and the real scanned documents were replaced with synthetic ones to avoid sharing any employer data. 

Bugs fixed in the rewrite:

1. The inbox was scanned in full on every cycle with no record of what had been handled, so the same attachments were re-downloaded and re-processed every few seconds. This meant that a manual operator would have to be alert and aware as to how far along in the document scanning the script was. Filtering now happens inside Outlook, and handled messages are marked read.

2. A suffix-stripping branch ran even when the pattern before it had not matched, calling .split() on None and terminating the process.

3. The extraction patterns were specific to a certain layout and wasn't robust enough to adapt to errors. 

4. Every OCR pass left a _OCR.pdf file behind, and the searchable output was discarded in favour of the original image-only scan. This resulted in there being many OCR (intermediary) files. Now tempfile deletes this.  
