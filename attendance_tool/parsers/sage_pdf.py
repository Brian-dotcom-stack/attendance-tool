import pdfplumber
import os
import re
from datetime import datetime

DATE_RANGE = re.compile(r"on (\d{2}/\d{2}/\d{4})(?: - (\d{2}/\d{2}/\d{4}))?")

def parse_sage_pdf(folder_path):
    cancellations = []

    files = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]

    for file in files:
        file_path = os.path.join(folder_path, file)

        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split("\n")

                current_name = None

                for line in lines:
                    line = line.strip()

                    # Capture names (heuristic: capitalised lines not containing "Holidays")
                    if line and "Holidays" not in line and "days" not in line and "Document" not in line:
                        if re.match(r"^[A-Za-z ]+$", line):
                            current_name = line.strip()
                            continue

                    match = DATE_RANGE.search(line)
                    if match and current_name:
                        start = datetime.strptime(match.group(1), "%d/%m/%Y")
                        end = match.group(2)

                        if end:
                            end = datetime.strptime(end, "%d/%m/%Y")
                        else:
                            end = start

                        cancellations.append({
                            "name": current_name,
                            "start_date": start.date(),
                            "end_date": end.date()
                        })

                current_name = None

    print(f"\n✅ Parsed {len(cancellations)} leave records from {len(files)} PDFs")
    return cancellations