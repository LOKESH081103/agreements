import os
import re
import time
import pandas as pd
import pdfplumber

# Folder containing PDFs (current script folder)
FOLDER_PATH = os.path.dirname(os.path.abspath(__file__))


def extract_data_from_pdf(pdf_path):
    all_text = ""

    # 1. Extract text from all pages
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

    # Normalize line breaks across platforms
    text = all_text.replace("\r\n", "\n").replace("\r", "\n")

    # 2. Extract Agreement No / Loan Account No
    agreement_no = ""
    agreement_match = re.search(
        r"(?:loan\s+account\s+no|loan\s+a/c\s+no|loan\s+agreement\s+no|agreement\s+no)\.?:?\s*([\w\d]+)",
        text,
        re.IGNORECASE,
    )
    if agreement_match:
        agreement_no = agreement_match.group(1).strip()

    # 3. Extract Cheque Details (Cheque No, Cheque Date, Total Amount, Drawn On)
    cheque_no = ""
    cheque_date = ""
    total_amount = ""
    drawn_on = ""

    # --- FORMAT 1: Paragraph Pattern ---
    para_match = re.search(
        r"cheque\s+bearing\s+no\.?\s*([\w\d]+)\s*,\s*dated\s*[-:\s]*(\d{2}[-/\.]\d{2}[-/\.]\d{2,4})\s+for\s+a\s+sum\s+of\s+Rs\.?\s*([\d,]+(?:/\-)?)\s+drawn\s+on\s+(.*?)(?=\s+in\s+favo?u?r|\s+for|\s+vide|\.|\n|$)",
        text,
        re.IGNORECASE,
    )

    if para_match:
        cheque_no = para_match.group(1).strip()
        cheque_date = para_match.group(2).strip()
        total_amount = para_match.group(3).strip()
        drawn_on = para_match.group(4).strip()
    else:
        # --- FORMAT 2: Tabular / Pipe Delimited Pattern Fallback ---
        text_clean_pipes = re.sub(r"\n\s*\|\s*", " | ", text)
        pipe_match = re.search(
            r"(\d{2}[-/\.]\d{2}[-/\.]\d{2,4})\s*\|\s*(\w+)\s*\|\s*([^|]+?)\s*(?:\||\n)",
            text_clean_pipes,
        )
        if pipe_match:
            cheque_date = pipe_match.group(1).strip()
            cheque_no = pipe_match.group(2).strip()
            drawn_on = pipe_match.group(3).strip()

        amount_match = re.search(
            r"(?:TOTAL|AMOUNT)\s*[:\-]?\s*([\d,]+(?:/\-)?)", text, re.IGNORECASE
        )
        if amount_match:
            total_amount = amount_match.group(1).strip()

    # 4. Extract Recipient Block (To ... Subject/Dear Sir)
    to_match = re.search(
        r"\bTo\s*[:,-]?\s*(.*?)(?=\n\s*(?:Dear\s+Sir|Sub\s*:|Subject\s*:|Under\s+the\s+instructions))",
        text,
        re.DOTALL | re.IGNORECASE,
    )

    records = []
    if to_match:
        to_block = "\n" + to_match.group(1).strip()

        # Regex to capture each numbered recipient: "1. NAME \n Address..."
        pattern = r"\n\s*(\d+)[\.\)]\s*([^\n]+)(.*?)(?=(?:\n\s*\d+[\.\)]|\Z))"
        matches = re.findall(pattern, to_block, re.DOTALL)

        for p_num, name, addr_raw in matches:
            name = name.strip()

            # Filter out empty or invalid name matches (e.g., orphaned numbers 1., 2., 3.)
            if not name or len(name) < 2 or name.isdigit():
                continue

            addr_lines = []
            for line in addr_raw.strip().split("\n"):
                line_str = line.strip()
                if not line_str:
                    continue
                # Skip Speed Post consignment tracking numbers (e.g. EM054483479IN, EW054 4 7 1 8 55 IN)
                if re.search(
                    r"^[A-Z]{2}[\d\s]+[A-Z]{2}$", line_str, re.IGNORECASE
                ):
                    continue
                # Skip DATED headers inside recipient block
                if re.search(r"^DATED\s*:", line_str, re.IGNORECASE):
                    continue
                # Skip orphaned list numbers
                if re.match(r"^\d+[\.\)]?$", line_str):
                    continue
                addr_lines.append(line_str)

            address = ", ".join(addr_lines)

            records.append(
                {
                    "Party Number": p_num,
                    "Name": name,
                    "Address": address,
                    "Agreement No": agreement_no,
                    "Cheque No": cheque_no,
                    "Cheque Date": cheque_date,
                    "Cheque Amount": total_amount,
                    "Drawn On": drawn_on,
                    "Source File": os.path.basename(pdf_path),
                }
            )

    return records


# --- Main Processing Loop ---
all_data = []
pdf_files = [f for f in os.listdir(FOLDER_PATH) if f.lower().endswith(".pdf")]

print(f"Found {len(pdf_files)} PDF files to process...\n")

for pdf_file in pdf_files:
    pdf_full_path = os.path.join(FOLDER_PATH, pdf_file)
    try:
        extracted_records = extract_data_from_pdf(pdf_full_path)
        all_data.extend(extracted_records)
        print(
            f"SUCCESS [{pdf_file}]: Extracted {len(extracted_records)} recipients"
        )
    except Exception as e:
        print(f"ERROR [{pdf_file}]: {e}")

# --- Export to Excel ---
if all_data:
    output_path = os.path.join(FOLDER_PATH, "extracted_notices.xlsx")
    df = pd.DataFrame(all_data)

    try:
        df.to_excel(output_path, index=False)
        print(
            f"\nCOMPLETED: Extracted {len(all_data)} total records to:\n{output_path}"
        )
    except PermissionError:
        fallback_path = os.path.join(
            FOLDER_PATH, f"extracted_notices_{int(time.time())}.xlsx"
        )
        df.to_excel(fallback_path, index=False)
        print(
            f"\nNOTE: 'extracted_notices.xlsx' is currently open in Excel. Saved output to:\n{fallback_path}"
        )
else:
    print("\nFAILED: No data was extracted from the PDFs.")