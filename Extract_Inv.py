import os
import sys
import requests
import json
import re
import pandas as pd
import platform
from pypdf import PdfReader

# --- Cross-platform Configuration ---
def get_default_source_dir():
    """Get default source directory based on OS"""
    if platform.system() == 'Windows':
        return r"d:\Project\ocr\source"
    else:
        return os.path.expanduser("~/ocr/source")


def get_default_output_dir():
    """Get default output directory based on OS"""
    if platform.system() == 'Windows':
        return r"d:\Project\ocr\output"
    else:
        return os.path.expanduser("~/ocr/output")


# --- Configuration (supports environment variables) ---
# API Key from environment variable (recommended) or fallback to config file
API_KEY = os.environ.get("TYPHOON_API_KEY", "")

# If not set in environment, try to load from config file
if not API_KEY:
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                API_KEY = config.get('API_KEY', '')
        except:
            pass

# Script directory for relative paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VENDOR_MASTER_FILE = "Vendor_branch.xlsx"
TEMPLATES_FILE = "document_templates.json"

# Command line arguments or defaults
# Usage: python Extract_Inv.py <source_dir> <output_dir> <page_config> [document_type]
if len(sys.argv) >= 3:
    SOURCE_DIR = sys.argv[1]
    OUTPUT_DIR = sys.argv[2]
    PAGE_CONFIG = sys.argv[3] if len(sys.argv) > 3 else "All"
    DOC_TYPE = sys.argv[4] if len(sys.argv) > 4 else "auto"
else:
    SOURCE_DIR = get_default_source_dir()
    OUTPUT_DIR = get_default_output_dir()
    PAGE_CONFIG = "2"
    DOC_TYPE = "auto"

# Create output directory if not exists
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


# --- Load Document Templates ---
def load_templates():
    """Load document templates from JSON file"""
    path = os.path.join(SCRIPT_DIR, TEMPLATES_FILE)
    if not os.path.exists(path):
        print(f"Warning: Templates file not found: {TEMPLATES_FILE}")
        return None
    
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading templates: {e}")
        return None


def detect_document_type(text, templates):
    """Auto-detect document type based on keywords in text"""
    if not text or not templates:
        return "invoice"  # default
    
    text_lower = text.lower()
    scores = {}
    
    for doc_type, template in templates.get("templates", {}).items():
        keywords = template.get("detect_keywords", [])
        score = 0
        for keyword in keywords:
            if keyword.lower() in text_lower:
                score += 1
        if score > 0:
            scores[doc_type] = score
    
    if scores:
        # Return type with highest score
        return max(scores, key=scores.get)
    
    return "invoice"  # default fallback


def extract_field_by_patterns(text, patterns, options=None):
    """Extract field value using multiple regex patterns"""
    if not text or not patterns:
        return ""
    
    options = options or {}
    
    for pattern in patterns:
        try:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                # Get first capturing group or full match
                value = match.group(1) if match.lastindex and match.lastindex >= 1 else match.group(0)
                
                # Clean HTML if specified
                if options.get("clean_html"):
                    value = re.sub(r'<br\s*/?>', ' ', value)
                    value = re.sub(r'<[^>]+>', '', value)
                
                # Clean whitespace
                value = re.sub(r'[\r\n]+', ' ', value)
                value = re.sub(r'\s+', ' ', value).strip()
                
                # Clean non-digits if specified
                if options.get("clean_non_digits"):
                    value = re.sub(r'\D', '', value)
                    # Truncate to specified length
                    if options.get("length"):
                        value = value[:options["length"]]
                
                if value:
                    return value
        except Exception:
            continue
    
    return ""


def extract_common_fields(text, common_fields_config):
    """Extract common fields (tax_id, branch) that apply to all document types"""
    result = {"tax_id": "", "branch": ""}
    
    if not text or not common_fields_config:
        return result
    
    # Extract Tax ID
    tax_config = common_fields_config.get("tax_id", {})
    tax_patterns = tax_config.get("patterns", [])
    
    # First try to find 13-digit number directly
    all_tax_ids = re.findall(r"\b(\d{13})\b", text)
    if all_tax_ids:
        result["tax_id"] = all_tax_ids[0]
    else:
        # Try pattern with dashes
        tax_pattern_match = re.search(r"\b\d{1}-\d{4}-\d{5}-\d{2}-\d{1}\b", text)
        if tax_pattern_match:
            result["tax_id"] = re.sub(r"\D", "", tax_pattern_match.group(0))
        else:
            # Try keyword-based extraction
            for pattern in tax_patterns:
                value = extract_field_by_patterns(text, [pattern], {"clean_non_digits": True, "length": 13})
                if value and len(value) >= 10:
                    result["tax_id"] = value
                    break
    
    # Extract Branch
    branch_config = common_fields_config.get("branch", {})
    branch_patterns = branch_config.get("patterns", [])
    default_hq = branch_config.get("default_hq", "00000")
    pad_zeros = branch_config.get("pad_zeros", 5)
    
    # Check for Head Office keywords first
    ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
    if ho_match:
        result["branch"] = default_hq
    else:
        # Try to find branch number
        branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})", text, re.IGNORECASE)
        if branch_match:
            result["branch"] = branch_match.group(1).zfill(pad_zeros)
    
    return result


def parse_ocr_data_with_template(text, templates, doc_type="auto"):
    """Parse OCR text using document template patterns"""
    result = {
        "document_type": "",
        "document_type_name": "",
        "document_no": "",
        "date": "",
        "amount": "",
        "tax_id": "",
        "branch": "",
        "extra_fields": {}
    }
    
    if not text:
        return result
    
    # Load templates if not provided
    if not templates:
        templates = load_templates()
    
    if not templates:
        # Fallback to basic extraction
        return parse_ocr_data_basic(text)
    
    # Detect or use specified document type
    if doc_type == "auto":
        detected_type = detect_document_type(text, templates)
    else:
        detected_type = doc_type if doc_type in templates.get("templates", {}) else "invoice"
    
    result["document_type"] = detected_type
    
    # Get template for this document type
    template = templates.get("templates", {}).get(detected_type, {})
    result["document_type_name"] = template.get("name", detected_type)
    
    # Extract common fields (tax_id, branch) - always extracted for Vendor lookup
    common_fields = templates.get("common_fields", {})
    common_result = extract_common_fields(text, common_fields)
    result["tax_id"] = common_result["tax_id"]
    result["branch"] = common_result["branch"]
    
    # Extract template-specific fields
    fields_config = template.get("fields", {})
    
    for field_name, field_config in fields_config.items():
        patterns = field_config.get("patterns", [])
        options = {
            "clean_html": field_config.get("clean_html", False),
            "clean_non_digits": field_config.get("clean_non_digits", False),
            "length": field_config.get("length")
        }
        
        value = extract_field_by_patterns(text, patterns, options)
        
        # Handle fallback for amount fields
        if not value and field_config.get("fallback") == "last_amount":
            amounts = re.findall(r"([\d,]+\.\d{2})", text)
            value = amounts[-1] if amounts else ""
        
        # Store in appropriate location
        if field_name in ["document_no", "date", "amount"]:
            result[field_name] = value
        else:
            result["extra_fields"][field_name] = value
    
    return result


def parse_ocr_data_basic(text):
    """Basic OCR parsing without templates (fallback)"""
    result = {
        "document_type": "invoice",
        "document_type_name": "ใบกำกับภาษี/Invoice",
        "document_no": "",
        "date": "",
        "amount": "",
        "tax_id": "",
        "branch": "",
        "extra_fields": {}
    }
    
    if not text:
        return result
    
    # Document number
    inv_match = re.search(r"เลขที่\s*[:\.]?\s*([A-Za-z0-9\-\/]{3,})", text)
    result["document_no"] = inv_match.group(1) if inv_match else ""
    
    # Date
    date_match = re.search(r"วันที่\s*[:\.]?\s*(\d{1,2}\s+[^\s]+\s+\d{4}|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", text)
    result["date"] = date_match.group(1) if date_match else ""
    
    # Amount
    amount_match = re.search(r"(?:จำนวนเงินรวมทั้งสิ้น|รวมเงินทั้งสิ้น|GRAND TOTAL)\s*[:\.]?\s*([\d,]+\.\d{2})", text, re.IGNORECASE)
    if amount_match:
        result["amount"] = amount_match.group(1)
    else:
        amounts = re.findall(r"([\d,]+\.\d{2})", text)
        result["amount"] = amounts[-1] if amounts else ""
    
    # Tax ID
    all_tax_ids = re.findall(r"\b(\d{13})\b", text)
    if all_tax_ids:
        result["tax_id"] = all_tax_ids[0]
    else:
        tax_pattern_match = re.search(r"\b\d{1}-\d{4}-\d{5}-\d{2}-\d{1}\b", text)
        if tax_pattern_match:
            result["tax_id"] = re.sub(r"\D", "", tax_pattern_match.group(0))
    
    # Branch
    ho_match = re.search(r"(?:สำนักงานใหญ่|สนญ\.?|Head\s*Office|H\.?O\.?)", text, re.IGNORECASE)
    if ho_match:
        result["branch"] = "00000"
    else:
        branch_match = re.search(r"(?:สาขา(?:ที่)?|Branch(?:\s*No\.?)?)\s*[:\.]?\s*(\d{1,5})", text, re.IGNORECASE)
        if branch_match:
            result["branch"] = branch_match.group(1).zfill(5)
    
    return result


# --- Load Vendor Master (Excel) ---
def load_vendor_master():
    """Load vendor master data from Excel file"""
    path = os.path.join(SCRIPT_DIR, VENDOR_MASTER_FILE)
    if not os.path.exists(path):
        print(f"Warning: Vendor master file not found: {VENDOR_MASTER_FILE} in {SCRIPT_DIR}")
        return None
    
    try:
        print(f"Loading Vendor Master from: {path}")
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        
        req_cols = ['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']
        if not all(col in df.columns for col in req_cols):
            print(f"Error: Missing columns in Master file (required: {req_cols})")
            return None

        df['เลขประจำตัวผู้เสียภาษี'] = df['เลขประจำตัวผู้เสียภาษี'].fillna('').str.replace(r'\D', '', regex=True)
        
        def clean_branch(x):
            x = str(x).strip()
            if x.isdigit():
                return x.zfill(5)
            return x
        
        df['สาขา'] = df['สาขา'].apply(clean_branch)
        
        # Also get company name if available
        cols_to_return = ['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']
        if 'ชื่อบริษัท' in df.columns:
            cols_to_return.append('ชื่อบริษัท')
        
        return df[cols_to_return]
        
    except Exception as e:
        print(f"Error reading Vendor file: {e}")
        return None


# --- Calculate target pages ---
def get_target_pages(selection_str, total_pages):
    """Parse page selection string and return list of pages to process"""
    selection_str = str(selection_str).lower().replace(" ", "")
    pages_to_process = set()

    if selection_str == 'all':
        return list(range(1, total_pages + 1))

    parts = selection_str.split(',')
    for part in parts:
        if '-' in part:
            start_s, end_s = part.split('-')
            start = int(start_s)
            end = total_pages if end_s == 'n' else int(end_s)
            end = min(end, total_pages)
            if start <= end:
                pages_to_process.update(range(start, end + 1))
        else:
            if part.isdigit():
                p = int(part)
                if 1 <= p <= total_pages:
                    pages_to_process.add(p)
    
    return sorted(list(pages_to_process))


# --- Typhoon OCR API ---
def extract_text_from_image(file_path, api_key, pages_list):
    """Extract text from PDF using Typhoon OCR API"""
    url = "https://api.opentyphoon.ai/v1/ocr"
    
    data = {
        'model': 'typhoon-ocr',
        'task_type': 'default', 
        'max_tokens': '16000',
        'temperature': '0.1',
        'top_p': '0.6',
        'repetition_penalty': '1.1'
    }
    
    if pages_list:
        data['pages'] = json.dumps(pages_list)
    
    headers = {'Authorization': f'Bearer {api_key}'}

    try:
        with open(file_path, 'rb') as file:
            files = {'file': file}
            response = requests.post(url, files=files, data=data, headers=headers)

        if response.status_code == 200:
            result = response.json()
            extracted_texts = []
            
            for page_result in result.get('results', []):
                if page_result.get('success'):
                    content = page_result['message']['choices'][0]['message']['content']
                    try:
                        parsed = json.loads(content)
                        text = parsed.get('natural_text', content)
                    except json.JSONDecodeError:
                        text = content
                    extracted_texts.append(text)
            
            return '\n'.join(extracted_texts)
        else:
            print(f"Error API: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Error processing file: {e}")
        return None


# --- Main Logic ---
def main():
    print(f"--- Start Processing ---")
    print(f"Platform: {platform.system()} {platform.release()}")
    print(f"Source: {SOURCE_DIR}")
    print(f"Output: {OUTPUT_DIR}")
    print(f"Page Config: {PAGE_CONFIG}")
    print(f"Document Type: {DOC_TYPE}")
    
    if not API_KEY:
        print("[ERROR] API Key not set. Please set TYPHOON_API_KEY environment variable or update config.json")
        return
    
    # Load templates
    templates = load_templates()
    if templates:
        available_types = list(templates.get("templates", {}).keys())
        print(f"Loaded templates: {available_types}")
    
    # Load Vendor Master
    vendor_df = load_vendor_master()
    
    data_rows = []
    
    if not os.path.exists(SOURCE_DIR):
        print(f"Error: Source directory not found: {SOURCE_DIR}")
        return

    files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".pdf")]
    
    if not files:
        print("No PDF files found.")
        return

    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        print(f"\nProcessing: {filename}")

        try:
            reader = PdfReader(file_path)
            total_pages = len(reader.pages)
            target_pages = get_target_pages(PAGE_CONFIG, total_pages)
            print(f"   -> Total Pages: {total_pages}, Target: {target_pages}")

            for page_num in target_pages:
                print(f"      Reading Page {page_num}...")
                page_text = extract_text_from_image(file_path, API_KEY, pages_list=[page_num])
                
                if page_text:
                    # Save raw OCR text
                    txt_filename = f"{os.path.splitext(filename)[0]}_page{page_num}.txt"
                    txt_path = os.path.join(OUTPUT_DIR, txt_filename)
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write(page_text)
                    
                    # Parse using templates
                    parsed = parse_ocr_data_with_template(page_text, templates, DOC_TYPE)
                    
                    print(f"      Detected Type: {parsed['document_type_name']}")
                    
                    hyperlink_formula = f'=HYPERLINK("{file_path}", "{filename} (Page {page_num})")'

                    row_data = {
                        "Link PDF": hyperlink_formula,
                        "Page": page_num,
                        "Document Type": parsed["document_type_name"],
                        "VendorID_OCR": parsed["tax_id"],
                        "Branch_OCR": parsed["branch"],
                        "Document No": parsed["document_no"],
                        "Date": parsed["date"],
                        "Amount": parsed["amount"],
                    }
                    
                    # Add extra fields from template
                    for field_name, value in parsed.get("extra_fields", {}).items():
                        # Convert field_name to readable label
                        label = field_name.replace("_", " ").title()
                        row_data[label] = value
                    
                    data_rows.append(row_data)
                else:
                    print(f"      Warning: Failed to read page {page_num}")

        except Exception as e:
            print(f"   Error reading PDF file: {e}")

    # Save and merge data
    if data_rows:
        df = pd.DataFrame(data_rows)
        
        if vendor_df is not None:
            print("\nMapping Vendor Code...")
            df = pd.merge(
                df, 
                vendor_df, 
                left_on=['VendorID_OCR', 'Branch_OCR'], 
                right_on=['เลขประจำตัวผู้เสียภาษี', 'สาขา'], 
                how='left'
            )
            df.rename(columns={'Vendor code SAP': 'Vendor code'}, inplace=True)
            df.drop(columns=['เลขประจำตัวผู้เสียภาษี', 'สาขา'], inplace=True, errors='ignore')
        else:
            df['Vendor code'] = ""

        # Reorder columns - put important ones first
        priority_cols = [
            "Link PDF", "Page", "Document Type",
            "VendorID_OCR", "Branch_OCR", "Vendor code", "ชื่อบริษัท",
            "Document No", "Date", "Amount"
        ]
        
        # Get all columns, prioritizing the defined order
        all_cols = df.columns.tolist()
        final_cols = [col for col in priority_cols if col in all_cols]
        final_cols += [col for col in all_cols if col not in final_cols]
        
        df = df[final_cols]

        output_excel_path = os.path.join(OUTPUT_DIR, "summary_ocr.xlsx")
        
        try:
            with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            print(f"\nSuccess! Output saved at: {output_excel_path}")
            print(f"Total rows: {len(df)}")
        except Exception as e:
            print(f"Error saving Excel: {e}")
    else:
        print("No data extracted.")


if __name__ == "__main__":
    main()
