import os
import sys
import requests
import json
import re
import pandas as pd
import base64
import io
import gc
import platform
import shutil
from pypdf import PdfReader
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance

# --- Cross-platform Configuration ---
def get_default_poppler_path():
    """Get Poppler path based on operating system"""
    system = platform.system()
    
    if system == 'Windows':
        # Common Windows install locations
        possible_paths = [
            r"C:\poppler\Library\bin",
            r"C:\poppler\bin",
            r"C:\Program Files\poppler\bin",
            os.path.expanduser(r"~\poppler\bin"),
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None  # Will try system PATH
    
    elif system == 'Darwin':  # macOS
        # Homebrew install location
        possible_paths = [
            "/opt/homebrew/bin",
            "/usr/local/bin",
            "/opt/homebrew/Cellar/poppler",
        ]
        for path in possible_paths:
            if os.path.exists(path) and shutil.which("pdftoppm"):
                return None  # Use system PATH
        return None
    
    else:  # Linux
        # Check if poppler-utils is installed
        if shutil.which("pdftoppm"):
            return None  # Use system PATH
        return None


def get_default_source_dir():
    """Get default source directory based on OS"""
    system = platform.system()
    
    if system == 'Windows':
        return r"d:\Project\ocr\source"
    else:
        # Use home directory for Linux/Mac
        return os.path.expanduser("~/ocr/source")


def get_default_output_dir():
    """Get default output directory based on OS"""
    system = platform.system()
    
    if system == 'Windows':
        return r"d:\Project\ocr\output"
    else:
        # Use home directory for Linux/Mac
        return os.path.expanduser("~/ocr/output")


# --- Configuration (supports environment variables) ---
OLLAMA_API_URL = os.environ.get("OLLAMA_API_URL", "http://localhost:11434/api/generate")
MODEL_NAME = os.environ.get("OCR_MODEL_NAME", "scb10x/typhoon-ocr1.5-3b:latest")
POPPLER_PATH = os.environ.get("POPPLER_PATH", get_default_poppler_path())

# Script directory for relative paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_FILE = "document_templates.json"

# Command line arguments or defaults
# Usage: python Extract_Inv_local.py <source_dir> <output_dir> <page_config> [document_type]
if len(sys.argv) >= 3:
    SOURCE_DIR, OUTPUT_DIR, PAGE_CONFIG = sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "All"
    DOC_TYPE = sys.argv[4] if len(sys.argv) > 4 else "auto"
else:
    SOURCE_DIR = get_default_source_dir()
    OUTPUT_DIR = get_default_output_dir()
    PAGE_CONFIG = "2"
    DOC_TYPE = "auto"

os.makedirs(OUTPUT_DIR, exist_ok=True)


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


def check_ollama_connection():
    """Check if Ollama is running and accessible"""
    try:
        response = requests.get(
            os.environ.get("OLLAMA_API_URL", "http://localhost:11434/api/tags").replace("/api/generate", "/api/tags"),
            timeout=5
        )
        return response.status_code == 200
    except:
        return False


def load_vendor_master():
    """Load vendor master data from Excel file"""
    path = os.path.join(SCRIPT_DIR, "Vendor_branch.xlsx")
    if not os.path.exists(path):
        return None
    try:
        df = pd.read_excel(path, dtype=str)
        df.columns = df.columns.str.strip()
        df['เลขประจำตัวผู้เสียภาษี'] = df['เลขประจำตัวผู้เสียภาษี'].fillna('').str.replace(r'\D', '', regex=True)
        df['สาขา'] = df['สาขา'].fillna('').str.strip()
        
        # Clean branch - pad with zeros
        def clean_branch(x):
            x = str(x).strip()
            if x.isdigit():
                return x.zfill(5)
            return x
        df['สาขา'] = df['สาขา'].apply(clean_branch)
        
        # Return columns needed
        cols_to_return = ['เลขประจำตัวผู้เสียภาษี', 'สาขา', 'Vendor code SAP']
        if 'ชื่อบริษัท' in df.columns:
            cols_to_return.append('ชื่อบริษัท')
        
        return df[cols_to_return]
    except:
        return None


def get_target_pages(selection_str, total_pages):
    """Parse page selection string and return list of pages to process"""
    selection_str = str(selection_str).lower().replace(" ", "")
    if selection_str == 'all':
        return list(range(1, total_pages + 1))
    
    pages = set()
    for part in selection_str.split(','):
        if '-' in part:
            try:
                s, e = part.split('-')
                pages.update(range(int(s), (total_pages if e == 'n' else int(e)) + 1))
            except:
                pass
        elif part.isdigit():
            pages.add(int(part))
    return sorted([p for p in pages if 1 <= p <= total_pages])


def preprocess_image(image, max_size=1280):
    """Preprocess image for better OCR results"""
    if image.mode != 'RGB':
        image = image.convert('RGB')
    image = ImageEnhance.Contrast(image).enhance(1.8)
    if max(image.size) > max_size:
        image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
    return image


def extract_text_from_image(file_path, pages_list):
    """Extract text from PDF pages using Ollama OCR"""
    extracted_pages = []
    
    # Use poppler_path if it exists, otherwise None (use system PATH)
    poppler = POPPLER_PATH if POPPLER_PATH and os.path.exists(POPPLER_PATH) else None
    
    for page_num in pages_list:
        try:
            print(f"   [Step 1] Rendering Page {page_num}...")
            images = convert_from_path(
                file_path,
                first_page=page_num,
                last_page=page_num,
                poppler_path=poppler,
                dpi=300
            )
            if not images:
                continue
                
            img = preprocess_image(images[0])
            buffered = io.BytesIO()
            img.save(buffered, format="PNG")
            img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
            
            print(f"   [Step 2] Sending to AI...")
            payload = {
                "model": MODEL_NAME,
                "prompt": "Extract text from image. Return clean Markdown only.",
                "images": [img_str],
                "stream": False,
                "options": {
                    "temperature": 0,
                    "num_ctx": 4096,
                    "num_predict": 1024
                }
            }
            response = requests.post(OLLAMA_API_URL, json=payload, timeout=300)
            
            if response.status_code == 200:
                raw_content = response.json().get("response", "").strip()
                if "Instructions:" in raw_content:
                    raw_content = raw_content.split("Instructions:")[-1]
                cleaned = clean_ocr_text(raw_content)
                extracted_pages.append((page_num, cleaned))
                print(f"   [Step 3] Page {page_num} Processed.")
            
            del img, img_str, buffered, images
            gc.collect()
            
        except Exception as e:
            print(f"   [Error] Page {page_num}: {e}")
    
    return extracted_pages


def clean_ocr_text(text):
    """Clean OCR extracted text"""
    if not text:
        return ""
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n', '\n\n', text)
    return text.strip()


def main():
    print(f"--- OCR Processing (Cross-Platform Mode) ---")
    print(f"Platform: {platform.system()} {platform.release()}")
    print(f"Source: {SOURCE_DIR}")
    print(f"Output: {OUTPUT_DIR}")
    print(f"Page Config: {PAGE_CONFIG}")
    print(f"Document Type: {DOC_TYPE}")
    
    if not check_ollama_connection():
        print("[ERROR] Cannot connect to Ollama. Please ensure Ollama is running.")
        return
    
    # Load templates
    templates = load_templates()
    if templates:
        available_types = list(templates.get("templates", {}).keys())
        print(f"Loaded templates: {available_types}")
    
    vendor_df = load_vendor_master()
    data_rows = []
    
    if not os.path.exists(SOURCE_DIR):
        print(f"[ERROR] Source directory not found: {SOURCE_DIR}")
        return
    
    files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".pdf")]
    
    if not files:
        print("No PDF files found.")
        return
    
    for filename in files:
        file_path = os.path.join(SOURCE_DIR, filename)
        print(f"\n[File] {filename}")
        try:
            reader = PdfReader(file_path)
            ocr_results = extract_text_from_image(
                file_path,
                get_target_pages(PAGE_CONFIG, len(reader.pages))
            )
            
            for p_num, raw_text in ocr_results:
                # Save raw OCR text
                txt_path = os.path.join(OUTPUT_DIR, f"{os.path.splitext(filename)[0]}_page{p_num}.txt")
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write(raw_text)
                
                # Parse using templates
                parsed = parse_ocr_data_with_template(raw_text, templates, DOC_TYPE)
                
                print(f"   Detected Type: {parsed['document_type_name']}")
                
                row_data = {
                    "Link PDF": f'=HYPERLINK("{file_path}", "{filename}")',
                    "Page": p_num,
                    "Document Type": parsed["document_type_name"],
                    "VendorID_OCR": parsed["tax_id"],
                    "Branch_OCR": parsed["branch"],
                    "Document No": parsed["document_no"],
                    "Date": parsed["date"],
                    "Amount": parsed["amount"],
                }
                
                # Add extra fields from template
                for field_name, value in parsed.get("extra_fields", {}).items():
                    label = field_name.replace("_", " ").title()
                    row_data[label] = value
                
                data_rows.append(row_data)
                
        except Exception as e:
            print(f"   [Error] {filename}: {e}")

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
        
        # Reorder columns
        priority_cols = [
            "Link PDF", "Page", "Document Type",
            "VendorID_OCR", "Branch_OCR", "Vendor code", "ชื่อบริษัท",
            "Document No", "Date", "Amount"
        ]
        
        all_cols = df.columns.tolist()
        final_cols = [col for col in priority_cols if col in all_cols]
        final_cols += [col for col in all_cols if col not in final_cols]
        
        df = df[final_cols]
        
        excel_path = os.path.join(OUTPUT_DIR, "summary_ocr_local.xlsx")
        df.to_excel(excel_path, index=False)
        print(f"\n[Success] Created Excel: {excel_path}")
        print(f"Total rows: {len(df)}")


if __name__ == "__main__":
    main()
