import os
import io
import re
import json
import time
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from pypdf import PdfReader, PdfWriter
import pdfplumber
from google import genai
from google.genai import types
import docx

def normalize_kc_string(text):
    """Strips numeric/KC prefixes such as 'KC1. ', 'KC 1: ', '1. ', 'KC-1 ' for semantic topic alignment."""
    cleaned = re.sub(r'^(?:KC\s*\d+[\.\:\-\s]*|\d+[\.\:\-\s]+)', '', str(text).strip(), flags=re.IGNORECASE).strip()
    return cleaned if cleaned else str(text).strip()

def match_kc_topic(topic_raw, question_no, master_kcs):
    """
    Smart KC Resolver:
    1. Exact & Normalized Match: Matches returned topic string to master KCs (handling prefixes like KC1. / KC 1:).
    2. Case-Insensitive Match: Matches 'density' -> 'Density'.
    3. Normalized Prefix Match: Matches 'Basics of respiration' -> 'KC1. Basics of respiration, its types'.
    4. Index Match: Matches 'KC1' or 'KC 1' to master KC starting with 'KC1' / 'KC 1' with word boundaries (avoiding KC10 collision).
    5. Substring Match: Matches normalized substring keywords between model topic and master KCs.
    6. Miscellaneous Reason Labeler: Ensures every true Miscellaneous item has an explicit bracketed sub-reason in question_no:
       - (Concept missing from syllabus list)
       - (Unreadable / Damaged text in PDF)
    7. Clean Question Identifier: Strips any embedded page tags since page numbers now reside in dedicated columns.
    """
    q_no_str = str(question_no).strip()
    # Strip any existing page tags like (Page 123) or [Page 123]
    q_no_str = re.sub(r'\s*[\(\[]\s*[Pp]age\s*\d+\s*[\)\]]', '', q_no_str, flags=re.IGNORECASE).strip()
    
    # Check if topic_raw is empty, null, or mistakenly contains a question identifier
    raw_str = str(topic_raw or '').strip()
    is_echoed_qno = bool(
        raw_str and (
            raw_str.lower() == q_no_str.lower() or
            re.search(r'-(?:Q\d+|Section-[A-Z]-Q\d+|Part-[A-Z]-Q\d+)$', raw_str, re.IGNORECASE) or
            re.match(r'^(?:Q\d+|Question\s*\d+)$', raw_str, re.IGNORECASE)
        )
    )

    if not topic_raw or raw_str in ["0", "0.0", "null", "None", "Unmapped", ""] or is_echoed_qno:
        q_text_lower = q_no_str.lower()
        for kc in master_kcs:
            kc_clean_lower = normalize_kc_string(kc).lower()
            if kc_clean_lower in q_text_lower or kc.lower() in q_text_lower:
                return kc, q_no_str
                
        # Unmapped fallback -> Label explicit reason if not already present
        if "concept missing" not in q_text_lower and "unreadable" not in q_text_lower:
            if "damaged" in q_text_lower or "blur" in q_text_lower or "cut off" in q_text_lower:
                reason_tag = " (Unreadable / Damaged text in PDF)"
            else:
                reason_tag = " (Concept missing from syllabus list)"
            q_no_str += reason_tag
                
        return "Miscellaneous", q_no_str
        
    topic_str = str(topic_raw).strip()
    topic_clean = normalize_kc_string(topic_str)
    topic_lower = topic_str.lower()
    topic_clean_lower = topic_clean.lower()
    
    # 1. Exact Match
    for kc in master_kcs:
        if topic_str == kc:
            return kc, q_no_str
            
    # 2. Case-Insensitive Match
    for kc in master_kcs:
        if topic_lower == kc.lower():
            return kc, q_no_str

    # 3. Normalized Prefix Match (e.g. 'Basics of respiration' -> 'KC1. Basics of respiration, its types')
    for kc in master_kcs:
        kc_clean = normalize_kc_string(kc)
        if topic_clean_lower == kc_clean.lower():
            return kc, q_no_str

    # 4. KC Index Match (e.g. 'KC1' -> 'KC1. Basics...' with word boundary to avoid 'KC10' collision)
    kc_idx_match = re.match(r'^KC\s*(\d+)$', topic_str.strip(), re.IGNORECASE)
    if kc_idx_match:
        kc_num = kc_idx_match.group(1)
        for kc in master_kcs:
            if re.match(rf'^KC\s*{kc_num}\b', kc.strip(), re.IGNORECASE):
                return kc, q_no_str

    # 5. Normalized Substring Match (length >= 3 to avoid noise)
    if len(topic_clean_lower) >= 3:
        for kc in master_kcs:
            kc_clean_lower = normalize_kc_string(kc).lower()
            if topic_clean_lower in kc_clean_lower or kc_clean_lower in topic_clean_lower:
                return kc, q_no_str

    # 6. Raw Substring Match
    for kc in master_kcs:
        if (len(topic_lower) >= 3 and topic_lower in kc.lower()) or (len(kc) >= 3 and kc.lower() in topic_lower):
            return kc, q_no_str
            
    # 7. Fallback: If topic does not match any Master KC, route cleanly to Miscellaneous
    q_text_lower = q_no_str.lower()
    if "concept missing" not in q_text_lower and "unreadable" not in q_text_lower:
        reason_tag = " (Concept missing from syllabus list)"
        q_no_str += reason_tag
    return "Miscellaneous", q_no_str

def get_user_authorized_models(api_key):
    """
    Returns verified production text generation models supported by Gemini API.
    Prioritizes latest production Flash models (gemini-3.8-flash, gemini-3.6-flash, gemini-3.5-flash-lite, gemini-flash-latest).
    Strictly filters out non-text, audio, transcription, and embedding models.
    """
    try:
        client = genai.Client(api_key=api_key)
        available_names = set()
        for m in client.models.list():
            actions = getattr(m, 'supported_actions', None) or []
            if actions and 'generateContent' not in actions:
                continue
            clean_name = m.name.replace("models/", "").strip()
            name_lower = clean_name.lower()
            if any(bad in name_lower for bad in ["transcribe", "embed", "imagen", "tts", "audio", "speech", "gemma", "exp", "thinking"]):
                continue
            available_names.add(clean_name)
            
        priority_order = [
            "gemini-flash-lite-latest",
            "gemini-3.5-flash-lite",
            "gemini-3.8-flash",
            "gemini-3.6-flash",
            "gemini-flash-latest",
            "gemini-2.5-flash",
            "gemini-2.0-flash"
        ]
        
        selected = [m for m in priority_order if m in available_names]
        if selected:
            return selected
        return ["gemini-flash-lite-latest", "gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-3.6-flash", "gemini-flash-latest"]
    except Exception as e:
        print(f"Error querying client.models.list(): {e}")
        return ["gemini-flash-lite-latest", "gemini-3.5-flash-lite", "gemini-3.8-flash", "gemini-3.6-flash", "gemini-flash-latest"]

def extract_kcs_from_docx(docx_file_source):
    """
    Extracts Knowledge Component (KC) names from an uploaded Word document (.docx).
    Supports:
    1. Table-based layouts (with columns for KC No, Topic, List of KC, etc.)
    2. Paragraph / bullet-list based syllabus layouts (e.g. 'KC 1. Introduction', '1. Topic', etc.)
    """
    if hasattr(docx_file_source, 'seek'):
        docx_file_source.seek(0)
    doc = docx.Document(docx_file_source)
    kcs_list = []

    # Strategy 1: Check Tables first
    if doc.tables:
        for table in doc.tables:
            if not table.rows:
                continue
            
            kc_no_col = None
            topic_col = None
            start_row = 0
            
            # Scan top rows for headers
            for r_idx, row in enumerate(table.rows[:5]):
                row_texts = [c.text.strip().lower() for c in row.cells]
                row_kc_no = None
                row_explicit_topic = None
                row_kc_standalone = None
                for c_idx, cell_text in enumerate(row_texts):
                    if not cell_text:
                        continue
                    if ('kc' in cell_text and ('no' in cell_text or '#' in cell_text or 'num' in cell_text)) or cell_text in ['s.no', 'sl.no', 'sr.no', 's no', 'no.', '#']:
                        row_kc_no = c_idx
                    elif cell_text in ['topic', 'kc topic', 'list of kc', 'knowledge components', 'kc name', 'kc names', 'topics', 'syllabus', 'list of kcs', 'kc list']:
                        row_explicit_topic = c_idx
                    elif cell_text in ['kc', 'kcs']:
                        row_kc_standalone = c_idx
                if row_explicit_topic is not None or row_kc_standalone is not None:
                    if row_explicit_topic is not None:
                        topic_col = row_explicit_topic
                        kc_no_col = row_kc_no if row_kc_no is not None else row_kc_standalone
                    else:
                        topic_col = row_kc_standalone
                        kc_no_col = row_kc_no
                    start_row = r_idx + 1
                    break
            
            if topic_col is None and len(table.columns) >= 2:
                first_c0 = table.rows[0].cells[0].text.strip().lower()
                first_c1 = table.rows[0].cells[1].text.strip().lower()
                if 'kc' in first_c0 or 'no' in first_c0 or 'topic' in first_c1 or '#' in first_c0:
                    kc_no_col = 0
                    topic_col = 1
                    start_row = 1
                else:
                    kc_no_col = 0
                    topic_col = 1
                    start_row = 0
            elif topic_col is None and len(table.columns) == 1:
                topic_col = 0
                start_row = 1 if ('topic' in table.rows[0].cells[0].text.strip().lower() or 'kc' in table.rows[0].cells[0].text.strip().lower()) else 0

            # Extract from table
            if topic_col is not None:
                for row in table.rows[start_row:]:
                    cells = row.cells
                    if len(cells) <= topic_col:
                        continue
                    c_topic = cells[topic_col].text.strip()
                    c_kc_no = cells[kc_no_col].text.strip() if (kc_no_col is not None and len(cells) > kc_no_col) else ""
                    
                    if not c_topic or c_topic.lower() in ['none', 'topic', 'list of kc', 'unnamed', 'null', 'kc', 'kcs', 'kc list']:
                        continue
                        
                    if c_kc_no and c_kc_no.lower() not in ['none', 'null', 'kc no.', 'kc no', '#']:
                        clean_kc_no = re.sub(r'\s+', ' ', c_kc_no).strip()
                        clean_topic = re.sub(r'\s+', ' ', c_topic).strip()
                        if not clean_topic.lower().startswith(clean_kc_no.lower()):
                            kcs_list.append(f"{clean_kc_no}. {clean_topic}")
                        else:
                            kcs_list.append(clean_topic)
                    else:
                        kcs_list.append(re.sub(r'\s+', ' ', c_topic).strip())

    # Strategy 2: If tables yielded nothing or document is paragraph/bullet-based
    if not kcs_list:
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            
            # Skip document titles and headings if they contain generic syllabus header words
            text_lower = text.lower()
            if any(h in text_lower for h in ['syllabus list', 'master kc', 'knowledge components list', 'table of contents', 'list of kc']):
                continue
            if getattr(p, 'style', None) and getattr(p.style, 'name', '') in ['Title', 'Subtitle', 'Heading 1'] and not re.match(r'^(?:KC|\d+[\.\:])', text, re.IGNORECASE):
                continue
                
            # Match KC patterns: "KC 1. Topic", "KC-1: Topic", "1. Topic", "• Topic", or plain topic text
            kc_match = re.match(r'^(?:KC\s*\d+[\.\:\-\s]+|\d+[\.\:\-\s]+|[\•\-\*]\s*)(.+)$', text, re.IGNORECASE)
            if kc_match or text.lower().startswith("kc"):
                clean_item = re.sub(r'\s+', ' ', text).strip()
                kcs_list.append(clean_item)
            elif len(text) > 3 and not text.isupper() and len(text.split()) <= 20:
                kcs_list.append(re.sub(r'\s+', ' ', text).strip())

    # Deduplicate while preserving sequence
    seen = set()
    unique_kcs = []
    for item in kcs_list:
        clean_item = re.sub(r'\s+', ' ', str(item)).strip()
        if clean_item and clean_item not in seen and clean_item.lower() not in ['none', 'list of kc', 'topic', 'kc topic', 'unnamed', 'kc', 'kcs', 'kc list']:
            seen.add(clean_item)
            unique_kcs.append(clean_item)
            
    return unique_kcs

def extract_kcs_from_csv(csv_file_source):
    """Extracts Knowledge Component (KC) names from an uploaded CSV file."""
    if hasattr(csv_file_source, 'seek'):
        csv_file_source.seek(0)
    df = pd.read_csv(csv_file_source)
    kc_list = []
    for col in df.columns:
        col_values = df[col].dropna().astype(str).str.strip().tolist()
        for idx, val in enumerate(col_values):
            if val.lower() in ['list of kc', 'kc topic', 'kc name', 'knowledge components', 'topic', 'topics', 'kc list', 'list of kcs', 'kc', 'kcs']:
                kc_list = col_values[idx + 1:]
                break
        if kc_list:
            break
            
    if not kc_list:
        for col in reversed(df.columns):
            if 'kc' in str(col).lower() or 'topic' in str(col).lower():
                kc_list = df[col].dropna().astype(str).str.strip().tolist()
                break

    clean_kcs = []
    for kc in kc_list:
        kc_clean = re.sub(r'\s+', ' ', str(kc)).strip()
        if kc_clean and kc_clean.lower() not in ['list of kc', 'kc topic', 'kc name', 'unnamed', 'topic', 'topics', 'kc list', 'list of kcs', 'kc', 'kcs']:
            if not kc_clean.isdigit():
                clean_kcs.append(kc_clean)
                
    seen = set()
    unique_kcs = []
    for item in clean_kcs:
        if item not in seen:
            seen.add(item)
            unique_kcs.append(item)
    return unique_kcs

def extract_kcs_from_file(skeleton_file_source, filename=""):
    """
    Universal multi-format Master KC extractor:
    Accepts Word (.docx), Excel (.xlsx, .xls), or CSV (.csv) files.
    """
    file_name_str = (filename or (skeleton_file_source if isinstance(skeleton_file_source, str) else getattr(skeleton_file_source, 'name', ''))).lower()
    
    if file_name_str.endswith('.docx') or file_name_str.endswith('.doc'):
        return extract_kcs_from_docx(skeleton_file_source)
    elif file_name_str.endswith('.csv'):
        return extract_kcs_from_csv(skeleton_file_source)
    else:
        try:
            return extract_kcs_from_skeleton(skeleton_file_source)
        except Exception:
            try:
                return extract_kcs_from_docx(skeleton_file_source)
            except Exception:
                return extract_kcs_from_csv(skeleton_file_source)

def extract_kcs_from_skeleton(skeleton_file_source):
    """
    Flexibly extracts Knowledge Component (KC) names from an uploaded skeleton Excel sheet.
    Searches for 'List of KC', 'KC Topic', or 'Topic' headers across columns.
    Handles multi-line wrapped cells and joined (KC No. + Topic) layouts cleanly.
    """
    try:
        if hasattr(skeleton_file_source, 'seek'):
            skeleton_file_source.seek(0)
        wb = openpyxl.load_workbook(skeleton_file_source, data_only=True)
        ws = wb.active
        
        kc_no_col = None
        topic_col = None
        start_row = 1

        for r in range(1, min(15, ws.max_row + 1)):
            row_kc_no = None
            row_explicit_topic = None
            row_kc_standalone = None
            
            for c in range(1, min(15, ws.max_column + 1)):
                val = str(ws.cell(row=r, column=c).value or '').strip().lower()
                if not val:
                    continue
                if ('kc' in val and ('no' in val or '#' in val or 'num' in val)) or val in ['s.no', 'sl.no', 'sr.no', 's no', 'no.', '#']:
                    row_kc_no = c
                elif val in ['topic', 'kc topic', 'list of kc', 'knowledge components', 'kc name', 'kc names', 'topics', 'syllabus', 'list of kcs', 'kc list']:
                    row_explicit_topic = c
                elif val in ['kc', 'kcs']:
                    row_kc_standalone = c
            
            if row_explicit_topic is not None or row_kc_standalone is not None:
                if row_explicit_topic is not None:
                    topic_col = row_explicit_topic
                    kc_no_col = row_kc_no if row_kc_no is not None else row_kc_standalone
                else:
                    topic_col = row_kc_standalone
                    kc_no_col = row_kc_no
                start_row = r + 1
                break

        kcs_list = []
        if topic_col:
            current_kc_no = None
            current_topic_parts = []
            
            for r in range(start_row, ws.max_row + 1):
                raw_c1 = ws.cell(row=r, column=kc_no_col).value if kc_no_col else None
                c1_val = str(raw_c1).strip() if raw_c1 is not None else ''
                if c1_val.lower() in ['none', 'null', 'nan']:
                    c1_val = ''
                
                raw_c2 = ws.cell(row=r, column=topic_col).value if topic_col else None
                c2_val = str(raw_c2).strip() if raw_c2 is not None else ''
                if c2_val.lower() in ['none', 'null', 'nan', 'unnamed']:
                    c2_val = ''
                
                # Check if this row represents a new KC
                if kc_no_col:
                    is_num_or_kc = bool(re.match(r'^(?:KC\s*[-.]*\s*)?\d+(?:\.0*)?$', c1_val, re.IGNORECASE))
                    is_new_kc = is_num_or_kc and bool(c2_val)
                else:
                    is_new_kc = bool(c2_val and c2_val.lower() not in ['none', 'topic', 'list of kc', 'kc', 'kcs', 'kc topic', 'kc list'])
                
                if is_new_kc:
                    if current_topic_parts:
                        full_topic = re.sub(r'\s+', ' ', ' '.join(current_topic_parts)).strip()
                        if current_kc_no:
                            if full_topic.lower().startswith(current_kc_no.lower()) or re.match(r'^(?:KC\s*\d+|\d+)[\.\:\-\s]+', full_topic, re.IGNORECASE):
                                kcs_list.append(full_topic)
                            else:
                                kcs_list.append(f"{current_kc_no}. {full_topic}")
                        else:
                            kcs_list.append(full_topic)
                    
                    if kc_no_col and c1_val:
                        if re.match(r'^KC', c1_val, re.IGNORECASE):
                            clean_no = re.sub(r'\s+', '', c1_val).rstrip('.')
                        else:
                            num = int(float(c1_val)) if re.match(r'^\d+(?:\.0*)?$', c1_val) else c1_val
                            clean_no = f"KC{num}"
                        current_kc_no = clean_no
                    else:
                        current_kc_no = None
                    current_topic_parts = [c2_val] if c2_val else []
                elif c2_val:
                    if current_topic_parts:
                        current_topic_parts.append(c2_val)
                    else:
                        current_topic_parts = [c2_val]
                        
            if current_topic_parts:
                full_topic = re.sub(r'\s+', ' ', ' '.join(current_topic_parts)).strip()
                if current_kc_no:
                    if full_topic.lower().startswith(current_kc_no.lower()) or re.match(r'^(?:KC\s*\d+|\d+)[\.\:\-\s]+', full_topic, re.IGNORECASE):
                        kcs_list.append(full_topic)
                    else:
                        kcs_list.append(f"{current_kc_no}. {full_topic}")
                else:
                    kcs_list.append(full_topic)

        if kcs_list:
            seen = set()
            unique_kcs = []
            for item in kcs_list:
                clean_item = re.sub(r'\s+', ' ', str(item)).strip()
                if clean_item and clean_item not in seen and clean_item.lower() not in ['none', 'list of kc', 'topic', 'kc topic', 'unnamed', 'kc', 'kcs', 'kc list']:
                    seen.add(clean_item)
                    unique_kcs.append(clean_item)
            return unique_kcs

    except Exception as e:
        print(f"openpyxl direct KC extraction fallback: {e}")

    # Fallback to pandas dataframe search
    if hasattr(skeleton_file_source, 'seek'):
        skeleton_file_source.seek(0)
    df = pd.read_excel(skeleton_file_source)
    kc_list = []
    
    for col in df.columns:
        col_values = df[col].dropna().astype(str).str.strip().tolist()
        for idx, val in enumerate(col_values):
            if val.lower() in ['list of kc', 'kc topic', 'kc name', 'knowledge components', 'topic', 'topics', 'kc list', 'list of kcs', 'kc', 'kcs']:
                kc_list = col_values[idx + 1:]
                break
        if kc_list:
            break
            
    if not kc_list:
        for col in reversed(df.columns):
            if 'kc' in str(col).lower() or 'topic' in str(col).lower():
                kc_list = df[col].dropna().astype(str).str.strip().tolist()
                if kc_list and kc_list[0].lower() in ['list of kc', 'kc topic', 'kc name', 'topic', 'topics', 'kc list', 'list of kcs', 'kc', 'kcs']:
                    kc_list = kc_list[1:]
                break

    clean_kcs = []
    for kc in kc_list:
        kc_clean = re.sub(r'\s+', ' ', str(kc)).strip()
        if kc_clean and kc_clean.lower() not in ['list of kc', 'kc topic', 'kc name', 'unnamed', 'topic', 'topics', 'kc list', 'list of kcs', 'kc', 'kcs', 'allen']:
            if not kc_clean.isdigit():
                clean_kcs.append(kc_clean)
                
    seen = set()
    unique_kcs = []
    for item in clean_kcs:
        if item not in seen:
            seen.add(item)
            unique_kcs.append(item)
            
    return unique_kcs

def extract_pages_from_docx(docx_file_source):
    """
    Extracts text content from an uploaded Word document (.docx) and structures it into
    virtual pages (~350-400 words) for the fast digital text-mode categorization pipeline.
    """
    if hasattr(docx_file_source, 'seek'):
        docx_file_source.seek(0)
    doc = docx.Document(docx_file_source)
    
    content_blocks = []
    for block in doc.paragraphs:
        txt = block.text.strip()
        if txt:
            content_blocks.append(txt)
            
    for table in doc.tables:
        table_lines = []
        for row in table.rows:
            row_vals = [c.text.strip() for c in row.cells if c.text.strip()]
            if row_vals:
                table_lines.append(" | ".join(row_vals))
        if table_lines:
            content_blocks.append("\n".join(table_lines))
            
    if not content_blocks:
        return [], 0, False, 1

    pages_data = []
    current_page_text = []
    current_word_count = 0
    page_counter = 1
    
    WORDS_PER_VIRTUAL_PAGE = 350
    
    for block in content_blocks:
        block_words = len(block.split())
        if current_word_count > 0 and (current_word_count + block_words > WORDS_PER_VIRTUAL_PAGE):
            pages_data.append({
                'page_num': page_counter,
                'printed_page': page_counter,
                'text': "\n\n".join(current_page_text)
            })
            page_counter += 1
            current_page_text = [block]
            current_word_count = block_words
        else:
            current_page_text.append(block)
            current_word_count += block_words
            
    if current_page_text:
        pages_data.append({
            'page_num': page_counter,
            'printed_page': page_counter,
            'text': "\n\n".join(current_page_text)
        })
        
    total_pages = len(pages_data)
    is_scanned_vision_mode = False
    detected_start_page = 1
    
    return pages_data, total_pages, is_scanned_vision_mode, detected_start_page

def extract_pages_from_document(file_source, filename=""):
    """
    Universal Question Paper Document Loader:
    Detects whether the file is a Word document (.docx/.doc) or PDF (.pdf) and
    dispatches to the optimal parser.
    """
    file_name_str = (filename or (file_source if isinstance(file_source, str) else getattr(file_source, 'name', ''))).lower()
    
    if file_name_str.endswith('.docx') or file_name_str.endswith('.doc'):
        return extract_pages_from_docx(file_source)
    else:
        try:
            return extract_pages_from_pdf(file_source)
        except Exception:
            return extract_pages_from_docx(file_source)

def extract_pages_from_pdf(pdf_file_source):
    """
    Fast, optimized PDF text extraction and layout classifier:
    - Performs an instant sample scan on the first 5 pages using pypdf.
    - If average characters < 150 (Scanned PDF), immediately activates Scanned Vision Mode in ~1 second.
    - If digital text PDF, runs full text extraction and sequence intelligence margin detection.
    """
    if hasattr(pdf_file_source, 'seek'):
        pdf_file_source.seek(0)
    reader = PdfReader(pdf_file_source)
    total_pages = len(reader.pages)
    
    if total_pages == 0:
        return [], 0, False, 1

    # Fast 5-page sample character density check (<1 second)
    sample_pages = min(5, total_pages)
    sample_chars = 0
    for i in range(sample_pages):
        try:
            txt = reader.pages[i].extract_text() or ""
            sample_chars += len(txt.strip())
        except Exception:
            pass

    avg_sample_chars = sample_chars / max(1, sample_pages)
    is_scanned_vision_mode = (avg_sample_chars < 150)

    if is_scanned_vision_mode:
        print(f"⚡ Fast PDF Classifier: Scanned Document Detected ({avg_sample_chars:.1f} avg chars/page across sample). Activating Vision Mode in ~1s.", flush=True)
        return [], total_pages, True, 1

    # Digital Text PDF -> Extract all pages with pypdf
    pages_data = []
    for idx, page in enumerate(reader.pages):
        try:
            page_text = page.extract_text() or ""
            pages_data.append({
                'page_num': idx + 1,
                'text': page_text
            })
        except Exception:
            pages_data.append({
                'page_num': idx + 1,
                'text': ""
            })

    # --- Pure Consecutive Chain Sequence Intelligence Engine (10-Page Window) for Text PDFs ---
    scan_limit = min(10, total_pages)
    page_margin_candidates = []
    for p in pages_data[:scan_limit]:
        lines = [line.strip() for line in p['text'].split('\n') if line.strip()]
        margin_lines = lines[:5] + lines[-5:] if lines else []
        combined = " ".join(margin_lines)
        raw_numbers = [int(n) for n in re.findall(r'\b(\d{1,4})\b', combined) if 0 <= int(n) <= 5000]
        page_margin_candidates.append({
            'page_num': p['page_num'],
            'numbers': set(raw_numbers)
        })

    # Extract all candidate start offsets
    potential_offsets = set()
    for c in page_margin_candidates:
        p_num = c['page_num']
        for n in c['numbers']:
            offset = n - (p_num - 1)
            if offset >= 0:
                potential_offsets.add(offset)

    # Score consecutive chain length for each candidate offset
    detected_start_page = 1
    max_chain_length = 0
    found_sequence = False

    for O in potential_offsets:
        max_run = 0
        curr_run = 0
        for c in page_margin_candidates:
            expected_num = O + (c['page_num'] - 1)
            if expected_num in c['numbers']:
                curr_run += 1
                if curr_run > max_run:
                    max_run = curr_run
            else:
                curr_run = 0
        if max_run > max_chain_length:
            max_chain_length = max_run
            detected_start_page = O

    if max_chain_length >= 2:
        found_sequence = True
        print(f"Sequence Intelligence locked Global Start Offset={detected_start_page} with {max_chain_length}-page consecutive chain!", flush=True)
    else:
        # Fallback to single-pair matching if no chain >= 2
        for i in range(len(page_margin_candidates) - 1):
            curr_p = page_margin_candidates[i]['page_num']
            curr_nums = page_margin_candidates[i]['numbers']
            next_p = page_margin_candidates[i+1]['page_num']
            next_nums = page_margin_candidates[i+1]['numbers']
            page_diff = next_p - curr_p
            for n1 in curr_nums:
                for n2 in next_nums:
                    if n2 - n1 == page_diff:
                        start_offset = max(0, n1 - (curr_p - 1))
                        detected_start_page = start_offset
                        found_sequence = True
                        break
                if found_sequence:
                    break
            if found_sequence:
                break

    # Per-Page Map Construction based on pure sequence lock
    per_page_printed_map = {}
    
    if found_sequence:
        # Pure sequence baseline: every page p gets (detected_start_page + p - 1)
        for idx in range(total_pages):
            pdf_p = idx + 1
            per_page_printed_map[pdf_p] = max(0, detected_start_page + (pdf_p - 1))
    else:
        # Fallback if no sequence matched: default 1-to-1 page numbering
        for idx in range(total_pages):
            pdf_p = idx + 1
            per_page_printed_map[pdf_p] = pdf_p

    # Attach per_page_printed_map to pages_data
    for p in pages_data:
        p['printed_page'] = per_page_printed_map.get(p['page_num'], p['page_num'])

    return pages_data, total_pages, is_scanned_vision_mode, detected_start_page

def get_pdf_chunk_bytes(pdf_file_source, start_idx, end_idx):
    """
    Creates a binary PDF chunk slice for Gemini Multimodal Vision API.
    """
    if hasattr(pdf_file_source, 'seek'):
        pdf_file_source.seek(0)
    reader = PdfReader(pdf_file_source)
    writer = PdfWriter()
    for i in range(start_idx, min(end_idx, len(reader.pages))):
        writer.add_page(reader.pages[i])
        
    chunk_io = io.BytesIO()
    writer.write(chunk_io)
    chunk_io.seek(0)
    return chunk_io.getvalue()

# ==============================================================================
# INTELLIGENT SECTION PRE-FILTER & BOUNDARY SLICER
# ==============================================================================

SOLUTION_HEADER_PATTERNS = [
    r'^\s*[:\-\s]*\b(?:ANSWER\s*KEY|ANSWERS|KEY)\b[:\-\s]*$',
    r'^\s*[:\-\s]*\b(?:HINTS?\s*(?:AND|&)\s*SOLUTIONS?|SOLUTIONS?|HINTS?|EXPLANATIONS?)\b[:\-\s]*$',
    r'^\s*[:\-\s]*\b(?:ANSWERS?\s*(?:AND|&)\s*SOLUTIONS?)\b[:\-\s]*$',
    r'^\s*[:\-\s]*\b(?:DETAILED\s*SOLUTIONS?|STEP\s*BY\s*STEP\s*SOLUTIONS?)\b[:\-\s]*$',
]

QUESTION_HEADER_PATTERNS = [
    r'^\s*[:\-\s]*(?:EXERCISE(?:\s*[-–—:]?\s*(?:\d+|[A-Z]))?)\s*[:\-\s]*$',
    r'^\s*[:\-\s]*(?:SECTION\b(?:\s*[-–—:]?\s*[A-Z0-9]+)?|PART\b(?:\s+[-–—:]?\s*[A-Z0-9]+|\s*[-–—:]\s*[A-Z0-9]+)?)\s*[:\-\s]*$',
    r'^\s*[:\-\s]*(?:LEVEL\s*[-–—:]?\s*(?:[IVXLCDM]+|\d+)|ASSIGNMENT\b(?:\s*[-–—:]?\s*\d+)?)\s*[:\-\s]*$',
    r'^\s*[:\-\s]*(?:PRACTICE\s*(?:PROBLEMS|QUESTIONS)|TOPICWISE\s*PRACTICE|QUESTION\s*BANK)\s*[:\-\s]*$',
    r'^\s*[:\-\s]*(?:OBJECTIVE\s*QUESTIONS|SINGLE\s*CORRECT\s*(?:ANSWER\s*)?TYPE|MULTIPLE\s*CHOICE\s*QUESTIONS)\s*[:\-\s]*$',
    r'^\s*[:\-\s]*(?:ASSERTION\s*(?:AND|&|-|–|—)?\s*REASON(?:ING)?(?:\s*TYPE)?)\s*[:\-\s]*$',
]

def is_solution_delimiter(line):
    line_clean = line.strip()
    if not line_clean or len(line_clean.split()) > 7:
        return False
    for pat in SOLUTION_HEADER_PATTERNS:
        if re.search(pat, line_clean, re.IGNORECASE):
            return True
    return False

def is_question_delimiter(line):
    line_clean = line.strip()
    if not line_clean or len(line_clean.split()) > 7:
        return False
    for pat in QUESTION_HEADER_PATTERNS:
        if re.search(pat, line_clean, re.IGNORECASE):
            return True
    return False

def is_dense_answer_grid(text):
    """Detects pages that consist almost entirely of an answer key grid (e.g. 1) a 2) d 3) a 4) d)."""
    matches = re.findall(r'\b\d{1,4}\s*[\)\.\-]\s*[a-dA-D]\b', text)
    if len(matches) >= 15:
        words = text.split()
        if len(words) > 0 and (len(matches) * 2 / len(words)) > 0.35:
            return True
    return False

def filter_document_pages(pages_data):
    """
    Intelligent Pre-Filter & Section Slicing:
    Tracks section state transitions line-by-line across pages.
    Excludes standalone Answer Keys, Hints & Solutions, while retaining
    real questions and handling mid-page transitions seamlessly.
    """
    filtered_pages = []
    in_question_mode = True # Default: document starts with questions
    excluded_count = 0
    
    for p in pages_data:
        p_num = p['page_num']
        raw_text = p.get('text', '')
        
        # Fast check: if page is a dense answer key grid, exclude it
        if is_dense_answer_grid(raw_text):
            in_question_mode = False
            excluded_count += 1
            continue
            
        lines = raw_text.split('\n')
        current_page_kept_lines = []
        
        for line in lines:
            # Check for transition to solutions/answers
            if is_solution_delimiter(line):
                in_question_mode = False
                continue
                
            # Check for transition back to questions
            if is_question_delimiter(line):
                in_question_mode = True
                current_page_kept_lines.append(line)
                continue
                
            # If in question mode, keep line
            if in_question_mode:
                current_page_kept_lines.append(line)
                
        kept_text = "\n".join(current_page_kept_lines).strip()
        if kept_text:
            filtered_p = dict(p)
            filtered_p['text'] = kept_text
            filtered_pages.append(filtered_p)
        else:
            excluded_count += 1
            
    return filtered_pages, excluded_count

def check_unclosed_question_at_page_bottom(page_text):
    """
    Analyzes the bottom of a page to determine if it ends with an unclosed question
    (e.g. question stem or partial options without closing options like (c)/(d) or (3)/(4)).
    Returns (is_unclosed, last_q_num, trailing_lines).
    """
    lines = [l.strip() for l in page_text.split('\n') if l.strip()]
    if not lines:
        return False, None, []
        
    q_pat = re.compile(r'^\s*(?:Q(?:uestion)?[\s\.\:\-]*\s*)?(\d{1,4})[\.\)\:\-]\s+[A-Za-z]')
    last_q_line_idx = -1
    last_q_num = None
    for i in range(len(lines) - 1, -1, -1):
        m = q_pat.match(lines[i])
        if m:
            last_q_line_idx = i
            last_q_num = m.group(1)
            break
            
    if last_q_line_idx == -1:
        return False, None, []
        
    trailing_lines = lines[last_q_line_idx:]
    trailing_text = " ".join(trailing_lines)
    has_closing_options = bool(re.search(r'(?:\b[cdCD]\)|\([cdCD]\)|\b[34]\)|\([34]\))', trailing_text))
    return not has_closing_options, last_q_num, trailing_lines

def extract_page_leading_continuation(page_text):
    """
    Extracts leading continuation lines (e.g. options or sub-parts of a question split across pages)
    from the top of a page before the first new question or section starts.
    Preserves top page headers (e.g. Page | X) in the remaining page text.
    Returns (continuation_text, remaining_page_text).
    """
    lines = page_text.split('\n')
    q_start_pat = re.compile(r'^\s*(?:Q(?:uestion)?[\s\.\:\-]*\s*)?(\d{1,4})[\.\)\:\-]\s+[A-Za-z]')
    header_pat = re.compile(r'^\s*P\s*a\s*g\s*e\s*\|\s*\d+\s*$', re.IGNORECASE)
    
    start_idx = 0
    while start_idx < len(lines) and (not lines[start_idx].strip() or header_pat.match(lines[start_idx].strip())):
        start_idx += 1
        
    first_q_idx = -1
    for i in range(start_idx, len(lines)):
        line_clean = lines[i].strip()
        if not line_clean:
            continue
        if q_start_pat.match(line_clean) or is_solution_delimiter(line_clean) or is_question_delimiter(line_clean):
            first_q_idx = i
            break
            
    if first_q_idx != -1 and first_q_idx > start_idx:
        header_lines = lines[:start_idx]
        cont_lines = lines[start_idx:first_q_idx]
        rem_lines = header_lines + lines[first_q_idx:]
        return "\n".join(cont_lines).strip(), "\n".join(rem_lines).strip()
    elif first_q_idx == start_idx:
        return "", page_text.strip()
    else:
        return "", page_text.strip()

def build_gemini_prompt(chunk_pages_or_none, master_kcs, start_page_num=1, end_page_num=1, start_printed_page=1, is_neet=False):
    """
    Constructs the prompt for Gemini API classification based on Stage 1 & Stage 2 rules.
    Maps relative chunk page indices (Slice Page 1, 2, 3...) to Absolute Printed Book Page Numbers per page.
    When is_neet=True, calibrates competency tiers (B/I/A) solely on intrinsic cognitive depth without historical paper benchmarking.
    """
    kcs_formatted = "\n".join([f"- {kc}" for kc in master_kcs])
    
    page_count = (end_page_num - start_page_num + 1)
    mapping_lines = []
    chunk_text = ""
    
    if chunk_pages_or_none:
        for idx, p in enumerate(chunk_pages_or_none):
            printed_p = p.get('printed_page', start_printed_page + (p['page_num'] - 1))
            mapping_lines.append(f"  * Slice Page {idx + 1} = PDF FILE PAGE {p['page_num']} = PRINTED BOOK PAGE {printed_p}")
            chunk_text += f"\n\n================================================================================\n<<< START OF PRINTED BOOK PAGE {printed_p} (PDF FILE PAGE {p['page_num']}) >>>\n================================================================================\n{p['text']}\n"
    else:
        for idx in range(page_count):
            pdf_p_index = start_page_num + idx
            mapping_lines.append(f"  * Slice Page {idx + 1} = PDF FILE PAGE {pdf_p_index}")
            
        page_mapping_table = "\n".join(mapping_lines)
        chunk_text = f"""Perform visual optical character recognition (OCR) and question parsing directly from the attached scanned PDF pages.

### PDF FILE PAGE MAPPING TABLE FOR THIS ATTACHED PDF CHUNK:
{page_mapping_table}

### MANDATORY VISUAL PRINTED PAGE MARGIN OCR:
- For EACH attached page image slice, visually inspect the top headers and bottom footers (outer corners) to find the PRINTED BOOK PAGE NUMBER (e.g. 185, 186, etc.).
- In your JSON response, populate `"page_margin_ocr"` with the detected printed page integer for each slice (e.g. `"slice_1_printed_page": 185`).
- If a specific page image has NO visible page number (e.g. cover page, unnumbered full-page diagram, or cropped margin), set its value to null (e.g. `"slice_2_printed_page": null`). Python will automatically interpolate unnumbered pages using bidirectional (+1 / -1) math.
- For each question in `"questions"`, set `"printed_page"` to the printed book page number of that slice, and `"pdf_page"` to the corresponding PDF FILE PAGE number from the mapping table above.
"""

    if is_neet:
        competency_instructions = """4. **Competency Level (`competency_level`) — PURE COGNITIVE DEPTH CALIBRATION (NEET SYLLABUS MODE):**
   - Assess the complexity of each question SOLELY based upon its intrinsic cognitive depth, reasoning steps, and conceptual demand.
   - **MANDATORY NEET CALIBRATION RULE:** DO NOT look into or benchmark against past / older examination papers to decide question complexity. Assess strictly on intrinsic cognitive depth:
     * `B` (Basic): Direct cognitive recall, single-step factual questions, straightforward definitions, and direct single-step formula substitutions.
     * `I` (Intermediate): Multi-statement conceptual evaluation, Assertion-Reason questions, Match the Following matrices, two-step procedural calculations, and relational concept linking.
     * `A` (Advanced): High cognitive depth — deep analytical & graphical interpretation, multi-concept cross-topic synthesis, multi-step problem solving, and complex conceptual reasoning.
   - DO NOT decide based on book section titles, chapter labels, or printed headings; evaluate the inherent cognitive depth of the question itself."""
    else:
        competency_instructions = """4. **Competency Level (`competency_level`):**
   - Evaluate each question against your extensive knowledge base of recent examination papers (2020 to 2024 or higher). DO NOT decide based on book section titles, chapter labels, or printed headings; evaluate the inherent standard and cognitive depth of the question itself:
     * `B` (Basic): Class Board Exam Standard (CBSE / State Board level, 2020 to 2024 or higher — direct recall, single-step factual questions, standard definition and formula applications).
     * `I` (Intermediate): National Competitive Exam Standard (JEE Main level, 2020 to 2024 or higher — multi-statement questions, Assertion-Reason, Match the Following, 2-step numerical applications).
     * `A` (Advanced): Premier Competitive Exam Standard (JEE Advanced / Olympiad level, 2020 to 2024 or higher — deep analytical & graphical interpretation, multi-concept cross-topic synthesis, multi-correct verification, complex multi-step reasoning).
   - If a question is an original practice problem that did not appear in a past exam, evaluate and categorize it based on its equivalent cognitive demand matching the above tiers."""

    prompt = f"""You are an expert educational content analyst and curriculum designer.
Your task is to exhaustively parse, classify, and map EVERY SINGLE QUESTION from the provided PDF content (text or scanned document images) into a structured JSON dataset matching the preloaded Master Knowledge Components (KCs).

### MASTER KNOWLEDGE COMPONENTS (KCs) LIST:
{kcs_formatted}

---

### CORE SEMANTIC EXTRACTION PRINCIPLE (UNIVERSAL):
1. **Sentence-by-Sentence Semantic Question Detection:**
   - Read the document text sequentially. Any standalone question, exercise problem, practice item, numbered question stem, conceptual evaluation task, calculation, True/False statement, Fill-in-the-blank, or assessment problem is semantically a **QUESTION**.
   - While standard exercise titles (e.g. Exercise-1, Section A, Practice Problems) do not dictate question conformity, you MUST STRICTLY EXCLUDE items appearing under section headers such as 'ANSWER KEY', 'HINTS & SOLUTIONS', 'HINTS AND SOLUTIONS', 'SOLUTIONS', 'ANSWERS', or 'EXPLANATIONS'.
   - If an inquiry/question in an exercise, practice problem, or numbered illustration is followed by a solution text (e.g. `Sol.`), you MUST extract ONLY the **question stem** itself.

2. **The ONLY Content to Exclude (Not Questions):**
   - Pure narrative expository textbook theory paragraphs, lecture rules with bullet points, and inline demonstration examples embedded directly in instructional text (e.g. `Application of I-effect (A) Stability of carbocation: (i) If... (ii) If... Example: Stability order: 3° > 2° > 1° | Solution. Reason: More no. of +I group`). These are lecture textbook explanations, NOT assessment questions.
   - Standalone answer keys or answer tables (pure grids of option numbers/letters like `1) a 2) d 3) a 4) d`).
   - Standalone end-of-module solution / explanation notes (e.g. sections titled 'HINTS AND SOLUTIONS', 'SOLUTIONS', 'EXPLANATIONS', or lines formatted like '120 (b) Biotic potential is...', '119 (c) Ecology is...', '1. Self explanatory', '2. Acidic strength increases...'). These are answer explanations, NOT questions. NEVER extract items from standalone solution sections.

3. **Complete Sub-Question & Multi-Part Extraction:**
   - For multi-part questions, passage-based sub-questions, True/False sequences, or Fill-in-the-blank sequences, extract ALL items sequentially starting from item #1 (e.g. 1, 2, 3... or (i), (ii), (iii)...). Never skip the first or last sub-questions.

4. **Questions Spanning Across Pages & Boundary Continuations:**
   - Questions frequently break across consecutive pages. You MUST stitch them together into single, coherent questions.
   - If a page ends with a question stem or partial options and is followed by a `<<< PAGE TRANSITION CONTINUATION (Page X -> Page Y) >>>` block, those lines complete the question started on Page X. Extract it as a complete question (tagged with Page X).
   - If a page starts directly with options or roman numerals without its own question number (orphan continuation), DO NOT create a phantom question for those lines; they belong to the preceding page's question.

---

### CLASSIFICATION & ROUTING RULES:

1. **UNIVERSAL RULE — PRINTED HEADINGS ARE 100% IMMATERIAL TO QUESTION TYPE:**
   - The printed section header, exercise title, or chapter label under which a question appears (e.g. "PART - I : SUBJECTIVE QUESTIONS", "Exercise-1", "Practice Problems", "Short Answer Type", "Section A") is **100% IMMATERIAL** and must NEVER dictate or bias the question format or conformity!
   - You MUST assess EVERY question purely on its own **intrinsic structural form, options, and response type**:

   * **CONFORMING (`is_conforming: true`) — Strictly standard competitive exam formats:**
     - `MCQ`: Any question that presents distinct, mutually exclusive **Multiple-Choice Answer Options** to choose from with 1 correct choice (e.g., options labeled (A)-(D), (1)-(4), (a)-(d), binary True/False choices, or Match the Column with choice options).
       * **CRITICAL MCQ DISAMBIGUATION (ANSWER OPTIONS VS. IN-STEM TARGET ENTITIES):**
         - Often, a question stem lists multiple molecules, chemical structures, reactions, statements, or test items labeled (1), (2), (3) or (i), (ii), (iii) or (A), (B), (C), (D) or (P), (Q), (R), (S) **inside the problem stem** to be evaluated or arranged (e.g. "Arrange the following carbocations in stability order: (1)... (2)... (3)...", "Which of the following compounds (A), (B), (C), (D) is most basic?", "Find the order of reactivity: (1)... (2)... (3)...").
         - If SEPARATE multiple-choice answer choices are provided to pick from (e.g., Options: `(A) 1 > 2 > 3`, `(B) 3 > 2 > 1`, `(C) 3 > 1 > 2`, `(D) 2 > 3 > 1` OR `(A)`, `(B)`, `(C)`, `(D)` containing alternative choices), it is strictly an **`MCQ`** (`is_conforming: true`).
         - If NO separate multiple-choice answer choices are provided (i.e. the student must write out the arrangement, explanation, or mechanism from scratch), the items in the stem are NOT answer choices! The question is strictly a **`Subjective Question`** (`is_conforming: false`, `non_conforming_reason: "Subjective open-ended question without multiple-choice answer options"`).
       * CRITICAL RULE FOR BLANKS WITH OPTIONS: If a question stem contains a blank (e.g. "Root hairs develop from ______") BUT is followed by multiple-choice options ((1), (2), (3), (4) or (A), (B), (C), (D)), it is strictly an `MCQ` and MUST be marked `is_conforming: true`.
     - `Multi correct`: Any question that presents multiple options / choices where one or more choices can be correct (e.g. options (A)-(D), (a)-(d), extended options (A)-(H), or multi-statement evaluation).
     - `Integer type`: Any question whose answer evaluates to a single integer value from 0 to 9999 inclusive (including single digits 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, or counting questions like "How many of the following compounds...").

   * **NON-CONFORMING (`is_conforming: false`) — All other question formats:**
     - `Subjective Question`: Truly open-ended questions that have ZERO multiple-choice answer options provided and NO integer answer (where the student must write an open response from scratch, e.g. "Derive the formula for...", "Draw the mechanism...", "Explain why...", "Arrange the following compounds in order of reactivity: (1)... (2)... (3)..." without MCQ options, "Write the tautomers of the following...").
     - `Drawing Task`: Any diagram drawing or plotting task with zero options.
     - `Fill-in-the-Blank`: Open-ended fill-in-the-blank questions with NO multiple-choice options provided.
     - `Decimal Numerical`: Any calculation resulting in a decimal/floating-point answer (e.g. 3.14).
     - `Negative Numerical`: Any calculation resulting in a negative answer (e.g. -5).
     - `Continuous Range`: Any answer spanning a continuous range (e.g. 10m to 20m).

2. **KC Topic Alignment & Multi-Concept (Culminating KC) Rule:**
   - Map each question to the most relevant KC from the Master List above.
   - Match topic names EXACTLY as written in the Master List.
   - `kc_topic` MUST be the full text string name from the Master List. NEVER output digits or "0".
   - **STRICT PROHIBITION AGAINST ECHOING QUESTION NUMBERS:** NEVER copy or echo the `question_no` into `kc_topic` (e.g. NEVER output "Exercise-1-Q5" or "Master-Your-Understanding-Q7" as the `kc_topic`). `kc_topic` MUST strictly be one of the syllabus topics from the Master List above.
   - MULTI-CONCEPT RULE: If a question spans multiple concepts, map it to the LAST / CULMINATING concept tested where the final answer or conclusion is reached.
   - MISCELLANEOUS LAST-RESORT RULE: Only map to "Miscellaneous" if the topic is completely absent from the syllabus list.

3. **Mandatory Unique Hierarchical Identifier (`question_no`) & Exact Page Tagging (`printed_page`, `pdf_page`):**
   - **MANDATORY UNIQUE PREFIX REQUIREMENT:** Every `question_no` MUST be uniquely identifiable across the entire book. You MUST prepend the enclosing top-level section name, box name, exercise title, or sub-heading to the question number.
   - **Formula:** `[Enclosing Section/Exercise/Box Name]-[Subsection/Type]-Q[Number]`
   - **UNIVERSAL HEADER RULE:** Derive the prefix strictly from the pedagogical exercise, section, or topic heading (e.g. `Exercise-1`, `Concept-Check`, `Match-The-Column-MCQs`, `Solved-Examples`). Do NOT use repetitive inline citations, book series tags, or footer metadata printed beneath individual questions as prefixes.
   - **STRICT PROHIBITION:** NEVER output bare question numbers like `1`, `2`, `Q1`, `Q2`, `Q3` alone. If a question #1 appears under a box or heading (e.g. "Train Your Brain"), it MUST be named `Train-Your-Brain-Q1`. If question #1 is under "Concept Application" in "True and False", it MUST be named `Concept-Application-True-False-Q1`. If under "Exercise-1", `Exercise-1-Q1`. If under "Exercise-2 Section A", `Exercise-2-Section-A-Q1`.
   - DO NOT append page numbers to `question_no` (page numbers belong strictly in the dedicated `printed_page` and `pdf_page` fields).
   - `printed_page`: Integer value representing the Printed Book Page Number.
   - `pdf_page`: Integer value representing the 1-based PDF File Page Number.
   - EXACT PAGE BOUNDARY RULE FOR MULTI-PART QUESTIONS: Tag each sub-question with the exact page where that specific sub-question is printed based on the `<<< START OF PRINTED BOOK PAGE ... >>>` banners.

{competency_instructions}

---

### MANDATORY JSON OUTPUT FORMAT:
You MUST return a valid JSON object with the exact following schema for every question extracted:
```json
{{
  "page_margin_ocr": {{
    "slice_1_printed_page": 185,
    "slice_2_printed_page": 186,
    "slice_3_printed_page": null,
    "slice_4_printed_page": 188,
    "slice_5_printed_page": 189
  }},
  "questions": [
    {{
      "kc_topic": "Exact Topic Name from Master List",
      "question_no": "Train-Your-Brain-Q1",
      "printed_page": 185,
      "pdf_page": 1,
      "question_type": "Subjective Question",
      "is_conforming": false,
      "non_conforming_reason": "Subjective in-text conceptual question"
    }},
    {{
      "kc_topic": "Exact Topic Name from Master List",
      "question_no": "Concept-Application-True-False-Q1",
      "printed_page": 186,
      "pdf_page": 2,
      "question_type": "MCQ",
      "is_conforming": true,
      "competency_level": "B"
    }},
    {{
      "kc_topic": "Exact Topic Name from Master List",
      "question_no": "Exercise-1-Q4",
      "printed_page": 188,
      "pdf_page": 4,
      "question_type": "MCQ",
      "is_conforming": true,
      "competency_level": "B"
    }}
  ]
}}
```
**CRITICAL REQUIREMENT:** Both `"printed_page"` (integer Printed Book Page Number) and `"pdf_page"` (integer PDF File Page Number) are MANDATORY fields in every question object. You must extract them directly from the `<<< START OF PRINTED BOOK PAGE [X] (PDF FILE PAGE [Y]) >>>` header immediately above that question.

---

### INPUT PDF CONTENT TO PROCESS:
{chunk_text}
"""
    return prompt

def parse_gemini_json_robust(raw_text):
    """
    Safely parses JSON responses from Gemini API.
    Handles None inputs, markdown code fences, unescaped control characters,
    and auto-heals truncated JSON output streams.
    """
    if not raw_text:
        return []
        
    text = str(raw_text).strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
        
    try:
        return json.loads(text)
    except Exception:
        pass
        
    try:
        return json.loads(text, strict=False)
    except Exception:
        pass

    # Auto-heal truncated JSON by closing brackets at last complete object
    last_brace_idx = text.rfind("}")
    if last_brace_idx != -1:
        truncated_slice = text[:last_brace_idx + 1].strip()
        candidate = truncated_slice
        if candidate.count("[") > candidate.count("]"):
            candidate += "]"
        if candidate.count("{") > candidate.count("}"):
            candidate += "}"
        try:
            return json.loads(candidate, strict=False)
        except Exception:
            pass

    # Regex extraction of individual valid question objects
    objects = []
    for m in re.finditer(r"\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", text):
        try:
            obj = json.loads(m.group(0), strict=False)
            if isinstance(obj, dict) and ("kc_topic" in obj or "question_no" in obj):
                objects.append(obj)
        except Exception:
            pass
    if objects:
        return {"questions": objects}

    return []

def resilient_countdown_sleep(seconds, progress_callback=None, message="Cooldown active"):
    """
    Sleeps for the specified seconds while pulsing every 1.0s to update progress_callback
    and maintain the live WebSocket connection without freezing the process.
    """
    t_end = time.time() + float(seconds)
    while time.time() < t_end:
        rem = int(round(t_end - time.time()))
        if progress_callback and rem > 0:
            try:
                progress_callback(None, f"⏳ {message} ({rem}s remaining)...")
            except Exception:
                pass
        time.sleep(min(1.0, max(0.1, t_end - time.time())))

def process_chunk_with_gemini(api_key, candidate_models, chunk_pages_or_bytes, master_kcs, is_vision=False, start_p=1, end_p=1, start_printed_page=1, progress_callback=None, is_neet=False):
    """
    Sends a page chunk to Gemini API (Text Mode or Multimodal Vision Mode).
    Supports single API key or multi-key pool (comma-separated or list) with instant rotation on 429.
    Enforces 100% Python-Driven Deterministic Printed and PDF Page Tagging.
    Returns (questions_list, actual_model_used_name).
    """
    if isinstance(api_key, (list, tuple)):
        key_list = [str(k).strip() for k in api_key if str(k).strip()]
    elif isinstance(api_key, str):
        key_list = [k.strip() for k in api_key.replace(";", ",").split(",") if k.strip()]
    else:
        key_list = [str(api_key).strip()]
        
    if not key_list:
        raise ValueError("No valid API Key provided.")
        
    key_idx = 0
    client = genai.Client(api_key=key_list[key_idx])
    
    if is_vision:
        prompt = build_gemini_prompt(None, master_kcs, start_page_num=start_p, end_page_num=end_p, start_printed_page=start_printed_page, is_neet=is_neet)
        pdf_part = types.Part.from_bytes(data=chunk_pages_or_bytes, mime_type="application/pdf")
        contents_payload = [pdf_part, prompt]
    else:
        prompt = build_gemini_prompt(chunk_pages_or_bytes, master_kcs, start_page_num=start_p, end_page_num=end_p, start_printed_page=start_printed_page, is_neet=is_neet)
        contents_payload = prompt

    response = None
    used_model = None
    errors = []
    
    # Prioritize previously successful model if available
    active_candidates = list(candidate_models)
    global GLOBAL_PRIMARY_MODEL_NAME
    if 'GLOBAL_PRIMARY_MODEL_NAME' in globals() and GLOBAL_PRIMARY_MODEL_NAME in active_candidates:
        active_candidates.remove(GLOBAL_PRIMARY_MODEL_NAME)
        active_candidates.insert(0, GLOBAL_PRIMARY_MODEL_NAME)
    
    for cycle in range(2):
        for candidate in active_candidates:
            model_succeeded = False
            last_candidate_error = None
            for attempt in range(2):
                try:
                    res = client.models.generate_content(
                        model=candidate,
                        contents=contents_payload,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            max_output_tokens=8192
                        )
                    )
                    if res and hasattr(res, 'text') and res.text:
                        response = res
                        used_model = candidate
                        GLOBAL_PRIMARY_MODEL_NAME = candidate
                        model_succeeded = True
                        print(f"✅ [Chunk Pages {start_p}..{end_p}] Successfully evaluated using model: {candidate} (attempt {attempt + 1})", flush=True)
                        break
                    else:
                        last_candidate_error = f"{candidate}: empty response text"
                        time.sleep(1.0)
                except Exception as e:
                    err_str = str(e)
                    last_candidate_error = f"{candidate} (attempt {attempt + 1}): {err_str[:160]}"
                    print(f"Model candidate {candidate} attempt {attempt + 1} failed: {err_str[:120]}", flush=True)
                    if any(k in err_str.lower() for k in ["429", "resource_exhausted", "quota"]):
                        # If multi-key pool is available, rotate key instantly with 0s cooldown!
                        if len(key_list) > 1:
                            key_idx = (key_idx + 1) % len(key_list)
                            print(f"--> Rate limit on key. Instantly rotating to API key {key_idx + 1} of {len(key_list)} without delay...", flush=True)
                            if progress_callback:
                                try:
                                    progress_callback(None, f"🔄 Rate limit reached on Key. Auto-switching to API Key {key_idx + 1}/{len(key_list)}...")
                                except Exception:
                                    pass
                            client = genai.Client(api_key=key_list[key_idx])
                            time.sleep(0.5)
                            continue
                        else:
                            print(f"--> Rate limit (429) on {candidate}. Instantly cascading in 0s to next model family with fresh quota pool...", flush=True)
                            break
                    elif any(k in err_str.lower() for k in ["503", "unavailable", "high demand"]):
                        if attempt < 1:
                            print(f"--> Transient server spike ({candidate}). Quick 2s retry before cascade...", flush=True)
                            time.sleep(2.0)
                            continue
                        else:
                            print(f"--> Severe 503 load on {candidate}. Cascading immediately to next cluster model...", flush=True)
                            break
                    # For non-transient errors (e.g. 404), break to next candidate
                    break
            if model_succeeded:
                break
            else:
                if last_candidate_error:
                    errors.append(last_candidate_error)
        if response and getattr(response, 'text', None):
            break
        elif cycle == 0:
            print("--> All model family quotas momentarily exhausted. Performing single 62s reset cycle...", flush=True)
            resilient_countdown_sleep(62.0, progress_callback, "All model quotas exhausted. Waiting 62s for full reset...")

    if not response or not getattr(response, 'text', None):
        err_msg = " | ".join(errors) if errors else "No response returned"
        raise ValueError(f"All authorized Gemini models failed for your API key. Details: {err_msg}")

    data = parse_gemini_json_robust(response.text)
    
    raw_questions = []
    if isinstance(data, dict):
        raw_questions = data.get("questions", [])
    elif isinstance(data, list):
        raw_questions = data

    # Python-Driven Deterministic Page Tagging Engine
    # Map relative slice indices, PDF file pages, and printed book pages bidirectionally
    pdf_to_printed_map = {}
    printed_to_pdf_map = {}
    slice_to_pdf_map = {}
    slice_to_printed_map = {}

    if not is_vision and chunk_pages_or_bytes and isinstance(chunk_pages_or_bytes, list):
        for idx, p in enumerate(chunk_pages_or_bytes):
            pdf_p = p['page_num']
            printed_p = p.get('printed_page', start_printed_page + (pdf_p - 1))
            slice_p = idx + 1
            
            slice_to_pdf_map[slice_p] = pdf_p
            slice_to_printed_map[slice_p] = printed_p
            pdf_to_printed_map[pdf_p] = printed_p
            printed_to_pdf_map[printed_p] = pdf_p
        sorted_pdf_pages = [p['page_num'] for p in chunk_pages_or_bytes]
        sorted_printed_pages = [pdf_to_printed_map[p] for p in sorted_pdf_pages]
    else:
        # Vision Mode: map relative slice indices (1..5) to PDF file pages (start_p..end_p)
        page_count = (end_p - start_p + 1)
        for idx in range(page_count):
            pdf_p = start_p + idx
            slice_p = idx + 1
            slice_to_pdf_map[slice_p] = pdf_p
        sorted_pdf_pages = [start_p + idx for idx in range(page_count)]
        sorted_printed_pages = []

        # --- Bidirectional (+1 / -1) Chunk Sequence Interpolation Engine ---
        margin_ocr = data.get("page_margin_ocr", {}) if isinstance(data, dict) else {}
        anchors = []
        for s_idx in range(1, page_count + 1):
            for key_variant in [f"slice_{s_idx}_printed_page", f"slice_{s_idx}", f"page_{s_idx}", f"slice{s_idx}"]:
                if key_variant in margin_ocr and margin_ocr[key_variant] is not None:
                    val_str = str(margin_ocr[key_variant]).strip()
                    if val_str.isdigit() and int(val_str) > 0:
                        anchors.append((s_idx, int(val_str)))
                        break

        if anchors:
            # Multi-Anchor Consensus
            base_offsets = [anchor_val - (s_i - 1) for (s_i, anchor_val) in anchors]
            from collections import Counter
            chunk_base_offset = Counter(base_offsets).most_common(1)[0][0]
            for s_idx in range(1, page_count + 1):
                interp_printed_p = chunk_base_offset + (s_idx - 1)
                slice_to_printed_map[s_idx] = interp_printed_p
                pdf_p = start_p + (s_idx - 1)
                pdf_to_printed_map[pdf_p] = interp_printed_p
                printed_to_pdf_map[interp_printed_p] = pdf_p

    resolved_questions = []
    num_questions = len(raw_questions)

    for q_idx, q in enumerate(raw_questions):
        raw_topic = q.get("kc_topic", "")
        q_no = str(q.get("question_no", "")).strip()
        
        target_printed_page = None
        target_pdf_page = None

        # 1. Check if model provided explicit printed_page and pdf_page fields
        raw_p_val = q.get("printed_page")
        raw_pdf_val = q.get("pdf_page")

        if is_vision:
            # In Vision Mode: Resolve PDF page from slice mapping
            if raw_pdf_val is not None and str(raw_pdf_val).isdigit():
                val_pdf = int(raw_pdf_val)
                if val_pdf in sorted_pdf_pages:
                    target_pdf_page = val_pdf
                elif val_pdf in slice_to_pdf_map:
                    target_pdf_page = slice_to_pdf_map[val_pdf]
                else:
                    target_pdf_page = start_p
            else:
                page_idx_pos = min(len(sorted_pdf_pages) - 1, int((q_idx / max(1, num_questions)) * len(sorted_pdf_pages))) if sorted_pdf_pages else 0
                target_pdf_page = sorted_pdf_pages[page_idx_pos] if sorted_pdf_pages else start_p

            # In Vision Mode: Resolve Printed page using interpolated slice map first
            slice_num = (target_pdf_page - start_p) + 1
            if slice_num in slice_to_printed_map:
                target_printed_page = slice_to_printed_map[slice_num]
            elif raw_p_val is not None and str(raw_p_val).isdigit() and int(raw_p_val) > 0:
                target_printed_page = int(raw_p_val)
            else:
                target_printed_page = target_pdf_page
        else:
            # In Text Mode: Resolve against calculated PDF page ground truth first
            # 1. Primary Anchor: If model provided valid PDF File Page
            if raw_pdf_val is not None and str(raw_pdf_val).isdigit():
                val_pdf = int(raw_pdf_val)
                if val_pdf in pdf_to_printed_map:
                    target_pdf_page = val_pdf
                    target_printed_page = pdf_to_printed_map[val_pdf]
                elif val_pdf in slice_to_pdf_map:
                    target_pdf_page = slice_to_pdf_map[val_pdf]
                    target_printed_page = slice_to_printed_map[val_pdf]

            # 2. Secondary Anchor: If model provided printed page (or slice index)
            if target_pdf_page is None and raw_p_val is not None and str(raw_p_val).isdigit():
                val_p = int(raw_p_val)
                if val_p in printed_to_pdf_map:
                    target_printed_page = val_p
                    target_pdf_page = printed_to_pdf_map[val_p]
                elif val_p in pdf_to_printed_map:
                    target_pdf_page = val_p
                    target_printed_page = pdf_to_printed_map[val_p]
                elif val_p in slice_to_pdf_map:
                    target_pdf_page = slice_to_pdf_map[val_p]
                    target_printed_page = slice_to_printed_map[val_p]

            # 3. Check embedded (Page X) in question_no
            page_match = re.search(r'[\(\[]\s*[Pp]age\s*(\d+)\s*[\)\]]', q_no, re.IGNORECASE)
            if target_pdf_page is None and page_match:
                extracted_p = int(page_match.group(1))
                if extracted_p in pdf_to_printed_map:
                    target_pdf_page = extracted_p
                    target_printed_page = pdf_to_printed_map[extracted_p]
                elif extracted_p in printed_to_pdf_map:
                    target_printed_page = extracted_p
                    target_pdf_page = printed_to_pdf_map[extracted_p]
                elif extracted_p in slice_to_pdf_map:
                    target_pdf_page = slice_to_pdf_map[extracted_p]
                    target_printed_page = slice_to_printed_map[extracted_p]

            # 4. Deterministic Positional Fallback if all lookups failed
            if target_pdf_page is None:
                if sorted_pdf_pages:
                    page_idx_pos = min(len(sorted_pdf_pages) - 1, int((q_idx / max(1, num_questions)) * len(sorted_pdf_pages)))
                    target_pdf_page = sorted_pdf_pages[page_idx_pos]
                    target_printed_page = sorted_printed_pages[page_idx_pos]
                else:
                    target_pdf_page = start_p
                    target_printed_page = start_printed_page

        # Clean question_no by stripping any embedded page tags
        q_no = re.sub(r'\s*[\(\[]\s*[Pp]age\s*\d+\s*[\)\]]', '', q_no, flags=re.IGNORECASE).strip()

        resolved_topic, resolved_q_no = match_kc_topic(raw_topic, q_no, master_kcs)
        q["kc_topic"] = resolved_topic
        q["question_no"] = resolved_q_no
        q["printed_page"] = target_printed_page
        q["pdf_page"] = target_pdf_page
        resolved_questions.append(q)
        
    return resolved_questions, used_model

def format_excel_outputs(conforming_questions, non_conforming_questions, master_kcs):
    """
    Creates the Master Dual-Sheet Excel Workbook ([Doc_Name]_Q_Analysis.xlsx) via openpyxl
    containing Sheet 1 (Categorized Questions) and Sheet 2 (Non-Conforming Questions),
    and returns (master_bytes, metrics).
    """
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    body_font = Font(name="Calibri", size=11, bold=False, color="000000")
    data_font = body_font
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_center_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    left_top_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    center_top_align = Alignment(horizontal="center", vertical="top", wrap_text=True)
    
    thin_border_side = Side(border_style="thin", color="D3D3D3")
    medium_border_side = Side(border_style="medium", color="1E293B")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    header_border = Border(left=medium_border_side, right=medium_border_side, top=medium_border_side, bottom=medium_border_side)

    def get_kc_block_border(r, c, start_r, end_r, min_col=1, max_col=8):
        top = medium_border_side if r == start_r else thin_border_side
        bottom = medium_border_side if r == end_r else thin_border_side
        left = medium_border_side
        right = medium_border_side
        return Border(top=top, bottom=bottom, left=left, right=right)

    # --- Global Sequence Normalization Guardrail across All Pages ---
    all_raw = list(conforming_questions or []) + list(non_conforming_questions or [])
    offsets = []
    for q in all_raw:
        p_p = q.get("printed_page")
        pdf_p = q.get("pdf_page")
        if p_p is not None and pdf_p is not None and isinstance(p_p, int) and isinstance(pdf_p, int):
            if p_p != pdf_p and p_p > 1:
                offsets.append(p_p - (pdf_p - 1))
    
    if offsets:
        from collections import Counter
        dominant_offset = Counter(offsets).most_common(1)[0][0]
        for q in all_raw:
            pdf_p = q.get("pdf_page")
            if pdf_p is not None and isinstance(pdf_p, int):
                q["printed_page"] = dominant_offset + (pdf_p - 1)

    # --- 1. Master_Categorized.xlsx ---
    wb_master = openpyxl.Workbook()
    ws_master = wb_master.active
    ws_master.title = "Categorized Questions"
    
    headers_master = [
        "KC Topic", "Question No", "Printed Page No", "PDF Page No",
        "Question Type", "Competency Level", "Verified by: ________", "Remarks"
    ]
    ws_master.append(headers_master)
    
    for col_idx in range(1, 9):
        cell = ws_master.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = header_border
        
    # --- Deterministic 1:1 KC Deduplication Guardrail on (question_no, printed_page) ---
    unique_conforming = []
    seen_conforming_keys = set()
    for q in conforming_questions:
        q_key = (str(q.get("question_no", "")).strip().lower(), q.get("pdf_page"), q.get("printed_page"))
        if q_key not in seen_conforming_keys:
            seen_conforming_keys.add(q_key)
            unique_conforming.append(q)

    # Group conforming questions strictly by KC Topic
    from collections import defaultdict
    kc_to_questions = defaultdict(list)
    for q in unique_conforming:
        topic = str(q.get("kc_topic", "")).strip()
        kc_to_questions[topic].append(q)

    mapped_kcs_found = set(kc_to_questions.keys())
    
    # Process Master KCs in exact input sequence from the skeleton sheet
    row_idx = 2
    for kc in master_kcs:
        questions = kc_to_questions.get(kc, [])
        if questions:
            start_r = row_idx
            for item in questions:
                ws_master.cell(row=row_idx, column=1, value=str(kc))
                ws_master.cell(row=row_idx, column=2, value=str(item.get("question_no", "")))
                ws_master.cell(row=row_idx, column=3, value=item.get("printed_page", ""))
                ws_master.cell(row=row_idx, column=4, value=item.get("pdf_page", ""))
                ws_master.cell(row=row_idx, column=5, value=str(item.get("question_type", "")))
                ws_master.cell(row=row_idx, column=6, value=str(item.get("competency_level", "")))
                ws_master.cell(row=row_idx, column=7, value="Not Checked")
                ws_master.cell(row=row_idx, column=8, value="")
                
                ws_master.cell(row=row_idx, column=1).alignment = center_align
                ws_master.cell(row=row_idx, column=2).alignment = left_center_align
                ws_master.cell(row=row_idx, column=3).alignment = center_align
                ws_master.cell(row=row_idx, column=4).alignment = center_align
                ws_master.cell(row=row_idx, column=5).alignment = center_align
                ws_master.cell(row=row_idx, column=6).alignment = center_align
                ws_master.cell(row=row_idx, column=7).alignment = center_align
                ws_master.cell(row=row_idx, column=8).alignment = left_center_align
                row_idx += 1
            
            end_r = row_idx - 1
            # Apply solid perimeter border to the entire KC block
            for r in range(start_r, end_r + 1):
                for c in range(1, 9):
                    cell = ws_master.cell(row=r, column=c)
                    cell.font = data_font
                    cell.border = get_kc_block_border(r, c, start_r, end_r, 1, 8)

            # Merge KC Topic cell across multi-question rows
            if end_r > start_r:
                ws_master.merge_cells(start_row=start_r, end_row=end_r, start_column=1, end_column=1)
                top_cell = ws_master.cell(row=start_r, column=1)
                top_cell.alignment = center_align
                ws_master.cell(row=start_r, column=1).border = Border(left=medium_border_side, right=medium_border_side, top=medium_border_side, bottom=medium_border_side)
                ws_master.cell(row=end_r, column=1).border = Border(left=medium_border_side, right=medium_border_side, top=thin_border_side, bottom=medium_border_side)
        else:
            # Unmapped KC sitting in its exact sequential place in the syllabus
            ws_master.cell(row=row_idx, column=1, value=str(kc))
            for c in range(2, 9):
                if c == 7:
                    ws_master.cell(row=row_idx, column=c, value="Not Checked")
                else:
                    empty_cell = ws_master.cell(row=row_idx, column=c, value="")
                    empty_cell.data_type = 's'
                
            ws_master.cell(row=row_idx, column=1).alignment = center_align
            ws_master.cell(row=row_idx, column=7).alignment = center_align
            ws_master.cell(row=row_idx, column=8).alignment = left_center_align
            for c in range(1, 9):
                cell = ws_master.cell(row=row_idx, column=c)
                cell.font = data_font
                cell.border = get_kc_block_border(row_idx, c, row_idx, row_idx, 1, 8)
            row_idx += 1

    # Handle any rare remaining topics not in master_kcs (e.g. Miscellaneous)
    extra_topics = [t for t in kc_to_questions.keys() if t not in master_kcs]
    for extra_t in extra_topics:
        display_topic = "Miscellaneous" if re.search(r'-(?:Q\d+|Section-[A-Z]-Q\d+|Part-[A-Z]-Q\d+)$', str(extra_t), re.IGNORECASE) or str(extra_t).lower() in ["none", "null", "unmapped", "0"] else str(extra_t)
        questions = kc_to_questions[extra_t]
        start_r = row_idx
        for item in questions:
            ws_master.cell(row=row_idx, column=1, value=display_topic)
            ws_master.cell(row=row_idx, column=2, value=str(item.get("question_no", "")))
            ws_master.cell(row=row_idx, column=3, value=item.get("printed_page", ""))
            ws_master.cell(row=row_idx, column=4, value=item.get("pdf_page", ""))
            ws_master.cell(row=row_idx, column=5, value=str(item.get("question_type", "")))
            ws_master.cell(row=row_idx, column=6, value=str(item.get("competency_level", "")))
            ws_master.cell(row=row_idx, column=7, value="Not Checked")
            ws_master.cell(row=row_idx, column=8, value="")
            
            ws_master.cell(row=row_idx, column=1).alignment = center_align
            ws_master.cell(row=row_idx, column=2).alignment = left_center_align
            ws_master.cell(row=row_idx, column=3).alignment = center_align
            ws_master.cell(row=row_idx, column=4).alignment = center_align
            ws_master.cell(row=row_idx, column=5).alignment = center_align
            ws_master.cell(row=row_idx, column=6).alignment = center_align
            ws_master.cell(row=row_idx, column=7).alignment = center_align
            ws_master.cell(row=row_idx, column=8).alignment = left_center_align
            row_idx += 1
            
        end_r = row_idx - 1
        for r in range(start_r, end_r + 1):
            for c in range(1, 9):
                cell = ws_master.cell(row=r, column=c)
                cell.font = data_font
                cell.border = get_kc_block_border(r, c, start_r, end_r, 1, 8)

        if end_r > start_r:
            ws_master.merge_cells(start_row=start_r, end_row=end_r, start_column=1, end_column=1)
            top_cell = ws_master.cell(row=start_r, column=1)
            top_cell.alignment = center_align
            ws_master.cell(row=start_r, column=1).border = Border(left=medium_border_side, right=medium_border_side, top=medium_border_side, bottom=medium_border_side)
            ws_master.cell(row=end_r, column=1).border = Border(left=medium_border_side, right=medium_border_side, top=thin_border_side, bottom=medium_border_side)

    # Total unmapped KCs for metrics / completion report
    unmapped_kcs = [kc for kc in master_kcs if kc not in mapped_kcs_found]

    column_widths_master = {1: 35, 2: 45, 3: 18, 4: 16, 5: 18, 6: 20, 7: 24, 8: 32}
    for col_i, width in column_widths_master.items():
        ws_master.column_dimensions[get_column_letter(col_i)].width = width

    # Freeze Panes on Row 1 (headers) and Column A (KC Topic)
    ws_master.freeze_panes = 'B2'

    # In-cell Data Validation dropdown for Verified by column
    if row_idx > 2:
        dv_master = DataValidation(type="list", formula1='"Not Checked,Checked"', allow_blank=True)
        ws_master.add_data_validation(dv_master)
        dv_master.add(f"G2:G{row_idx - 1}")

    # --- 2. Non-Conforming Questions Sheet (Tab 2 in the same Workbook) ---
    ws_outliers = wb_master.create_sheet(title="Non-Conforming Questions")
    
    headers_outliers = [
        "Question No", "Printed Page No", "PDF Page No",
        "Detected Format", "Exclusion Reason", "Verified by: ________", "Remarks"
    ]
    ws_outliers.append(headers_outliers)
    
    for col_idx in range(1, 8):
        cell = ws_outliers.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = header_border

    # Deduplicate non-conforming items and exclude any already in conforming
    unique_non_conforming = []
    seen_outlier_keys = set()
    for q in non_conforming_questions:
        q_key = (str(q.get("question_no", "")).strip().lower(), q.get("pdf_page"), q.get("printed_page"))
        if q_key not in seen_outlier_keys and q_key not in seen_conforming_keys:
            seen_outlier_keys.add(q_key)
            unique_non_conforming.append(q)

    r_outlier = 2
    total_outliers = len(unique_non_conforming)
    for item in unique_non_conforming:
        ws_outliers.cell(row=r_outlier, column=1, value=item.get("question_no", "Unlabeled Item"))
        ws_outliers.cell(row=r_outlier, column=2, value=item.get("printed_page", ""))
        ws_outliers.cell(row=r_outlier, column=3, value=item.get("pdf_page", ""))
        
        # Cleanly resolve contradiction: Output actual outlier type instead of 'Integer type'
        raw_type = str(item.get("question_type", "Non-Conforming")).strip()
        reason = str(item.get("non_conforming_reason", "")).lower()
        
        if raw_type in ["Integer type", "MCQ", "Multi correct"]:
            if "decimal" in reason or "floating" in reason:
                clean_format = "Decimal Numerical"
            elif "proof" in reason or "derivation" in reason or "subjective" in reason:
                clean_format = "Subjective Proof"
            elif "draw" in reason or "diagram" in reason:
                clean_format = "Drawing Task"
            elif "range" in reason or "continuous" in reason:
                clean_format = "Continuous Range"
            elif "negative" in reason:
                clean_format = "Negative Numerical"
            else:
                clean_format = "Non-Standard Outlier"
        else:
            clean_format = raw_type
            
        ws_outliers.cell(row=r_outlier, column=4, value=clean_format)
        ws_outliers.cell(row=r_outlier, column=5, value=item.get("non_conforming_reason", "Does not conform to MCQ/Multi-correct/Integer standard"))
        ws_outliers.cell(row=r_outlier, column=6, value="Not Checked")
        ws_outliers.cell(row=r_outlier, column=7, value="")
        
        ws_outliers.cell(row=r_outlier, column=1).alignment = left_center_align
        ws_outliers.cell(row=r_outlier, column=2).alignment = center_align
        ws_outliers.cell(row=r_outlier, column=3).alignment = center_align
        ws_outliers.cell(row=r_outlier, column=4).alignment = center_align
        ws_outliers.cell(row=r_outlier, column=5).alignment = left_center_align
        ws_outliers.cell(row=r_outlier, column=6).alignment = center_align
        ws_outliers.cell(row=r_outlier, column=7).alignment = left_center_align
        
        is_last_outlier = (r_outlier == total_outliers + 1)
        outlier_border = Border(
            left=medium_border_side,
            right=medium_border_side,
            top=thin_border_side,
            bottom=medium_border_side if is_last_outlier else thin_border_side
        )
        for c in range(1, 8):
            cell = ws_outliers.cell(row=r_outlier, column=c)
            cell.font = data_font
            cell.border = outlier_border
        r_outlier += 1

    column_widths_outliers = {1: 24, 2: 16, 3: 14, 4: 20, 5: 45, 6: 24, 7: 32}
    for col_i, width in column_widths_outliers.items():
        ws_outliers.column_dimensions[get_column_letter(col_i)].width = width

    # Freeze Panes on Row 1 (headers) and Column A (Question No)
    ws_outliers.freeze_panes = 'B2'

    # In-cell Data Validation dropdown for Verified by column
    if r_outlier > 2:
        dv_outliers = DataValidation(type="list", formula1='"Not Checked,Checked"', allow_blank=True)
        ws_outliers.add_data_validation(dv_outliers)
        dv_outliers.add(f"F2:F{r_outlier - 1}")

    # Convert Single Unified Workbook to in-memory bytes
    master_io = io.BytesIO()
    wb_master.save(master_io)
    master_io.seek(0)
    master_bytes = master_io.getvalue()

    total_conforming = len(unique_conforming)
    total_non_conforming = len(unique_non_conforming)
    total_scanned = total_conforming + total_non_conforming
    mapped_kcs_count = len([kc for kc in master_kcs if kc_to_questions.get(kc)])
    unmapped_kcs = [kc for kc in master_kcs if not kc_to_questions.get(kc)]
    
    metrics = {
        "total_scanned": total_scanned,
        "conforming": total_conforming,
        "non_conforming": total_non_conforming,
        "mapped_kcs": mapped_kcs_count,
        "unmapped_kcs": len(unmapped_kcs)
    }

    return master_bytes, metrics

GLOBAL_ACTIVE_RUN_ID = None

def run_categorization_pipeline(
    api_key,
    pdf_file_source,
    skeleton_file_source,
    progress_callback=None,
    run_id=None,
    start_chunk_idx=0,
    existing_conforming=None,
    existing_non_conforming=None,
    chunk_completed_callback=None,
    is_neet=None
):
    """
    Main orchestration function:
    - Unified Chunk Size of 3 pages per chunk across both Text and Vision modes.
    - Single-Run Mutex: Automatically terminates stale background threads if superseded by a newer run_id.
    - Resumable: Supports starting at start_chunk_idx and accumulating questions from earlier checkpoints.
    - Returns (master_bytes, metrics).
    """
    global GLOBAL_ACTIVE_RUN_ID
    if not run_id:
        import uuid
        run_id = str(uuid.uuid4())
    GLOBAL_ACTIVE_RUN_ID = run_id
    if isinstance(api_key, (list, tuple)):
        key_list = [str(k).strip() for k in api_key if str(k).strip()]
    elif isinstance(api_key, str):
        key_list = [k.strip() for k in api_key.replace(";", ",").split(",") if k.strip()]
    else:
        key_list = [str(api_key).strip()]
        
    if not key_list:
        raise ValueError("No valid Gemini API Key provided.")
        
    print(f"🎬 Initiating Categorization Run ID: {run_id} (Starting Chunk: {start_chunk_idx + 1}, Key Pool Size: {len(key_list)})", flush=True)

    candidate_models = get_user_authorized_models(key_list[0])
    if not candidate_models:
        raise ValueError("No accessible Gemini models were found for this API key. Please check your key at aistudio.google.com/app/apikey.")
        
    print(f"Verified Model Priority: {candidate_models}")

    skeleton_filename = getattr(skeleton_file_source, 'name', str(skeleton_file_source) if isinstance(skeleton_file_source, str) else '')
    if is_neet is None:
        is_neet = bool(re.search(r'\bNEET\b', skeleton_filename, re.IGNORECASE))

    if is_neet:
        print(f"🎯 NEET KC Document Detected ('{skeleton_filename}') → Calibrating competency purely on cognitive depth (zero historical paper benchmarking).", flush=True)
    else:
        print(f"📘 Standard KC Document ('{skeleton_filename}') → Calibrating competency against National Exam Benchmarks (CBSE / JEE Main / JEE Advanced 2020-2024+).", flush=True)

    master_kcs = extract_kcs_from_file(skeleton_file_source, filename=skeleton_filename)
    if not master_kcs:
        raise ValueError("Could not find Knowledge Components (KCs) in the uploaded KC list file. Please ensure it contains a 'List of KC' / 'Topic' column or valid syllabus items.")

    pages_data, total_pages, is_scanned_vision_mode, detected_start_page = extract_pages_from_document(pdf_file_source, filename=getattr(pdf_file_source, 'name', ''))
    if total_pages == 0:
        raise ValueError("The uploaded Question Paper document appears to be empty or unreadable.")

    all_conforming = list(existing_conforming or [])
    all_non_conforming = list(existing_non_conforming or [])

    if is_scanned_vision_mode:
        # Scanned PDF Vision Mode -> Unified constant of 3 pages per chunk
        vision_chunk_size = 3
        total_chunks = (total_pages + vision_chunk_size - 1) // vision_chunk_size
        
        for idx in range(start_chunk_idx, total_chunks):
            if GLOBAL_ACTIVE_RUN_ID != run_id:
                print(f"🛑 Run {run_id} superseded by active run {GLOBAL_ACTIVE_RUN_ID}. Exiting stale thread immediately.", flush=True)
                return None, None
            start_p_idx = idx * vision_chunk_size
            end_p_idx = min((idx + 1) * vision_chunk_size, total_pages)
            
            chunk_bytes = get_pdf_chunk_bytes(pdf_file_source, start_p_idx, end_p_idx)
            
            print(f"[Vision Chunk {idx + 1}/{total_chunks}] Processing scanned pages {start_p_idx + 1} to {end_p_idx} of {total_pages}...", flush=True)
            
            if progress_callback:
                progress_callback(
                    (idx + 1) / total_chunks,
                    f"Processing Scanned PDF pages {start_p_idx + 1} to {end_p_idx} of {total_pages} (Vision Mode Chunk {idx + 1}/{total_chunks})...."
                )
                
            questions, actual_model_used = process_chunk_with_gemini(
                key_list, candidate_models, chunk_bytes, master_kcs, is_vision=True, start_p=start_p_idx + 1, end_p=end_p_idx, start_printed_page=detected_start_page, progress_callback=progress_callback, is_neet=is_neet
            )
            
            for q in questions:
                if q.get("is_conforming", True):
                    all_conforming.append(q)
                else:
                    all_non_conforming.append(q)

            if chunk_completed_callback:
                chunk_completed_callback(idx + 1, total_chunks, all_conforming, all_non_conforming, end_p_idx, master_kcs)

            if idx < total_chunks - 1:
                resilient_countdown_sleep(4.0, progress_callback, "Pacing buffer before next chunk")
    else:
        # Standard Text-Based Document Mode -> Filter out Answer Keys, Hints & Solutions
        filtered_pages_data, total_excluded_pages = filter_document_pages(pages_data)
        if filtered_pages_data:
            print(f"🧹 Document Pre-Filter: Kept {len(filtered_pages_data)} question pages (excluded {total_excluded_pages} answer/solution pages).", flush=True)
            active_pages = filtered_pages_data
        else:
            print(f"⚠️ Document Pre-Filter: No pages kept after filtering. Falling back to original pages.", flush=True)
            active_pages = pages_data

        text_chunk_size = 3
        chunks = [active_pages[i:i + text_chunk_size] for i in range(0, len(active_pages), text_chunk_size)]
        total_chunks = len(chunks)
        total_active_pages = len(active_pages)
        
        # --- Cross-Chunk Question Boundary Stitching ---
        # When a question begins at the bottom of the last page of Chunk N and its options
        # or completion appear at the top of Chunk N+1, look ahead and stitch them together.
        for c_idx in range(len(chunks) - 1):
            curr_chunk = chunks[c_idx]
            next_chunk = chunks[c_idx + 1]
            last_p = curr_chunk[-1]
            next_first_p = next_chunk[0]
            
            is_unclosed, q_num, trailing = check_unclosed_question_at_page_bottom(last_p.get('text', ''))
            if is_unclosed:
                cont_text, rem_text = extract_page_leading_continuation(next_first_p.get('text', ''))
                if cont_text:
                    p_last_num = last_p.get('page_num', '?')
                    p_next_num = next_first_p.get('page_num', '?')
                    print(f"🔗 [Chunk Seam Stitcher] Stitching split Q{q_num} across Chunk {c_idx + 1} (PDF Page {p_last_num}) and Chunk {c_idx + 2} (PDF Page {p_next_num})", flush=True)
                    
                    stitched_block = (
                        f"\n\n================================================================================\n"
                        f"<<< PAGE TRANSITION CONTINUATION (PDF Page {p_last_num} -> PDF Page {p_next_num}) >>>\n"
                        f"(The following continuation lines complete the unclosed Question {q_num} that began on PDF Page {p_last_num}.\n"
                        f" Tag Question {q_num} with PDF Page {p_last_num} and its corresponding printed page number.)\n"
                        f"--------------------------------------------------------------------------------\n"
                        f"{cont_text}\n"
                        f"================================================================================\n"
                    )
                    last_p['text'] = last_p['text'] + stitched_block
                    next_first_p['text'] = rem_text

        for idx, chunk in enumerate(chunks):
            if idx < start_chunk_idx:
                continue
            if GLOBAL_ACTIVE_RUN_ID != run_id:
                print(f"🛑 Run {run_id} superseded by active run {GLOBAL_ACTIVE_RUN_ID}. Exiting stale thread immediately.", flush=True)
                return None, None
            start_p_val = chunk[0]['page_num']
            end_p_val = chunk[-1]['page_num']
            
            # --- Dynamic Per-Chunk Sequence Inspection (Jump Detection) ---
            chunk_candidates = []
            for p in chunk:
                lines = [l.strip() for l in p.get('text', '').split('\n') if l.strip()]
                margin_lines = lines[:5] + lines[-5:] if lines else []
                combined = " ".join(margin_lines)
                nums = [int(n) for n in re.findall(r'\b(\d{1,4})\b', combined) if 0 <= int(n) <= 5000]
                chunk_candidates.append({'page_num': p['page_num'], 'numbers': set(nums)})
                
            chunk_potential_offsets = set()
            for c in chunk_candidates:
                p_num = c['page_num']
                for n in c['numbers']:
                    off = n - (p_num - 1)
                    if off >= 0:
                        chunk_potential_offsets.add(off)
                        
            chunk_best_offset = None
            chunk_best_chain = 0
            for O in chunk_potential_offsets:
                curr_run = 0
                max_run = 0
                for c in chunk_candidates:
                    expected_num = O + (c['page_num'] - 1)
                    if expected_num in c['numbers']:
                        curr_run += 1
                        if curr_run > max_run:
                            max_run = curr_run
                    else:
                        curr_run = 0
                if max_run > chunk_best_chain:
                    chunk_best_chain = max_run
                    chunk_best_offset = O
                    
            if chunk_best_chain >= 3 and chunk_best_offset is not None and chunk_best_offset > 1 and chunk_best_offset != detected_start_page:
                effective_chunk_start_offset = chunk_best_offset
                print(f"[Text Chunk {idx + 1}/{total_chunks}] 🔀 Local Jump Detected! Adapted Offset to {chunk_best_offset} (Chain: {chunk_best_chain} pages)", flush=True)
            else:
                effective_chunk_start_offset = detected_start_page

            for p in chunk:
                p['printed_page'] = max(0, effective_chunk_start_offset + (p['page_num'] - 1))

            first_printed_p = chunk[0]['printed_page']
            last_printed_p = chunk[-1]['printed_page']
            print(f"[Text Chunk {idx + 1}/{total_chunks}] Processing PDF pages {start_p_val}..{end_p_val} (Printed: {first_printed_p}..{last_printed_p}, Offset={effective_chunk_start_offset})...", flush=True)
            
            if progress_callback:
                progress_callback(
                    (idx + 1) / total_chunks,
                    f"Processing pages {start_p_val} to {end_p_val} of {total_active_pages} (Text Chunk {idx + 1}/{total_chunks} • Offset {effective_chunk_start_offset} • PDF {start_p_val} → Printed {first_printed_p})...."
                )
                
            questions, actual_model_used = process_chunk_with_gemini(
                key_list, candidate_models, chunk, master_kcs, is_vision=False, start_p=start_p_val, end_p=end_p_val, start_printed_page=effective_chunk_start_offset, progress_callback=progress_callback, is_neet=is_neet
            )
            
            for q in questions:
                if q.get("is_conforming", True):
                    all_conforming.append(q)
                else:
                    all_non_conforming.append(q)

            if chunk_completed_callback:
                chunk_completed_callback(idx + 1, total_chunks, all_conforming, all_non_conforming, end_p_val, master_kcs)

            if idx < total_chunks - 1:
                resilient_countdown_sleep(3.0, progress_callback, "Pacing buffer before next chunk")

    master_bytes, metrics = format_excel_outputs(all_conforming, all_non_conforming, master_kcs)
    
    return master_bytes, metrics

