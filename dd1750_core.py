"""
DD1750 Core - Packing List Generator from BOM PDFs

This module extracts items from GCSS-Army Component Listing / Hand Receipt PDFs
and generates DD Form 1750 Packing Lists.

Supported BOM formats:
1. Standard GCSS-Army Component Listing with LV column (e.g., B49.pdf)
2. Equipment Property Record format (epp.pdf style)

Note: Handwritten BOMs are NOT supported. Users should obtain clean digital
BOMs from GCSS-Army through their supply teams.
"""

import io
import math
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Any
from enum import Enum

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# OCR support is optional. If pdf2image / pytesseract aren't installed, the
# OCR fallback path is silently disabled and we fall back to form-field
# extraction only. The user-facing warning message tells the user to verify
# missing items manually in that case.
try:
    from pdf2image import convert_from_path as _convert_from_path
    import pytesseract as _pytesseract
    from pytesseract import Output as _OCR_Output
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False
    _convert_from_path = None
    _pytesseract = None
    _OCR_Output = None


# DD1750 Form Layout Constants (Letter size: 612 x 792 points)
# These measurements are from the official DD FORM 1750, SEP 70 (EG)
ROWS_PER_PAGE = 18
PAGE_W, PAGE_H = 612.0, 792.0

# Column boundaries (x coordinates in points from left edge)
# Derived from official template analysis
X_BOX_L, X_BOX_R = 45.0, 88.2           # Box Number column
X_CONTENT_L, X_CONTENT_R = 88.2, 365.4  # Contents (Stock # and Nomenclature)
X_UOI_L, X_UOI_R = 365.4, 408.6         # Unit of Issue
X_INIT_L, X_INIT_R = 408.6, 453.6       # Initial Operation
X_SPARES_L, X_SPARES_R = 453.6, 514.8   # Running Spares
X_TOTAL_L, X_TOTAL_R = 514.8, 567.0     # Total

# Row layout (PDF coordinates: 0 at bottom)
Y_TABLE_TOP = 616.0      # Top of table content area
Y_TABLE_BOTTOM = 89.1    # Bottom of table content area
ROW_H = (Y_TABLE_TOP - Y_TABLE_BOTTOM) / ROWS_PER_PAGE  # ~29.27 points
PAD_X = 3.0  # Horizontal padding from column edge


class BomFormat(Enum):
    """Enumeration of supported BOM formats."""
    GCSS_ARMY_STANDARD = "gcss_army_standard"  # Has LV column, standard Component Listing
    EPP_FORMAT = "epp_format"                   # Equipment Property Record format
    DA_2062 = "da_2062"                         # DA Form 2062 Hand Receipt
    UNKNOWN = "unknown"


@dataclass
class BomItem:
    """Represents a single item from a Bill of Materials."""
    line_no: int
    description: str
    nsn: str = ""           # National Stock Number (9-digit NIIN)
    qty: int = 1            # Authorized quantity
    unit_of_issue: str = "EA"
    material_number: str = ""  # Full material/part number
    oh_qty: int = -1        # On-hand quantity (-1 = not specified, 0 = zero, >0 = has qty)
    
    # For user review/editing
    is_editable: bool = True
    original_description: str = ""
    
    def __post_init__(self):
        if not self.original_description:
            self.original_description = self.description


@dataclass
class HeaderInfo:
    """Header information for DD1750 form."""
    packed_by: str = ""
    num_boxes: str = "1"
    requisition_no: str = ""
    order_no: str = ""
    end_item: str = ""
    date: str = ""
    # Page numbers are auto-calculated


@dataclass
class BomMetadata:
    """Metadata extracted from BOM header."""
    end_item_niin: str = ""
    end_item_description: str = ""
    lin: str = ""
    pub_num: str = ""
    pub_date: str = ""
    serial_equip_no: str = ""
    uic: str = ""
    fe: str = ""
    bom_format: BomFormat = BomFormat.UNKNOWN


@dataclass
class ExtractionResult:
    """Result of BOM extraction including items and metadata."""
    items: List[BomItem] = field(default_factory=list)
    metadata: BomMetadata = field(default_factory=BomMetadata)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    pages_processed: int = 0
    format_detected: BomFormat = BomFormat.UNKNOWN


def detect_bom_format(tables: List[List[List[str]]], page_text: str) -> BomFormat:
    """
    Detect the format of the BOM based on table structure and page content.
    
    Args:
        tables: Extracted tables from the page
        page_text: Full text content of the page
        
    Returns:
        Detected BomFormat enum value
    """
    page_upper = page_text.upper()
    
    # Check for DA Form 2062 (Hand Receipt/Shortage Listing)
    # Can be identified by "DA FORM 2062", "HAND RECEIPT/SHORTAGE LISTING", or specific column structure
    if "DA FORM 2062" in page_upper or "HAND RECEIPT/SHORTAGE LISTING" in page_upper:
        # Verify it has the 2062 table structure
        for table in tables:
            if table and len(table) > 0:
                for row in table[:5]:
                    row_text = ' '.join(str(cell or '') for cell in row).upper()
                    # Check for either "STOCK NUMBER" or "MATERIAL NUMBER" with "ITEM DESCRIPTION"
                    if ('STOCK NUMBER' in row_text or 'MATERIAL NUMBER' in row_text) and 'ITEM DESCRIPTION' in row_text:
                        return BomFormat.DA_2062
    
    # Check for GCSS-Army standard format markers
    if "COMPONENT LISTING" in page_upper or "HAND RECEIPT" in page_upper:
        # Look for LV column in headers
        for table in tables:
            if table and len(table) > 0:
                header = table[0]
                header_text = ' '.join(str(cell or '') for cell in header).upper()
                if 'LV' in header_text or 'LEVEL' in header_text:
                    return BomFormat.GCSS_ARMY_STANDARD
        
        # Even without LV column, if it has the standard structure
        if "AUTH" in page_upper and "QTY" in page_upper:
            return BomFormat.GCSS_ARMY_STANDARD
    
    # Check for EPP format markers
    if "PWR PLANT" in page_upper or "OPERATIONAL SUPPORT" in page_upper:
        return BomFormat.EPP_FORMAT
    
    # Default to standard format if we see Material and Description columns
    for table in tables:
        if table and len(table) > 0:
            header = table[0]
            header_text = ' '.join(str(cell or '') for cell in header).upper()
            if 'MATERIAL' in header_text and 'DESCRIPTION' in header_text:
                return BomFormat.GCSS_ARMY_STANDARD
    
    return BomFormat.UNKNOWN


def find_column_indices(header: List[str]) -> Dict[str, Optional[int]]:
    """
    Find column indices from header row.
    
    Args:
        header: List of header cell values
        
    Returns:
        Dictionary mapping column names to their indices
    """
    indices = {
        'lv': None,
        'description': None,
        'material': None,
        'auth_qty': None,    # Authorized Quantity - THIS IS WHAT WE USE
        'oh_qty': None,      # On-Hand Quantity (last column, often handwritten)
        'ui': None,
        'image': None,
        'ciic': None,        # CIIC column - if has a letter, row is valid item
    }
    
    for i, cell in enumerate(header):
        if not cell:
            continue
        text = str(cell).upper().strip()
        # Also check for multi-line headers
        text_joined = text.replace('\n', ' ')
        
        # Level column
        if text in ('LV', 'LEVEL') or 'LV' in text.split():
            indices['lv'] = i
        # Description column
        elif 'DESC' in text:
            indices['description'] = i
        # Material column
        elif 'MATERIAL' in text or text == 'MAT':
            indices['material'] = i
        # Authorized quantity - check for "AUTH" and "QTY" together
        elif ('AUTH' in text_joined and 'QTY' in text_joined) or text_joined == 'AUTH QTY':
            indices['auth_qty'] = i
        # On-Hand quantity (rightmost qty column)
        elif ('OH' in text_joined and 'QTY' in text_joined) or text_joined == 'OH QTY':
            indices['oh_qty'] = i
        # CIIC column
        elif text == 'CIIC' or 'CIIC' in text_joined:
            indices['ciic'] = i
        # Unit of Issue
        elif text == 'UI' or text == 'UNIT':
            indices['ui'] = i
        # Image column (usually first)
        elif 'IMAGE' in text or text == 'IMG':
            indices['image'] = i
    
    return indices


def extract_nsn_from_material(material_text: str) -> str:
    """
    Extract 9-character NIIN from material/part number field.
    
    Handles various formats found in GCSS-Army BOMs:
    - Direct 9-digit NIIN: 002643796
    - With line breaks: 002643796\nC_19207 ~ 11655778-5
    - Full NSN format: 6545-00-922-1200
    - Material number with NIIN: C_89875 ~ 6545-00-922-1200
    - Alphanumeric "C-prefix" NIIN: 01C079749 (digits + 1 letter + digits = 9 chars)
    - NIIN on second line if first is a part number
    
    Args:
        material_text: Text from material column
        
    Returns:
        9-character NIIN string or empty string if not found
    """
    if not material_text:
        return ""
    
    text = str(material_text).strip()
    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
    
    # Pattern A: 9-digit NIIN at start of any line (most common GCSS format)
    for line in lines:
        match = re.match(r'^(\d{9})(?:\b|$)', line)
        if match:
            return match.group(1)
    
    # Pattern B: Alphanumeric 9-char NIIN at start of any line
    # Format: digits + letter(s) + digits = exactly 9 chars (e.g., 01C079749)
    # Must have at least 2 digits at the start to avoid matching part numbers
    for line in lines:
        match = re.match(r'^(\d{2}[A-Z]\d{6}|\d{2}[A-Z]{2}\d{5}|\d{3}[A-Z]\d{5}|\d{2}[A-Z]\d{2}[A-Z]\d{3})(?:\b|$)', line)
        if match:
            return match.group(1)
    
    # Pattern C: Full NSN format anywhere (XXXX-XX-XXX-XXXX) - extract NIIN portion
    nsn_match = re.search(r'\b(\d{4})-(\d{2})-(\d{3})-(\d{4})\b', text)
    if nsn_match:
        # NIIN is the last 9 digits: FSC-NIIN format
        return nsn_match.group(2) + nsn_match.group(3) + nsn_match.group(4)
    
    # Pattern D: Any 9-digit number in the text (last-ditch fallback)
    # Avoid matching obvious part numbers (preceded by letters/dashes)
    for line in lines:
        # Skip lines that look like part numbers (contain : ~ - prominently)
        match = re.search(r'(?:^|[\s])(\d{9})(?:\b|$)', line)
        if match:
            return match.group(1)
    
    return ""


def clean_description(desc_text: str) -> str:
    """
    Clean and normalize description text.
    
    Args:
        desc_text: Raw description text
        
    Returns:
        Cleaned description string
    """
    if not desc_text:
        return ""
    
    lines = str(desc_text).strip().split('\n')
    
    # Often the second line is the actual description
    description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
    
    # Remove parenthetical content (often contains codes)
    if '(' in description:
        description = description.split('(')[0].strip()
    
    # Remove trailing codes that sometimes appear
    codes_pattern = r'\s+(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G|9B|9T|2B|2E|2W|2T|85|7K|7B)$'
    description = re.sub(codes_pattern, '', description, flags=re.IGNORECASE)
    
    # Normalize whitespace
    description = re.sub(r'\s+', ' ', description).strip()
    
    return description


def extract_quantity(qty_cell: Any) -> int:
    """
    Extract numeric quantity from cell value.
    
    Args:
        qty_cell: Cell value (may be string, int, or None)
        
    Returns:
        Integer quantity (defaults to 1 if extraction fails)
    """
    if not qty_cell:
        return 1
    
    qty_str = str(qty_cell).strip()
    
    # Find first number in the string
    match = re.search(r'(\d+)', qty_str)
    if match:
        return int(match.group(1))
    
    return 1


def extract_items_gcss_standard(tables: List[List[List[str]]]) -> List[BomItem]:
    """
    Extract items from GCSS-Army standard format BOM.
    
    Standard format has:
    - Image, Material, LV, Description, WTY, ARC, CIIC, UI, SCMC, Auth Qty, OH Qty
    - Items with LV="B" are components to extract
    - LV="A" items are category headers
    - Uses Auth Qty column for quantities
    - Always uses EA for unit of issue
    
    Args:
        tables: List of tables extracted from PDF
        
    Returns:
        List of BomItem objects
    """
    items = []
    
    for table in tables:
        if not table or len(table) < 2:
            continue
        
        header = table[0]
        indices = find_column_indices(header)
        
        # Need at least description column
        if indices['description'] is None:
            # Try to find description column by looking at header content
            for i, cell in enumerate(header):
                if cell:
                    text = str(cell).upper()
                    if 'DESC' in text:
                        indices['description'] = i
                        break
        
        if indices['description'] is None:
            continue
        
        for row_num, row in enumerate(table[1:]):
            # Skip empty rows
            if not any(cell for cell in row if cell):
                continue
            
            # PRIMARY CHECK: Use LV column to identify valid items
            # LV='B' = component (the items we want)
            # LV='A' = category header (skip)
            # LV empty + has Material/Description = also valid (some EPP-style rows)
            #
            # CIIC column is informational only - it can be a letter (U, M, J, Y)
            # OR a digit (7, 9) for sensitive items. Both are valid.
            lv_value = ""
            if indices['lv'] is not None and indices['lv'] < len(row):
                lv_cell = row[indices['lv']]
                lv_value = str(lv_cell).strip().upper() if lv_cell else ""
            
            ciic_value = ""
            if indices['ciic'] is not None and indices['ciic'] < len(row):
                ciic_cell = row[indices['ciic']]
                ciic_value = str(ciic_cell).strip().upper() if ciic_cell else ""
            
            # Skip "A" level items (category headers like COEI/BII)
            if lv_value == 'A':
                continue
            
            # If neither LV nor CIIC has content, this is probably a separator/blank row
            # Real item rows have at least one of: LV='B', non-empty CIIC, or both
            if not lv_value and not ciic_value:
                # Allow rows with no LV and no CIIC ONLY if they have material AND description
                # (some EPP-format rows have empty LV and CIIC)
                has_material_data = (indices['material'] is not None 
                                     and indices['material'] < len(row) 
                                     and row[indices['material']]
                                     and str(row[indices['material']]).strip())
                if not has_material_data:
                    continue
            
            # If LV is set, it must be 'B' (or some other component code, NOT 'A')
            if lv_value and lv_value not in ('B', 'C', 'D', 'E'):
                continue
            
            # Extract description - ALWAYS use the FIRST LINE
            # The first line contains the clean nomenclature (e.g., "CHAIN ASSEMBLY,SINGLE LEG")
            # Lower lines may have additional details but can be truncated/fragmented
            desc_cell = row[indices['description']] if indices['description'] < len(row) else None
            description = ""
            if desc_cell:
                lines = str(desc_cell).strip().split('\n')
                # Use the first non-empty line
                for line in lines:
                    line = line.strip()
                    if line and len(line) >= 3:
                        description = line
                        break
                
                # Clean up
                description = re.sub(r'\s+', ' ', description).strip()  # Normalize whitespace
                description = re.sub(r'[/\\]+\s*$', '', description)    # Remove trailing slashes
            
            if not description or len(description) < 3:
                continue
            
            # Skip category descriptions and header rows
            skip_patterns = [
                'COMPONENT OF END ITEM', 'BASIC ISSUE ITEMS', 
                'COEI-', 'BII-', 'OPERATIONAL SUPPORT',
            ]
            if any(pat in description.upper() for pat in skip_patterns):
                continue
            
            # Skip if description looks like an end item ID code, NOT a regular nomenclature.
            # An ID code looks like "WH12B0" or "T59652-014120143" - it has digits OR a dash.
            # Pure alphabetical descriptions like "ANTENNA", "HANDSET" are valid item names!
            desc_upper = description.upper()
            looks_like_id = (
                len(description) < 20
                and re.match(r'^[\dA-Z\-]+$', desc_upper)  # Only digits/letters/dashes (no spaces, no commas)
                and (any(c.isdigit() for c in desc_upper) or '-' in desc_upper)  # Has digits OR dash
            )
            if looks_like_id:
                continue
            
            # Extract NSN from material column
            nsn = ""
            if indices['material'] is not None and indices['material'] < len(row):
                mat_cell = row[indices['material']]
                nsn = extract_nsn_from_material(mat_cell)
            
            # Extract quantity from Auth Qty column
            qty = 1  # Default
            if indices['auth_qty'] is not None and indices['auth_qty'] < len(row):
                qty_cell = row[indices['auth_qty']]
                if qty_cell:
                    qty = extract_quantity(qty_cell)
            
            # Always use EA for unit of issue
            items.append(BomItem(
                line_no=len(items) + 1,
                description=description[:100],  # Limit length
                nsn=nsn,
                qty=qty,
                unit_of_issue="EA"  # Always EA
            ))
    
    return items


def extract_items_epp_format(tables: List[List[List[str]]], page_text: str) -> List[BomItem]:
    """
    Extract items from EPP (Equipment Property Record) format BOM.
    
    EPP format typically has:
    - Material column with NIIN/part numbers
    - Description column
    - Auth Qty column
    - OH Qty column (THIS IS WHAT WE USE)
    - May not have LV column
    
    Uses OH Qty for quantities, always uses EA for unit of issue.
    Skips items with 0 quantity.
    
    Args:
        tables: List of tables extracted from PDF
        page_text: Full page text for fallback parsing
        
    Returns:
        List of BomItem objects
    """
    items = []
    
    for table in tables:
        if not table or len(table) < 2:
            continue
        
        header = table[0]
        indices = find_column_indices(header)
        
        # EPP format detection: has Material and Description but may not have LV
        has_material = indices['material'] is not None
        has_description = indices['description'] is not None
        has_lv = indices['lv'] is not None
        
        # If no standard columns found, try to detect by content
        if not has_description:
            # Try to find columns by position/content
            for i, cell in enumerate(header):
                if not cell:
                    continue
                text = str(cell).upper()
                # Sometimes Description is just "DESCRIPTION" or contains it
                if 'DESCR' in text or text == 'DESC':
                    indices['description'] = i
                    has_description = True
        
        if not has_description:
            continue
        
        for row in table[1:]:
            if not any(cell for cell in row if cell):
                continue
            
            # If LV column exists, check for 'B' level items
            # But EPP format often doesn't have LV column
            if has_lv and indices['lv'] is not None:
                lv_cell = row[indices['lv']] if indices['lv'] < len(row) else None
                if lv_cell and str(lv_cell).strip().upper() == 'A':
                    # Skip category headers (A level)
                    continue
            
            # Extract description
            desc_cell = row[indices['description']] if indices['description'] < len(row) else None
            description = clean_description(desc_cell)
            
            if not description:
                continue
            
            # Skip obvious header/category rows (substring match, not exact)
            skip_patterns = [
                'COMPONENT OF END ITEM', 'BASIC ISSUE ITEMS', 
                'OPERATIONAL SUPPORT', 'COEI-', 'BII-',
            ]
            if any(pat in description.upper() for pat in skip_patterns):
                continue
            
            # Skip ID-like descriptions (e.g., "WH12B0", "T59652-014120143")
            desc_upper = description.upper()
            looks_like_id = (
                len(description) < 20
                and re.match(r'^[\dA-Z\-]+$', desc_upper)
                and (any(c.isdigit() for c in desc_upper) or '-' in desc_upper)
            )
            if looks_like_id:
                continue
            
            # Extract NSN from material column
            nsn = ""
            if indices['material'] is not None and indices['material'] < len(row):
                nsn = extract_nsn_from_material(row[indices['material']])
            
            # Extract quantity from Auth Qty column
            qty = 1
            if indices['auth_qty'] is not None and indices['auth_qty'] < len(row):
                qty = extract_quantity(row[indices['auth_qty']])
            
            # Always use EA for unit of issue
            items.append(BomItem(
                line_no=len(items) + 1,
                description=description[:100],
                nsn=nsn,
                qty=qty,
                unit_of_issue="EA"
            ))
    
    return items


def extract_items_da2062(tables: List[List[List[str]]], page_text: str) -> List[BomItem]:
    """
    Extract items from DA Form 2062 (Hand Receipt/Shortage Listing).
    
    DA 2062 can have different column layouts:
    - STOCK NUMBER or MATERIAL NUMBER column (NSN)
    - ITEM DESCRIPTION column
    - UI column
    - QTY AUTH column
    
    Args:
        tables: List of tables extracted from the page
        page_text: Full page text (used for validation)
        
    Returns:
        List of BomItem objects
    """
    items = []
    
    for table in tables:
        if not table or len(table) < 4:
            continue
        
        # Find the header row and column indices
        # Look for row with "STOCK NUMBER" or "MATERIAL NUMBER" as column header
        # (not "END ITEM STOCK NUMBER" which is a different row)
        header_row_idx = -1
        nsn_col = -1
        desc_col = -1
        
        for i, row in enumerate(table[:8]):
            # Check each cell for column header patterns
            for col_idx, cell in enumerate(row):
                cell_text = str(cell or '').upper().strip()
                
                # Look for STOCK NUMBER or MATERIAL NUMBER as a column header
                # (starts with it, not contains "END ITEM")
                if cell_text.startswith('STOCK NUMBER') or cell_text.startswith('MATERIAL NUMBER'):
                    if 'END ITEM' not in cell_text:
                        nsn_col = col_idx
                        header_row_idx = i
                
                # Look for ITEM DESCRIPTION column
                if cell_text.startswith('ITEM DESCRIPTION'):
                    desc_col = col_idx
                    if header_row_idx < 0:
                        header_row_idx = i
            
            # If we found both columns in this row, stop looking
            if nsn_col >= 0 and desc_col >= 0:
                break
        
        if header_row_idx < 0 or (nsn_col < 0 and desc_col < 0):
            continue
        
        # Default column indices if not found in header
        if nsn_col < 0:
            nsn_col = 0
        if desc_col < 0:
            desc_col = 2 if nsn_col < 2 else 3
        
        # Process rows after header (skip subheader row if present)
        start_row = header_row_idx + 1
        if start_row < len(table) and table[start_row]:
            # Check if this is a subheader row (contains A, B, C... or a., b., c., etc.)
            row_text = ' '.join(str(cell or '') for cell in table[start_row]).upper().strip()
            if re.match(r'^[A-F\s\.]+$', row_text) or not any(c.isalpha() and len(c) > 2 for c in [str(x or '') for x in table[start_row]]):
                start_row += 1
        
        for row in table[start_row:]:
            if not row or len(row) < max(nsn_col, desc_col) + 1:
                continue
            
            # Skip empty rows
            if not any(cell for cell in row if cell and str(cell).strip()):
                continue
            
            # Extract stock number - parse to 9-digit NIIN format for consistency
            nsn = ""
            if nsn_col < len(row) and row[nsn_col]:
                nsn = extract_nsn_from_material(str(row[nsn_col]))
            
            # Also check adjacent cells for NSN if not found
            if not nsn:
                for check_col in range(max(0, nsn_col-1), min(len(row), nsn_col+2)):
                    if row[check_col]:
                        candidate = extract_nsn_from_material(str(row[check_col]))
                        if candidate:
                            nsn = candidate
                            break
            
            # Extract description
            description = ""
            if desc_col < len(row) and row[desc_col]:
                desc_text = str(row[desc_col]).strip()
                # Take first line, remove reference numbers
                lines = desc_text.split('\n')
                for line in lines:
                    line = line.strip()
                    # Skip lines that are just numbers/references
                    if line and not re.match(r'^[\d\(\)\s\-]+$', line):
                        description = line
                        break
                
                # Clean up description
                description = re.sub(r'\s+', ' ', description).strip()
                description = re.sub(r'\s*\([^)]*\)\s*$', '', description)  # Remove trailing (...)
            
            if not description or len(description) < 3:
                continue
            
            # Skip footer/header text
            skip_patterns = ['NOTHING FOLLOWS', 'DA FORM', 'HAND RECEIPT', 'PAGE']
            if any(pat in description.upper() for pat in skip_patterns):
                continue
            
            # Extract quantity (try columns 9, 10 for Auth qty)
            qty = 1
            for col_idx in [9, 10, 11]:
                if len(row) > col_idx and row[col_idx]:
                    qty_val = extract_quantity(row[col_idx])
                    if qty_val > 0:
                        qty = qty_val
                        break
            
            items.append(BomItem(
                line_no=len(items) + 1,
                description=description[:100],
                nsn=nsn,
                qty=qty,
                unit_of_issue="EA"
            ))
    
    return items


def extract_metadata(page_text: str) -> BomMetadata:
    """
    Extract metadata from BOM header text.
    
    Args:
        page_text: Full text content of the first page
        
    Returns:
        BomMetadata object with extracted values
    """
    metadata = BomMetadata()
    
    # END ITEM NIIN
    match = re.search(r'END\s*ITEM\s*NIIN[:\s]*(\d{9})', page_text, re.IGNORECASE)
    if match:
        metadata.end_item_niin = match.group(1)
    
    # LIN
    match = re.search(r'LIN[:\s]*([A-Z0-9]+)', page_text, re.IGNORECASE)
    if match:
        metadata.lin = match.group(1)
    
    # End-Item Description (the DESC that comes after LIN:, not the unit DESC)
    # The BOM header has two DESC fields:
    #   "FE: ... UIC: ... DESC: <unit>"               (we don't want this)
    #   "END ITEM NIIN: ... LIN: ... DESC: <model>"   (this is the one)
    # Anchor on the LIN: prefix to grab the right one. Stop at newline.
    match = re.search(
        r'LIN[:\s]*[A-Z0-9]+\s+DESC[:\s]*([^\n\r]+)',
        page_text, re.IGNORECASE,
    )
    if match:
        # Trim trailing whitespace and stop at large gaps (which indicate
        # the next column on the same printed line)
        desc = match.group(1).strip()
        # If there's a 3+ space gap, take only what's before it
        gap_match = re.search(r'\s{3,}', desc)
        if gap_match:
            desc = desc[:gap_match.start()].strip()
        metadata.end_item_description = desc[:50]
    
    # Serial/Equipment Number (stop at newline / multi-space gap)
    match = re.search(r'SER/EQUIP\s*NO[:\s]*([A-Z0-9]+)', page_text, re.IGNORECASE)
    if match:
        val = match.group(1)
        # If the SER/EQUIP NO field is empty in the source, our match may have
        # captured the next adjacent field label (e.g. "TO", "FROM"). Skip those.
        # We also check whether the very next non-space char is ":", which is
        # a strong signal we matched a label, not a value.
        next_chars = page_text[match.end(1):match.end(1) + 3].lstrip()
        is_label = (val.upper() in {'TO', 'FROM', 'SLOC', 'PUB', 'DATE'}
                    or next_chars.startswith(':'))
        if not is_label:
            metadata.serial_equip_no = val
    
    # UIC
    match = re.search(r'UIC[:\s]*([A-Z0-9]+)', page_text, re.IGNORECASE)
    if match:
        metadata.uic = match.group(1)
    
    # FE
    match = re.search(r'FE[:\s]*(\d+)', page_text, re.IGNORECASE)
    if match:
        metadata.fe = match.group(1)
    
    return metadata


def extract_metadata_from_form_fields(pdf_path: str, metadata: BomMetadata) -> BomMetadata:
    """
    Fill in metadata fields by scanning PDF form fields.
    
    Used as a fallback for form-only BOMs where the page text stream is empty
    and `extract_metadata` couldn't pull values from text. Only fills in fields
    that are currently empty — never overrides values already set from text.
    
    Args:
        pdf_path: Path to the PDF file
        metadata: Existing metadata (will be mutated and returned)
        
    Returns:
        The same BomMetadata, with empty fields populated from form data when found
    """
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields() or {}
        if not fields:
            return metadata
        
        # END ITEM NIIN: extracted from COEI<NIIN> or BII<NIIN> field name patterns
        if not metadata.end_item_niin:
            for name, field in fields.items():
                tip = str(field.get('/TU', '') or '')
                # Look for "COEI-015342228" in tooltip OR "COEI015342228" in name
                m = re.search(r'(?:COEI|BII)[\s\-]*(\d{9})', tip, re.IGNORECASE)
                if not m:
                    m = re.search(r'(?:COEI|BII)[\s\-]*(\d{9})', name, re.IGNORECASE)
                if m:
                    metadata.end_item_niin = m.group(1)
                    break
        
        # SER/EQUIP NO: appears as the value in fields named "undefined"
        # (the form designer didn't give the SER/EQUIP cell a meaningful name).
        # We accept any value that looks like a serial/equipment number:
        # alphanumeric, length 5-25, contains at least one digit.
        if not metadata.serial_equip_no:
            for name, field in fields.items():
                if name.lower().startswith('undefined'):
                    val = str(field.get('/V', '') or '').strip()
                    if (val and 5 <= len(val) <= 25
                            and re.match(r'^[A-Z0-9]+$', val.upper())
                            and any(c.isdigit() for c in val)):
                        metadata.serial_equip_no = val
                        break
        
        # SLOC: directly in a field named "SLOC"
        # We don't have a dedicated metadata slot for it; store it in a field
        # the caller can read if useful. For now, only used informally.
        
    except Exception:
        # Best-effort - don't fail the whole extraction over metadata
        pass
    
    return metadata


def extract_items_via_ocr(pdf_path: str, dpi: int = 250) -> List[BomItem]:
    """
    Extract BOM items by running OCR on every page.
    
    This is the last-resort extraction path used when both text extraction
    (pdfplumber) and form-field extraction yielded incomplete results — e.g.
    when a BOM was generated as form-only with table data flattened into
    page images (common with newer GCSS-Army exports).
    
    Strategy:
      1. Render each page to an image at the given DPI.
      2. Use Tesseract's image_to_data to locate every word with its X/Y
         bounding box.
      3. Find 9-character NIIN tokens in the Material column (left half of page).
      4. Group words into rows by Y proximity, then per row, take the
         description column words (middle of page) as the description.
      5. Skip header rows (page 1 always has end-item NIIN at the very top)
         and totals/footer rows.
    
    Args:
        pdf_path: Path to BOM PDF
        dpi: Render DPI - 250 is a sweet spot for accuracy vs speed
        
    Returns:
        List of BomItem objects, one per detected row.
        Empty list if OCR is unavailable or fails.
    """
    if not OCR_AVAILABLE:
        return []
    
    items: List[BomItem] = []
    
    # NIIN token patterns (must match the entire word)
    nine_digit  = re.compile(r'^\d{9}$')
    alphanum_9  = re.compile(r'^(\d{2}[A-Z]\d{6}|\d{2}[A-Z]{2}\d{5}|\d{3}[A-Z]\d{5})$')
    full_nsn    = re.compile(r'^\d{4}-\d{2}-\d{3}-\d{4}$')
    
    try:
        images = _convert_from_path(pdf_path, dpi=dpi)
    except Exception:
        return []
    
    for page_num, img in enumerate(images, 1):
        try:
            data = _pytesseract.image_to_data(
                img, output_type=_OCR_Output.DICT, config='--psm 6',
            )
        except Exception:
            continue
        
        img_w, img_h = img.size
        
        # Approximate column boundaries (proportional to image width).
        # Real GCSS-Army Component Listings have columns at:
        #   Image | Material | LV | Description | WTY | ARC | CIIC | UI | SCMC | AuthQty | OHQty
        # The Material column starts ~20% across, Description column ~45%,
        # Qty columns at the far right ~88%+.
        material_x_lo = int(img_w * 0.18)
        material_x_hi = int(img_w * 0.42)
        desc_x_lo     = int(img_w * 0.43)
        desc_x_hi     = int(img_w * 0.78)
        oh_qty_x_lo   = int(img_w * 0.92)
        
        # Skip the page header (above the table) and the page footer
        header_y_cutoff = int(img_h * 0.13)
        footer_y_cutoff = int(img_h * 0.95)
        
        # Build word list with positions
        words = []
        for i, txt in enumerate(data['text']):
            t = (txt or '').strip()
            if not t:
                continue
            try:
                conf = int(data['conf'][i])
            except (ValueError, IndexError):
                conf = 0
            words.append({
                'y': int(data['top'][i]),
                'x': int(data['left'][i]),
                'h': int(data['height'][i]),
                'text': t,
                'conf': conf,
            })
        
        # Detect signature block boundary. The DD1750 / GCSS Component Listing
        # has an "ISSUED BY / RECEIVED BY / SIGNATURE" footer that only appears
        # on the LAST page. Anything past this block is signature metadata, not
        # items — including digital-signature ID numbers that look like NIINs
        # (e.g. "151478505" picked up from "HOLLAND.JAMES.M.1514795050").
        # Find the highest Y of any of these marker words; treat that as the
        # effective footer cutoff for item extraction on this page.
        SIGNATURE_MARKERS = {
            'SIGNATURE', 'SIGNATURE:', 'ISSUED', 'RECEIVED',
            'HOLLAND.JAME', 'HOLLAND.', 'RABATIN.', 'GRADE:', 'GRADE',
        }
        signature_y = footer_y_cutoff
        for w in words:
            ut = w['text'].upper().rstrip(':')
            if ut in SIGNATURE_MARKERS or 'SIGNATURE' in ut or '.JAME' in ut.upper():
                # Stop extraction a bit ABOVE the signature line so we don't
                # catch printed names from the ISSUED BY / RECEIVED BY row
                signature_y = min(signature_y, w['y'] - 30)
        
        effective_footer_y = min(footer_y_cutoff, signature_y)
        
        # Find NIIN words inside the Material column.
        # Use effective_footer_y (which respects the signature-block boundary)
        # rather than the raw footer_y_cutoff so we don't catch digital
        # signature IDs that happen to be 9 digits.
        material_niins = []
        for w in words:
            if w['y'] < header_y_cutoff or w['y'] > effective_footer_y:
                continue
            if not (material_x_lo <= w['x'] <= material_x_hi):
                continue
            t = w['text']
            niin = None
            if nine_digit.match(t):
                niin = t
            elif alphanum_9.match(t):
                niin = t
            elif full_nsn.match(t):
                # Full NSN: drop the 4-digit FSC, keep the 9-digit NIIN
                parts = t.split('-')
                niin = parts[1] + parts[2] + parts[3]
            
            if niin:
                material_niins.append({
                    'y_center': w['y'] + w['h'] // 2,
                    'y_top':    w['y'],
                    'nsn':      niin,
                })
        
        if not material_niins:
            continue
        
        material_niins.sort(key=lambda r: r['y_center'])
        
        # Compute average row spacing (used as a window for matching descriptions).
        # Each table row is typically 200-250 pixels tall at 250 DPI.
        if len(material_niins) >= 2:
            spacings = [material_niins[i+1]['y_center'] - material_niins[i]['y_center']
                        for i in range(len(material_niins) - 1)]
            spacings.sort()
            row_spacing = spacings[len(spacings) // 2]  # median
        else:
            row_spacing = int(img_h * 0.10)
        
        # The description for a row can start ABOVE the NIIN (the description
        # column has multi-line text starting at the top of the cell, while
        # the NIIN sits roughly in the vertical middle of the cell).
        # Bias the window: more space above than below.
        window_above = max(70, int(row_spacing * 0.65))
        window_below = max(70, int(row_spacing * 0.55))
        
        for niin_row in material_niins:
            yc = niin_row['y_center']
            y_min = yc - window_above
            y_max = yc + window_below
            
            # Don't let this row's window overlap the NEXT row's NIIN
            # (or the previous row's NIIN), so descriptions stay within their cells.
            niin_index = material_niins.index(niin_row)
            if niin_index > 0:
                prev_y = material_niins[niin_index - 1]['y_center']
                # Cut the upper window so we don't dip into the previous row
                y_min = max(y_min, (prev_y + yc) // 2)
            if niin_index < len(material_niins) - 1:
                next_y = material_niins[niin_index + 1]['y_center']
                y_max = min(y_max, (yc + next_y) // 2)
            
            # Description: words in the description column within the Y window.
            # We take the FIRST line of the description (highest words within the window)
            desc_words = [w for w in words
                          if y_min <= (w['y'] + w['h']//2) <= y_max
                          and desc_x_lo <= w['x'] <= desc_x_hi]
            if not desc_words:
                continue
            
            # Strip OCR garbage that often appears between rows of the table:
            #  - column headers ("LV", "Description", "WTY", etc.)
            #  - single non-letter glyphs ("P", "|", ".", ",")
            #  - "B" by itself (the LV column code that bleeds into desc col on
            #    pages where Tesseract misjudges column boundaries)
            COLUMN_HEADERS = {
                'IMAGE', 'MATERIAL', 'LV', 'DESCRIPTION', 'WTY', 'ARC',
                'CIIC', 'UI', 'SCMC', 'AUTH', 'OH', 'QTY',
            }
            
            def is_garbage(w_text: str) -> bool:
                t = w_text.strip()
                if not t:
                    return True
                if t.upper() in COLUMN_HEADERS:
                    return True
                # Single character that isn't part of a real word
                if len(t) == 1 and not t.isalpha():
                    return True
                # Common single-letter LV-column noise
                if t in ('B', 'P', '|', 'a', 'b', 'I', 'l'):
                    return True
                return False
            
            desc_words = [w for w in desc_words if not is_garbage(w['text'])]
            if not desc_words:
                continue
            
            desc_words.sort(key=lambda w: (w['y'], w['x']))
            
            # Take the topmost row of words (within ~30 px of the highest)
            top_y = desc_words[0]['y']
            first_line = [w for w in desc_words if abs(w['y'] - top_y) <= 30]
            first_line.sort(key=lambda w: w['x'])
            description = ' '.join(w['text'] for w in first_line)
            
            # Clean up common OCR artifacts
            description = re.sub(r'\s+', ' ', description).strip()
            description = re.sub(r'^[^A-Z]+', '', description)  # leading non-letters
            description = re.sub(r'[._]+\s*$', '', description)  # trailing dots
            
            if not description or len(description) < 3:
                continue
            
            # Skip header-row remnants
            if any(skip in description.upper() for skip in
                   ['END ITEM', 'COMPONENT OF', 'BASIC ISSUE', 'NOTHING FOLLOWS']):
                continue
            
            # Find OH Qty for this row (rightmost column, same Y window)
            qty = 1
            oh_qty = -1
            qty_words = [w for w in words
                         if y_min <= (w['y'] + w['h']//2) <= y_max
                         and w['x'] >= oh_qty_x_lo
                         and re.match(r'^\d+$', w['text'])]
            if qty_words:
                # Pick the one closest to the row's Y center
                qty_words.sort(key=lambda w: abs((w['y'] + w['h']//2) - yc))
                try:
                    oh_qty = int(qty_words[0]['text'])
                    if oh_qty > 0:
                        qty = oh_qty
                except ValueError:
                    pass
            
            items.append(BomItem(
                line_no=len(items) + 1,
                description=description[:100],
                nsn=niin_row['nsn'],
                qty=qty,
                unit_of_issue="EA",
                oh_qty=oh_qty,
            ))
    
    return items


def extract_items_from_form_fields(pdf_path: str) -> List[BomItem]:
    """
    Extract BOM items from PDF form fields when text extraction fails.
    
    Some GCSS-Army BOMs are "form-only" PDFs where the page text stream is
    empty and all content is stored as form field annotations. Form fields
    appear in document order following this pattern for each item:
    
        MATERIAL field   - tooltip starts with 9-char NIIN, e.g. "011661384 1766590W:C_75Q65"
        WTY field        - tooltip starts with "WTY_", contains the description
        OH Qty field     - holds the on-hand quantity value
    
    Category headers (COEI-XXXXXXXXX, BII-XXXXXXXXX) follow a different
    pattern with no NIIN and should be skipped.
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        List of BomItem objects with NSN populated where available
    """
    items = []
    
    try:
        reader = PdfReader(pdf_path)
        fields = reader.get_fields() or {}
        
        if not fields:
            return items
        
        # Walk fields in document order, tracking state as we go.
        # Python preserves dict insertion order (3.7+), and PdfReader.get_fields()
        # returns fields in their PDF document order.
        pending_nsn = ""
        pending_material_text = ""
        last_item_idx = -1  # index into `items` of the most recently created item
        
        # Patterns for recognizing material field tooltips/names
        # GCSS material fields look like: "011661384 1766590W:C 75Q65 -"
        # or with alphanumeric NIIN: "01C079749 ..."
        nsn_at_start_re = re.compile(
            r'^(\d{9}|\d{2}[A-Z]\d{6}|\d{2}[A-Z]{2}\d{5}|\d{3}[A-Z]\d{5})\b'
        )
        # Pattern for skipping category-header field tooltips
        is_category = lambda t: bool(re.search(r'\b(COEI|BII)-\d', t, re.IGNORECASE))
        
        for name, field in fields.items():
            tooltip = str(field.get('/TU', '') or '')
            value = str(field.get('/V', '') or '')
            tooltip_stripped = tooltip.strip()
            name_stripped = name.strip()
            
            # ---- Skip metadata/non-item fields ----
            if not tooltip_stripped and not name_stripped:
                continue
            
            # Top-level metadata fields - skip
            metadata_names = {
                'SLOC', 'TO', 'FROM', 'DATE', 'GRADE', 'SIGNATURE',
                'undefined', 'PUB NUM', 'PUB/BOM', 'EA',
            }
            if name_stripped in metadata_names or tooltip_stripped in metadata_names:
                continue
            
            # Category header fields (COEI-XXXX, BII-XXXX) - skip and reset state
            # so the next material field starts fresh
            if is_category(tooltip_stripped) or is_category(name_stripped):
                pending_nsn = ""
                pending_material_text = ""
                continue
            
            # ---- Detect MATERIAL fields ----
            # Tooltip or name starts with a 9-char NIIN pattern
            mat_match = nsn_at_start_re.match(tooltip_stripped) or nsn_at_start_re.match(name_stripped)
            if mat_match:
                pending_nsn = mat_match.group(1)
                pending_material_text = tooltip_stripped or name_stripped
                continue
            
            # Material fields without a NIIN (just part numbers) — record but no NSN
            # Example: "T25050T:C_0WFM3" or "13632952-CBLE:C_18876"
            # These look like part-number-with-cage format
            looks_like_part_only = (
                re.match(r'^[A-Z0-9][\w\-]+\s*:\s*C[_ ]\w+', tooltip_stripped) or
                re.match(r'^[A-Z0-9][\w\-]+\s*:\s*C[_ ]\w+', name_stripped)
            )
            if looks_like_part_only:
                pending_nsn = ""
                pending_material_text = tooltip_stripped or name_stripped
                continue
            
            # ---- Detect WTY/DESC fields (the description field) ----
            # Three known prefix patterns in GCSS-Army form-only BOMs:
            #   "WTY_..."         - normal pattern
            #   "9_..."           - alternate pattern
            #   "Description ..." - condensed pattern (some pages)
            desc_prefix = None
            if tooltip_stripped.startswith('WTY_'):
                desc_prefix = 'WTY_'
            elif tooltip_stripped.startswith('9_'):
                desc_prefix = '9_'
            elif tooltip_stripped.startswith('Description '):
                desc_prefix = 'Description '
            
            if desc_prefix:
                desc = tooltip_stripped[len(desc_prefix):].strip()
                
                # The tooltip often duplicates the nomenclature
                # Format: "SHORT,NAME FULL DESCRIPTION..." where SHORT NAME repeats
                # Try to detect repetition and keep just the first instance
                parts = desc.split()
                if len(parts) > 1:
                    first_word = parts[0].replace(',', '').upper()
                    for i, part in enumerate(parts[1:], 1):
                        if part.replace(',', '').upper() == first_word:
                            desc = ' '.join(parts[:i])
                            break
                
                # Clean up
                desc = re.sub(r'\s+', ' ', desc).strip()
                desc = re.sub(r',+', ',', desc)
                
                if not desc or len(desc) < 3:
                    continue
                
                # Skip if it still looks like a category header
                if is_category(desc):
                    continue
                
                # Create the item, pairing it with the most recent NSN
                items.append(BomItem(
                    line_no=len(items) + 1,
                    description=desc[:100],
                    nsn=pending_nsn,
                    qty=1,  # Default; updated when we see the OH Qty field
                    unit_of_issue="EA",
                    material_number=pending_material_text,
                ))
                last_item_idx = len(items) - 1
                
                # Reset NIIN so it doesn't accidentally bind to the next item
                pending_nsn = ""
                pending_material_text = ""
                continue
            
            # ---- Detect OH Qty fields and attach to most recent item ----
            is_oh_qty = (
                'OH Qty' in name or 'OH Qty' in tooltip or
                'oh_qty' in name.lower() or 'oh qty' in tooltip.lower()
            )
            if is_oh_qty and last_item_idx >= 0:
                if value and str(value).strip().isdigit():
                    qty_val = int(str(value).strip())
                    items[last_item_idx].oh_qty = qty_val
                    if qty_val > 0:
                        items[last_item_idx].qty = qty_val
                continue
        
    except Exception:
        # Silently fail - caller will fall back to other methods
        pass
    
    return items


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> ExtractionResult:
    """
    Extract BOM items from a PDF file.
    
    Supports multiple BOM formats from GCSS-Army:
    - Standard Component Listing / Hand Receipt with LV column
    - EPP format
    - Form-only PDFs (content in form fields, not extractable text)
    
    Args:
        pdf_path: Path to the BOM PDF file
        start_page: Page number to start extraction (0-based)
        
    Returns:
        ExtractionResult containing items, metadata, and any warnings/errors
    """
    result = ExtractionResult()
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if start_page >= len(pdf.pages):
                result.errors.append(f"Start page {start_page} exceeds document length ({len(pdf.pages)} pages)")
                return result
            
            # Get first page text for metadata and format detection
            first_page = pdf.pages[start_page]
            first_page_text = first_page.extract_text() or ""
            first_page_tables = first_page.extract_tables()
            
            # Check if PDF has extractable text
            has_text = len(first_page_text.strip()) > 50
            
            # Detect format
            result.format_detected = detect_bom_format(first_page_tables, first_page_text)
            result.metadata = extract_metadata(first_page_text)
            
            # If page text is empty (form-only PDF), fill in metadata from form fields
            if not has_text:
                result.metadata = extract_metadata_from_form_fields(pdf_path, result.metadata)
            
            result.metadata.bom_format = result.format_detected
            
            # Extract items from all pages
            all_items = []
            
            if has_text:
                # Normal extraction from text/tables
                for page_num, page in enumerate(pdf.pages[start_page:], start=start_page):
                    result.pages_processed += 1
                    tables = page.extract_tables()
                    page_text = page.extract_text() or ""
                    
                    if result.format_detected == BomFormat.GCSS_ARMY_STANDARD:
                        page_items = extract_items_gcss_standard(tables)
                    elif result.format_detected == BomFormat.EPP_FORMAT:
                        page_items = extract_items_epp_format(tables, page_text)
                    elif result.format_detected == BomFormat.DA_2062:
                        page_items = extract_items_da2062(tables, page_text)
                    else:
                        # Try standard format as fallback
                        page_items = extract_items_gcss_standard(tables)
                        if not page_items:
                            page_items = extract_items_epp_format(tables, page_text)
                        if not page_items:
                            page_items = extract_items_da2062(tables, page_text)
                    
                    all_items.extend(page_items)
            
            # If no items extracted via normal methods, this is likely a
            # form-only or scanned PDF. Run BOTH form-field extraction and OCR,
            # then merge: OCR gives us every visible row (including ones where
            # the data is rasterized into images), and form-fields give us the
            # cleanest descriptions where the data is editable text. We always
            # run both because partial form-field coverage doesn't tell us how
            # many items we missed — only OCR can see all rows on the page.
            if not all_items:
                form_items = extract_items_from_form_fields(pdf_path)
                ocr_items: List[BomItem] = []
                
                if OCR_AVAILABLE:
                    try:
                        ocr_items = extract_items_via_ocr(pdf_path)
                    except Exception as e:
                        result.warnings.append(f"OCR fallback failed: {e}")
                        ocr_items = []
                
                # Merge: when OCR is available and found items, OCR is the
                # source of truth for the row count (it sees every visible row,
                # including ones with rasterized data). Form-field descriptions
                # are usually cleaner than OCR text, so when an OCR NIIN matches
                # a form-field NIIN, prefer the form-field description and qty.
                if ocr_items and form_items:
                    form_by_nsn = {it.nsn: it for it in form_items if it.nsn}
                    for it in ocr_items:
                        if it.nsn and it.nsn in form_by_nsn:
                            ff = form_by_nsn[it.nsn]
                            if ff.description and len(ff.description) >= 3:
                                it.description = ff.description
                            if ff.qty and ff.qty > 0:
                                it.qty = ff.qty
                            if ff.oh_qty != -1:
                                it.oh_qty = ff.oh_qty
                    all_items = ocr_items
                    extraction_source = 'ocr+form'
                elif ocr_items:
                    all_items = ocr_items
                    extraction_source = 'ocr'
                elif form_items:
                    all_items = form_items
                    extraction_source = 'form'
                else:
                    extraction_source = 'none'
                
                if all_items:
                    result.format_detected = BomFormat.GCSS_ARMY_STANDARD
                    result.pages_processed = len(pdf.pages)
                    
                    final_with_nsn = sum(1 for it in all_items if it.nsn)
                    if extraction_source == 'ocr+form':
                        result.warnings.append(
                            f"This PDF stored data as images. Used OCR + form fields to extract "
                            f"{len(all_items)} items ({final_with_nsn} with NSN). "
                            f"Please verify against the original PDF before printing."
                        )
                    elif extraction_source == 'ocr':
                        result.warnings.append(
                            f"This PDF stored data as images. Used OCR to extract "
                            f"{len(all_items)} items ({final_with_nsn} with NSN). "
                            f"OCR is best-effort — please verify each item against the original PDF "
                            f"before printing, especially NSN digits and descriptions."
                        )
                    elif extraction_source == 'form':
                        result.warnings.append(
                            f"This PDF stores data as form fields. Extracted "
                            f"{len(all_items)} items ({final_with_nsn} with NSN). "
                            f"Some items may have been rendered as images that couldn't be "
                            f"recovered — please verify against the original PDF and add any "
                            f"missing items in the review screen."
                        )
            
            # Renumber items
            for i, item in enumerate(all_items):
                item.line_no = i + 1
            
            result.items = all_items
            
            if not result.items:
                result.warnings.append("No items extracted. This PDF may be a scanned image - try using a digital BOM from GCSS-Army.")
            
    except Exception as e:
        result.errors.append(f"Failed to process PDF: {str(e)}")
    
    return result


def generate_dd1750_overlay(
    items: List[BomItem], 
    page_num: int, 
    total_pages: int,
    header: Optional[HeaderInfo] = None
) -> io.BytesIO:
    """
    Generate a PDF overlay with item data for a single DD1750 page.
    
    Fills in:
    - Page numbers (automatically calculated)
    - Table items
    
    Form fields are added separately after the merge in generate_dd1750_from_items.
    
    Args:
        items: List of items for this page (max 18)
        page_num: Current page number (1-based)
        total_pages: Total number of pages
        header: Optional header information (not used - kept for API compatibility)
        
    Returns:
        BytesIO buffer containing the overlay PDF
    """
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
    
    # === HEADER FIELDS ===
    # PAGE NUMBERS - Always fill these in as static text
    can.setFont("Helvetica", 10)
    can.drawCentredString(472, PAGE_H - 132, str(page_num))      # Current page
    can.drawCentredString(520, PAGE_H - 132, str(total_pages))   # Total pages
    
    # When a header is provided with pre-filled values, draw them directly on
    # the canvas. The caller is responsible for telling generate_dd1750_from_items
    # to skip creating empty form fields for these positions.
    if header is not None:
        # Packed By (single line; left-aligned)
        if header.packed_by:
            can.setFont("Helvetica", 9)
            can.drawString(95, 736, header.packed_by[:60])
        
        # Number of Boxes (centered)
        if header.num_boxes:
            can.setFont("Helvetica", 9)
            can.drawString(285, 736, str(header.num_boxes)[:10])
        
        # Requisition / Order Number
        if header.requisition_no:
            can.setFont("Helvetica", 9)
            can.drawString(408, 736, str(header.requisition_no)[:30])
        if header.order_no:
            can.setFont("Helvetica", 9)
            can.drawString(408, 716, str(header.order_no)[:30])
        
        # End Item (may be multi-line: NOMENCLATURE / MODEL / SERIAL NUMBER)
        if header.end_item:
            can.setFont("Helvetica", 8)
            lines = header.end_item.split('\n')
            # First line at the standard position; subsequent lines below
            y_top = 696
            for i, line in enumerate(lines[:3]):
                can.drawString(95, y_top - (i * 10), line[:55])
        
        # Date
        if header.date:
            can.setFont("Helvetica", 9)
            can.drawString(450, 696, str(header.date)[:20])
    
    # === TABLE CONTENT ===
    for i, item in enumerate(items):
        # Calculate Y position for this row (rows go top to bottom)
        row_top = Y_TABLE_TOP - (i * ROW_H)
        y_line1 = row_top - 10.0    # First line (description)
        y_line2 = row_top - 20.0    # Second line (NSN)
        
        # Line number in Box column (centered)
        can.setFont("Helvetica", 9)
        box_center_x = (X_BOX_L + X_BOX_R) / 2
        can.drawCentredString(box_center_x, y_line1, str(item.line_no))
        
        # Description (left-aligned with padding)
        can.setFont("Helvetica", 8)
        desc = item.description[:55] if len(item.description) > 55 else item.description
        can.drawString(X_CONTENT_L + PAD_X, y_line1, desc)
        
        # NSN on second line if present
        if item.nsn:
            can.setFont("Helvetica", 7)
            can.drawString(X_CONTENT_L + PAD_X, y_line2, f"NSN: {item.nsn}")
        
        # Unit of Issue (centered) - Always EA
        can.setFont("Helvetica", 9)
        uoi_center_x = (X_UOI_L + X_UOI_R) / 2
        can.drawCentredString(uoi_center_x, y_line1, "EA")
        
        # Initial Operation quantity (centered)
        init_center_x = (X_INIT_L + X_INIT_R) / 2
        can.drawCentredString(init_center_x, y_line1, str(item.qty))
        
        # Running Spares (centered) - always 0
        spares_center_x = (X_SPARES_L + X_SPARES_R) / 2
        can.drawCentredString(spares_center_x, y_line1, "0")
        
        # Total (centered)
        total_center_x = (X_TOTAL_L + X_TOTAL_R) / 2
        can.drawCentredString(total_center_x, y_line1, str(item.qty))
    
    # === "NOTHING FOLLOWS" MARKER ===
    # Drawn on the last page, on the row immediately after the last item.
    # If the page is completely full (18 items), we skip it — the user can
    # add a final blank page manually if needed.
    if page_num == total_pages and len(items) < ROWS_PER_PAGE:
        marker_row_top = Y_TABLE_TOP - (len(items) * ROW_H)
        marker_y = marker_row_top - 10.0
        marker_center_x = (X_CONTENT_L + X_CONTENT_R) / 2
        can.setFont("Helvetica-Bold", 8)
        can.drawCentredString(
            marker_center_x,
            marker_y,
            "------------------- NOTHING FOLLOWS -------------------"
        )
    
    can.save()
    packet.seek(0)
    return packet


def generate_dd1750_from_items(
    items: List[BomItem],
    template_path: str,
    output_path: str,
    header: Optional[HeaderInfo] = None
) -> Tuple[str, int]:
    """
    Generate DD1750 PDF from a list of items.
    
    Args:
        items: List of BomItem objects
        template_path: Path to blank DD1750 template PDF
        output_path: Path for output PDF
        header: Optional header information (packed by, date, etc.)
        
    Returns:
        Tuple of (output_path, item_count)
    """
    from pypdf.generic import (
        DictionaryObject, ArrayObject, NameObject, 
        TextStringObject, NumberObject, FloatObject
    )
    from pypdf.annotations import FreeText
    
    if not items:
        # Return blank template if no items
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
    
    total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
    writer = PdfWriter()
    
    for page_num in range(total_pages):
        start_idx = page_num * ROWS_PER_PAGE
        end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
        page_items = items[start_idx:end_idx]
        
        # Generate overlay with header info
        overlay_buffer = generate_dd1750_overlay(
            page_items, 
            page_num + 1, 
            total_pages,
            header
        )
        overlay = PdfReader(overlay_buffer)
        
        # Merge with template
        template_page = PdfReader(template_path).pages[0]
        template_page.merge_page(overlay.pages[0])
        writer.add_page(template_page)
    
    # Add fillable form fields to the first page — but ONLY for fields that
    # weren't pre-filled via the HeaderInfo (those are drawn on the canvas
    # by the overlay generator above and don't need editable form fields).
    # Define form field positions (x, y, width, height) based on DD1750 layout.
    all_form_fields = [
        {'name': 'packed_by', 'rect': (92, 732, 230, 746), 'tooltip': 'Packed By',
         'header_attr': 'packed_by'},
        {'name': 'no_boxes', 'rect': (282, 732, 332, 746), 'tooltip': 'Number of Boxes',
         'header_attr': 'num_boxes'},
        {'name': 'req_no', 'rect': (405, 732, 566, 746), 'tooltip': 'Requisition Number',
         'header_attr': 'requisition_no'},
        {'name': 'order_no', 'rect': (405, 712, 566, 726), 'tooltip': 'Order Number',
         'header_attr': 'order_no'},
        {'name': 'end_item', 'rect': (92, 689, 370, 703), 'tooltip': 'End Item',
         'header_attr': 'end_item'},
        {'name': 'date', 'rect': (447, 689, 566, 703), 'tooltip': 'Date',
         'header_attr': 'date'},
        {'name': 'typed_name', 'rect': (92, 46, 290, 60), 'tooltip': 'Typed Name and Title',
         'header_attr': None},
    ]
    
    # Filter out fields whose value was already drawn on the canvas
    form_fields = []
    for f in all_form_fields:
        attr = f.get('header_attr')
        if header is not None and attr and getattr(header, attr, ''):
            # Skip — value is already baked into the page via the overlay
            continue
        form_fields.append(f)
    
    # Create AcroForm for the document
    writer._root_object[NameObject("/AcroForm")] = DictionaryObject({
        NameObject("/Fields"): ArrayObject([]),
        NameObject("/NeedAppearances"): NameObject("/true")
    })
    
    # Add text fields to first page
    page = writer.pages[0]
    
    for field_def in form_fields:
        # Create text field annotation
        field = DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/FT"): NameObject("/Tx"),  # Text field
            NameObject("/T"): TextStringObject(field_def['name']),
            NameObject("/Rect"): ArrayObject([
                FloatObject(field_def['rect'][0]),
                FloatObject(field_def['rect'][1]),
                FloatObject(field_def['rect'][2]),
                FloatObject(field_def['rect'][3])
            ]),
            NameObject("/F"): NumberObject(4),  # Print flag
            NameObject("/Ff"): NumberObject(0),  # Field flags (editable)
            NameObject("/DA"): TextStringObject("/Helv 9 Tf 0 g"),  # Default appearance
            NameObject("/TU"): TextStringObject(field_def['tooltip']),  # Tooltip
            NameObject("/V"): TextStringObject(""),  # Initial value
            NameObject("/DV"): TextStringObject(""),  # Default value
        })
        
        # Add to page annotations
        if "/Annots" not in page:
            page[NameObject("/Annots")] = ArrayObject([])
        page[NameObject("/Annots")].append(field)
        
        # Add to AcroForm fields
        writer._root_object["/AcroForm"]["/Fields"].append(field)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path, len(items)


def format_packed_by(name: str, rank: str = "", unit: str = "") -> str:
    """
    Format packer information into a single-line PACKED BY string.
    
    Examples:
        format_packed_by("James M. Holland", "CPT", "B BTY 4-3 ADA")
            -> "CPT JAMES M. HOLLAND, B BTY 4-3 ADA"
        format_packed_by("Holland, James M.")
            -> "HOLLAND, JAMES M."
    """
    name = (name or "").strip()
    rank = (rank or "").strip()
    unit = (unit or "").strip()
    
    parts = []
    if rank:
        parts.append(rank.upper())
    if name:
        parts.append(name.upper())
    head = ' '.join(parts)
    
    if unit:
        return f"{head}, {unit.upper()}" if head else unit.upper()
    return head


def format_end_item(nomenclature: str, model: str = "", serial_number: str = "") -> str:
    """
    Format end-item info as the three-line block used on the DD1750.
    
    Output:
        NOMENCLATURE: <nom>
        MODEL: <model>
        SERIAL NUMBER: <sn>
    
    Empty values are still shown with the label so the form looks complete;
    callers can pass "" to leave a line blank.
    """
    return (
        f"NOMENCLATURE: {nomenclature or ''}\n"
        f"MODEL: {model or ''}\n"
        f"SERIAL NUMBER: {serial_number or ''}"
    )


@dataclass
class BatchBomEntry:
    """A single BOM in a batch redeployment job."""
    items: List[BomItem]
    nomenclature: str = ""        # User's nickname (e.g., "B49")
    model: str = ""               # End-item DESC from BOM (e.g., "TRK CGO W/W M985A4GMT")
    serial_number: str = ""       # SER/EQUIP NO from BOM
    end_item_niin: str = ""       # END ITEM NIIN from BOM (informational)
    source_filename: str = ""     # Original BOM filename, for ordering


def generate_batch_dd1750(
    entries: List[BatchBomEntry],
    template_path: str,
    output_path: str,
    packed_by: str = "",
    date: str = "",
) -> Tuple[str, int, int]:
    """
    Generate one combined DD1750 PDF covering multiple BOMs.
    
    Each BOM gets its own DD1750 (potentially spanning multiple pages),
    and they are concatenated back-to-back in the output PDF. The PACKED BY
    field is the same on every DD1750 (from the unit's packer); the End Item
    box is filled per-BOM with NOMENCLATURE / MODEL / SERIAL NUMBER.
    
    Args:
        entries: List of BatchBomEntry, one per piece of equipment
        template_path: Path to blank DD1750 template
        output_path: Path for combined output PDF
        packed_by: PACKED BY value (same across all DD1750s)
        date: Date string (same across all DD1750s)
    
    Returns:
        Tuple of (output_path, num_boms_processed, total_items)
    """
    from pypdf.generic import (
        DictionaryObject, ArrayObject, NameObject,
        TextStringObject, NumberObject, FloatObject,
    )
    
    if not entries:
        # Nothing to do — return a blank template
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0, 0
    
    writer = PdfWriter()
    total_items = 0
    
    for entry in entries:
        if not entry.items:
            continue
        
        header = HeaderInfo(
            packed_by=packed_by,
            num_boxes="1",
            end_item=format_end_item(
                entry.nomenclature,
                entry.model,
                entry.serial_number,
            ),
            date=date,
        )
        
        bom_pages = math.ceil(len(entry.items) / ROWS_PER_PAGE)
        
        for page_num in range(bom_pages):
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(entry.items))
            page_items = entry.items[start_idx:end_idx]
            
            overlay_buffer = generate_dd1750_overlay(
                page_items,
                page_num + 1,
                bom_pages,
                header,
            )
            overlay = PdfReader(overlay_buffer)
            
            template_page = PdfReader(template_path).pages[0]
            template_page.merge_page(overlay.pages[0])
            writer.add_page(template_page)
        
        total_items += len(entry.items)
    
    # Batch output is final/non-editable — we don't add form fields, since
    # everything is already pre-filled on the canvas.
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path, len(entries), total_items


def generate_dd1750_from_pdf(
    bom_path: str,
    template_path: str,
    output_path: str,
    start_page: int = 0
) -> Tuple[str, int]:
    """
    Generate DD1750 from a BOM PDF file.
    
    This is the main entry point for the conversion process.
    
    Args:
        bom_path: Path to the input BOM PDF
        template_path: Path to blank DD1750 template
        output_path: Path for output PDF
        start_page: Page to start extraction (0-based)
        
    Returns:
        Tuple of (output_path, item_count)
    """
    try:
        result = extract_items_from_pdf(bom_path, start_page)
        
        if result.errors:
            print(f"Errors during extraction: {result.errors}")
        
        if result.warnings:
            print(f"Warnings: {result.warnings}")
        
        print(f"Format detected: {result.format_detected.value}")
        print(f"Items found: {len(result.items)}")
        print(f"Pages processed: {result.pages_processed}")
        
        return generate_dd1750_from_items(result.items, template_path, output_path)
        
    except Exception as e:
        print(f"Critical error: {e}")
        import traceback
        traceback.print_exc()
        
        # Return blank template on error
        try:
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
        except:
            pass
        
        return output_path, 0


# Export for API use
__all__ = [
    'BomItem',
    'BomMetadata',
    'ExtractionResult',
    'BomFormat',
    'HeaderInfo',
    'BatchBomEntry',
    'extract_items_from_pdf',
    'generate_dd1750_from_items',
    'generate_dd1750_from_pdf',
    'generate_batch_dd1750',
    'format_packed_by',
    'format_end_item',
    'OCR_AVAILABLE',
]
