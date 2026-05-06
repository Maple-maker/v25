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
    
    # Description (after DESC:)
    match = re.search(r'DESC[:\s]*([A-Z0-9\s/\-]+)', page_text, re.IGNORECASE)
    if match:
        metadata.end_item_description = match.group(1).strip()[:50]
    
    # Serial/Equipment Number
    match = re.search(r'SER/EQUIP\s*NO[:\s]*([A-Z0-9]+)', page_text, re.IGNORECASE)
    if match:
        metadata.serial_equip_no = match.group(1)
    
    # UIC
    match = re.search(r'UIC[:\s]*([A-Z0-9]+)', page_text, re.IGNORECASE)
    if match:
        metadata.uic = match.group(1)
    
    # FE
    match = re.search(r'FE[:\s]*(\d+)', page_text, re.IGNORECASE)
    if match:
        metadata.fe = match.group(1)
    
    return metadata


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
            
            # If no items extracted via normal methods, try form field extraction
            if not all_items:
                form_items = extract_items_from_form_fields(pdf_path)
                if form_items:
                    all_items = form_items
                    result.format_detected = BomFormat.GCSS_ARMY_STANDARD
                    result.pages_processed = len(pdf.pages)
                    
                    # Form-only PDFs sometimes have pages where item data was flattened
                    # to image pixels and is no longer extractable. Warn the user so
                    # they can double-check the result against the original.
                    items_with_nsn = sum(1 for item in form_items if item.nsn)
                    items_without_nsn = len(form_items) - items_with_nsn
                    result.warnings.append(
                        f"This PDF stores its data as form fields rather than text. "
                        f"Extracted {len(form_items)} items ({items_with_nsn} with NSN). "
                        f"Some pages of this BOM type may have items rendered as images "
                        f"that cannot be extracted automatically — please verify against "
                        f"the original PDF and add any missing items in the review screen."
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
    
    # Add fillable form fields to the first page
    # Define form field positions (x, y, width, height) based on DD1750 layout
    form_fields = [
        {'name': 'packed_by', 'rect': (92, 732, 230, 746), 'tooltip': 'Packed By'},
        {'name': 'no_boxes', 'rect': (282, 732, 332, 746), 'tooltip': 'Number of Boxes'},
        {'name': 'req_no', 'rect': (405, 732, 566, 746), 'tooltip': 'Requisition Number'},
        {'name': 'order_no', 'rect': (405, 712, 566, 726), 'tooltip': 'Order Number'},
        {'name': 'end_item', 'rect': (92, 689, 370, 703), 'tooltip': 'End Item'},
        {'name': 'date', 'rect': (447, 689, 566, 703), 'tooltip': 'Date'},
        {'name': 'typed_name', 'rect': (92, 46, 290, 60), 'tooltip': 'Typed Name and Title'},
    ]
    
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
    'extract_items_from_pdf',
    'generate_dd1750_from_items',
    'generate_dd1750_from_pdf',
]
