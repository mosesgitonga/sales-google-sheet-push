"""
sheets/services.py
==================
All Google Sheets interaction lives here.

Sheet layout (horizontal sections):
  - Each section occupies SECTION_WIDTH columns (9 data cols).
  - Sections are separated by 1 blank column.
  - So section offsets (0-indexed col): 0, 10, 20, 30, ...
  - Row 1 (index 0): page name written in section's first column
  - Row 2 (index 1): headers
  - Row 3+ (index 2+): data rows

Headers (in order):
  DATE | USERNAME | FORM OF PAYMENT | GROSS PRICE | NET PRICE | COMMISSION 20% | NOTES | TOTAL COMMISSION | TOTAL NET
"""

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

from apps.accounts.services import get_valid_credentials

# ── Constants ─────────────────────────────────────────────────────────────────

SECTION_WIDTH = 9          # number of data columns per section
SECTION_GAP   = 1          # blank columns between sections
SECTION_STEP  = SECTION_WIDTH + SECTION_GAP   # 10

HEADERS = [
    'DATE', 'USERNAME', 'FORM OF PAYMENT',
    'GROSS PRICE', 'NET PRICE', 'COMMISSION 25%',
    'NOTES', 'TOTAL COMMISSION', 'TOTAL NET',
]

ROW_PAGE_NAME = 1   # 1-indexed: row where page name is written
ROW_HEADERS   = 2   # 1-indexed: row where headers are written
ROW_DATA_START = 3  # 1-indexed: first data row


# ── Helpers ───────────────────────────────────────────────────────────────────

def col_index_to_letter(idx: int) -> str:
    """Convert 0-based column index to A1 notation letter(s). e.g. 0→A, 25→Z, 26→AA."""
    letters = ''
    idx += 1
    while idx:
        idx, remainder = divmod(idx - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def section_start_col(section_index: int) -> int:
    """Return the 0-based column index of a section's first column."""
    return section_index * SECTION_STEP


def section_range(section_index: int, start_row: int, end_row: int) -> str:
    """Return A1 range string for a full section between two rows (1-indexed)."""
    start_col = section_start_col(section_index)
    end_col   = start_col + SECTION_WIDTH - 1
    return f'{col_index_to_letter(start_col)}{start_row}:{col_index_to_letter(end_col)}{end_row}'


def build_sheets_service(user):
    creds   = get_valid_credentials(user)
    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    return service.spreadsheets()


def build_drive_service(user):
    creds   = get_valid_credentials(user)
    service = build('drive', 'v3', credentials=creds, cache_discovery=False)
    return service


# ── Drive: list spreadsheets ──────────────────────────────────────────────────

def list_spreadsheets(user) -> list[dict]:
    """Return a list of the user's Google Sheets files."""
    drive = build_drive_service(user)
    result = drive.files().list(
        q="mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
        fields='files(id,name)',
        orderBy='modifiedTime desc',
        pageSize=50,
    ).execute()
    return result.get('files', [])


# ── Sheets: list tabs ─────────────────────────────────────────────────────────

def list_tabs(user, spreadsheet_id: str) -> list[dict]:
    """Return all sheet tabs in a spreadsheet as [{id, title}]."""
    sheets = build_sheets_service(user)
    meta   = sheets.get(spreadsheetId=spreadsheet_id, fields='sheets(properties)').execute()
    return [
        {'id': s['properties']['sheetId'], 'title': s['properties']['title']}
        for s in meta.get('sheets', [])
    ]


# ── Sheet state: read all sections ────────────────────────────────────────────

def read_sheet_state(user, spreadsheet_id: str, sheet_title: str) -> list[dict]:
    """
    Read the sheet and return a list of section descriptors:
    [
      {
        'section_index': 0,
        'page_name': 'heather free',   # or None if empty
        'has_headers': True,
        'data_rows': 3,                # number of filled data rows
        'last_data_row': 5,            # 1-indexed sheet row of last data row
      },
      ...
    ]
    We read up to MAX_SECTIONS sections.
    """
    MAX_SECTIONS = 20
    sheets = build_sheets_service(user)

    # Read a large range to cover all sections
    end_col   = col_index_to_letter(MAX_SECTIONS * SECTION_STEP)
    range_str = f"'{sheet_title}'!A1:{end_col}1000"

    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=range_str,
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()

    all_values = result.get('values', [])

    # Pad rows to consistent length
    max_col = MAX_SECTIONS * SECTION_STEP + SECTION_WIDTH
    padded  = []
    for row in all_values:
        padded.append(row + [''] * (max_col - len(row)))

    sections = []
    for sec_idx in range(MAX_SECTIONS):
        start_col = section_start_col(sec_idx)

        # Row 1 (index 0): page name
        page_name = ''
        if padded:
            page_name = str(padded[0][start_col]).strip() if len(padded[0]) > start_col else ''

        # Row 2 (index 1): headers
        has_headers = False
        if len(padded) > 1:
            header_val = str(padded[1][start_col]).strip().upper() if len(padded[1]) > start_col else ''
            has_headers = header_val == 'DATE'

        # Count filled data rows from row 3 (index 2) onward
        data_rows     = 0
        last_data_row = ROW_HEADERS  # default: last row = header row

        for row_idx in range(2, len(padded)):  # index 2 = row 3
            date_val = str(padded[row_idx][start_col]).strip() if len(padded[row_idx]) > start_col else ''
            user_val = str(padded[row_idx][start_col + 1]).strip() if len(padded[row_idx]) > start_col + 1 else ''
            if date_val or user_val:
                data_rows    += 1
                last_data_row = row_idx + 1  # convert to 1-indexed
            else:
                # Stop at first empty row in this section
                break

        sections.append({
            'section_index': sec_idx,
            'page_name':     page_name,
            'has_headers':   has_headers,
            'data_rows':     data_rows,
            'last_data_row': last_data_row,
        })

        # Stop scanning once we hit fully empty sections (2 consecutive empty)
        if not page_name and not has_headers and sec_idx > 0:
            prev = sections[sec_idx - 1]
            if not prev['page_name'] and not prev['has_headers']:
                break

    return sections


# ── Find or allocate a section for a page ─────────────────────────────────────

def find_section_for_page(sections: list[dict], page_name: str) -> dict | None:
    """Return the section dict that owns this page name, or None."""
    page_lower = page_name.strip().lower()
    for sec in sections:
        if sec['page_name'].strip().lower() == page_lower:
            return sec
    return None


def find_empty_section(sections: list[dict]) -> dict | None:
    """Return the first section with no page name and no headers."""
    for sec in sections:
        if not sec['page_name'] and not sec['has_headers']:
            return sec
    return None


def next_overflow_target(sections: list[dict], page_name: str) -> dict:
    """
    All sections are filled with OTHER pages.
    Find the section for this page if it exists (overflow allowed),
    otherwise pick section 0 as the overflow target.
    The caller handles the 3-row gap skip.
    """
    existing = find_section_for_page(sections, page_name)
    if existing:
        return existing
    # Return first section for overflow (could be enhanced to rotate)
    return sections[0]


# ── Write helpers ─────────────────────────────────────────────────────────────

def ensure_section_headers(sheets_api, spreadsheet_id: str, sheet_id: int,
                            sheet_title: str, section_index: int, page_name: str):
    """Write page name on row 1 and headers on row 2 for this section."""
    start_col = section_start_col(section_index)

    # Page name in row 1, col A of section
    page_range = f"'{sheet_title}'!{col_index_to_letter(start_col)}{ROW_PAGE_NAME}"
    sheets_api.values().update(
        spreadsheetId=spreadsheet_id,
        range=page_range,
        valueInputOption='RAW',
        body={'values': [[page_name]]},
    ).execute()

    # Headers in row 2
    header_range = (
        f"'{sheet_title}'!"
        f"{col_index_to_letter(start_col)}{ROW_HEADERS}:"
        f"{col_index_to_letter(start_col + SECTION_WIDTH - 1)}{ROW_HEADERS}"
    )
    sheets_api.values().update(
        spreadsheetId=spreadsheet_id,
        range=header_range,
        valueInputOption='RAW',
        body={'values': [HEADERS]},
    ).execute()


def insert_row_in_section(sheets_api, spreadsheet_id: str, sheet_id: int,
                           insert_at_row: int):
    """Insert a blank row at insert_at_row (1-indexed), shifting rows down."""
    sheets_api.batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            'requests': [{
                'insertDimension': {
                    'range': {
                        'sheetId':    sheet_id,
                        'dimension':  'ROWS',
                        'startIndex': insert_at_row - 1,   # 0-indexed
                        'endIndex':   insert_at_row,
                    },
                    'inheritFromBefore': True,
                }
            }]
        }
    ).execute()


def write_data_row(sheets_api, spreadsheet_id: str, sheet_title: str,
                   section_index: int, row_number: int, row_data: list):
    """Write a single data row at row_number (1-indexed) in this section."""
    start_col  = section_start_col(section_index)
    range_str  = (
        f"'{sheet_title}'!"
        f"{col_index_to_letter(start_col)}{row_number}:"
        f"{col_index_to_letter(start_col + SECTION_WIDTH - 1)}{row_number}"
    )
    sheets_api.values().update(
        spreadsheetId=spreadsheet_id,
        range=range_str,
        valueInputOption='USER_ENTERED',
        body={'values': [row_data]},
    ).execute()


def update_section_totals(sheets_api, spreadsheet_id: str, sheet_title: str,
                           section_index: int, sections_state: list[dict]):
    """
    Recalculate TOTAL COMMISSION and TOTAL NET for this section
    by summing columns F (commission) and E (net) from row 3 downward,
    and write them in the TOTAL COMMISSION and TOTAL NET cells.
    We place running totals in col H and I of the section's row 2.
    """
    # Re-read the section to get current totals
    sec       = next((s for s in sections_state if s['section_index'] == section_index), None)
    if not sec or sec['last_data_row'] < ROW_DATA_START:
        return

    start_col  = section_start_col(section_index)
    net_col    = col_index_to_letter(start_col + 4)   # E offset = NET PRICE
    comm_col   = col_index_to_letter(start_col + 5)   # F offset = COMMISSION
    tc_col     = col_index_to_letter(start_col + 7)   # H offset = TOTAL COMMISSION
    tn_col     = col_index_to_letter(start_col + 8)   # I offset = TOTAL NET
    last_row   = sec['last_data_row']

    total_comm_formula = f'=SUM({comm_col}{ROW_DATA_START}:{comm_col}{last_row})'
    total_net_formula  = f'=SUM({net_col}{ROW_DATA_START}:{net_col}{last_row})'

    # Write totals into row 2 cols H, I of this section
    tc_range = f"'{sheet_title}'!{tc_col}{ROW_HEADERS}"
    tn_range = f"'{sheet_title}'!{tn_col}{ROW_HEADERS}"

    sheets_api.values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {'range': tc_range, 'values': [[total_comm_formula]]},
                {'range': tn_range, 'values': [[total_net_formula]]},
            ]
        }
    ).execute()


# ── Main entry: push rows ─────────────────────────────────────────────────────

def push_rows(user, spreadsheet_id: str, sheet_id: int, sheet_title: str,
              page_name: str, rows: list[dict]) -> dict:
    """
    Push a list of row dicts to the correct section for page_name.

    row dict keys: date, username, payment, gross, net, commission, notes

    Returns: { 'inserted': int, 'section_index': int, 'tab': str }
    """
    sheets_api = build_sheets_service(user)
    sections   = read_sheet_state(user, spreadsheet_id, sheet_title)

    existing_sec = find_section_for_page(sections, page_name)

    if existing_sec:
        # Section found — append rows after last data row, inserting sheet rows
        sec           = existing_sec
        section_index = sec['section_index']
        insert_row    = sec['last_data_row'] + 1

        for row_data in rows:
            # Insert a new row at insert_row to push existing content down
            insert_row_in_section(sheets_api, spreadsheet_id, sheet_id, insert_row)
            write_data_row(
                sheets_api, spreadsheet_id, sheet_title,
                section_index, insert_row,
                _format_row(row_data, page_name)
            )
            insert_row += 1

    else:
        # No existing section — find an empty one or overflow
        empty_sec = find_empty_section(sections)

        if empty_sec:
            section_index = empty_sec['section_index']
            insert_row    = ROW_DATA_START
        else:
            # All sections have OTHER pages — overflow back to a section
            # We use the section with the most content (section 0) and skip 3 rows
            overflow_sec  = sections[0]
            section_index = overflow_sec['section_index']
            insert_row    = overflow_sec['last_data_row'] + 4  # skip 3 blank rows

        # Write page name + headers
        ensure_section_headers(
            sheets_api, spreadsheet_id, sheet_id,
            sheet_title, section_index, page_name
        )

        for row_data in rows:
            insert_row_in_section(sheets_api, spreadsheet_id, sheet_id, insert_row)
            write_data_row(
                sheets_api, spreadsheet_id, sheet_title,
                section_index, insert_row,
                _format_row(row_data, page_name)
            )
            insert_row += 1

    # Refresh state and update totals
    updated_sections = read_sheet_state(user, spreadsheet_id, sheet_title)
    update_section_totals(
        sheets_api, spreadsheet_id, sheet_title,
        section_index, updated_sections 
    )

    return {
        'inserted':      len(rows),
        'section_index': section_index,
        'tab':           sheet_title,
    }


def _format_row(row_data: dict, page_name: str) -> list:
    """Convert a row dict into the ordered list matching HEADERS."""
    return [
        row_data.get('date', ''),
        row_data.get('username', ''),
        row_data.get('payment', 'Message purchase'),
        row_data.get('gross', ''),
        row_data.get('net', ''),
        row_data.get('commission', ''),
        row_data.get('notes', page_name),
        '',   # TOTAL COMMISSION — filled by formula
        '',   # TOTAL NET — filled by formula
    ]
