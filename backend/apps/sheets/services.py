"""
sheets/services.py
==================
Google Sheets interaction.

Layout (horizontal sections):
  - Each section = 9 data columns + 1 gap column = 10 columns per section
  - Section start columns (0-indexed): 0, 10, 20, 30, ...
  - Row 1 : page name
  - Row 2 : headers
  - Row 3 : TOTAL COMMISSION and TOTAL NET formulas
  - Row 4+: data rows

Column order per section:
  0: DATE  1: USERNAME  2: FORM OF PAYMENT  3: GROSS PRICE  4: NET PRICE
  5: COMMISSION 25%  6: NOTES  7: TOTAL COMMISSION  8: TOTAL NET

Row placement rules:
  1. Find section whose row-1 matches page name.
     - Scan data rows (row 4+). If a row has NOTES = page name but
       DATE + USERNAME empty -> REPLACE it.
     - No empty rows left -> write to next empty row after last filled row.
  2. No matching section -> find section with empty data rows -> claim it.
  3. No space -> skip 3 rows after last filled content, write new block.

Key: NO insertDimension calls. We only write to specific column ranges
     so other sections are never affected.
"""

from googleapiclient.discovery import build
from apps.accounts.services import get_valid_credentials

# ── Constants ─────────────────────────────────────────────────────────────────

SECTION_WIDTH  = 9
SECTION_GAP    = 1
SECTION_STEP   = SECTION_WIDTH + SECTION_GAP   # 10

COL_DATE       = 0
COL_USERNAME   = 1
COL_PAYMENT    = 2
COL_GROSS      = 3
COL_NET        = 4
COL_COMMISSION = 5
COL_NOTES      = 6
COL_TOTAL_COMM = 7
COL_TOTAL_NET  = 8

HEADERS = [
    'DATE', 'USERNAME', 'FORM OF PAYMENT',
    'GROSS PRICE', 'NET PRICE', 'COMMISSION 25%',
    'NOTES', 'TOTAL COMMISSION', 'TOTAL NET',
]

ROW_PAGE_NAME  = 1   # 1-indexed
ROW_HEADERS    = 2
ROW_TOTALS     = 3   # totals live here
ROW_DATA_START = 4   # data starts here


# ── Column helpers ────────────────────────────────────────────────────────────

def col_letter(idx: int) -> str:
    """0-based column index to A1 letter. 0->A, 25->Z, 26->AA."""
    letters = ''
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def sec_start(sec: int) -> int:
    return sec * SECTION_STEP


# ── API builders ──────────────────────────────────────────────────────────────

def get_sheets_api(user):
    creds = get_valid_credentials(user)
    return build('sheets', 'v4', credentials=creds, cache_discovery=False).spreadsheets()


def get_drive_api(user):
    creds = get_valid_credentials(user)
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


# ── Grid limits ───────────────────────────────────────────────────────────────

def get_grid_limits(user, spreadsheet_id: str, sheet_id: int) -> dict:
    api  = get_sheets_api(user)
    meta = api.get(
        spreadsheetId=spreadsheet_id,
        fields='sheets(properties(sheetId,gridProperties))'
    ).execute()
    for s in meta.get('sheets', []):
        if s['properties']['sheetId'] == sheet_id:
            gp = s['properties']['gridProperties']
            return {
                'rows': gp.get('rowCount', 1000),
                'cols': gp.get('columnCount', 26),
            }
    return {'rows': 1000, 'cols': 26}


# ── Drive ─────────────────────────────────────────────────────────────────────

def list_spreadsheets(user) -> list[dict]:
    drive  = get_drive_api(user)
    result = drive.files().list(
        q="mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
        fields='files(id,name)',
        orderBy='modifiedTime desc',
        pageSize=50,
    ).execute()
    return result.get('files', [])


def list_tabs(user, spreadsheet_id: str) -> list[dict]:
    api  = get_sheets_api(user)
    meta = api.get(spreadsheetId=spreadsheet_id, fields='sheets(properties)').execute()
    return [
        {'id': s['properties']['sheetId'], 'title': s['properties']['title']}
        for s in meta.get('sheets', [])
    ]


# ── Safe 2-D value reader ─────────────────────────────────────────────────────

def read_all_values(api, spreadsheet_id: str, sheet_title: str,
                    max_col_idx: int, max_row: int = 2000) -> list[list]:
    end_col   = col_letter(max_col_idx)
    range_str = f"'{sheet_title}'!A1:{end_col}{max_row}"
    result    = api.values().get(
        spreadsheetId=spreadsheet_id,
        range=range_str,
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    raw   = result.get('values', [])
    width = max_col_idx + 1
    return [row + [''] * (width - len(row)) for row in raw]


def cell(values: list[list], row_idx: int, col_idx: int) -> str:
    """Safe 0-indexed access, always returns str."""
    if row_idx >= len(values):
        return ''
    row = values[row_idx]
    return str(row[col_idx]).strip() if col_idx < len(row) else ''


# ── Read sheet state ──────────────────────────────────────────────────────────

def read_sheet_state(user, spreadsheet_id: str, sheet_title: str,
                     sheet_id: int = None) -> list[dict]:
    """
    Return list of section descriptors.

    Each descriptor:
    {
      'section_index':   int,
      'page_name':       str,
      'has_headers':     bool,
      'last_filled_row': int,   # 1-indexed; ROW_TOTALS if no data yet
      'data_rows': [
        {
          'sheet_row': int,   # 1-indexed, starts at ROW_DATA_START (4)
          'date':      str,
          'username':  str,
          'notes':     str,
          'is_empty':  bool,  # True when date AND username both blank
        },
        ...
      ],
    }
    """
    api = get_sheets_api(user)

    if sheet_id is not None:
        limits   = get_grid_limits(user, spreadsheet_id, sheet_id)
        max_secs = max(1, limits['cols'] // SECTION_STEP)
    else:
        max_secs = 10

    max_col_idx = max_secs * SECTION_STEP + SECTION_WIDTH - 1
    values      = read_all_values(api, spreadsheet_id, sheet_title, max_col_idx)

    sections          = []
    consecutive_empty = 0

    for sec_idx in range(max_secs):
        sc = sec_start(sec_idx)

        page_name   = cell(values, 0, sc)                   # row 1
        has_headers = cell(values, 1, sc).upper() == 'DATE' # row 2

        # Data rows start at row index 3 (sheet row 4, skipping totals row)
        data_rows        = []
        last_filled_row  = ROW_TOTALS   # default: no data yet
        consecutive_none = 0

        for row_idx in range(3, len(values)):   # row_idx 3 = sheet row 4
            date_val  = cell(values, row_idx, sc + COL_DATE)
            user_val  = cell(values, row_idx, sc + COL_USERNAME)
            notes_val = cell(values, row_idx, sc + COL_NOTES)
            is_empty  = not date_val and not user_val

            data_rows.append({
                'sheet_row': row_idx + 1,   # 1-indexed
                'date':      date_val,
                'username':  user_val,
                'notes':     notes_val,
                'is_empty':  is_empty,
            })

            if not is_empty:
                last_filled_row  = row_idx + 1
                consecutive_none = 0
            else:
                consecutive_none += 1
                if consecutive_none >= 3:
                    break

        sections.append({
            'section_index':   sec_idx,
            'page_name':       page_name,
            'has_headers':     has_headers,
            'data_rows':       data_rows,
            'last_filled_row': last_filled_row,
        })

        sec_is_empty      = not page_name and not has_headers and last_filled_row == ROW_TOTALS
        consecutive_empty = consecutive_empty + 1 if sec_is_empty else 0
        if consecutive_empty >= 2:
            break

    return sections


# ── Write helpers ─────────────────────────────────────────────────────────────

def write_page_and_headers(api, spreadsheet_id: str, sheet_title: str,
                            sec: int, page_name: str):
    """Write page name on row 1 and headers on row 2."""
    sc = sec_start(sec)
    ec = sc + SECTION_WIDTH - 1
    api.values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            'valueInputOption': 'RAW',
            'data': [
                {
                    'range':  f"'{sheet_title}'!{col_letter(sc)}{ROW_PAGE_NAME}",
                    'values': [[page_name]],
                },
                {
                    'range':  f"'{sheet_title}'!{col_letter(sc)}{ROW_HEADERS}:{col_letter(ec)}{ROW_HEADERS}",
                    'values': [HEADERS],
                },
            ]
        }
    ).execute()


def write_page_name_only(api, spreadsheet_id: str, sheet_title: str,
                          sec: int, page_name: str):
    sc = sec_start(sec)
    api.values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_title}'!{col_letter(sc)}{ROW_PAGE_NAME}",
        valueInputOption='RAW',
        body={'values': [[page_name]]},
    ).execute()


def write_data_row(api, spreadsheet_id: str, sheet_title: str,
                   sec: int, sheet_row: int, values: list):
    """
    Write values into a specific row of a section's column range only.
    This never touches other sections because we target exact columns.
    """
    sc = sec_start(sec)
    ec = sc + SECTION_WIDTH - 1
    api.values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{sheet_title}'!{col_letter(sc)}{sheet_row}:{col_letter(ec)}{sheet_row}",
        valueInputOption='USER_ENTERED',
        body={'values': [values]},
    ).execute()


def update_totals(api, spreadsheet_id: str, sheet_title: str,
                  sec: int, last_data_row: int):
    """
    Write SUM formulas into ROW_TOTALS (row 3) for this section only.
    Targets TOTAL COMMISSION (col H of section) and TOTAL NET (col I of section).
    """
    sc       = sec_start(sec)
    net_col  = col_letter(sc + COL_NET)
    comm_col = col_letter(sc + COL_COMMISSION)
    tc_col   = col_letter(sc + COL_TOTAL_COMM)
    tn_col   = col_letter(sc + COL_TOTAL_NET)

    api.values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {
                    'range':  f"'{sheet_title}'!{tc_col}{ROW_TOTALS}",
                    'values': [[f'=SUM({comm_col}{ROW_DATA_START}:{comm_col}{last_data_row})']],
                },
                {
                    'range':  f"'{sheet_title}'!{tn_col}{ROW_TOTALS}",
                    'values': [[f'=SUM({net_col}{ROW_DATA_START}:{net_col}{last_data_row})']],
                },
            ]
        }
    ).execute()


def format_row(row: dict, page_name: str) -> list:
    return [
        row.get('date', ''),
        row.get('username', ''),
        row.get('payment', 'Message purchase'),
        row.get('gross', ''),
        row.get('net', ''),
        row.get('commission', ''),
        row.get('notes', page_name),
        '',   # TOTAL COMMISSION col — left blank, formula is in row 3
        '',   # TOTAL NET col — left blank, formula is in row 3
    ]


# ── Placement within a known section ─────────────────────────────────────────

def place_rows_in_section(api, spreadsheet_id: str, sheet_title: str,
                           sec_desc: dict, page_name: str,
                           incoming: list[dict]) -> int:
    """
    Write incoming rows into sec_desc. No row inserts — we only write
    to specific column ranges so other sections are never shifted.

    For each incoming row:
      1. Find empty slot (date+username blank) whose NOTES matches
         page_name or is empty -> overwrite it.
      2. No empty slots left -> write to next row after last filled row.

    Returns 1-indexed sheet row of the last written row.
    """
    sec          = sec_desc['section_index']
    last_written = sec_desc['last_filled_row']
    page_lower   = page_name.strip().lower()

    empty_slots = [
        r for r in sec_desc['data_rows']
        if r['is_empty'] and (
            r['notes'].strip().lower() == page_lower
            or r['notes'] == ''
        )
    ]

    for row_data in incoming:
        values = format_row(row_data, page_name)

        if empty_slots:
            slot = empty_slots.pop(0)
            write_data_row(api, spreadsheet_id, sheet_title,
                           sec, slot['sheet_row'], values)
            last_written = max(last_written, slot['sheet_row'])
        else:
            # Write to the next row — no insert, just target the next empty row
            next_row = last_written + 1
            write_data_row(api, spreadsheet_id, sheet_title,
                           sec, next_row, values)
            last_written = next_row

    return last_written


# ── Main entry ────────────────────────────────────────────────────────────────

def push_rows(user, spreadsheet_id: str, sheet_id: int,
              sheet_title: str, page_name: str, rows: list[dict]) -> dict:
    """
    Push rows into the correct section. No full-row inserts ever.

    1. Section whose row-1 matches page_name -> replace empty rows or append.
    2. No match -> section with empty data rows -> claim and fill.
    3. No space -> write after last content with 3-row gap.
    """
    api        = get_sheets_api(user)
    sections   = read_sheet_state(user, spreadsheet_id, sheet_title, sheet_id)
    page_lower = page_name.strip().lower()

    # ── 1. Matching section ────────────────────────────────────────────────────
    match = next(
        (s for s in sections if s['page_name'].strip().lower() == page_lower),
        None
    )
    if match:
        last_row = place_rows_in_section(
            api, spreadsheet_id, sheet_title, match, page_name, rows
        )
        update_totals(api, spreadsheet_id, sheet_title,
                      match['section_index'], last_row)
        return {
            'inserted':      len(rows),
            'section_index': match['section_index'],
            'tab':           sheet_title,
        }

    # ── 2. Empty section available ─────────────────────────────────────────────
    empty_sec = next(
        (s for s in sections
         if not s['page_name']
         and any(r['is_empty'] for r in s['data_rows'])),
        None
    )
    if empty_sec:
        sec_idx = empty_sec['section_index']
        if not empty_sec['has_headers']:
            write_page_and_headers(api, spreadsheet_id,
                                   sheet_title, sec_idx, page_name)
        else:
            write_page_name_only(api, spreadsheet_id,
                                 sheet_title, sec_idx, page_name)

        last_row = place_rows_in_section(
            api, spreadsheet_id, sheet_title, empty_sec, page_name, rows
        )
        update_totals(api, spreadsheet_id, sheet_title, sec_idx, last_row)
        return {
            'inserted':      len(rows),
            'section_index': sec_idx,
            'tab':           sheet_title,
        }

    # ── 3. Overflow ────────────────────────────────────────────────────────────
    # All sections taken — write after section 0's last content with 3-row gap
    base_sec  = sections[0]
    sec_idx   = base_sec['section_index']
    start_row = base_sec['last_filled_row'] + 4   # 3 gap rows + 1

    write_page_and_headers(api, spreadsheet_id,
                           sheet_title, sec_idx, page_name)

    # Data starts 2 rows after start_row (row 1 = page name, row 2 = headers)
    data_row = start_row + 2
    for row_data in rows:
        write_data_row(api, spreadsheet_id, sheet_title, sec_idx,
                       data_row, format_row(row_data, page_name))
        data_row += 1

    last_row = data_row - 1
    update_totals(api, spreadsheet_id, sheet_title, sec_idx, last_row)
    return {
        'inserted':      len(rows),
        'section_index': sec_idx,
        'tab':           sheet_title,
    }