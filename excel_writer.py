"""
Flaming News - Excel Writer
Builds the styled, print-ready Flaming News workbook from parsed report data.
"""
import io
import zipfile
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

RED_FLAG = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
HEADER_FILL = PatternFill(start_color='FF3C3C3C', end_color='FF3C3C3C', fill_type='solid')
TITLE_FONT = Font(name='Calibri', size=16, bold=True)
HEADER_FONT = Font(name='Calibri', size=12, bold=True, color='FFFFFFFF')
SECTION_FONT = Font(name='Calibri', size=13, bold=True)
DATA_FONT = Font(name='Calibri', size=11)
THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def _section_header(ws, row, text, span=7):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    for c in range(1, span + 1):
        ws.cell(row=row, column=c).fill = HEADER_FILL
    cell.font = Font(name='Calibri', size=13, bold=True, color='FFFFFFFF')
    return row + 1


def _table_header(ws, row, headers):
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    return row + 1


def _flag_if_blank(cell, value):
    cell.value = value if value else None
    cell.border = BORDER
    cell.alignment = CENTER
    cell.font = DATA_FONT
    if not value:
        cell.fill = RED_FLAG
    return cell


def _match_photo(zf, room, name):
    """Try to find an image in the zip matching this guest's room number or name."""
    if zf is None:
        return None
    names = zf.namelist()
    room_clean = (room or '').lstrip('0')
    for n in names:
        base = n.rsplit('/', 1)[-1]
        base_no_ext = re.sub(r'\.(jpg|jpeg|png)$', '', base, flags=re.I)
        if room and (base_no_ext == room or base_no_ext.lstrip('0') == room_clean):
            return n
    # fallback: match by surname (first token before comma)
    surname = (name or '').split(',')[0].strip().lower()
    for n in names:
        base = n.rsplit('/', 1)[-1].lower()
        if surname and surname in base:
            return n
    return None


def build_workbook(report_date, forecast, arrivals, in_house, departures, photos_zip_bytes=None):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Accomodation'

    zf = zipfile.ZipFile(io.BytesIO(photos_zip_bytes)) if photos_zip_bytes else None

    row = 1
    title = ws.cell(row=row, column=1, value=f'FLAMING NEWS FOR {report_date}')
    title.font = TITLE_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    row += 2

    # ---------------- 7 Day Forecast ----------------
    row = _section_header(ws, row, '7 Day Forecast', span=9)
    metrics = [
        ('Occupancy %', 'occupancy_pct', '0.0%'),
        ('Rooms Occupied', 'rooms_occupied', '0'),
        ('ADR', 'adr', '#,##0.00'),
        ('Arrivals', 'arrivals', '0'),
        ('Departures', 'departures', '0'),
    ]
    date_row = row
    ws.cell(row=date_row, column=1, value='DATE:').font = SECTION_FONT
    for i, f in enumerate(forecast[:7]):
        c = ws.cell(row=date_row, column=2 + i, value=f['date'])
        c.font = DATA_FONT
        c.alignment = CENTER
        c.border = BORDER
    row += 1
    for label, key, fmt in metrics:
        ws.cell(row=row, column=1, value=label).font = DATA_FONT
        for i, f in enumerate(forecast[:7]):
            c = ws.cell(row=row, column=2 + i, value=f.get(key))
            c.number_format = fmt
            c.alignment = CENTER
            c.border = BORDER
        row += 1
    row += 1

    # ---------------- VIP color legend ----------------
    row = _section_header(ws, row, 'VIP Code Legend', span=7)
    legend = [('T3', 'ALL GOLD', 'FF000000', 'FFFFFFFF'),
              ('T4', 'ALL Platinum', 'FFFF0000', 'FFFFFFFF'),
              ('T5', 'ALL Diamond', 'FFFFFFFF', 'FF000000'),
              ('T6', 'ALL Limitless', 'FF000000', 'FFFFFFFF'),
              ('SA', 'Birthday/Anniversary/Recovery', 'FF00B050', 'FFFFFFFF'),
              ('DV', 'Fairmont Gold', 'FFFFFF00', 'FF000000'),
              ('V1', '>6 Booking.com res. -> aggregated', None, 'FF000000')]
    for code, desc, fill, font_col in legend:
        c1 = ws.cell(row=row, column=1, value=code)
        c1.font = Font(bold=True, color=font_col if fill else 'FF000000')
        if fill:
            c1.fill = PatternFill(start_color=fill, end_color=fill, fill_type='solid')
        c1.alignment = CENTER
        c1.border = BORDER
        c2 = ws.cell(row=row, column=2, value=desc)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
        c2.font = DATA_FONT
        row += 1
    row += 1

    # ---------------- VIP Arrivals (with Photo column) ----------------
    row = _section_header(ws, row, 'VIP Arrivals', span=8)
    headers = ['Guest', 'Code', 'Room', 'ETA', 'Remarks', 'Company', 'Photo']
    header_row = row
    row = _table_header(ws, row, headers)
    photo_col = 7
    ws.column_dimensions[get_column_letter(photo_col)].width = 14

    for rec in arrivals:
        r = row
        ws.row_dimensions[r].height = 65
        ws.cell(row=r, column=1, value=rec['name']); ws.cell(row=r, column=1).font = DATA_FONT
        ws.cell(row=r, column=1).border = BORDER
        code_cell = ws.cell(row=r, column=2, value=rec.get('vip_code'))
        code_cell.border = BORDER
        code_cell.alignment = CENTER
        if rec.get('fill_color'):
            code_cell.fill = PatternFill(start_color=rec['fill_color'], end_color=rec['fill_color'], fill_type='solid')
            code_cell.font = Font(bold=True, color=rec.get('font_color') or 'FF000000')
        else:
            code_cell.font = DATA_FONT
        _flag_if_blank(ws.cell(row=r, column=3), rec.get('room'))
        eta_cell = ws.cell(row=r, column=4, value=rec.get('eta') or None)
        eta_cell.border = BORDER
        eta_cell.alignment = CENTER
        eta_cell.font = DATA_FONT  # blank ETA is normal, not flagged
        rem_cell = ws.cell(row=r, column=5, value=rec.get('remark') or None)
        rem_cell.border = BORDER
        rem_cell.font = DATA_FONT
        comp_cell = ws.cell(row=r, column=6, value=rec.get('company') or None)
        comp_cell.border = BORDER
        comp_cell.font = DATA_FONT

        # Photo
        photo_cell_ref = f'{get_column_letter(photo_col)}{r}'
        match = _match_photo(zf, rec.get('room'), rec.get('name'))
        if match:
            try:
                img_bytes = zf.read(match)
                img = XLImage(io.BytesIO(img_bytes))
                img.width, img.height = 90, 85
                ws.add_image(img, photo_cell_ref)
            except Exception:
                ws.cell(row=r, column=photo_col).fill = RED_FLAG
        else:
            ws.cell(row=r, column=photo_col).fill = RED_FLAG
        ws.cell(row=r, column=photo_col).border = BORDER
        row += 1
    row += 1

    # ---------------- VIP In-House Guests ----------------
    row = _section_header(ws, row, 'VIP In-House Guests', span=7)
    row = _table_header(ws, row, ['Guest', 'Code', 'Room', 'Departure day', 'Remarks', 'Company', ''])
    for rec in in_house:
        r = row
        ws.cell(row=r, column=1, value=rec['name']); ws.cell(row=r, column=1).font = DATA_FONT
        ws.cell(row=r, column=1).border = BORDER
        code_cell = ws.cell(row=r, column=2, value=rec.get('vip_code'))
        code_cell.border = BORDER
        code_cell.alignment = CENTER
        if rec.get('fill_color'):
            code_cell.fill = PatternFill(start_color=rec['fill_color'], end_color=rec['fill_color'], fill_type='solid')
            code_cell.font = Font(bold=True, color=rec.get('font_color') or 'FF000000')
        else:
            code_cell.font = DATA_FONT
        _flag_if_blank(ws.cell(row=r, column=3), rec.get('room'))
        _flag_if_blank(ws.cell(row=r, column=4), rec.get('dep_date'))
        rem_cell = ws.cell(row=r, column=5, value=rec.get('remark') or None)
        rem_cell.border = BORDER
        rem_cell.font = DATA_FONT
        comp_cell = ws.cell(row=r, column=6, value=rec.get('company') or None)
        comp_cell.border = BORDER
        comp_cell.font = DATA_FONT
        row += 1
    row += 1

    # ---------------- VIP Departures ----------------
    row = _section_header(ws, row, 'VIP Departures', span=7)
    row = _table_header(ws, row, ['Guest', 'Code', 'Room', 'Departure day', 'Departure Time', 'Company', ''])
    for rec in departures:
        r = row
        ws.cell(row=r, column=1, value=rec['name']); ws.cell(row=r, column=1).font = DATA_FONT
        ws.cell(row=r, column=1).border = BORDER
        code_cell = ws.cell(row=r, column=2, value=rec.get('vip_code'))
        code_cell.border = BORDER
        code_cell.alignment = CENTER
        if rec.get('fill_color'):
            code_cell.fill = PatternFill(start_color=rec['fill_color'], end_color=rec['fill_color'], fill_type='solid')
            code_cell.font = Font(bold=True, color=rec.get('font_color') or 'FF000000')
        else:
            code_cell.font = DATA_FONT
        _flag_if_blank(ws.cell(row=r, column=3), rec.get('room'))
        _flag_if_blank(ws.cell(row=r, column=4), rec.get('dep_date'))
        time_cell = ws.cell(row=r, column=5, value=rec.get('dep_time') or None)
        time_cell.border = BORDER
        time_cell.alignment = CENTER
        time_cell.font = DATA_FONT  # blank departure time is normal
        comp_cell = ws.cell(row=r, column=6, value=rec.get('company') or None)
        comp_cell.border = BORDER
        comp_cell.font = DATA_FONT
        row += 1

    # Column widths
    widths = {1: 26, 2: 8, 3: 10, 4: 16, 5: 24, 6: 26}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    # Print setup - both sheets same paper/orientation for duplex printing
    ws.print_area = f'A1:I{row+2}'
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Blank F&B Section sheet (manual entry for now)
    fb = wb.create_sheet('F&B Section')
    fb.cell(row=1, column=1, value=f'FLAMING NEWS for date {report_date}').font = TITLE_FONT
    fb.cell(row=3, column=1, value='FOOD & BEVERAGE').font = SECTION_FONT
    fb.page_setup.paperSize = 9
    fb.page_setup.orientation = 'portrait'
    fb.page_setup.fitToWidth = 1
    fb.page_setup.fitToHeight = 1
    fb.sheet_properties.pageSetUpPr.fitToPage = True

    if zf:
        zf.close()

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
