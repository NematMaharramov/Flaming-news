"""
Flaming News - PDF Parsers
Parses the 4 Fairmont PMS report PDFs into structured Python dicts.
"""
import re
import pdfplumber

VIP_CODES = {'T3', 'T4', 'T5', 'T6', 'DV', 'SA', 'V1'}


def _extract_text(file_stream):
    """Return list of raw text lines across all pages of a PDF."""
    lines = []
    with pdfplumber.open(file_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend(text.split('\n'))
    return lines


# ---------------------------------------------------------------------------
# 1. History and Forecast -> 7 Day Forecast section
# ---------------------------------------------------------------------------
def parse_forecast(file_stream):
    lines = _extract_text(file_stream)
    records = []
    pattern = re.compile(
        r'^(\d{2}\.\d{2}\.\d{2})\s+\w{3}\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+'
        r'([\d.]+)%\s+([\d,]+\.\d{2})\s+([\d.]+)\s+(\d+)'
    )
    for line in lines:
        m = pattern.match(line)
        if m:
            date, tot_occ, arr, comp, house, occ_pct, rev, adr, dep = m.groups()
            records.append({
                'date': date,
                'occupancy_pct': float(occ_pct) / 100,
                'rooms_occupied': int(tot_occ) - int(house) - int(comp),
                'adr': float(adr),
                'arrivals': int(arr),
                'departures': int(dep),
            })
    return records


# ---------------------------------------------------------------------------
# 2. Guests INH - By Room -> VIP In-House Guests section
# ---------------------------------------------------------------------------
def parse_in_house(file_stream):
    lines = _extract_text(file_stream)
    records = []
    current = None
    main_re = re.compile(
        r'^(\d{4}(?:\s*,\s*\d{4})*)\s+([A-Za-z][^0-9]+?)\s+'
        r'(?:(?:T-|C-|S-)\s*(.+?)\s+)?'
        r'(\d{2}\.\d{2}\.\d{2})(\d{2}\.\d{2}\.\d{2})([A-Z0-9]+)\s+(\d+)\s+(\d+)\s*([A-Z]{2})\s+(\S+)?$'
    )
    skip_words = ('VIP', 'SA', 'V1', 'DV', 'T3', 'T4', 'T5', 'T6', 'Accompanying', 'Res.')
    for i, raw in enumerate(lines):
        line = raw.rstrip()
        m = main_re.match(line)
        if m:
            room, name, company, arr, dep, roomtype, adl, chl, paymth, rate = m.groups()
            company = (company or '').strip()
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if (re.match(r'^[A-Za-z .()]+$', nxt) and
                        not any(nxt.startswith(w) or nxt == w for w in skip_words)):
                    company = (company + ' ' + nxt).strip()
            current = {
                'room': room.split(',')[0].strip(),
                'name': name.strip(),
                'company': company,
                'arr_date': arr,
                'dep_date': dep,
                'vip_code': None,
                'raw_specials': [],
            }
            records.append(current)
            continue
        if current and re.match(r'^(T3|T4|T5|T6|DV|SA|V1)$', line.strip()):
            current['vip_code'] = line.strip()
        if current and line.strip().startswith('Specials:'):
            current['raw_specials'].append(line.strip().replace('Specials:', '').strip())
    return [r for r in records if r['vip_code']]


# ---------------------------------------------------------------------------
# 3. Arrivals: Detailed -> VIP Arrivals section
# ---------------------------------------------------------------------------
def parse_arrivals(file_stream):
    lines = _extract_text(file_stream)
    records = []
    current = None
    last_room = ''
    main_re = re.compile(
        r'^(\d{4})?\s*([A-Za-z][^0-9]+?)\s+'
        r'(?:(?:T-|C-|S-)\s*(.+?)\s+)?'
        r'(\d{2}\.\d{2}\.\d{2})\s+(\d{2}\.\d{2}\.\d{2})\s+([A-Z0-9]+)\s+'
        r'(\d+)\s+(\d+)\s+(\d+)\s+\S+\s+\S+\s+\S+\s+(\S+)\s+[A-Z]{3}\s+[\d.,]+\S*\s+[\d.]+$'
    )
    vip_line_re = re.compile(r'^(\d+)\s+(T3|T4|T5|T6|DV|SA|V1)(?:\s+(\d{1,2}:\d{2}))?')
    skip_words = ('VIP', 'SA', 'V1', 'DV', 'T3', 'T4', 'T5', 'T6', 'Accompanying',
                  'Reservation', 'RESERVATIONS', 'CASHIER', 'Specials', 'Membership',
                  'Traces', 'Routing')
    for i, raw in enumerate(lines):
        line = raw.rstrip()
        m = main_re.match(line)
        if m:
            room, name, company, arr, dep, roomtype, adl, chl, rms, rate = m.groups()
            company = (company or '').strip()
            same_guest = bool(records) and records[-1]['name'].strip() == name.strip()
            if room:
                last_room = room
                room_val = room
            elif same_guest:
                room_val = last_room
            else:
                room_val = ''  # different guest, room unreadable in source PDF -> leave blank/flag
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if (re.match(r'^[A-Za-z .()]+$', nxt) and
                        not any(nxt.startswith(w) for w in skip_words) and
                        not vip_line_re.match(nxt)):
                    company = (company + ' ' + nxt).strip()
            current = {
                'room': room_val,
                'name': name.strip(),
                'company': company,
                'arr_date': arr,
                'eta': '',
                'vip_code': None,
                'raw_specials': [],
            }
            records.append(current)
            continue
        vm = vip_line_re.match(line.strip())
        if current and vm:
            current['vip_code'] = vm.group(2)
            current['eta'] = vm.group(3) or ''
        if current and line.strip().startswith('Specials:'):
            current['raw_specials'].append(line.strip().replace('Specials:', '').strip())
    return [r for r in records if r['vip_code']]


# ---------------------------------------------------------------------------
# 4. Departures -> VIP Departures section
# ---------------------------------------------------------------------------
def parse_departures(file_stream):
    lines = _extract_text(file_stream)
    records = []
    current = None
    main_re = re.compile(
        r'^(\d{4})\s+([A-Za-z][^0-9]+?)\s+'
        r'(?:(?:T-|C-|S-)\s*(.+?)\s+)?'
        r'(T3|T4|T5|T6|DV|SA|V1)\s+'
        r'(\d{2}\.\d{2}\.\d{2})\s+(\d{2}\.\d{2}\.\d{2})\s+'
        r'(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([A-Z0-9]+)\s+(\S+)\s+([A-Z]{2,4})'
        r'(?:\s+(\d{1,2}:\d{2}))?\s+([A-Z]{2})$'
    )
    skip_words = ('G-', 'Res.Comments', 'Reservation', 'General', 'Specials',
                  'Membership', 'Profile')
    for i, raw in enumerate(lines):
        line = raw.rstrip()
        m = main_re.match(line)
        if m:
            (room, name, company, vip_code, arr, dep, adl, chl, rms, nts,
             roomtype, rate, resstatus, dep_time, paymth) = m.groups()
            company = (company or '').strip()
            if i + 1 < len(lines):
                nxt = lines[i + 1].strip()
                if (re.match(r'^\(?[A-Za-z .()]+\)?$', nxt) and
                        not any(nxt.startswith(w) for w in skip_words)):
                    company = (company + ' ' + nxt).strip()
            current = {
                'room': room,
                'name': name.strip(),
                'company': company,
                'vip_code': vip_code,
                'dep_date': dep,
                'dep_time': dep_time or '',
                'raw_specials': [],
            }
            records.append(current)
            continue
        if current and line.strip().startswith('Specials:'):
            current['raw_specials'].append(
                line.strip().replace('Specials:', '').strip())
    return records


# ---------------------------------------------------------------------------
# Clean company prefixes (T-/C-/S-) — used post-parse
# ---------------------------------------------------------------------------
def clean_company(name):
    if not name:
        return ''
    return re.sub(r'^(T-|C-|S-)\s*', '', name).strip()
