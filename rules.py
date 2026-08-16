"""
Flaming News - VIP Business Rules Engine
Derives the Remarks text and cell colour for each guest, based on VIP code,
and filters the list down to genuine VIP guests per management rules:

  - DV guests are never included in Flaming News.
  - Booking.com-sourced guests are never classified/shown as V1 VIPs.
  - Only genuine VIP codes are kept (T3/T4/T5/T6/SA/V1); anything else is dropped.
"""

GENUINE_VIP_CODES = {'T3', 'T4', 'T5', 'T6', 'SA', 'V1'}
EXCLUDED_CODES = {'DV'}

FIXED_REMARKS = {
    'T3': 'ALL Gold',
    'T4': 'ALL Platinum',
    'T5': 'ALL Diamond',
    'T6': 'ALL Limitless',
}

# Exact colours taken from the source Excel template (Accomodation!B28:B52 etc.)
# openpyxl-friendly ARGB hex codes.
COLOR_FILL = {
    'T3': 'FF000000',   # Black
    'T4': 'FFFF0000',   # Red
    'T5': None,         # No highlight (white)
    'T6': 'FF000000',   # Black (same tier family as T3)
    'SA': 'FF70AD47',   # Green (template accent6)
    'V1': 'FFFF0000',   # Red
}

FONT_COLOR = {
    'T3': 'FFFFFFFF',   # white text on black
    'T4': 'FF000000',   # black text on red (matches template)
    'T5': 'FF000000',   # black text, no fill
    'T6': 'FFFFFFFF',   # white text on black
    'SA': 'FF000000',   # black text on green (matches template)
    'V1': 'FF000000',   # black text on red (matches template)
}

SA_ABBREV = {'bd': 'Birthday', 'an': 'Anniversary', 'rt': 'Recovery'}
SA_KEYWORDS = ['Birthday', 'Anniversary', 'Recovery', 'Special Attention']


def _sa_remark(raw_specials_text, fallback=''):
    text = raw_specials_text.lower()
    found = []
    for kw in SA_KEYWORDS:
        if kw.lower() in text:
            found.append(kw)
    tokens = [t.strip().lower() for t in raw_specials_text.replace(',', ' ').split()]
    for tok, full in SA_ABBREV.items():
        if tok in tokens and full not in found:
            found.append(full)
    return ', '.join(found) if found else fallback


def is_booking_dot_com(company):
    return 'booking.com' in (company or '').lower()


def filter_vip_records(records):
    """
    Drop DV guests entirely, and drop V1 guests sourced from Booking.com
    (OTA-driven loyalty tags, not genuine VIPs the team should treat
    specially). Anything with a code outside the genuine VIP set is also
    dropped, since Flaming News should only ever show real VIPs.
    """
    kept = []
    for r in records:
        code = (r.get('vip_code') or '').strip().upper()
        if code in EXCLUDED_CODES:
            continue
        if code not in GENUINE_VIP_CODES:
            continue
        if code == 'V1' and is_booking_dot_com(r.get('company')):
            continue
        kept.append(r)
    return kept


def apply_business_rules(records):
    """
    Filters to genuine VIPs (see filter_vip_records) then mutates each
    remaining record dict in-place, adding:
      - 'remark'      -> derived Remarks text for the template
      - 'fill_color'  -> ARGB hex string or None
      - 'font_color'  -> ARGB hex string or None
    Returns the filtered, enriched list (does not mutate the input list itself).
    """
    records = filter_vip_records(records)

    for r in records:
        code = (r.get('vip_code') or '').strip().upper()
        specials_text = ' '.join(r.get('raw_specials', []))

        if code in FIXED_REMARKS:
            r['remark'] = FIXED_REMARKS[code]
        elif code == 'SA':
            r['remark'] = _sa_remark(specials_text, fallback=r.get('remark', ''))
        elif code == 'V1':
            r['remark'] = r.get('remark', '') or 'VIP Guest'
        else:
            r['remark'] = specials_text

        r['vip_code'] = code
        r['fill_color'] = COLOR_FILL.get(code)
        r['font_color'] = FONT_COLOR.get(code)

    return records


def color_for_code(code):
    code = (code or '').strip().upper()
    return COLOR_FILL.get(code), FONT_COLOR.get(code)
