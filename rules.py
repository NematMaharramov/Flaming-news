"""
Flaming News - VIP Business Rules Engine
Derives the Remarks text and cell colour for each guest, based on VIP code.
"""

FIXED_REMARKS = {
    'T3': 'ALL GOLD',
    'T4': 'ALL Platinum',
    'T5': 'ALL Diamond',
    'T6': 'ALL Limitless',
}

# openpyxl-friendly ARGB hex codes
COLOR_FILL = {
    'T3': 'FF000000',   # Black
    'T4': 'FFFF0000',   # Red
    'T5': 'FFFFFFFF',   # White
    'T6': 'FF000000',   # Black
    'SA': 'FF00B050',   # Green
    'DV': 'FFFFFF00',   # Yellow
    'V1': None,
}

# Font colour chosen for contrast against each fill
FONT_COLOR = {
    'T3': 'FFFFFFFF',   # white text on black
    'T4': 'FFFFFFFF',   # white text on red
    'T5': 'FF000000',   # black text on white
    'T6': 'FFFFFFFF',   # white text on black
    'SA': 'FFFFFFFF',   # white text on green
    'DV': 'FF000000',   # black text on yellow
    'V1': 'FF000000',
}

# Abbreviation codes occasionally used instead of full words in "Specials:"
SA_ABBREV = {'bd': 'Birthday', 'an': 'Anniversary', 'rt': 'Recovery'}
SA_KEYWORDS = ['Birthday', 'Anniversary', 'Recovery']


def _sa_remark(raw_specials_text):
    text = raw_specials_text.lower()
    found = []
    for kw in SA_KEYWORDS:
        if kw.lower() in text:
            found.append(kw)
    # also check abbreviation-style comma lists e.g. "BD,SO,TD"
    tokens = [t.strip().lower() for t in raw_specials_text.replace(',', ' ').split()]
    for tok, full in SA_ABBREV.items():
        if tok in tokens and full not in found:
            found.append(full)
    return ', '.join(found)


def _dv_remark(raw_specials_text):
    text = raw_specials_text.lower()
    tokens = [t.strip().lower() for t in raw_specials_text.replace(',', ' ').split()]
    if 'fg' in tokens or 'fairmont gold' in text:
        return 'Fairmont Gold'
    return ''


def apply_business_rules(records):
    """
    Mutates each record dict in-place, adding:
      - 'remark'      -> derived Remarks text for the template
      - 'fill_color'  -> ARGB hex string or None
      - 'font_color'  -> ARGB hex string or None
    V1 booking.com aggregation is computed across the given list (per-section).
    """
    bcom_count = sum(
        1 for r in records
        if r.get('vip_code') == 'V1' and 'booking.com' in (r.get('company') or '').lower()
    )

    for r in records:
        code = r.get('vip_code')
        specials_text = ' '.join(r.get('raw_specials', []))

        if code in FIXED_REMARKS:
            r['remark'] = FIXED_REMARKS[code]
        elif code == 'SA':
            r['remark'] = _sa_remark(specials_text)
        elif code == 'DV':
            r['remark'] = _dv_remark(specials_text)
        elif code == 'V1':
            if 'booking.com' in (r.get('company') or '').lower() and bcom_count > 6:
                r['remark'] = f'{bcom_count} Rooms Booking.com'
            else:
                r['remark'] = ''
        else:
            r['remark'] = specials_text

        r['fill_color'] = COLOR_FILL.get(code)
        r['font_color'] = FONT_COLOR.get(code)

    return records
