"""
Flaming News - VIP Business Rules Engine
Derives the Remarks text and cell colour for each guest, based on VIP code,
and filters the list down to genuine VIP guests per management rules:

  - DV guests are never included in Flaming News.
  - Booking.com-sourced guests are never classified/shown as V1 VIPs.
  - Only genuine VIP codes are kept (T3/T4/T5/T6/SA/V1); anything else is dropped.
"""
import codes

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


def _sa_remark(raw_specials_lines):
    """
    Derive the SA (Special Attention) remark from the guest's raw
    'Specials:' code lines, using the hotel's code table.
    Returns (remark_text, needs_review).
      - If a recognisable occasion code is found (Birthday, Anniversary,
        Honeymooners, Recovery, ...), remark is built from it and no review
        is needed.
      - Otherwise remark defaults to 'Special Attention' (SA's own
        definition) but is flagged needs_review=True so a person checks
        whether there's a more specific reason buried in free-text notes
        the parser can't read.
    """
    tokens = []
    for line in raw_specials_lines:
        tokens.extend(t.strip().upper() for t in line.replace(',', ' ').split() if t.strip())
    labels, matched, unmatched = codes.match_occasions(tokens)
    if labels:
        return ', '.join(labels), False
    return 'Special Attention', True


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
      - 'remark'                 -> derived Remarks text for the template
      - 'remarks_needs_review'   -> True if the remark is a best-effort
                                     guess the person should double-check
      - 'fill_color'  -> ARGB hex string or None
      - 'font_color'  -> ARGB hex string or None
    Returns the filtered, enriched list (does not mutate the input list itself).
    """
    records = filter_vip_records(records)

    for r in records:
        code = (r.get('vip_code') or '').strip().upper()
        needs_review = False

        if code in FIXED_REMARKS:
            remark = FIXED_REMARKS[code]
        elif code == 'SA':
            remark, needs_review = _sa_remark(r.get('raw_specials', []))
        elif code == 'V1':
            remark = r.get('remark', '') or 'VIP Guest'
        else:
            remark = ' '.join(r.get('raw_specials', []))

        r['remark'] = remark
        r['remarks_needs_review'] = needs_review
        r['vip_code'] = code
        r['fill_color'] = COLOR_FILL.get(code)
        r['font_color'] = FONT_COLOR.get(code)

    return records


def color_for_code(code):
    code = (code or '').strip().upper()
    return COLOR_FILL.get(code), FONT_COLOR.get(code)
