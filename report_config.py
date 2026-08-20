"""
Flaming News - Report Configuration Engine

Simple file-based (no database) configuration system per the admin
requirements:

  report-config.json  -> the currently PUBLISHED config (what the live
                          User-Facing HTML/Excel/PDF actually use)
  draft-config.json   -> what the Admin is currently editing/previewing;
                          not live until explicitly Published
  history/<ts>.json   -> every previously-published version, kept forever
                          so it can be previewed and restored

Key ideas:
  - Draft/Publish: editing never touches the live config until Publish.
  - History: every Publish snapshots the outgoing config first, so nothing
    is ever lost.
  - Static/Default values: a small set of "rarely-changing" fields (e.g.
    Fairmont Goals) live in the config. Every time a NEW day is created
    (see data_store.load_today_or_carry_forward), these defaults are
    stamped onto that day's data. Once a day is saved, its own values are
    authoritative — changing the config default later only affects days
    created AFTER that change, and a manual per-day edit never rewrites
    the global default. This matches: "dəyişiklik yalnız yeni günlərə
    tətbiq olunur, köhnə saxlanmış günlərə toxunmur".
"""
import json
import os
import glob
from datetime import datetime

CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'config')
HISTORY_DIR = os.path.join(CONFIG_DIR, 'history')
PUBLISHED_PATH = os.path.join(CONFIG_DIR, 'report-config.json')
DRAFT_PATH = os.path.join(CONFIG_DIR, 'draft-config.json')

os.makedirs(HISTORY_DIR, exist_ok=True)


def default_config():
    """
    The factory-default configuration. static_defaults keys are dotted
    paths into a day's data dict (same shape as data_store.empty_day),
    applied only at day-creation time (see data_store.py).

    sections: ordered list describing every part of the report.
    - type='table': fully driven by `columns` — admin can add/remove/
      resize/style columns, or add/remove whole sections, without touching
      application code. Each column may carry `header_style` (the label
      row) and `cell_style` (the data row) separately — see Part F1. A
      legacy single `style` key is still read as a fallback for either
      when the specific one isn't set, so older configs keep rendering
      unchanged.
    - Every style object (section.style, column.header_style/cell_style,
      fixed-row .style, title_row.style, quote_row.style) shares the same
      shape: bg, color, font_family, font_size, font_weight (or the older
      boolean `bold`), italic, underline, align, padding, line_height,
      border_width, border_color, border_style. Consumers (index.html,
      admin_editor.html) apply whichever keys are present and ignore the
      rest, so new keys can be added without a schema migration.
    - type='parallel_lines': the Birthday/Anniversary style (two free-text
      columns, no headers per line).
    - type='fixed': matrix-shaped blocks (7 Day Forecast, the AM MOD/
      Weather/Goals block) that aren't repeating rows in the same sense as
      a table. As of schema v3 these are still driven by config too:
      'forecast' has `metric_rows` (one row per forecast metric — Occupancy
      %, ADR, etc, each mapped to a key in a day's forecast.* arrays) and
      'am_pm_mod'/'goals' have `fields` (one row per label/value pair,
      each mapped to a dotted data path). An admin can add/edit/delete/
      reorder these rows from the Visual Editor without code changes —
      e.g. adding a 'RevPAR' forecast metric or an 'NPS Goal' field.
    """
    return {
        'schema_version': 4,
        'updated_at': None,
        'static_defaults': {
            'fairmont_goals.ces_goal': 90,
            'fairmont_goals.lqa_goal': 90,
            'fairmont_goals.rps_goal': 0.929,
        },
        'occupancy_color_rule': {
            'enabled': True,
            'low_pct': 0,
            'low_color': '#FF0000',
            'high_pct': 100,
            'high_color': '#00B050',
        },
        # The two banner rows repeated at the top of every sheet. Previously
        # hardcoded ('FLAMING NEWS FOR ...' / 'Quote of the Day : ...') in
        # templates/index.html; now config-driven (Part F1) so the Visual
        # Editor can restyle and re-word them like any other section.
        # label_template placeholders: {date} for title_row, {quote} for
        # quote_row — substituted client-side in templates/index.html.
        'title_row': {
            'label_template': 'FLAMING NEWS FOR {date}',
            'style': {
                'bg': '#E7E6E6', 'color': '#FF0000', 'bold': True,
                'font_size': 20, 'align': 'center', 'padding': 10,
            },
        },
        'quote_row': {
            'label_template': 'Quote of the Day : \u201c{quote}\u201d',
            'style': {
                'bg': '#E7E6E6', 'color': '#FF0000', 'bold': True,
                'font_size': 16, 'align': 'center', 'padding': 10,
            },
        },
        'sections': [
            {
                'id': 'forecast', 'sheet': 1, 'title': '7 Day Forecast', 'type': 'fixed',
                'enabled': True, 'order': 1,
                'metric_rows': [
                    {'key': 'occupancy_pct', 'label': 'Occupancy %'},
                    {'key': 'rooms_occupied', 'label': 'Rooms Occupied (excl house use & comp)'},
                    {'key': 'adr', 'label': 'ADR'},
                    {'key': 'arrivals', 'label': 'Arrivals'},
                    {'key': 'departures', 'label': 'Departures'},
                ],
            },
            {
                'id': 'am_pm_mod', 'sheet': 1, 'title': 'AM MOD / PM MOD / House Status / Weather / Enrollments',
                'type': 'fixed', 'enabled': True, 'order': 2,
                'fields': [
                    {'key': 'am_mod', 'label': 'AM MOD:', 'data_path': 'am_mod'},
                    {'key': 'pm_mod', 'label': 'PM MOD:', 'data_path': 'pm_mod'},
                    {'key': 'nm', 'label': 'NM:', 'data_path': 'nm'},
                    {'key': 'weekend_eod', 'label': 'Weekend EOD:', 'data_path': 'weekend_eod'},
                    {'key': 'oos_ooo', 'label': 'OOS/ OOO', 'data_path': 'house_status.oos_ooo'},
                    {'key': 'no_show', 'label': 'No Show', 'data_path': 'house_status.no_show'},
                    {'key': 'comp_house_use', 'label': 'Comp./house use', 'data_path': 'house_status.comp_house_use'},
                    {'key': 'weather', 'label': 'Weather Today', 'data_path': 'weather'},
                    {'key': 'enrollments_goal', 'label': 'ALL Enrollments Goal', 'data_path': 'enrollments_goal'},
                    {'key': 'enrollments_ytd', 'label': 'ALL Enrollments YTD', 'data_path': 'enrollments_ytd'},
                ],
            },
            {
                'id': 'goals', 'sheet': 1, 'title': 'Fairmont Baku 2026 Goals', 'type': 'fixed',
                'enabled': True, 'order': 3,
                'fields': [
                    {'key': 'ces_goal', 'label': 'CES Goal', 'data_path': 'fairmont_goals.ces_goal'},
                    {'key': 'ces_actual', 'label': 'CES Actual', 'data_path': 'fairmont_goals.ces_actual'},
                    {'key': 'lqa_goal', 'label': 'LQA Goal', 'data_path': 'fairmont_goals.lqa_goal'},
                    {'key': 'lqa_actual', 'label': 'LQA Actual', 'data_path': 'fairmont_goals.lqa_actual'},
                    {'key': 'rps_goal', 'label': 'RPS Goal', 'data_path': 'fairmont_goals.rps_goal'},
                    {'key': 'rps_mtd', 'label': 'RPS MTD', 'data_path': 'fairmont_goals.rps_mtd'},
                    {'key': 'rps_ytd', 'label': 'RPS YTD', 'data_path': 'fairmont_goals.rps_ytd'},
                ],
            },
            {
                'id': 'site_inspections', 'sheet': 1, 'title': 'Site Inspections', 'type': 'table',
                'enabled': True, 'order': 4, 'data_key': 'site_inspections',
                'columns': [
                    {'key': 'time', 'label': 'Time', 'width_pct': 10, 'colspan': 1, 'kind': 'time'},
                    {'key': 'guest', 'label': 'Guest', 'width_pct': 16, 'colspan': 1, 'kind': 'text'},
                    {'key': 'position_company', 'label': 'Position / Company', 'width_pct': 52, 'colspan': 5, 'kind': 'text'},
                    {'key': 'sales_contact', 'label': 'Sales Contact', 'width_pct': 22, 'colspan': 2, 'kind': 'text'},
                ],
            },
            {
                'id': 'vip_arrivals', 'sheet': 1, 'title': 'VIP Arrivals', 'type': 'table',
                'enabled': True, 'order': 5, 'data_key': 'vip_arrivals', 'vip_guest_row': True,
                'columns': [
                    {'key': 'guest', 'label': 'Guest', 'width_pct': 20, 'colspan': 1, 'kind': 'guest_with_photo'},
                    {'key': 'code', 'label': 'Code', 'width_pct': 7, 'colspan': 1, 'kind': 'vip_code'},
                    {'key': 'room', 'label': 'Room', 'width_pct': 8, 'colspan': 1, 'kind': 'text'},
                    {'key': 'eta', 'label': 'ETA', 'width_pct': 8, 'colspan': 1, 'kind': 'time'},
                    {'key': 'remarks', 'label': 'Remarks', 'width_pct': 22, 'colspan': 2, 'kind': 'remarks'},
                    {'key': 'company', 'label': 'Company', 'width_pct': 25, 'colspan': 2, 'kind': 'text'},
                ],
            },
            {
                'id': 'vip_inhouse', 'sheet': 1, 'title': 'VIP In-House Guests', 'type': 'table',
                'enabled': True, 'order': 6, 'data_key': 'vip_inhouse', 'vip_guest_row': True,
                'columns': [
                    {'key': 'guest', 'label': 'Guest', 'width_pct': 20, 'colspan': 1, 'kind': 'guest_with_photo'},
                    {'key': 'code', 'label': 'Code', 'width_pct': 7, 'colspan': 1, 'kind': 'vip_code'},
                    {'key': 'room', 'label': 'Room', 'width_pct': 8, 'colspan': 1, 'kind': 'text'},
                    {'key': 'departure_day', 'label': 'Departure day', 'width_pct': 10, 'colspan': 1, 'kind': 'date'},
                    {'key': 'remarks', 'label': 'Remarks', 'width_pct': 18, 'colspan': 2, 'kind': 'remarks'},
                    {'key': 'company', 'label': 'Company', 'width_pct': 25, 'colspan': 2, 'kind': 'text'},
                ],
            },
            {
                'id': 'vip_departures', 'sheet': 1, 'title': 'VIP Departures', 'type': 'table',
                'enabled': True, 'order': 7, 'data_key': 'vip_departures', 'vip_guest_row': True,
                'columns': [
                    {'key': 'guest', 'label': 'Guest', 'width_pct': 20, 'colspan': 1, 'kind': 'guest_with_photo'},
                    {'key': 'code', 'label': 'Code', 'width_pct': 7, 'colspan': 1, 'kind': 'vip_code'},
                    {'key': 'room', 'label': 'Room', 'width_pct': 8, 'colspan': 1, 'kind': 'text'},
                    {'key': 'departure_day', 'label': 'Departure day', 'width_pct': 10, 'colspan': 1, 'kind': 'date'},
                    {'key': 'departure_time', 'label': 'Departure Time', 'width_pct': 16, 'colspan': 2, 'kind': 'time'},
                    {'key': 'company', 'label': 'Company', 'width_pct': 25, 'colspan': 2, 'kind': 'text'},
                ],
            },
            {
                'id': 'events', 'sheet': 2, 'title': 'Events', 'type': 'table',
                'enabled': True, 'order': 8, 'data_key': 'events',
                'columns': [
                    {'key': 'time', 'label': 'Time', 'width_pct': 10, 'colspan': 1, 'kind': 'time'},
                    {'key': 'meeting_room', 'label': 'Meeting Room', 'width_pct': 18, 'colspan': 2, 'kind': 'text'},
                    {'key': 'company', 'label': 'Company', 'width_pct': 36, 'colspan': 3, 'kind': 'text'},
                    {'key': 'sales_contact', 'label': 'Sales contact', 'width_pct': 36, 'colspan': 3, 'kind': 'text'},
                ],
            },
            {
                'id': 'birthday_anniversary', 'sheet': 2, 'title': 'Birthday / Anniversary', 'type': 'parallel_lines',
                'enabled': True, 'order': 9,
                'left': {'label': 'Birthday', 'data_key': 'birthday', 'colspan': 3},
                'right': {'label': 'Anniversary', 'data_key': 'anniversary', 'colspan': 6},
            },
            {
                'id': 'fb_performance', 'sheet': 2, 'title': 'Food & Beverage Performance', 'type': 'table',
                'enabled': True, 'order': 10, 'data_key': 'fb_performance',
                'columns': [
                    {'key': 'outlet', 'label': 'Outlet', 'width_pct': 18, 'colspan': 1, 'kind': 'text_left'},
                    {'key': 'revenue', 'label': 'Revenue', 'width_pct': 14, 'colspan': 1, 'kind': 'text'},
                    {'key': 'gih', 'label': 'GİH', 'width_pct': 9, 'colspan': 1, 'kind': 'text'},
                    {'key': 'external_guests', 'label': 'External guests', 'width_pct': 14, 'colspan': 1, 'kind': 'text'},
                    {'key': 'special_promotions', 'label': 'Special promotions', 'width_pct': 45, 'colspan': 5, 'kind': 'text_left'},
                ],
            },
        ],
    }


def _migrate(cfg):
    """Bring an older on-disk config up to the current schema_version,
    chaining through each step. Each step re-merges the admin's actual
    choices (enabled/title/order/style, and now custom sections) onto the
    current factory section definitions, rather than trying to guess
    structure a prior version never had."""
    version = cfg.get('schema_version', 1)

    if version < 2:
        old_by_id = {s.get('id'): s for s in cfg.get('sections', [])}
        fresh = default_config()
        for s in fresh['sections']:
            old = old_by_id.get(s['id'])
            if old:
                s['enabled'] = old.get('enabled', s['enabled'])
                if old.get('title'):
                    s['title'] = old['title']
        cfg['sections'] = fresh['sections']
        cfg['schema_version'] = 2
        version = 2

    if version < 3:
        # v2 'fixed' sections had no metric_rows/fields; v3 adds them.
        # Any custom (admin-added) sections from v2 carry forward as-is.
        old_by_id = {s.get('id'): s for s in cfg.get('sections', [])}
        fresh = default_config()
        fresh_ids = {s['id'] for s in fresh['sections']}
        merged = []
        for s in fresh['sections']:
            old = old_by_id.get(s['id'])
            if old:
                s['enabled'] = old.get('enabled', s['enabled'])
                s['order'] = old.get('order', s['order'])
                if old.get('title'):
                    s['title'] = old['title']
                if old.get('style'):
                    s['style'] = old['style']
                if s.get('type') == 'table' and old.get('columns'):
                    s['columns'] = old['columns']
            merged.append(s)
        for old in cfg.get('sections', []):
            if old.get('id') not in fresh_ids:
                merged.append(old)  # preserve admin-added custom sections
        cfg['sections'] = merged
        cfg['schema_version'] = 3
        version = 3

    if version < 4:
        # Part F1 — Visual Editor Style System.
        # (a) title_row/quote_row are new; stamp in the factory defaults for
        #     configs that predate them.
        # (b) table columns now support separate header_style/cell_style
        #     (previously a single 'style' applied to both header and data
        #     cell). Any admin-set 'style' is copied into both so existing
        #     designs look identical after the upgrade; 'style' itself is
        #     left in place as a legacy fallback (see report_config.py
        #     column-style resolution used by index.html/admin_editor.html).
        fresh = default_config()
        cfg.setdefault('title_row', fresh['title_row'])
        cfg.setdefault('quote_row', fresh['quote_row'])
        for s in cfg.get('sections', []):
            if s.get('type') == 'table':
                for col in s.get('columns', []):
                    if col.get('style') and 'header_style' not in col and 'cell_style' not in col:
                        col['header_style'] = dict(col['style'])
                        col['cell_style'] = dict(col['style'])
        cfg['schema_version'] = 4
        version = 4

    return cfg


def _read(path):
    if not os.path.exists(path):
        return None
    with open(path, 'r', encoding='utf-8') as f:
        return _migrate(json.load(f))


def _write(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_published():
    """The live config. Created with factory defaults on first use."""
    cfg = _read(PUBLISHED_PATH)
    if cfg is None:
        cfg = default_config()
        cfg['updated_at'] = datetime.now().isoformat(timespec='seconds')
        _write(PUBLISHED_PATH, cfg)
    return cfg


def load_draft():
    """The admin's working copy. Starts as a copy of the published config
    the first time someone opens the admin panel."""
    cfg = _read(DRAFT_PATH)
    if cfg is None:
        cfg = load_published()
        _write(DRAFT_PATH, cfg)
    return cfg


def save_draft(new_config):
    """Save without publishing — visible only in the admin preview."""
    new_config['updated_at'] = datetime.now().isoformat(timespec='seconds')
    _write(DRAFT_PATH, new_config)
    return new_config


def publish_draft():
    """
    Snapshot the current published config into history, then promote the
    draft to be the new published config. The draft file remains (as a
    copy of what's now published) so admin editing continues seamlessly.
    """
    current_published = load_published()
    ts = datetime.now().strftime('%Y%m%dT%H%M%S')
    history_path = os.path.join(HISTORY_DIR, f'{ts}.json')
    _write(history_path, current_published)

    draft = load_draft()
    draft['updated_at'] = datetime.now().isoformat(timespec='seconds')
    _write(PUBLISHED_PATH, draft)
    return draft


def list_history():
    """Newest first: [{'id': '20260817T142233', 'updated_at': ...}, ...]"""
    entries = []
    for path in glob.glob(os.path.join(HISTORY_DIR, '*.json')):
        entry_id = os.path.splitext(os.path.basename(path))[0]
        cfg = _read(path)
        entries.append({'id': entry_id, 'updated_at': cfg.get('updated_at') if cfg else None})
    entries.sort(key=lambda e: e['id'], reverse=True)
    return entries


def get_history_version(entry_id):
    """Full config for one history entry, or None."""
    safe_id = os.path.basename(entry_id)  # no path traversal
    return _read(os.path.join(HISTORY_DIR, f'{safe_id}.json'))


def restore_history_version(entry_id):
    """
    Load a past version as the new DRAFT (not live yet — admin previews
    then must Publish, same as any other edit).
    """
    cfg = get_history_version(entry_id)
    if cfg is None:
        return None
    return save_draft(cfg)


def get_static_defaults():
    """Dotted-path -> value map from the PUBLISHED config, used when a
    brand new day is created."""
    return load_published().get('static_defaults', {})


def apply_static_defaults(data):
    """
    Stamp the published config's static/default values onto a day's data
    dict (in place), for dotted paths like 'fairmont_goals.lqa_goal'.
    Only called at day-CREATION time — see data_store.py — never on an
    already-saved day, so per-day manual overrides are never clobbered.
    """
    defaults = get_static_defaults()
    for dotted_path, value in defaults.items():
        parts = dotted_path.split('.')
        node = data
        for p in parts[:-1]:
            if not isinstance(node.get(p), dict):
                node[p] = {}
            node = node[p]
        node[parts[-1]] = value
    return data


# ---------------------------------------------------------------------------
# Occupancy % gradient (red at low_pct -> green at high_pct), Part E.
# Shared by excel_writer.py and the PDF print template so all three
# surfaces (live HTML, Excel, PDF) compute the exact same colour for the
# exact same percentage.
# ---------------------------------------------------------------------------
def _hex_to_rgb(hex_color):
    h = (hex_color or '#000000').lstrip('#')
    if len(h) != 6:
        return (0, 0, 0)
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def occupancy_color(pct_value, rule=None):
    """
    pct_value: a number (e.g. 61.54) — NOT a fraction, matches how
    occupancy is stored/displayed ('61.54%' strings, parsed to 61.54).
    rule: an occupancy_color_rule dict (falls back to the published
    config's rule if not given). Returns an '#RRGGBB' string, or None if
    the rule is disabled or pct_value can't be read.
    """
    if rule is None:
        rule = load_published().get('occupancy_color_rule', {})
    if not rule.get('enabled', True) or pct_value is None:
        return None
    try:
        pct = float(pct_value)
    except (TypeError, ValueError):
        return None
    lo_pct = rule.get('low_pct', 0)
    hi_pct = rule.get('high_pct', 100)
    if hi_pct == lo_pct:
        return rule.get('high_color')
    t = max(0.0, min(1.0, (pct - lo_pct) / (hi_pct - lo_pct)))
    lo = _hex_to_rgb(rule.get('low_color', '#FF0000'))
    hi = _hex_to_rgb(rule.get('high_color', '#00B050'))
    r = round(lo[0] + (hi[0] - lo[0]) * t)
    g = round(lo[1] + (hi[1] - lo[1]) * t)
    b = round(lo[2] + (hi[2] - lo[2]) * t)
    return '#{:02X}{:02X}{:02X}'.format(r, g, b)


def parse_pct_string(value):
    """'61.54%' -> 61.54 ; '61.54' -> 61.54 ; '' / None -> None."""
    if value is None:
        return None
    s = str(value).strip().rstrip('%')
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None
