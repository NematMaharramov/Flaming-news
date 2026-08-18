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

    sections: ordered list describing every repeating/table-like part of
    the report. type='table' sections are fully driven by `columns` (the
    HTML editor renders them generically — see static/report-render.js) so
    an admin can add/remove/resize columns, or add/remove whole sections,
    without touching application code. type='parallel_lines' is the
    Birthday/Anniversary style (two free-text columns, no headers per
    line). type='fixed' sections (forecast, the AM MOD/weather/goals
    block) are NOT yet config-driven — their layout is still hardcoded in
    the template; that's a larger follow-up since they're matrix-shaped
    rather than repeating rows.
    """
    return {
        'schema_version': 2,
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
        'sections': [
            {'id': 'forecast', 'sheet': 1, 'title': '7 Day Forecast', 'type': 'fixed', 'enabled': True, 'order': 1},
            {'id': 'am_pm_mod', 'sheet': 1, 'title': 'AM MOD / PM MOD / House Status / Weather / Enrollments', 'type': 'fixed', 'enabled': True, 'order': 2},
            {'id': 'goals', 'sheet': 1, 'title': 'Fairmont Baku 2026 Goals', 'type': 'fixed', 'enabled': True, 'order': 3},
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
    """Bring an older on-disk config up to the current schema_version.
    v1 -> v2: 'sections' gained 'type'/'columns'/'order' for table-driven
    rendering. Since v1 sections only had id/sheet/title/enabled, the
    safest migration is to take the admin's enabled/disabled choices and
    title edits (if any) and re-merge them onto the current factory
    section definitions, rather than trying to guess columns for a v1
    section that never had any."""
    if cfg.get('schema_version', 1) >= 2:
        return cfg
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
