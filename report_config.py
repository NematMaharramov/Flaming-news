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
    """
    return {
        'schema_version': 1,
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
        # Section order/visibility for Sheet 1 / Sheet 2 — placeholder for
        # the dynamic-sections work (later part); today it just documents
        # what exists so the admin UI has something real to list.
        'sections': [
            {'id': 'forecast', 'sheet': 1, 'title': '7 Day Forecast', 'enabled': True},
            {'id': 'am_pm_mod', 'sheet': 1, 'title': 'AM MOD / PM MOD / House Status / Weather / Enrollments', 'enabled': True},
            {'id': 'goals', 'sheet': 1, 'title': 'Fairmont Baku 2026 Goals', 'enabled': True},
            {'id': 'site_inspections', 'sheet': 1, 'title': 'Site Inspections', 'enabled': True},
            {'id': 'vip_arrivals', 'sheet': 1, 'title': 'VIP Arrivals', 'enabled': True},
            {'id': 'vip_inhouse', 'sheet': 1, 'title': 'VIP In-House Guests', 'enabled': True},
            {'id': 'vip_departures', 'sheet': 1, 'title': 'VIP Departures', 'enabled': True},
            {'id': 'events', 'sheet': 2, 'title': 'Events', 'enabled': True},
            {'id': 'birthday_anniversary', 'sheet': 2, 'title': 'Birthday / Anniversary', 'enabled': True},
            {'id': 'fb_performance', 'sheet': 2, 'title': 'Food & Beverage Performance', 'enabled': True},
        ],
    }


def _read(path):
    if not os.path.exists(path):
        return None
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


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
