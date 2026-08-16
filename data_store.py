"""
Flaming News - Daily data persistence.

Each day's Flaming News content (both sheets, fully editable) is stored as one
JSON file under DATA_DIR, named by date: data/2026-08-16.json

This is what makes "previous day's data auto-loads" and "only edit what's
different today" work: opening the app on a new day with no file yet copies
forward the most recent existing day's JSON as the starting point.
"""
import json
import os
import glob
from datetime import date, datetime

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
os.makedirs(DATA_DIR, exist_ok=True)


def _path_for(iso_date):
    return os.path.join(DATA_DIR, f'{iso_date}.json')


def empty_day(iso_date):
    """A brand new, structurally-complete but empty day (used only if there is
    truly no prior data at all — e.g. very first run)."""
    return {
        'iso_date': iso_date,
        'quote': 'Be kind whenever possible. It is always possible. - Dalai Lama',
        'forecast': {
            'dates': ['', '', '', '', '', '', ''],
            'occupancy_pct': ['', '', '', '', '', '', ''],
            'rooms_occupied': ['', '', '', '', '', '', ''],
            'adr': ['', '', '', '', '', '', ''],
            'arrivals': ['', '', '', '', '', '', ''],
            'departures': ['', '', '', '', '', '', ''],
        },
        'am_mod': '', 'pm_mod': '', 'nm': '', 'weekend_eod': '',
        'house_status': {'oos_ooo': '', 'no_show': '', 'comp_house_use': ''},
        'weather': '',
        'enrollments_goal': '', 'enrollments_ytd': '',
        'fairmont_goals': {
            'ces_goal': '', 'ces_actual': '',
            'lqa_goal': '', 'lqa_actual': '',
            'rps_goal': '', 'rps_mtd': '', 'rps_ytd': '',
        },
        'site_inspections': [],
        'vip_arrivals': [],
        'vip_inhouse': [],
        'vip_departures': [],
        'events': [],
        'birthday': ['', '', '', ''],
        'anniversary': ['', '', '', ''],
        'fb_performance': [],
    }


def list_available_dates():
    """All dates that have saved data, sorted ascending (ISO yyyy-mm-dd strings)."""
    files = glob.glob(os.path.join(DATA_DIR, '*.json'))
    dates = sorted(os.path.splitext(os.path.basename(f))[0] for f in files)
    return dates


def most_recent_before(iso_date):
    """Latest saved date strictly before iso_date, or None."""
    dates = [d for d in list_available_dates() if d < iso_date]
    return dates[-1] if dates else None


def load_day(iso_date):
    """Load a specific day's saved data, or None if it doesn't exist."""
    path = _path_for(iso_date)
    if not os.path.exists(path):
        return None
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_today_or_carry_forward(iso_date=None):
    """
    Core "previous day auto-load" behaviour:
      1. If today's file already exists (already opened/edited today) -> return it as-is.
      2. Otherwise, find the most recent earlier day and carry its data forward,
         shifting the 7-day forecast window and clearing day-specific VIP lists
         that should come from fresh report uploads, while keeping every
         manually-maintained field (MOD names, weather, goals, events,
         birthdays/anniversaries the user hasn't cleared, quote, etc).
      3. If there is no prior data at all, return a blank structurally-correct day.
    Returns (data_dict, carried_forward: bool, source_date: str|None)
    """
    iso_date = iso_date or date.today().isoformat()
    existing = load_day(iso_date)
    if existing is not None:
        return existing, False, None

    prev_date = most_recent_before(iso_date)
    if prev_date is None:
        return empty_day(iso_date), False, None

    prev = load_day(prev_date)
    carried = json.loads(json.dumps(prev))  # deep copy
    carried['iso_date'] = iso_date

    # VIP guest lists and site inspections are day-specific — today's reports
    # (or manual entry) should populate them fresh rather than repeating
    # yesterday's guests. Keep everything else (MOD roster, weather, goals,
    # enrollments, events/birthday/anniversary/F&B the team plans ahead of
    # time, quote of the day) so only genuine changes need editing.
    carried['vip_arrivals'] = []
    carried['vip_inhouse'] = []
    carried['vip_departures'] = []
    carried['site_inspections'] = []
    carried['fb_performance'] = []
    carried['events'] = []
    carried['birthday'] = ['', '', '', '']
    carried['anniversary'] = ['', '', '', '']

    return carried, True, prev_date


def save_day(iso_date, data):
    data = dict(data)
    data['iso_date'] = iso_date
    data['_saved_at'] = datetime.now().isoformat(timespec='seconds')
    with open(_path_for(iso_date), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return data
