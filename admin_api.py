"""
Flask blueprint providing admin endpoints for:
- config save/load (draft)
- publish (draft -> published)
- history listing / restore
- xml upload and structure parse

Drop this file in the repo root and register the blueprint in app.py:
from admin_api import admin_bp
app.register_blueprint(admin_bp)
"""
import os
import json
import shutil
from datetime import datetime
from flask import Blueprint, request, jsonify, current_app, send_file, abort
import xml.etree.ElementTree as ET

try:
    import xmltodict
except Exception:
    xmltodict = None

admin_bp = Blueprint('admin_api', __name__, url_prefix='/api/admin')

# file layout (relative to repository root)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, 'config')
DRAFT_DIR = os.path.join(CONFIG_DIR, 'draft')
PUBLISHED_DIR = os.path.join(CONFIG_DIR, 'published')
HISTORY_DIR = os.path.join(CONFIG_DIR, 'history')
HISTORY_INDEX = os.path.join(CONFIG_DIR, 'report-history.json')


def ensure_dirs():
    for d in (CONFIG_DIR, DRAFT_DIR, PUBLISHED_DIR, HISTORY_DIR):
        os.makedirs(d, exist_ok=True)
    if not os.path.exists(HISTORY_INDEX):
        with open(HISTORY_INDEX, 'w', encoding='utf-8') as f:
            json.dump([], f, indent=2)

def load_json(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_json(path, data):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)

def version_stamp():
    return datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')

def history_push(template_rel='report-template.html', config_rel='report-config.json', note=None):
    """
    Copy current draft (or published) files into history with timestamp.
    """
    ensure_dirs()
    ts = version_stamp()
    suffix = ts
    hist_template = os.path.join(HISTORY_DIR, f'{suffix}_template.html')
    hist_config = os.path.join(HISTORY_DIR, f'{suffix}_config.json')
    # choose the most recent of draft/published for history snapshot
    src_template = os.path.join(DRAFT_DIR, template_rel) if os.path.exists(os.path.join(DRAFT_DIR, template_rel)) else os.path.join(PUBLISHED_DIR, template_rel)
    src_config = os.path.join(DRAFT_DIR, config_rel) if os.path.exists(os.path.join(DRAFT_DIR, config_rel)) else os.path.join(PUBLISHED_DIR, config_rel)
    if os.path.exists(src_template):
        shutil.copy2(src_template, hist_template)
    if os.path.exists(src_config):
        shutil.copy2(src_config, hist_config)
    history = load_json(HISTORY_INDEX, [])
    entry = {
        'id': suffix,
        'timestamp': datetime.utcnow().isoformat() + 'Z',
        'template': os.path.basename(hist_template),
        'config': os.path.basename(hist_config),
        'note': note or ''
    }
    history.insert(0, entry)
    save_json(HISTORY_INDEX, history)
    return entry

@admin_bp.route('/config', methods=['GET'])
def get_config():
    """
    Return current draft config and published metadata.
    """
    ensure_dirs()
    draft_cfg_path = os.path.join(DRAFT_DIR, 'report-config.json')
    published_cfg_path = os.path.join(PUBLISHED_DIR, 'report-config.json')
    draft_cfg = load_json(draft_cfg_path, {})
    published_cfg = load_json(published_cfg_path, {})
    return jsonify({'status':'ok', 'draft': draft_cfg, 'published': published_cfg})

@admin_bp.route('/config/save', methods=['POST'])
def save_config():
    """
    Save draft config (and template optionally).
    Accepts JSON body:
    { "config": {...}, "template_html": "<html>...</html>", "note":"optional note" }
    """
    ensure_dirs()
    body = request.get_json(force=True)
    cfg = body.get('config') or {}
    template_html = body.get('template_html')  # optional
    note = body.get('note')
    # push previous draft to history before overwrite
    history_push()
    # save draft config
    draft_cfg_path = os.path.join(DRAFT_DIR, 'report-config.json')
    save_json(draft_cfg_path, cfg)
    # save template if provided
    if template_html is not None:
        draft_template_path = os.path.join(DRAFT_DIR, 'report-template.html')
        tmp = draft_template_path + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            f.write(template_html)
        os.replace(tmp, draft_template_path)
    return jsonify({'status':'ok', 'saved_at': datetime.utcnow().isoformat() + 'Z'})

@admin_bp.route('/config/publish', methods=['POST'])
def publish_config():
    """
    Publish current draft: copy draft -> published, and archive current published to history.
    Optionally accepts {"note":"..."}.
    """
    ensure_dirs()
    body = request.get_json(force=True) or {}
    note = body.get('note')
    # archive current published into history
    if os.path.exists(os.path.join(PUBLISHED_DIR, 'report-config.json')) or os.path.exists(os.path.join(PUBLISHED_DIR, 'report-template.html')):
        history_push()
    # copy draft to published (if present)
    draft_cfg = os.path.join(DRAFT_DIR, 'report-config.json')
    draft_template = os.path.join(DRAFT_DIR, 'report-template.html')
    if os.path.exists(draft_cfg):
        shutil.copy2(draft_cfg, os.path.join(PUBLISHED_DIR, 'report-config.json'))
    if os.path.exists(draft_template):
        shutil.copy2(draft_template, os.path.join(PUBLISHED_DIR, 'report-template.html'))
    return jsonify({'status':'ok', 'published_at': datetime.utcnow().isoformat() + 'Z'})

@admin_bp.route('/history', methods=['GET'])
def history_list():
    ensure_dirs()
    hist = load_json(HISTORY_INDEX, [])
    return jsonify({'status':'ok', 'history': hist})

@admin_bp.route('/history/<hid>/restore', methods=['POST'])
def history_restore(hid):
    """
    Restore a history entry back into the draft (does not publish).
    """
    ensure_dirs()
    hist = load_json(HISTORY_INDEX, [])
    entry = next((e for e in hist if e.get('id')==hid), None)
    if not entry:
        return jsonify({'status':'error', 'message':'History id not found'}), 404
    # map file names to history files
    hist_template = os.path.join(HISTORY_DIR, entry['template'])
    hist_config = os.path.join(HISTORY_DIR, entry['config'])
    if os.path.exists(hist_template):
        shutil.copy2(hist_template, os.path.join(DRAFT_DIR, 'report-template.html'))
    if os.path.exists(hist_config):
        shutil.copy2(hist_config, os.path.join(DRAFT_DIR, 'report-config.json'))
    return jsonify({'status':'ok', 'restored': hid})

@admin_bp.route('/upload/xml', methods=['POST'])
def upload_xml():
    """
    Accept XML upload, parse into JSON, and return the parsed structure (a tree of paths).
    """
    ensure_dirs()
    if 'xml' not in request.files:
        return jsonify({'status':'error', 'message':'missing file field "xml"'}), 400
    f = request.files['xml']
    try:
        raw = f.read()
        # Try xmltodict for robust conversion if available, else fallback to ElementTree
        try:
            if xmltodict:
                parsed = xmltodict.parse(raw)
            else:
                raise Exception('xmltodict not available')
            # return JSON parsed tree and a flattened list of paths (keys)
            def collect_paths(obj, base=''):
                out = []
                if isinstance(obj, dict):
                    for k,v in obj.items():
                        p = f'{base}.{k}' if base else k
                        out.append(p)
                        out.extend(collect_paths(v, p))
                elif isinstance(obj, list):
                    for i, item in enumerate(obj):
                        p = f'{base}[{i}]'
                        out.append(p)
                        out.extend(collect_paths(item, p))
                return out
            paths = collect_paths(parsed)
            return jsonify({'status':'ok', 'tree': parsed, 'paths': paths})
        except Exception as e:
            # fallback
            root = ET.fromstring(raw)
            def elem_to_dict(e):
                d = {}
                for child in e:
                    tag = child.tag
                    val = elem_to_dict(child)
                    if tag in d:
                        if not isinstance(d[tag], list):
                            d[tag] = [d[tag]]
                        d[tag].append(val)
                    else:
                        d[tag] = val
                if e.text and e.text.strip():
                    d['_text'] = e.text.strip()
                return d
            parsed = {root.tag: elem_to_dict(root)}
            return jsonify({'status':'ok', 'tree': parsed, 'paths': list(parsed.keys())})
    except Exception as ex:
        return jsonify({'status':'error', 'message': str(ex)}), 500

@admin_bp.route('/config/download/<which>', methods=['GET'])
def config_download(which):
    """
    Download published or draft template/config.
    which = draft-template | draft-config | published-template | published-config
    """
    ensure_dirs()
    mapping = {
        'draft-template': os.path.join(DRAFT_DIR, 'report-template.html'),
        'draft-config': os.path.join(DRAFT_DIR, 'report-config.json'),
        'published-template': os.path.join(PUBLISHED_DIR, 'report-template.html'),
        'published-config': os.path.join(PUBLISHED_DIR, 'report-config.json'),
    }
    if which not in mapping:
        return abort(404)
    path = mapping[which]
    if not os.path.exists(path):
        return abort(404)
    return send_file(path, as_attachment=True)
