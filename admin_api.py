"""
Flaming News - Admin API
Flask Blueprint exposing report_config's draft/publish/history workflow.
No authentication (per requirement: open access, no login wall) — anyone
with the URL can edit. Mounted at /api/admin/... in app.py.
"""
from flask import Blueprint, jsonify, request

import report_config

admin_bp = Blueprint('admin_api', __name__, url_prefix='/api/admin')


@admin_bp.route('/config/published', methods=['GET'])
def get_published():
    return jsonify(report_config.load_published())


@admin_bp.route('/config/draft', methods=['GET'])
def get_draft():
    return jsonify(report_config.load_draft())


@admin_bp.route('/config/draft', methods=['POST'])
def post_draft():
    body = request.get_json(force=True)
    saved = report_config.save_draft(body)
    return jsonify({'status': 'ok', 'config': saved})


@admin_bp.route('/config/publish', methods=['POST'])
def post_publish():
    published = report_config.publish_draft()
    return jsonify({'status': 'ok', 'config': published})


@admin_bp.route('/config/history', methods=['GET'])
def get_history():
    return jsonify(report_config.list_history())


@admin_bp.route('/config/history/<entry_id>', methods=['GET'])
def get_history_entry(entry_id):
    cfg = report_config.get_history_version(entry_id)
    if cfg is None:
        return jsonify({'status': 'error', 'message': 'Not found'}), 404
    return jsonify(cfg)


@admin_bp.route('/config/history/<entry_id>/restore', methods=['POST'])
def post_restore(entry_id):
    cfg = report_config.restore_history_version(entry_id)
    if cfg is None:
        return jsonify({'status': 'error', 'message': 'Not found'}), 404
    return jsonify({'status': 'ok', 'config': cfg, 'note': 'Restored into DRAFT — click Publish to make it live.'})
