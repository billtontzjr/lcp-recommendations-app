"""Health check endpoint."""
from flask import Blueprint, jsonify

health_bp = Blueprint('health', __name__)


@health_bp.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint for Render."""
    return jsonify({'status': 'healthy', 'service': 'lcp-recommendations'}), 200
