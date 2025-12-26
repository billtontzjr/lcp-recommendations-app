"""Main Flask application."""
import os
import sys

# Ensure repo root is in Python path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, render_template
from app.config import Config
from app.routes.health import health_bp
from app.routes.api import api_bp


def create_app():
    """Create and configure the Flask application."""
    app = Flask(
        __name__,
        template_folder=os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates'),
        static_folder=os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static')
    )

    app.config.from_object(Config)

    # Ensure upload folder exists
    os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)

    # Register blueprints
    app.register_blueprint(health_bp)
    app.register_blueprint(api_bp)

    # Main page route
    @app.route('/')
    def index():
        return render_template('index.html')

    return app


# For gunicorn
app = create_app()

if __name__ == '__main__':
    # For local development only - production uses gunicorn
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
