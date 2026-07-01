import sys, os, traceback
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask

app = Flask(__name__)

try:
    from app import app as _app
    app = _app
except Exception as e:
    @app.route('/')
    @app.route('/<path:path>')
    def error_page(path=''):
        return f"<h2>Import Error</h2><pre>{traceback.format_exc()}</pre>", 500
