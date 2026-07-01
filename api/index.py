import sys, os, traceback
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask

app = Flask(__name__)
import_error = None

try:
    from app import app as _app
    app = _app
except Exception:
    import_error = traceback.format_exc()

if import_error:
    @app.route('/')
    @app.route('/<path:path>')
    def error_page(path=''):
        return f"<h2>Import Error</h2><pre>{import_error}</pre>", 500
