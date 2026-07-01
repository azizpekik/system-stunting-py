import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from app import app
except Exception as e:
    import traceback
    from flask import Flask
    app = Flask(__name__)

    @app.route('/')
    @app.route('/<path:path>')
    def error_page(path=''):
        return (
            f"<h2>Import Error</h2>"
            f"<pre>{traceback.format_exc()}</pre>",
            500
        )
