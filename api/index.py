import sys, os, traceback
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import pandas
    msg = f"pandas OK: {pandas.__version__}"
except Exception as e:
    msg = f"pandas FAIL: {type(e).__name__}: {str(e)}: {e.args}"

try:
    import openpyxl
    msg += f"\nopenpyxl OK: {openpyxl.__version__}"
except Exception as e:
    msg += f"\nopenpyxl FAIL: {type(e).__name__}: {str(e)}: {e.args}"

try:
    import flask_session
    msg += f"\nflask_session OK"
except Exception as e:
    msg += f"\nflask_session FAIL: {type(e).__name__}: {str(e)}: {e.args}"

from flask import Flask
app = Flask(__name__)

@app.route('/')
@app.route('/<path:path>')
def catch_all(path=''):
    return msg, 200
