#!/usr/bin/env python3
"""
Minimal test Flask app for Railway deployment
"""
import os
from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/')
def index():
    return jsonify({
        'status': 'running',
        'service': 'SiTrack Stunting Test',
        'message': 'Minimal test app is working'
    })

@app.route('/health')
def health():
    """Health check endpoint for Railway"""
    return jsonify({
        'status': 'healthy',
        'service': 'SiTrack Stunting Test',
        'version': '0.1.0'
    })

if __name__ == '__main__':
    # Minimal logging
    print("=== Starting Minimal Test App ===")
    port = int(os.environ.get('PORT', 5000))
    host = os.environ.get('HOST', '0.0.0.0')

    print(f"PORT: {port}")
    print(f"HOST: {host}")
    print(f"Starting Flask app on {host}:{port}")

    try:
        app.run(host=host, port=port, debug=False)
    except Exception as e:
        print(f"ERROR: Failed to start Flask app: {e}")
        import traceback
        traceback.print_exc()