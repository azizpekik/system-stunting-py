#!/usr/bin/env python3
"""
Gunicorn application factory for SiTrack Stunting
"""
import os
from app import app

# Standard Gunicorn application interface
application = app

if __name__ == "__main__":
    application.run()