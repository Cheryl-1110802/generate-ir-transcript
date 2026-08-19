# -*- coding: utf-8 -*-
"""
Production entry point for running on the shared server.

Flask's built-in `app.run()` (used when running app.py directly) is a
development server - not meant to stay up long-term or handle more than
one request at a time. This uses waitress instead, which is a pure-Python
WSGI server that works well on Windows and is fine for an internal tool
at this scale.

Usage:
    python serve.py
"""
from waitress import serve
from app import app, config

if __name__ == "__main__":
    srv = config["server"]
    print(f"Serving (waitress) on {srv['host']}:{srv['port']} ...")
    serve(app, host=srv["host"], port=srv["port"])
