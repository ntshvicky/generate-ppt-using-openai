"""
Vercel serverless entry point.
All requests are routed here via vercel.json.
Generated PPTX files go to /tmp (ephemeral) — the download route
regenerates them on-demand from slides_json if not present.
"""
import sys
import os

# Put the project root on the path so imports resolve correctly
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# Ensure /tmp/generated_ppts exists on cold start
os.makedirs('/tmp/generated_ppts', exist_ok=True)

from app import create_app  # noqa: E402

app = create_app()

# Vercel looks for an object named `app` (or `handler`) at module level
# No __main__ block needed — Vercel calls the WSGI app directly
