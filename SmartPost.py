"""
SmartPost.py
"""

from . import commands
from .lib import fusionAddInUtils as futil


def run(_context):
    """Start the Fusion add-in."""
    try:
        commands.start()
    except Exception:
        futil.handle_error("run")


def stop(_context):
    """Stop the Fusion add-in."""
    try:
        futil.clear_handlers()
        commands.stop()
    except Exception:
        futil.handle_error("stop")
